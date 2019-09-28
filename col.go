package xlsx

// Default column width in excel
const ColWidth = 9.5
const Excel2006MaxRowCount = 1048576
const Excel2006MaxRowIndex = Excel2006MaxRowCount - 1

type Col struct {
	Min             int
	Max             int
	Hidden          bool
	Width           float64
	Collapsed       bool
	OutlineLevel    uint8
	numFmt          string
	parsedNumFmt    *parsedNumberFormat
	style           *Style
	DataValidation  []*xlsxCellDataValidation
	defaultCellType *CellType
}

// SetType will set the format string of a column based on the type that you want to set it to.
// This function does not really make a lot of sense.
func (c *Col) SetType(cellType CellType) {
	switch cellType {
	case CellTypeString:
		c.numFmt = builtInNumFmt[builtInNumFmtIndex_STRING]
	case CellTypeNumeric:
		c.numFmt = builtInNumFmt[builtInNumFmtIndex_INT]
	case CellTypeBool:
		c.numFmt = builtInNumFmt[builtInNumFmtIndex_GENERAL] //TEMP
	case CellTypeInline:
		c.numFmt = builtInNumFmt[builtInNumFmtIndex_STRING]
	case CellTypeError:
		c.numFmt = builtInNumFmt[builtInNumFmtIndex_GENERAL] //TEMP
	case CellTypeDate:
		// Cells that are stored as dates are not properly supported in this library.
		// They should instead be stored as a Numeric with a date format.
		c.numFmt = builtInNumFmt[builtInNumFmtIndex_GENERAL]
	case CellTypeStringFormula:
		c.numFmt = builtInNumFmt[builtInNumFmtIndex_STRING]
	}
}

// SetCellMetadata sets the CellMetadata related attributes
// of a Col
func (c *Col) SetCellMetadata(cellMetadata CellMetadata) {
	c.defaultCellType = &cellMetadata.cellType
	c.SetStreamStyle(cellMetadata.streamStyle)
}

// GetStyle returns the Style associated with a Col
func (c *Col) GetStyle() *Style {
	return c.style
}

// SetStyle sets the style of a Col
func (c *Col) SetStyle(style *Style) {
	c.style = style
}

// SetDataValidation set data validation with zero based start and end.
// Set end to -1 for all rows.
func (c *Col) SetDataValidation(dd *xlsxCellDataValidation, start, end int) {
	if end < 0 {
		end = Excel2006MaxRowIndex
	}

	dd.minRow = start
	dd.maxRow = end

	tmpDD := make([]*xlsxCellDataValidation, 0)
	for _, item := range c.DataValidation {
		if item.maxRow < dd.minRow {
			tmpDD = append(tmpDD, item) //No intersection
		} else if item.minRow > dd.maxRow {
			tmpDD = append(tmpDD, item) //No intersection
		} else if dd.minRow <= item.minRow && dd.maxRow >= item.maxRow {
			continue //union , item can be ignored
		} else if dd.minRow >= item.minRow {
			//Split into three or two, Newly added object, intersect with the current object in the lower half
			tmpSplit := new(xlsxCellDataValidation)
			*tmpSplit = *item

			if dd.minRow > item.minRow { //header whetherneed to split
				item.maxRow = dd.minRow - 1
				tmpDD = append(tmpDD, item)
			}
			if dd.maxRow < tmpSplit.maxRow { //footer whetherneed to split
				tmpSplit.minRow = dd.maxRow + 1
				tmpDD = append(tmpDD, tmpSplit)
			}

		} else {
			item.minRow = dd.maxRow + 1
			tmpDD = append(tmpDD, item)
		}
	}
	tmpDD = append(tmpDD, dd)
	c.DataValidation = tmpDD
}

// SetDataValidationWithStart set data validation with a zero basd start row.
// This will apply to the rest of the rest of the column.
func (c *Col) SetDataValidationWithStart(dd *xlsxCellDataValidation, start int) {
	c.SetDataValidation(dd, start, -1)
}

// SetStreamStyle sets the style and number format id to the ones specified in the given StreamStyle
func (c *Col) SetStreamStyle(style StreamStyle) {
	c.style = style.style
	// TODO: `style.xNumFmtId` could be out of the range of the builtin map
	// returning "" which may not be a valid formatCode
	c.numFmt = builtInNumFmt[style.xNumFmtId]
}

func (c *Col) GetStreamStyle() StreamStyle {
	// TODO: Like `SetStreamStyle`, `numFmt` could be out of the range of the builtin inv map
	// returning 0 which maps to formatCode "general"
	return StreamStyle{builtInNumFmtInv[c.numFmt], c.style}
}

// copyToRange is an internal convenience function to make a copy of a
// Col with a different Min and Max value, it is not intended as a
// general purpose Col copying function as you must still insert the
// resulting Col into the ColStore.
func (c *Col) copyToRange(min, max int) *Col {
	return &Col{
		Min:             min,
		Max:             max,
		Hidden:          c.Hidden,
		Width:           c.Width,
		Collapsed:       c.Collapsed,
		OutlineLevel:    c.OutlineLevel,
		numFmt:          c.numFmt,
		parsedNumFmt:    c.parsedNumFmt,
		style:           c.style,
		DataValidation:  append([]*xlsxCellDataValidation{}, c.DataValidation...),
		defaultCellType: c.defaultCellType,
	}
}

type colStoreNode struct {
	Col  *Col
	Prev *colStoreNode
	Next *colStoreNode
}

//
func (csn *colStoreNode) findNodeForColNum(num int) *colStoreNode {
	switch {
	case num >= csn.Col.Min && num <= csn.Col.Max:
		return csn

	case num < csn.Col.Min:
		if csn.Prev == nil {
			return nil
		}
		if csn.Prev.Col.Max < num {
			return nil
		}
		return csn.Prev.findNodeForColNum(num)

	case num > csn.Col.Max:
		if csn.Next == nil {
			return nil
		}
		if csn.Next.Col.Min > num {
			return nil
		}
		return csn.Next.findNodeForColNum(num)
	}
	return nil
}

// ColStore is the working store of Col definitions, it will simplify all Cols added to it, to ensure there ar no overlapping definitions.
type ColStore struct {
	Root *colStoreNode
}

// Add a Col to the ColStore. If it overwrites all, or part of some
// existing Col's range of columns the that Col will be adjusted
// and/or split to make room for the new Col.
func (cs *ColStore) Add(col *Col) {
	newNode := &colStoreNode{Col: col}
	if cs.Root == nil {
		cs.Root = newNode
		return
	}
	cs.makeWay(cs.Root, newNode)
	return
}

//
func (cs *ColStore) findNodeForColNum(num int) *colStoreNode {
	if cs.Root == nil {
		return nil
	}
	return cs.Root.findNodeForColNum(num)
}

//
func (cs *ColStore) removeNode(node *colStoreNode) {
	if node.Prev != nil {
		if node.Next != nil {
			node.Prev.Next = node.Next
		} else {
			node.Prev.Next = nil
		}

	}
	if node.Next != nil {
		if node.Prev != nil {
			node.Next.Prev = node.Prev
		} else {
			node.Next.Prev = nil
		}
	}
	if cs.Root == node {
		switch {
		case node.Prev != nil:
			cs.Root = node.Prev
		case node.Next != nil:
			cs.Root = node.Next
		default:
			cs.Root = nil
		}
	}
	node.Next = nil
	node.Prev = nil
}

// makeWay will adjust the Min and Max of this colStoreNode's Col to
// make way for a new colStoreNode's Col. If necessary it will
// generate an additional colStoreNode with a new Col covering the
// "tail" portion of this colStoreNode's Col should the new node lay
// completely within the range of this one, but without reaching its
// maximum extent.
func (cs *ColStore) makeWay(node1, node2 *colStoreNode) {
	switch {
	case node1.Col.Max < node2.Col.Min:
		// The node2 starts after node1 ends, there's no overlap
		//
		// Node1 |----|
		// Node2        |----|
		if node1.Next != nil {
			next := node1.Next
			if next.Col.Min >= node2.Col.Min {
				node1.Next = node2
				node2.Prev = node1
			}
			cs.makeWay(next, node2)
			return
		}
		node1.Next = node2
		node2.Prev = node1
		return

	case node1.Col.Min > node2.Col.Max:
		// Node2 ends before node1 begins, there's no overlap
		//
		// Node1         |-----|
		// Node2  |----|
		if node1.Prev != nil {
			prev := node1.Prev
			if prev.Col.Max <= node2.Col.Max {
				node1.Prev = node2
				node2.Next = node1
			}
			cs.makeWay(prev, node2)
			return
		}
		node1.Prev = node2
		node2.Next = node1
		return

	case node1.Col.Min == node2.Col.Min && node1.Col.Max == node2.Col.Max:
		// Exact match
		//
		// Node1 |xxx|
		// Node2 |---|

		prev := node1.Prev
		next := node1.Next
		cs.removeNode(node1)
		if prev != nil {
			prev.Next = node2
			node2.Prev = prev
		}
		if next != nil {
			next.Prev = node2
			node2.Next = next
		}
		// Remove node may have set the root to nil
		if cs.Root == nil {
			cs.Root = node2
		}
		return

	case node1.Col.Min > node2.Col.Min && node1.Col.Max < node2.Col.Max:
		// Node2 envelopes node1
		//
		// Node1  |xx|
		// Node2 |----|

		prev := node1.Prev
		next := node1.Next
		cs.removeNode(node1)

		if prev != nil {
			prev.Next = nil
			cs.makeWay(prev, node2)
		}
		if next != nil {
			next.Prev = nil
			cs.makeWay(next, node2)
		}

		if cs.Root == nil {
			cs.Root = node2
		}

	case node1.Col.Min < node2.Col.Min && node1.Col.Max > node2.Col.Max:
		// Node2 bisects node1:
		//
		// Node1 |---xx---|
		// Node2    |--|
		newCol := node1.Col.copyToRange(node2.Col.Max+1, node1.Col.Max)
		newNode := &colStoreNode{Col: newCol, Prev: node2, Next: node1.Next}
		node1.Col.Max = node2.Col.Min - 1
		node1.Next = node2
		node2.Prev = node1
		node2.Next = newNode
		return

	case node1.Col.Max >= node2.Col.Min && node1.Col.Min < node2.Col.Min:
		// Node2 overlaps node1 at some point above it's minimum:
		//
		//  Node1  |----xx|
		//  Node2      |-------|
		node1.Col.Max = node2.Col.Min - 1
		if node1.Next != nil {
			// Break the link to this node, which prevents
			// us looping back and forth forever
			node1.Next.Prev = nil
			cs.makeWay(node1.Next, node2)
		}
		node1.Next = node2
		node2.Prev = node1
		return

	case node1.Col.Min <= node2.Col.Max && node1.Col.Min > node2.Col.Min:
		// Node2 overlaps node1 at some point below it's maximum:
		//
		// Node1:     |------|
		// Node2: |----xx|
		node1.Col.Min = node2.Col.Max + 1
		if node1.Prev != nil {
			// Break the link to this node, which prevents
			// us looping back and forth forever
			node1.Prev.Next = nil
			cs.makeWay(node1.Prev, node2)
		}
		node1.Prev = node2
		node2.Next = node1
		return
	}
	return
}
