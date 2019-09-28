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

// makeWay will adjust the Min and Max of this colStoreNode's Col to
// make way for a new colStoreNode's Col. If necessary it will
// generate an additional colStoreNode with a new Col covering the
// "tail" portion of this colStoreNode's Col should the new node lay
// completely within the range of this one, but without reaching its
// maximum extent.
func (csn *colStoreNode) makeWay(node *colStoreNode) {
	switch {
	case csn.Col.Max < node.Col.Min:
		// The new node starts after this one ends, there's no overlap
		//
		// Node1 |----|
		// Node2        |----|
		if csn.Next != nil {
			csn.Next.makeWay(node)
			return
		}
		csn.Next = node
		node.Prev = csn
		return

	case csn.Col.Min > node.Col.Max:
		// The new node ends before this one begins, there's no overlap
		//
		// Node1         |-----|
		// Node2  |----|
		if csn.Prev != nil {
			csn.Prev.makeWay(node)
			return
		}
		csn.Prev = node
		node.Next = csn
		return

	case csn.Col.Min < node.Col.Min && csn.Col.Max > node.Col.Max:
		// The new node bisects this one:
		//
		// Node1 |---xx---|
		// Node2    |--|
		newCol := csn.Col.copyToRange(node.Col.Max+1, csn.Col.Max)
		newNode := &colStoreNode{Col: newCol, Prev: node, Next: csn.Next}
		csn.Col.Max = node.Col.Min - 1
		csn.Next = node
		node.Prev = csn
		node.Next = newNode
		return

	case csn.Col.Max >= node.Col.Min && csn.Col.Min < node.Col.Min:
		// The new node overlaps this one at some point above it's minimum:
		//
		//  Node1  |----xx|
		//  Node2      |-------|
		csn.Col.Max = node.Col.Min - 1
		if csn.Next != nil {
			// Break the link to this node, which prevents
			// us looping back and forth forever
			csn.Next.Prev = nil
			csn.Next.makeWay(node)
		}
		csn.Next = node
		node.Prev = csn
		return

	case csn.Col.Min <= node.Col.Max && csn.Col.Min > node.Col.Min:
		// The new node overlaps this one at some point below it's maximum:
		//
		// Node1:     |------|
		// Node2: |----xx|
		csn.Col.Min = node.Col.Max + 1
		if csn.Prev != nil {
			// Break the link to this node, which prevents
			// us looping back and forth forever
			csn.Prev.Next = nil
			csn.Prev.makeWay(node)
		}
		csn.Prev = node
		node.Next = csn
		return
	}
	return
}

type ColStore struct {
	Root *colStoreNode
}

//
func (cs *ColStore) Add(col *Col) {
	newNode := &colStoreNode{Col: col}
	if cs.Root == nil {
		cs.Root = newNode
		return
	}
	cs.Root.makeWay(newNode)
	return
}
