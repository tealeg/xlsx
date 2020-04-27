package xlsx

// Default column width in excel
const ColWidth = 9.5
const Excel2006MaxRowCount = 1048576
const Excel2006MaxRowIndex = Excel2006MaxRowCount - 1

type Col struct {
	Min          int
	Max          int
	Hidden       *bool
	Width        *float64
	Collapsed    *bool
	OutlineLevel *uint8
	BestFit      *bool
	CustomWidth  *bool
	Phonetic     *bool
	numFmt       string
	parsedNumFmt *parsedNumberFormat
	style        *Style
	outXfID      int
}

// NewColForRange return a pointer to a new Col, which will apply to
// columns in the range min to max (inclusive).  Note, in order for
// this Col to do anything useful you must set some of its parameters
// and then apply it to a Sheet by calling sheet.SetColParameters.
func NewColForRange(min, max int) *Col {
	if max < min {
		// Nice try ;-)
		return &Col{Min: max, Max: min}
	}

	return &Col{Min: min, Max: max}
}

// SetWidth sets the width of columns that have this Col applied to
// them.  The width is expressed as the number of characters of the
// maximum digit width of the numbers 0-9 as rendered in the normal
// style's font.
func (c *Col) SetWidth(width float64) {
	c.Width = &width
	custom := true
	c.CustomWidth = &custom
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

// GetStyle returns the Style associated with a Col
func (c *Col) GetStyle() *Style {
	return c.style
}

// SetStyle sets the style of a Col
func (c *Col) SetStyle(style *Style) {
	c.style = style
}

func (c *Col) SetOutlineLevel(outlineLevel uint8) {
	c.OutlineLevel = &outlineLevel
}

// copyToRange is an internal convenience function to make a copy of a
// Col with a different Min and Max value, it is not intended as a
// general purpose Col copying function as you must still insert the
// resulting Col into the Col Store.
func (c *Col) copyToRange(min, max int) *Col {
	return &Col{
		Min:          min,
		Max:          max,
		Hidden:       c.Hidden,
		Width:        c.Width,
		Collapsed:    c.Collapsed,
		OutlineLevel: c.OutlineLevel,
		BestFit:      c.BestFit,
		CustomWidth:  c.CustomWidth,
		Phonetic:     c.Phonetic,
		numFmt:       c.numFmt,
		parsedNumFmt: c.parsedNumFmt,
		style:        c.style,
	}
}

type ColStoreNode struct {
	Col  *Col
	Prev *ColStoreNode
	Next *ColStoreNode
}

//
func (csn *ColStoreNode) findNodeForColNum(num int) *ColStoreNode {
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
	Root *ColStoreNode
	Len  int
}

// Add a Col to the ColStore. If it overwrites all, or part of some
// existing Col's range of columns the that Col will be adjusted
// and/or split to make room for the new Col.
func (cs *ColStore) Add(col *Col) *ColStoreNode {
	newNode := &ColStoreNode{Col: col}
	if cs.Root == nil {
		cs.Root = newNode
		cs.Len = 1
		return newNode
	}
	cs.makeWay(cs.Root, newNode)
	return newNode
}

func (cs *ColStore) FindColByIndex(index int) *Col {
	csn := cs.findNodeForColNum(index)
	if csn != nil {
		return csn.Col
	}
	return nil
}

func (cs *ColStore) findNodeForColNum(num int) *ColStoreNode {
	if cs.Root == nil {
		return nil
	}
	return cs.Root.findNodeForColNum(num)
}

func (cs *ColStore) removeNode(node *ColStoreNode) {
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
	cs.Len -= 1
}

// makeWay will adjust the Min and Max of this ColStoreNode's Col to
// make way for a new ColStoreNode's Col. If necessary it will
// generate an additional ColStoreNode with a new Col covering the
// "tail" portion of this ColStoreNode's Col should the new node lay
// completely within the range of this one, but without reaching its
// maximum extent.
func (cs *ColStore) makeWay(node1, node2 *ColStoreNode) {
	switch {
	case node1.Col.Max < node2.Col.Min:
		// The node2 starts after node1 ends, there's no overlap
		//
		// Node1 |----|
		// Node2        |----|
		if node1.Next != nil {
			if node1.Next.Col.Min <= node2.Col.Max {
				cs.makeWay(node1.Next, node2)
				return
			}
			cs.addNode(node1, node2, node1.Next)
			return
		}
		cs.addNode(node1, node2, nil)
		return

	case node1.Col.Min > node2.Col.Max:
		// Node2 ends before node1 begins, there's no overlap
		//
		// Node1         |-----|
		// Node2  |----|
		if node1.Prev != nil {
			if node1.Prev.Col.Max >= node2.Col.Min {
				cs.makeWay(node1.Prev, node2)
				return
			}
			cs.addNode(node1.Prev, node2, node1)
			return
		}
		cs.addNode(nil, node2, node1)
		return

	case node1.Col.Min == node2.Col.Min && node1.Col.Max == node2.Col.Max:
		// Exact match
		//
		// Node1 |xxx|
		// Node2 |---|

		prev := node1.Prev
		next := node1.Next
		cs.removeNode(node1)
		cs.addNode(prev, node2, next)
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
		switch {
		case prev == node2:
			node2.Next = next
		case next == node2:
			node2.Prev = prev
		default:
			cs.addNode(prev, node2, next)
		}

		if node2.Prev != nil && node2.Prev.Col.Max >= node2.Col.Min {
			cs.makeWay(prev, node2)
		}
		if node2.Next != nil && node2.Next.Col.Min <= node2.Col.Max {
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
		newNode := &ColStoreNode{Col: newCol}
		cs.addNode(node1, newNode, node1.Next)
		node1.Col.Max = node2.Col.Min - 1
		cs.addNode(node1, node2, newNode)
		return

	case node1.Col.Max >= node2.Col.Min && node1.Col.Min < node2.Col.Min:
		// Node2 overlaps node1 at some point above it's minimum:
		//
		//  Node1  |----xx|
		//  Node2      |-------|
		next := node1.Next
		node1.Col.Max = node2.Col.Min - 1
		if next == node2 {
			return
		}
		cs.addNode(node1, node2, next)
		if next != nil && next.Col.Min <= node2.Col.Max {
			cs.makeWay(next, node2)
		}
		return

	case node1.Col.Min <= node2.Col.Max && node1.Col.Min > node2.Col.Min:
		// Node2 overlaps node1 at some point below it's maximum:
		//
		// Node1:     |------|
		// Node2: |----xx|
		prev := node1.Prev
		node1.Col.Min = node2.Col.Max + 1
		if prev == node2 {
			return
		}
		cs.addNode(prev, node2, node1)
		if prev != nil && prev.Col.Max >= node2.Col.Min {
			cs.makeWay(node1.Prev, node2)
		}
		return
	}
	return
}

func (cs *ColStore) addNode(prev, this, next *ColStoreNode) {
	if prev != nil {
		prev.Next = this
	}
	this.Prev = prev
	this.Next = next
	if next != nil {
		next.Prev = this
	}
	cs.Len += 1
}

func (cs *ColStore) getOrMakeColsForRange(start *ColStoreNode, min, max int) []*Col {
	cols := []*Col{}
	var csn *ColStoreNode
	var newCol *Col
	switch {
	case start == nil:
		newCol = NewColForRange(min, max)
		csn = cs.Add(newCol)
	case start.Col.Min <= min && start.Col.Max >= min:
		csn = start
	case start.Col.Min < min && start.Col.Max < min:
		if start.Next != nil {
			return cs.getOrMakeColsForRange(start.Next, min, max)
		}
		newCol = NewColForRange(min, max)
		csn = cs.Add(newCol)
	case start.Col.Min > min:
		if start.Col.Min > max {
			newCol = NewColForRange(min, max)
		} else {
			newCol = NewColForRange(min, start.Col.Min-1)
		}
		csn = cs.Add(newCol)
	}

	cols = append(cols, csn.Col)
	if csn.Col.Max >= max {
		return cols
	}
	cols = append(cols, cs.getOrMakeColsForRange(csn.Next, csn.Col.Max+1, max)...)
	return cols
}

func chainOp(csn *ColStoreNode, fn func(idx int, col *Col)) {
	for csn.Prev != nil {
		csn = csn.Prev
	}

	var i int
	for i = 0; csn.Next != nil; i++ {
		fn(i, csn.Col)
		csn = csn.Next
	}
	fn(i+1, csn.Col)
}

// ForEach calls the function fn for each Col defined in the ColStore.
func (cs *ColStore) ForEach(fn func(idx int, col *Col)) {
	if cs.Root == nil {
		return
	}
	chainOp(cs.Root, fn)
}
