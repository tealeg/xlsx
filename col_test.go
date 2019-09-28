package xlsx

import (
	. "gopkg.in/check.v1"
)

type ColSuite struct{}

var _ = Suite(&ColSuite{})

func (cs *ColSuite) TestCopyToRange(c *C) {
	nf := &parsedNumberFormat{}
	s := &Style{}
	cdv1 := &xlsxCellDataValidation{}
	cdv2 := &xlsxCellDataValidation{}
	ct := CellTypeBool.Ptr()
	c1 := &Col{
		Min:             0,
		Max:             10,
		Hidden:          true,
		Width:           300.4,
		Collapsed:       true,
		OutlineLevel:    2,
		numFmt:          "-0.00",
		parsedNumFmt:    nf,
		style:           s,
		DataValidation:  []*xlsxCellDataValidation{cdv1, cdv2},
		defaultCellType: ct,
	}

	c2 := c1.copyToRange(4, 10)
	c.Assert(c2.Min, Equals, 4)
	c.Assert(c2.Max, Equals, 10)
	c.Assert(c2.Hidden, Equals, c1.Hidden)
	c.Assert(c2.Width, Equals, c1.Width)
	c.Assert(c2.Collapsed, Equals, c1.Collapsed)
	c.Assert(c2.OutlineLevel, Equals, c1.OutlineLevel)
	c.Assert(c2.numFmt, Equals, c1.numFmt)
	c.Assert(c2.parsedNumFmt, Equals, c1.parsedNumFmt)
	c.Assert(c2.style, Equals, c1.style)
	c.Assert(c2.DataValidation, HasLen, 2)
	c.Assert(c2.DataValidation[0], Equals, c1.DataValidation[0])
	c.Assert(c2.DataValidation[1], Equals, c1.DataValidation[1])
	c.Assert(c2.defaultCellType, Equals, c1.defaultCellType)
}

type ColStoreSuite struct{}

var _ = Suite(&ColStoreSuite{})

func (css *ColStoreSuite) TestAddRootNode(c *C) {
	col := &Col{Min: 0, Max: 1}
	cs := ColStore{}
	cs.Add(col)
	c.Assert(cs.Root.Col, Equals, col)
}

func (css *ColStoreSuite) TestMakeWay(c *C) {
	assertWayMade := func(cols []*Col, chainFunc func(root *colStoreNode)) {

		cs := ColStore{}
		for _, col := range cols {
			cs.Add(col)
		}
		chainFunc(cs.Root)
	}

	// Col1: |--|
	// Col2:    |--|
	assertWayMade([]*Col{&Col{Min: 0, Max: 1}, &Col{Min: 2, Max: 3}},
		func(root *colStoreNode) {
			c.Assert(root.Col.Min, Equals, 0)
			c.Assert(root.Col.Max, Equals, 1)
			c.Assert(root.Prev, IsNil)
			c.Assert(root.Next, NotNil)
			node2 := root.Next
			c.Assert(node2.Prev, Equals, root)
			c.Assert(node2.Next, IsNil)
			c.Assert(node2.Col.Min, Equals, 2)
			c.Assert(node2.Col.Max, Equals, 3)
		})

	// Col1:    |--|
	// Col2: |--|
	assertWayMade([]*Col{&Col{Min: 2, Max: 3}, &Col{Min: 0, Max: 1}},
		func(root *colStoreNode) {
			c.Assert(root.Col.Min, Equals, 2)
			c.Assert(root.Col.Max, Equals, 3)
			c.Assert(root.Prev, NotNil)
			c.Assert(root.Next, IsNil)
			node2 := root.Prev
			c.Assert(node2.Next, Equals, root)
			c.Assert(node2.Prev, IsNil)
			c.Assert(node2.Col.Min, Equals, 0)
			c.Assert(node2.Col.Max, Equals, 1)
		})

	// Col1: |--x|
	// Col2:   |--|
	assertWayMade([]*Col{&Col{Min: 0, Max: 2}, &Col{Min: 2, Max: 3}},
		func(root *colStoreNode) {
			c.Assert(root.Col.Min, Equals, 0)
			c.Assert(root.Col.Max, Equals, 1)
			c.Assert(root.Prev, IsNil)
			c.Assert(root.Next, NotNil)
			node2 := root.Next
			c.Assert(node2.Prev, Equals, root)
			c.Assert(node2.Next, IsNil)
			c.Assert(node2.Col.Min, Equals, 2)
			c.Assert(node2.Col.Max, Equals, 3)
		})

	// Col1:  |x-|
	// Col2: |--|
	assertWayMade([]*Col{&Col{Min: 1, Max: 2}, &Col{Min: 0, Max: 1}},
		func(root *colStoreNode) {
			c.Assert(root.Col.Min, Equals, 2)
			c.Assert(root.Col.Max, Equals, 2)
			c.Assert(root.Prev, NotNil)
			c.Assert(root.Next, IsNil)
			node2 := root.Prev
			c.Assert(node2.Next, Equals, root)
			c.Assert(node2.Prev, IsNil)
			c.Assert(node2.Col.Min, Equals, 0)
			c.Assert(node2.Col.Max, Equals, 1)
		})

	// Col1: |---xx---|
	// Col2:    |--|
	assertWayMade([]*Col{&Col{Min: 0, Max: 7}, &Col{Min: 3, Max: 4}},
		func(root *colStoreNode) {
			c.Assert(root.Prev, IsNil)
			c.Assert(root.Next, NotNil)
			node2 := root.Next
			c.Assert(node2.Prev, Equals, root)
			c.Assert(node2.Col.Min, Equals, 3)
			c.Assert(node2.Col.Max, Equals, 4)
			c.Assert(node2.Next, NotNil)
			node3 := node2.Next
			c.Assert(node3.Prev, Equals, node2)
			c.Assert(node3.Next, IsNil)
			c.Assert(node3.Col.Min, Equals, 5)
			c.Assert(node3.Col.Max, Equals, 7)
		})

	// Col1: |xx|
	// Col2: |--|
	assertWayMade([]*Col{&Col{Min: 0, Max: 1, Width: 40.1}, &Col{Min: 0, Max: 1, Width: 10.0}},
		func(root *colStoreNode) {
			c.Assert(root, NotNil)
			c.Assert(root.Prev, IsNil)
			c.Assert(root.Next, IsNil)
			c.Assert(root.Col.Min, Equals, 0)
			c.Assert(root.Col.Max, Equals, 1)
			// This is how we establish we have the new node, and not the old one
			c.Assert(root.Col.Width, Equals, 10.0)
		})

	// Col1:  |xx|
	// Col2: |----|
	assertWayMade([]*Col{&Col{Min: 1, Max: 2, Width: 40.1}, &Col{Min: 0, Max: 3, Width: 10.0}},
		func(root *colStoreNode) {
			c.Assert(root.Prev, IsNil)
			c.Assert(root.Next, IsNil)
			c.Assert(root.Col.Min, Equals, 0)
			c.Assert(root.Col.Max, Equals, 3)
			// This is how we establish we have the new node, and not the old one
			c.Assert(root.Col.Width, Equals, 10.0)
		})

	// Col1: |--|
	// Col2:          |--|
	// Col3:     |--|
	assertWayMade([]*Col{&Col{Min: 0, Max: 1}, &Col{Min: 9, Max: 10}, &Col{Min: 4, Max: 5}},
		func(root *colStoreNode) {
			c.Assert(root.Prev, IsNil)
			c.Assert(root.Next, NotNil)
			node2 := root.Next
			c.Assert(node2.Prev, Equals, root)
			c.Assert(node2.Col.Min, Equals, 4)
			c.Assert(node2.Col.Max, Equals, 5)
			c.Assert(node2.Next, NotNil)
			node3 := node2.Next
			c.Assert(node3.Prev, Equals, node2)
			c.Assert(node3.Next, IsNil)
			c.Assert(node3.Col.Min, Equals, 9)
			c.Assert(node3.Col.Max, Equals, 10)
		})

}

func (css *ColStoreSuite) TestFindNodeForCol(c *C) {

	assertNodeFound := func(cs *ColStore, num int, col *Col) {
		node := cs.findNodeForColNum(num)
		if col == nil {
			c.Assert(node, IsNil)
			return
		}
		c.Assert(node, NotNil)
		c.Assert(node.Col, Equals, col)
	}

	cs := &ColStore{}
	col0 := &Col{Min: 0, Max: 0}
	cs.Add(col0)
	col1 := &Col{Min: 1, Max: 1}
	cs.Add(col1)
	col2 := &Col{Min: 2, Max: 2}
	cs.Add(col2)
	col3 := &Col{Min: 3, Max: 3}
	cs.Add(col3)
	col4 := &Col{Min: 4, Max: 4}
	cs.Add(col4)
	col5 := &Col{Min: 100, Max: 125}
	cs.Add(col5)

	assertNodeFound(cs, -1, nil)
	assertNodeFound(cs, 0, col0)
	assertNodeFound(cs, 1, col1)
	assertNodeFound(cs, 2, col2)
	assertNodeFound(cs, 3, col3)
	assertNodeFound(cs, 4, col4)
	assertNodeFound(cs, 5, nil)
	assertNodeFound(cs, 99, nil)
	assertNodeFound(cs, 100, col5)
	assertNodeFound(cs, 110, col5)
	assertNodeFound(cs, 125, col5)
	assertNodeFound(cs, 126, nil)
}

func (css *ColStoreSuite) TestRemoveNode(c *C) {

	assertChain := func(cs *ColStore, chain []*Col) {
		node := cs.Root
		for _, col := range chain {
			c.Assert(node, NotNil)
			c.Assert(node.Col.Min, Equals, col.Min)
			c.Assert(node.Col.Max, Equals, col.Max)
			node = node.Next
		}
		c.Assert(node, IsNil)
	}

	cs := &ColStore{}
	col0 := &Col{Min: 0, Max: 0}
	cs.Add(col0)
	col1 := &Col{Min: 1, Max: 1}
	cs.Add(col1)
	col2 := &Col{Min: 2, Max: 2}
	cs.Add(col2)
	col3 := &Col{Min: 3, Max: 3}
	cs.Add(col3)
	col4 := &Col{Min: 4, Max: 4}
	cs.Add(col4)

	cs.removeNode(cs.findNodeForColNum(4))
	assertChain(cs, []*Col{col0, col1, col2, col3})

	cs.removeNode(cs.findNodeForColNum(0))
	assertChain(cs, []*Col{col1, col2, col3})
}
