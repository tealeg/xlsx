package xlsx

import (
	"testing"

	qt "github.com/frankban/quicktest"
	. "gopkg.in/check.v1"
)

var notNil = qt.Not(qt.IsNil)

func TestNewColForRange(t *testing.T) {
	c := qt.New(t)
	col := NewColForRange(30, 45)
	c.Assert(col, notNil)
	c.Assert(col.Min, qt.Equals, 30)
	c.Assert(col.Max, qt.Equals, 45)

	// Auto fix the min/max
	col = NewColForRange(45, 30)
	c.Assert(col, notNil)
	c.Assert(col.Min, qt.Equals, 30)
	c.Assert(col.Max, qt.Equals, 45)
}

func TestCol(t *testing.T) {
	c := qt.New(t)
	c.Run("SetType", func(c *qt.C) {
		expectations := map[CellType]string{
			CellTypeString:        builtInNumFmt[builtInNumFmtIndex_STRING],
			CellTypeNumeric:       builtInNumFmt[builtInNumFmtIndex_INT],
			CellTypeBool:          builtInNumFmt[builtInNumFmtIndex_GENERAL],
			CellTypeInline:        builtInNumFmt[builtInNumFmtIndex_STRING],
			CellTypeError:         builtInNumFmt[builtInNumFmtIndex_GENERAL],
			CellTypeDate:          builtInNumFmt[builtInNumFmtIndex_GENERAL],
			CellTypeStringFormula: builtInNumFmt[builtInNumFmtIndex_STRING],
		}

		assertSetType := func(cellType CellType, expectation string) {
			col := &Col{}
			col.SetType(cellType)
			c.Assert(col.numFmt, qt.Equals, expectation)
		}
		for k, v := range expectations {
			assertSetType(k, v)
		}
	})
	c.Run("SetWidth", func(c *qt.C) {
		col := &Col{}
		col.SetWidth(20.2)
		c.Assert(*col.Width, qt.Equals, 20.2)
		c.Assert(*col.CustomWidth, qt.Equals, true)
	})

	c.Run("copyToRange", func(c *qt.C) {
		nf := &parsedNumberFormat{}
		s := &Style{}
		c1 := &Col{
			Min:          1,
			Max:          11,
			Hidden:       bPtr(true),
			Width:        fPtr(300.4),
			Collapsed:    bPtr(true),
			OutlineLevel: u8Ptr(2),
			numFmt:       "-0.00",
			parsedNumFmt: nf,
			style:        s,
		}

		c2 := c1.copyToRange(4, 10)
		c.Assert(c2.Min, qt.Equals, 4)
		c.Assert(c2.Max, qt.Equals, 10)
		c.Assert(c2.Hidden, qt.Equals, c1.Hidden)
		c.Assert(c2.Width, qt.Equals, c1.Width)
		c.Assert(c2.Collapsed, qt.Equals, c1.Collapsed)
		c.Assert(c2.OutlineLevel, qt.Equals, c1.OutlineLevel)
		c.Assert(c2.numFmt, qt.Equals, c1.numFmt)
		c.Assert(c2.parsedNumFmt, qt.Equals, c1.parsedNumFmt)
		c.Assert(c2.style, qt.Equals, c1.style)
	})

}

type ColStoreSuite struct{}

var _ = Suite(&ColStoreSuite{})

func (css *ColStoreSuite) TestAddRootNode(c *C) {
	col := &Col{Min: 1, Max: 1}
	cs := ColStore{}
	cs.Add(col)
	c.Assert(cs.Len, Equals, 1)
	c.Assert(cs.Root.Col, Equals, col)
}

func TestMakeWay(t *testing.T) {
	c := qt.New(t)
	assertWayMade := func(cols []*Col, chainFunc func(*ColStore)) {

		cs := &ColStore{}
		for _, col := range cols {
			_ = cs.Add(col)
		}
		chainFunc(cs)
	}

	// Col1: |--|
	// Col2:    |--|
	assertWayMade([]*Col{{Min: 1, Max: 2}, {Min: 3, Max: 4}},
		func(cs *ColStore) {
			c.Assert(cs.Len, qt.Equals, 2)
			root := cs.Root
			c.Assert(root.Col.Min, qt.Equals, 1)
			c.Assert(root.Col.Max, qt.Equals, 2)
			c.Assert(root.Prev, qt.IsNil)
			c.Assert(root.Next, notNil)
			node2 := root.Next
			c.Assert(node2.Prev, qt.Equals, root)
			c.Assert(node2.Next, qt.IsNil)
			c.Assert(node2.Col.Min, qt.Equals, 3)
			c.Assert(node2.Col.Max, qt.Equals, 4)
		})

	// Col1:    |--|
	// Col2: |--|
	assertWayMade([]*Col{{Min: 3, Max: 4}, {Min: 1, Max: 2}},
		func(cs *ColStore) {
			root := cs.Root
			c.Assert(cs.Len, qt.Equals, 2)
			c.Assert(root.Col.Min, qt.Equals, 3)
			c.Assert(root.Col.Max, qt.Equals, 4)
			c.Assert(root.Prev, notNil)
			c.Assert(root.Next, qt.IsNil)
			node2 := root.Prev
			c.Assert(node2.Next, qt.Equals, root)
			c.Assert(node2.Prev, qt.IsNil)
			c.Assert(node2.Col.Min, qt.Equals, 1)
			c.Assert(node2.Col.Max, qt.Equals, 2)
		})

	// Col1: |--x|
	// Col2:   |--|
	assertWayMade([]*Col{{Min: 1, Max: 3}, {Min: 3, Max: 4}},
		func(cs *ColStore) {
			root := cs.Root
			c.Assert(cs.Len, qt.Equals, 2)
			c.Assert(root.Col.Min, qt.Equals, 1)
			c.Assert(root.Col.Max, qt.Equals, 2)
			c.Assert(root.Prev, qt.IsNil)
			c.Assert(root.Next, notNil)
			node2 := root.Next
			c.Assert(node2.Prev, qt.Equals, root)
			c.Assert(node2.Next, qt.IsNil)
			c.Assert(node2.Col.Min, qt.Equals, 3)
			c.Assert(node2.Col.Max, qt.Equals, 4)
		})

	// Col1:  |x-|
	// Col2: |--|
	assertWayMade([]*Col{{Min: 2, Max: 3}, {Min: 1, Max: 2}},
		func(cs *ColStore) {
			root := cs.Root
			c.Assert(cs.Len, qt.Equals, 2)
			c.Assert(root.Col.Min, qt.Equals, 3)
			c.Assert(root.Col.Max, qt.Equals, 3)
			c.Assert(root.Prev, notNil)
			c.Assert(root.Next, qt.IsNil)
			node2 := root.Prev
			c.Assert(node2.Next, qt.Equals, root)
			c.Assert(node2.Prev, qt.IsNil)
			c.Assert(node2.Col.Min, qt.Equals, 1)
			c.Assert(node2.Col.Max, qt.Equals, 2)
		})

	// Col1: |---xx---|
	// Col2:    |--|
	assertWayMade([]*Col{{Min: 1, Max: 8}, {Min: 4, Max: 5}},
		func(cs *ColStore) {
			root := cs.Root
			c.Assert(cs.Len, qt.Equals, 3)
			c.Assert(root.Prev, qt.IsNil)
			c.Assert(root.Next, notNil)
			c.Assert(root.Col.Min, qt.Equals, 1)
			c.Assert(root.Col.Max, qt.Equals, 3)
			node2 := root.Next
			c.Assert(node2.Prev, qt.Equals, root)
			c.Assert(node2.Col.Min, qt.Equals, 4)
			c.Assert(node2.Col.Max, qt.Equals, 5)
			c.Assert(node2.Next, notNil)
			node3 := node2.Next
			c.Assert(node3.Prev, qt.Equals, node2)
			c.Assert(node3.Next, qt.IsNil)
			c.Assert(node3.Col.Min, qt.Equals, 6)
			c.Assert(node3.Col.Max, qt.Equals, 8)
		})

	// Col1: |xx|
	// Col2: |--|
	assertWayMade([]*Col{{Min: 1, Max: 2, Width: fPtr(40.1)}, {Min: 1, Max: 2, Width: fPtr(10.0)}},
		func(cs *ColStore) {
			root := cs.Root
			c.Assert(cs.Len, qt.Equals, 1)
			c.Assert(root, notNil)
			c.Assert(root.Prev, qt.IsNil)
			c.Assert(root.Next, qt.IsNil)
			c.Assert(root.Col.Min, qt.Equals, 1)
			c.Assert(root.Col.Max, qt.Equals, 2)
			// This is how we establish we have the new node, and not the old one
			c.Assert(*root.Col.Width, qt.Equals, 10.0)
		})

	// Col1:  |xx|
	// Col2: |----|
	assertWayMade([]*Col{{Min: 2, Max: 3, Width: fPtr(40.1)}, {Min: 1, Max: 4, Width: fPtr(10.0)}},
		func(cs *ColStore) {
			root := cs.Root
			c.Assert(cs.Len, qt.Equals, 1)
			c.Assert(root.Prev, qt.IsNil)
			c.Assert(root.Next, qt.IsNil)
			c.Assert(root.Col.Min, qt.Equals, 1)
			c.Assert(root.Col.Max, qt.Equals, 4)
			// This is how we establish we have the new node, and not the old one
			c.Assert(*root.Col.Width, qt.Equals, 10.0)
		})

	// Col1: |--|
	// Col2:    |--|
	// Col3:       |--|
	assertWayMade([]*Col{{Min: 1, Max: 2}, {Min: 3, Max: 4}, {Min: 5, Max: 6}},
		func(cs *ColStore) {
			root := cs.Root
			c.Assert(cs.Len, qt.Equals, 3)
			c.Assert(root.Prev, qt.IsNil)
			c.Assert(root.Next, notNil)
			c.Assert(root.Col.Min, qt.Equals, 1)
			c.Assert(root.Col.Max, qt.Equals, 2)
			node2 := root.Next
			c.Assert(node2.Prev, qt.Equals, root)
			c.Assert(node2.Col.Min, qt.Equals, 3)
			c.Assert(node2.Col.Max, qt.Equals, 4)
			c.Assert(node2.Next, notNil)
			node3 := node2.Next
			c.Assert(node3.Prev, qt.Equals, node2)
			c.Assert(node3.Next, qt.IsNil)
			c.Assert(node3.Col.Min, qt.Equals, 5)
			c.Assert(node3.Col.Max, qt.Equals, 6)
		})

	// Col1:       |--|
	// Col2:    |--|
	// Col3: |--|
	assertWayMade([]*Col{{Min: 5, Max: 6}, {Min: 3, Max: 4}, {Min: 1, Max: 2}},
		func(cs *ColStore) {
			root := cs.Root
			c.Assert(cs.Len, qt.Equals, 3)
			c.Assert(root.Prev, notNil)
			c.Assert(root.Next, qt.IsNil)
			c.Assert(root.Col.Min, qt.Equals, 5)
			c.Assert(root.Col.Max, qt.Equals, 6)
			node2 := root.Prev
			c.Assert(node2.Next, qt.Equals, root)
			c.Assert(node2.Col.Min, qt.Equals, 3)
			c.Assert(node2.Col.Max, qt.Equals, 4)
			c.Assert(node2.Prev, notNil)
			node3 := node2.Prev
			c.Assert(node3.Next, qt.Equals, node2)
			c.Assert(node3.Prev, qt.IsNil)
			c.Assert(node3.Col.Min, qt.Equals, 1)
			c.Assert(node3.Col.Max, qt.Equals, 2)
		})

	// Col1: |--|
	// Col2:          |--|
	// Col3:     |--|
	assertWayMade([]*Col{{Min: 1, Max: 2}, {Min: 10, Max: 11}, {Min: 5, Max: 6}},
		func(cs *ColStore) {
			root := cs.Root
			c.Assert(cs.Len, qt.Equals, 3)
			c.Assert(root.Prev, qt.IsNil)
			c.Assert(root.Next, notNil)
			c.Assert(root.Col.Min, qt.Equals, 1)
			c.Assert(root.Col.Max, qt.Equals, 2)
			node2 := root.Next
			c.Assert(node2.Prev, qt.Equals, root)
			c.Assert(node2.Col.Min, qt.Equals, 5)
			c.Assert(node2.Col.Max, qt.Equals, 6)
			c.Assert(node2.Next, notNil)
			node3 := node2.Next
			c.Assert(node3.Prev, qt.Equals, node2)
			c.Assert(node3.Next, qt.IsNil)
			c.Assert(node3.Col.Min, qt.Equals, 10)
			c.Assert(node3.Col.Max, qt.Equals, 11)
		})

	// Col1: |-x|
	// Col2:        |x-|
	// Col3:  |-------|
	assertWayMade([]*Col{
		{Min: 1, Max: 2}, {Min: 8, Max: 9}, {Min: 2, Max: 8}},
		func(cs *ColStore) {
			root := cs.Root
			c.Assert(cs.Len, qt.Equals, 3)
			c.Assert(root.Prev, qt.IsNil)
			c.Assert(root.Next, notNil)
			c.Assert(root.Col.Min, qt.Equals, 1)
			c.Assert(root.Col.Max, qt.Equals, 1)
			node2 := root.Next
			c.Assert(node2.Prev, qt.Equals, root)
			c.Assert(node2.Next, notNil)
			c.Assert(node2.Col.Min, qt.Equals, 2)
			c.Assert(node2.Col.Max, qt.Equals, 8)
			node3 := node2.Next
			c.Assert(node3.Prev, qt.Equals, node2)
			c.Assert(node3.Next, qt.IsNil)
			c.Assert(node3.Col.Min, qt.Equals, 9)
			c.Assert(node3.Col.Max, qt.Equals, 9)
		})

	// Col1: |-x|
	// Col2:        |--|
	// Col3:  |-----|
	assertWayMade([]*Col{
		{Min: 1, Max: 2}, {Min: 8, Max: 9}, {Min: 2, Max: 7}},
		func(cs *ColStore) {
			root := cs.Root
			c.Assert(cs.Len, qt.Equals, 3)
			c.Assert(root.Prev, qt.IsNil)
			c.Assert(root.Next, notNil)
			c.Assert(root.Col.Min, qt.Equals, 1)
			c.Assert(root.Col.Max, qt.Equals, 1)
			node2 := root.Next
			c.Assert(node2.Prev, qt.Equals, root)
			c.Assert(node2.Next, notNil)
			c.Assert(node2.Col.Min, qt.Equals, 2)
			c.Assert(node2.Col.Max, qt.Equals, 7)
			node3 := node2.Next
			c.Assert(node3.Prev, qt.Equals, node2)
			c.Assert(node3.Next, qt.IsNil)
			c.Assert(node3.Col.Min, qt.Equals, 8)
			c.Assert(node3.Col.Max, qt.Equals, 9)
		})

	// Col1: |--|
	// Col2:        |x-|
	// Col3:    |-----|
	assertWayMade([]*Col{
		{Min: 1, Max: 2}, {Min: 8, Max: 9}, {Min: 3, Max: 8}},
		func(cs *ColStore) {
			root := cs.Root
			c.Assert(cs.Len, qt.Equals, 3)
			c.Assert(root.Prev, qt.IsNil)
			c.Assert(root.Next, notNil)
			c.Assert(root.Col.Min, qt.Equals, 1)
			c.Assert(root.Col.Max, qt.Equals, 2)
			node2 := root.Next
			c.Assert(node2.Prev, qt.Equals, root)
			c.Assert(node2.Next, notNil)
			c.Assert(node2.Col.Min, qt.Equals, 3)
			c.Assert(node2.Col.Max, qt.Equals, 8)
			node3 := node2.Next
			c.Assert(node3.Prev, qt.Equals, node2)
			c.Assert(node3.Next, qt.IsNil)
			c.Assert(node3.Col.Min, qt.Equals, 9)
			c.Assert(node3.Col.Max, qt.Equals, 9)
		})

	// Col1: |--|
	// Col2:   |xx|
	// Col3:     |--|
	// Col4:   |--|
	assertWayMade(
		[]*Col{
			{Min: 1, Max: 2},
			{Min: 3, Max: 4, Width: fPtr(1.0)},
			{Min: 5, Max: 6},
			{Min: 3, Max: 4, Width: fPtr(2.0)},
		},
		func(cs *ColStore) {
			root := cs.Root
			c.Assert(cs.Len, qt.Equals, 3)
			c.Assert(root.Prev, qt.IsNil)
			c.Assert(root.Next, notNil)
			c.Assert(root.Col.Min, qt.Equals, 1)
			c.Assert(root.Col.Max, qt.Equals, 2)
			node2 := root.Next
			c.Assert(node2.Prev, qt.Equals, root)
			c.Assert(node2.Next, notNil)
			c.Assert(node2.Col.Min, qt.Equals, 3)
			c.Assert(node2.Col.Max, qt.Equals, 4)
			c.Assert(*node2.Col.Width, qt.Equals, 2.0) // We have the later version
			node3 := node2.Next
			c.Assert(node3.Prev, qt.Equals, node2)
			c.Assert(node3.Next, qt.IsNil)
			c.Assert(node3.Col.Min, qt.Equals, 5)
			c.Assert(node3.Col.Max, qt.Equals, 6)
		})

	// Col1: |-x|
	// Col2:   |xx|
	// Col3:     |x-|
	// Col4:  |----|
	assertWayMade(
		[]*Col{
			{Min: 1, Max: 2, Width: fPtr(1.0)},
			{Min: 3, Max: 4, Width: fPtr(2.0)},
			{Min: 5, Max: 6, Width: fPtr(3.0)},
			{Min: 2, Max: 5, Width: fPtr(4.0)},
		},
		func(cs *ColStore) {
			root := cs.Root
			c.Assert(cs.Len, qt.Equals, 3)
			c.Assert(root.Prev, qt.IsNil)
			c.Assert(root.Next, notNil)
			c.Assert(root.Col.Min, qt.Equals, 1)
			c.Assert(root.Col.Max, qt.Equals, 1)
			c.Assert(*root.Col.Width, qt.Equals, 1.0)
			node2 := root.Next
			c.Assert(node2.Prev, qt.Equals, root)
			c.Assert(node2.Next, notNil)
			c.Assert(node2.Col.Min, qt.Equals, 2)
			c.Assert(node2.Col.Max, qt.Equals, 5)
			c.Assert(*node2.Col.Width, qt.Equals, 4.0)
			node3 := node2.Next
			c.Assert(node3.Prev, qt.Equals, node2)
			c.Assert(node3.Next, qt.IsNil)
			c.Assert(node3.Col.Min, qt.Equals, 6)
			c.Assert(node3.Col.Max, qt.Equals, 6)
			c.Assert(*node3.Col.Width, qt.Equals, 3.0)
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
	col0 := &Col{Min: 1, Max: 1}
	cs.Add(col0)
	col1 := &Col{Min: 2, Max: 2}
	cs.Add(col1)
	col2 := &Col{Min: 3, Max: 3}
	cs.Add(col2)
	col3 := &Col{Min: 4, Max: 4}
	cs.Add(col3)
	col4 := &Col{Min: 5, Max: 5}
	cs.Add(col4)
	col5 := &Col{Min: 100, Max: 125}
	cs.Add(col5)

	assertNodeFound(cs, 0, nil)
	assertNodeFound(cs, 1, col0)
	assertNodeFound(cs, 2, col1)
	assertNodeFound(cs, 3, col2)
	assertNodeFound(cs, 4, col3)
	assertNodeFound(cs, 5, col4)
	assertNodeFound(cs, 6, nil)
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
	col0 := &Col{Min: 1, Max: 1}
	cs.Add(col0)
	col1 := &Col{Min: 2, Max: 2}
	cs.Add(col1)
	col2 := &Col{Min: 3, Max: 3}
	cs.Add(col2)
	col3 := &Col{Min: 4, Max: 4}
	cs.Add(col3)
	col4 := &Col{Min: 5, Max: 5}
	cs.Add(col4)
	c.Assert(cs.Len, Equals, 5)

	cs.removeNode(cs.findNodeForColNum(5))
	c.Assert(cs.Len, Equals, 4)
	assertChain(cs, []*Col{col0, col1, col2, col3})

	cs.removeNode(cs.findNodeForColNum(1))
	c.Assert(cs.Len, Equals, 3)
	assertChain(cs, []*Col{col1, col2, col3})
}

func (css *ColStoreSuite) TestForEach(c *C) {
	cs := &ColStore{}
	col0 := &Col{Min: 1, Max: 1, Hidden: bPtr(true)}
	cs.Add(col0)
	col1 := &Col{Min: 2, Max: 2}
	cs.Add(col1)
	col2 := &Col{Min: 3, Max: 3}
	cs.Add(col2)
	col3 := &Col{Min: 4, Max: 4}
	cs.Add(col3)
	col4 := &Col{Min: 5, Max: 5}
	cs.Add(col4)
	cs.ForEach(func(index int, col *Col) {
		col.Phonetic = bPtr(true)
	})

	c.Assert(col0.Phonetic, Equals, true)
	c.Assert(col1.Phonetic, Equals, true)
	c.Assert(col2.Phonetic, Equals, true)
	c.Assert(col3.Phonetic, Equals, true)
	c.Assert(col4.Phonetic, Equals, true)
}

func (css *ColStoreSuite) TestGetOrMakeColsForRange(c *C) {
	assertCols := func(min, max int, initalCols, expectedCols []*Col) {
		cs := &ColStore{}
		for _, col := range initalCols {
			cs.Add(col)
		}
		result := cs.getOrMakeColsForRange(cs.Root, min, max)
		c.Assert(result, HasLen, len(expectedCols))
		for i := 0; i < len(expectedCols); i++ {
			got := result[i]
			expected := expectedCols[i]
			c.Assert(got.Min, Equals, expected.Min)
			c.Assert(got.Max, Equals, expected.Max)
		}
	}

	// make everything
	assertCols(1, 11, nil, []*Col{{Min: 1, Max: 11}})

	// get everything, one col
	assertCols(1, 11, []*Col{{Min: 1, Max: 11}}, []*Col{{Min: 1, Max: 11}})

	// get everything, many cols
	assertCols(1, 11,
		[]*Col{
			{Min: 1, Max: 4},
			{Min: 5, Max: 8},
			{Min: 9, Max: 11},
		},
		[]*Col{
			{Min: 1, Max: 4},
			{Min: 5, Max: 8},
			{Min: 9, Max: 11},
		},
	)

	// make missing col
	assertCols(1, 11,
		[]*Col{
			{Min: 1, Max: 4},
			{Min: 9, Max: 11},
		},
		[]*Col{
			{Min: 1, Max: 4},
			{Min: 5, Max: 8},
			{Min: 9, Max: 11},
		},
	)

}
