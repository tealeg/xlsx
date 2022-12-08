package xlsx

import (
	. "gopkg.in/check.v1"
)

type StreamCellSuite struct{}

var _ = Suite(&StreamCellSuite{})

// Test we can merge StreamCells
func (r *StreamCellSuite) TestMerge(c *C) {
	cell := NewStringStreamCell("First")

	c.Assert(cell.HMerge, Equals, 0)
	c.Assert(cell.VMerge, Equals, 0)

	cell.Merge(3, 4)

	c.Assert(cell.HMerge, Equals, 3)
	c.Assert(cell.VMerge, Equals, 4)
}

