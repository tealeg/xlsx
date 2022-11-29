package xlsx

import (
	. "gopkg.in/check.v1"
)

type StreamRowSuite struct{}

var _ = Suite(&StreamRowSuite{})

// Test we can create a new StreamRow
func (r *StreamRowSuite) TestNewStreamRow(c *C) {
	row := NewStreamRow([]StreamCell{})
	c.Assert(len(row.Cells), Equals, 0)

	c1 := NewStringStreamCell("First")
	c2 := NewStringStreamCell("Second")
	row2 := NewStreamRow([]StreamCell{c1, c2})
	c.Assert(len(row2.Cells), Equals, 2)
}

// Test we can set a custom height on a StreamRow
func (r *StreamRowSuite) TestSetRowHeight(c *C) {
	row := NewStreamRow([]StreamCell{})

	c.Assert(row.isCustom, Equals, false)

	row.SetHeight(123)
	c.Assert(row.isCustom, Equals, true)
	c.Assert(row.Height, Equals, float64(123))
}

// Test we can set a custom height on a StreamRow in cm
func (r *StreamRowSuite) TestSetRowHeightCM(c *C) {
	row := NewStreamRow([]StreamCell{})

	row.SetHeightCM(1)
	c.Assert(row.isCustom, Equals, true)
	c.Assert(row.Height, Equals, cmToPostscriptPts)
}
