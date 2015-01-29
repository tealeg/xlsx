package xlsx

import (
	. "gopkg.in/check.v1"
)

type WriteSuite struct{}

var _ = Suite(&WriteSuite{})

// Test if we can write a struct to a row
func (r *RowSuite) TestWriteStruct(c *C) {
	var f *File
	f = NewFile()
	sheet := f.AddSheet("Test1")
	row := sheet.AddRow()
	type e struct {
		FirstName string
		Age       int
		GPA       float64
		LikesPHP  bool
	}
	testStruct := e{
		"Eric",
		20,
		3.94,
		false,
	}
	row.WriteStruct(&testStruct, -1)
	c.Assert(row, NotNil)

	c0 := row.Cells[0].String()
	c1, e1 := row.Cells[1].Int()
	c2, e2 := row.Cells[2].Float()
	c3 := row.Cells[3].Bool()

	c.Assert(c0, Equals, "Eric")
	c.Assert(c1, Equals, 20)
	c.Assert(c2, Equals, 3.94)
	c.Assert(c3, Equals, false)

	c.Assert(e1, Equals, nil)
	c.Assert(e2, Equals, nil)
}
