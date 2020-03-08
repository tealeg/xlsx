package xlsx

import (
	"testing"

	qt "github.com/frankban/quicktest"
)

func TestRow(t *testing.T) {
	c := qt.New(t)
	// Test we can add a new Cell to a Row
	csRunO(c, "TestAddCell", func(c *qt.C, option FileOption) {
		var f *File
		f = NewFile(option)
		sheet, _ := f.AddSheet("MySheet")
		row := sheet.AddRow()
		cell := row.AddCell()
		c.Assert(cell, qt.Not(qt.IsNil))
		c.Assert(row.cellCount, qt.Equals, 1)
	})

	csRunO(c, "TestGetCell", func(c *qt.C, option FileOption) {
		var f *File
		f = NewFile(option)
		sheet, _ := f.AddSheet("MySheet")
		row := sheet.AddRow()
		cell := row.AddCell()
		cell.SetValue("foo")
		cell1 := row.AddCell()
		cell1.SetValue("bar")

		cell2 := row.GetCell(0)
		c.Assert(cell.Value, qt.Equals, cell2.Value)
	})

}
