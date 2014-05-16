package xlsx

import (
	. "gopkg.in/check.v1"
)

type SheetSuite struct{}

var _ = Suite(&SheetSuite{})

// Test we can add a Row to a Sheet
func (s *SheetSuite) TestAddRow(c *C) {
	var f *File
	f = NewFile()
	sheet := f.AddSheet("MySheet")
	row := sheet.AddRow()
	c.Assert(row, NotNil)
	c.Assert(len(sheet.Rows), Equals, 1)
}

func (s *SheetSuite) TestMakeXLSXSheetFromRows(c *C) {
	file := NewFile()
	sheet := file.AddSheet("Sheet1")
	row := sheet.AddRow()
	cell := row.AddCell()
	cell.Value = "A cell!"
	xSheet, err := sheet.makeXLSXSheet()
	c.Assert(err, IsNil)
	expectedXLSXSheet := `
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:A1"/>
  <sheetData>
    <row r="1">
      <c r="A1" t="s">
        <v>A cell!</v>
      </c>
    </row>
  </sheetData>
</worksheet>
`
	c.Assert(string(xSheet), Equals, expectedXLSXSheet)
}

