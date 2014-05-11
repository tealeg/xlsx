package xlsx

import (
	. "gopkg.in/check.v1"
)

type FileSuite struct {}

var _ = Suite(&FileSuite{})

// Test we can correctly open a XSLX file and return a xlsx.File
// struct.
func (l *FileSuite) TestOpenFile(c *C) {
	var xlsxFile *File
	var error error

	xlsxFile, error = OpenFile("testfile.xlsx")
	c.Assert(error, IsNil)
	c.Assert(xlsxFile, NotNil)
}

// Test we can create a File object from scratch
func (l *FileSuite) TestCreateFile(c *C) {
	var xlsxFile *File

	xlsxFile = NewFile()
	c.Assert(xlsxFile, NotNil)
}

// Test that when we open a real XLSX file we create xlsx.Sheet
// objects for the sheets inside the file and that these sheets are
// themselves correct.
func (l *FileSuite) TestCreateSheet(c *C) {
	var xlsxFile *File
	var err error
	var sheet *Sheet
	var row *Row
	xlsxFile, err = OpenFile("testfile.xlsx")
	c.Assert(err, IsNil)
	c.Assert(xlsxFile, NotNil)
	sheetLen := len(xlsxFile.Sheets)
	c.Assert(sheetLen, Equals, 3)
	sheet = xlsxFile.Sheets[0]
	rowLen := len(sheet.Rows)
	c.Assert(rowLen, Equals, 2)
	row = sheet.Rows[0]
	c.Assert(len(row.Cells), Equals, 2)
	cell := row.Cells[0]
	cellstring := cell.String()
	c.Assert(cellstring, Equals, "Foo")
}

// Test that we can add a sheet to a File
func (l *FileSuite) TestAddSheet(c *C) {
	var f *File
	f = NewFile()
	sheet := f.AddSheet("MySheet")
	c.Assert(sheet, NotNil)
	c.Assert(len(f.Sheets), Equals, 1)
	c.Assert(f.Sheets[0], Equals, sheet)
	c.Assert(len(f.Sheet), Equals, 1)
	c.Assert(f.Sheet["MySheet"], Equals, sheet)
}
