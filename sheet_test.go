package xlsx

import (
	"bytes"
	"encoding/xml"

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
	refTable := NewSharedStringRefTable()
	styles := &xlsxStyles{}
	xSheet := sheet.makeXLSXSheet(refTable, styles)
	c.Assert(xSheet.Dimension.Ref, Equals, "A1:A1")
	c.Assert(xSheet.SheetData.Row, HasLen, 1)
	xRow := xSheet.SheetData.Row[0]
	c.Assert(xRow.R, Equals, 1)
	c.Assert(xRow.Spans, Equals, "")
	c.Assert(xRow.C, HasLen, 1)
	xC := xRow.C[0]
	c.Assert(xC.R, Equals, "A1")
	c.Assert(xC.S, Equals, 0)
	c.Assert(xC.T, Equals, "s") // Shared string type
	c.Assert(xC.V, Equals, "0") // reference to shared string
	xSST := refTable.makeXLSXSST()
	c.Assert(xSST.Count, Equals, 1)
	c.Assert(xSST.UniqueCount, Equals, 1)
	c.Assert(xSST.SI, HasLen, 1)
	xSI := xSST.SI[0]
	c.Assert(xSI.T, Equals, "A cell!")
}

// When we create the xlsxSheet we also populate the xlsxStyles struct
// with style information.
func (s *SheetSuite) TestMakeXLSXSheetAlsoPopulatesXLSXSTyles(c *C) {
	file := NewFile()
	sheet := file.AddSheet("Sheet1")
	row := sheet.AddRow()
	cell := row.AddCell()
	cell.Value = "A cell!"
	style := *NewStyle()
	style.Font = *NewFont(10, "Verdana")
	style.Fill = *NewFill("solid", "FFFFFFFF", "00000000")
	style.Border = *NewBorder("none", "thin", "none", "thin")
	cell.SetStyle(style)
	refTable := NewSharedStringRefTable()
	styles := &xlsxStyles{}
	_ = sheet.makeXLSXSheet(refTable, styles)
	c.Assert(len(styles.Fonts), Equals, 1)
	c.Assert(styles.Fonts[0].Sz.Val, Equals, "10")
	c.Assert(styles.Fonts[0].Name.Val, Equals, "Verdana")
	c.Assert(len(styles.Fills), Equals, 1)
	c.Assert(styles.Fills[0].PatternFill.PatternType, Equals, "solid")
	c.Assert(styles.Fills[0].PatternFill.FgColor.RGB, Equals, "FFFFFFFF")
	c.Assert(styles.Fills[0].PatternFill.BgColor.RGB, Equals, "00000000")
	c.Assert(len(styles.Borders), Equals, 1)
	c.Assert(styles.Borders[0].Left.Style, Equals, "none")
	c.Assert(styles.Borders[0].Right.Style, Equals, "thin")
	c.Assert(styles.Borders[0].Top.Style, Equals, "none")
	c.Assert(styles.Borders[0].Bottom.Style, Equals, "thin")
	c.Assert(len(styles.CellStyleXfs), Equals, 1)
	c.Assert(styles.CellStyleXfs[0].FontId, Equals, 0)
	c.Assert(styles.CellStyleXfs[0].FillId, Equals, 0)
	c.Assert(styles.CellStyleXfs[0].BorderId, Equals, 0)
}

func (s *SheetSuite) TestMarshalSheet(c *C) {
	file := NewFile()
	sheet := file.AddSheet("Sheet1")
	row := sheet.AddRow()
	cell := row.AddCell()
	cell.Value = "A cell!"
	refTable := NewSharedStringRefTable()
	styles := &xlsxStyles{}
	xSheet := sheet.makeXLSXSheet(refTable, styles)

	output := bytes.NewBufferString(xml.Header)
	body, err := xml.MarshalIndent(xSheet, "  ", "  ")
	c.Assert(err, IsNil)
	c.Assert(body, NotNil)
	_, err = output.Write(body)
	c.Assert(err, IsNil)
	expectedXLSXSheet := `<?xml version="1.0" encoding="UTF-8"?>
  <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <dimension ref="A1:A1"></dimension>
    <cols></cols>
    <sheetData>
      <row r="1">
        <c r="A1" t="s">
          <v>0</v>
        </c>
      </row>
    </sheetData>
  </worksheet>`
	c.Assert(output.String(), Equals, expectedXLSXSheet)
}
