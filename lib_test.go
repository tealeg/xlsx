package xlsx

import (
	"bytes"
	"encoding/xml"
	// "strconv"
	. "gopkg.in/check.v1"
	"strings"
)

type LibSuite struct{}

var _ = Suite(&LibSuite{})

// Test we can correctly open a XSLX file and return a xlsx.File
// struct.
func (l *LibSuite) TestOpenFile(c *C) {
	var xlsxFile *File
	var error error

	xlsxFile, error = OpenFile("testfile.xlsx")
	c.Assert(error, IsNil)
	c.Assert(xlsxFile, NotNil)

}

// Test we can create a File object from scratch
func (l *LibSuite) TestCreateFile(c *C) {
	var xlsxFile *File

	xlsxFile = NewFile()
	c.Assert(xlsxFile, NotNil)
}

// Test that when we open a real XLSX file we create xlsx.Sheet
// objects for the sheets inside the file and that these sheets are
// themselves correct.
func (l *LibSuite) TestCreateSheet(c *C) {
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

func (l *LibSuite) TestGetNumberFormat(c *C) {
	var cell *Cell
	var cellXfs []xlsxXf
	var numFmt xlsxNumFmt
	var numFmts []xlsxNumFmt
	var xStyles *xlsxStyles
	var numFmtRefTable map[int]xlsxNumFmt

	cellXfs = make([]xlsxXf, 1)
	cellXfs[0] = xlsxXf{NumFmtId: 1}

	numFmts = make([]xlsxNumFmt, 1)
	numFmtRefTable = make(map[int]xlsxNumFmt)

	xStyles = &xlsxStyles{NumFmts: numFmts, CellXfs: cellXfs}

	cell = &Cell{Value: "123.123", numFmtRefTable: numFmtRefTable, styleIndex: 1, styles: xStyles}

	numFmt = xlsxNumFmt{NumFmtId: 1, FormatCode: "dd/mm/yy"}
	numFmts[0] = numFmt
	numFmtRefTable[1] = numFmt
	c.Assert(cell.GetNumberFormat(), Equals, "dd/mm/yy")
}

// We can return a string representation of the formatted data
func (l *LibSuite) TestFormattedValue(c *C) {
	var cell, earlyCell, negativeCell, smallCell *Cell
	var cellXfs []xlsxXf
	var numFmt xlsxNumFmt
	var numFmts []xlsxNumFmt
	var xStyles *xlsxStyles
	var numFmtRefTable map[int]xlsxNumFmt

	cellXfs = make([]xlsxXf, 1)
	cellXfs[0] = xlsxXf{NumFmtId: 1}

	numFmts = make([]xlsxNumFmt, 1)
	numFmtRefTable = make(map[int]xlsxNumFmt)

	xStyles = &xlsxStyles{NumFmts: numFmts, CellXfs: cellXfs}
	cell = &Cell{Value: "37947.7500001", numFmtRefTable: numFmtRefTable, styleIndex: 1, styles: xStyles}
	negativeCell = &Cell{Value: "-37947.7500001", numFmtRefTable: numFmtRefTable, styleIndex: 1, styles: xStyles}
	smallCell = &Cell{Value: "0.007", numFmtRefTable: numFmtRefTable, styleIndex: 1, styles: xStyles}
	earlyCell = &Cell{Value: "2.1", numFmtRefTable: numFmtRefTable, styleIndex: 1, styles: xStyles}
	setCode := func(code string) {
		numFmt = xlsxNumFmt{NumFmtId: 1, FormatCode: code}
		numFmts[0] = numFmt
		numFmtRefTable[1] = numFmt
	}

	setCode("general")
	c.Assert(cell.FormattedValue(), Equals, "37947.7500001")
	c.Assert(negativeCell.FormattedValue(), Equals, "-37947.7500001")

	setCode("0")
	c.Assert(cell.FormattedValue(), Equals, "37947")

	setCode("#,##0") // For the time being we're not doing this
	// comma formatting, so it'll fall back to
	// the related non-comma form.
	c.Assert(cell.FormattedValue(), Equals, "37947")

	setCode("0.00")
	c.Assert(cell.FormattedValue(), Equals, "37947.75")

	setCode("#,##0.00") // For the time being we're not doing this
	// comma formatting, so it'll fall back to
	// the related non-comma form.
	c.Assert(cell.FormattedValue(), Equals, "37947.75")

	setCode("#,##0 ;(#,##0)")
	c.Assert(cell.FormattedValue(), Equals, "37947")
	c.Assert(negativeCell.FormattedValue(), Equals, "(37947)")

	setCode("#,##0 ;[red](#,##0)")
	c.Assert(cell.FormattedValue(), Equals, "37947")
	c.Assert(negativeCell.FormattedValue(), Equals, "(37947)")

	setCode("0%")
	c.Assert(cell.FormattedValue(), Equals, "3794775%")

	setCode("0.00%")
	c.Assert(cell.FormattedValue(), Equals, "3794775.00%")

	setCode("0.00e+00")
	c.Assert(cell.FormattedValue(), Equals, "3.794775e+04")

	setCode("##0.0e+0") // This is wrong, but we'll use it for now.
	c.Assert(cell.FormattedValue(), Equals, "3.794775e+04")

	setCode("mm-dd-yy")
	c.Assert(cell.FormattedValue(), Equals, "11-22-03")

	setCode("d-mmm-yy")
	c.Assert(cell.FormattedValue(), Equals, "22-Nov-03")
	c.Assert(earlyCell.FormattedValue(), Equals, "1-Jan-00")

	setCode("d-mmm")
	c.Assert(cell.FormattedValue(), Equals, "22-Nov")
	c.Assert(earlyCell.FormattedValue(), Equals, "1-Jan")

	setCode("mmm-yy")
	c.Assert(cell.FormattedValue(), Equals, "Nov-03")

	setCode("h:mm am/pm")
	c.Assert(cell.FormattedValue(), Equals, "6:00 pm")
	c.Assert(smallCell.FormattedValue(), Equals, "12:14 am")

	setCode("h:mm:ss am/pm")
	c.Assert(cell.FormattedValue(), Equals, "6:00:00 pm")
	c.Assert(smallCell.FormattedValue(), Equals, "12:14:47 am")

	setCode("h:mm")
	c.Assert(cell.FormattedValue(), Equals, "18:00")
	c.Assert(smallCell.FormattedValue(), Equals, "00:14")

	setCode("h:mm:ss")
	c.Assert(cell.FormattedValue(), Equals, "18:00:00")
	// This is wrong, but there's no eary way aroud it in Go right now, AFAICT.
	c.Assert(smallCell.FormattedValue(), Equals, "00:14:47")

	setCode("m/d/yy h:mm")
	c.Assert(cell.FormattedValue(), Equals, "11/22/03 18:00")
	c.Assert(smallCell.FormattedValue(), Equals, "12/30/99 00:14") // Note, that's 1899
	c.Assert(earlyCell.FormattedValue(), Equals, "1/1/00 02:24")   // and 1900

	setCode("mm:ss")
	c.Assert(cell.FormattedValue(), Equals, "00:00")
	c.Assert(smallCell.FormattedValue(), Equals, "14:47")

	setCode("[h]:mm:ss")
	c.Assert(cell.FormattedValue(), Equals, "18:00:00")
	c.Assert(smallCell.FormattedValue(), Equals, "14:47")

	setCode("mmss.0") // I'm not sure about these.
	c.Assert(cell.FormattedValue(), Equals, "00.8640")
	c.Assert(smallCell.FormattedValue(), Equals, "1447.999997")

	setCode("yyyy\\-mm\\-dd")
	c.Assert(cell.FormattedValue(), Equals, "2003\\-11\\-22")

	setCode("dd/mm/yy")
	c.Assert(cell.FormattedValue(), Equals, "22/11/03")
	c.Assert(earlyCell.FormattedValue(), Equals, "01/01/00")

	setCode("hh:mm:ss")
	c.Assert(cell.FormattedValue(), Equals, "18:00:00")
	c.Assert(smallCell.FormattedValue(), Equals, "00:14:47")

	setCode("dd/mm/yy\\ hh:mm")
	c.Assert(cell.FormattedValue(), Equals, "22/11/03\\ 18:00")

	setCode("yy-mm-dd")
	c.Assert(cell.FormattedValue(), Equals, "03-11-22")

	setCode("d-mmm-yyyy")
	c.Assert(cell.FormattedValue(), Equals, "22-Nov-2003")
	c.Assert(earlyCell.FormattedValue(), Equals, "1-Jan-1900")

	setCode("m/d/yy")
	c.Assert(cell.FormattedValue(), Equals, "11/22/03")
	c.Assert(earlyCell.FormattedValue(), Equals, "1/1/00")

	setCode("m/d/yyyy")
	c.Assert(cell.FormattedValue(), Equals, "11/22/2003")
	c.Assert(earlyCell.FormattedValue(), Equals, "1/1/1900")

	setCode("dd-mmm-yyyy")
	c.Assert(cell.FormattedValue(), Equals, "22-Nov-2003")

	setCode("dd/mm/yyyy")
	c.Assert(cell.FormattedValue(), Equals, "22/11/2003")

	setCode("mm/dd/yy hh:mm am/pm")
	c.Assert(cell.FormattedValue(), Equals, "11/22/03 06:00 pm")

	setCode("mm/dd/yyyy hh:mm:ss")
	c.Assert(cell.FormattedValue(), Equals, "11/22/2003 18:00:00")
	c.Assert(smallCell.FormattedValue(), Equals, "12/30/1899 00:14:47")

	setCode("yyyy-mm-dd hh:mm:ss")
	c.Assert(cell.FormattedValue(), Equals, "2003-11-22 18:00:00")
	c.Assert(smallCell.FormattedValue(), Equals, "1899-12-30 00:14:47")
}

// Test that GetStyle correctly converts the xlsxStyle.Fonts.
func (l *LibSuite) TestGetStyleWithFonts(c *C) {
	var cell *Cell
	var style *Style
	var xStyles *xlsxStyles
	var fonts []xlsxFont
	var cellXfs []xlsxXf

	fonts = make([]xlsxFont, 1)
	fonts[0] = xlsxFont{
		Sz:   xlsxVal{Val: "10"},
		Name: xlsxVal{Val: "Calibra"}}

	cellXfs = make([]xlsxXf, 1)
	cellXfs[0] = xlsxXf{ApplyFont: true, FontId: 0}

	xStyles = &xlsxStyles{Fonts: fonts, CellXfs: cellXfs}

	cell = &Cell{Value: "123", styleIndex: 0, styles: xStyles}
	style = cell.GetStyle()
	c.Assert(style, NotNil)
	c.Assert(style.Font.Size, Equals, 10)
	c.Assert(style.Font.Name, Equals, "Calibra")
}

// Test that GetStyle correctly converts the xlsxStyle.Fills.
func (l *LibSuite) TestGetStyleWithFills(c *C) {
	var cell *Cell
	var style *Style
	var xStyles *xlsxStyles
	var fills []xlsxFill
	var cellXfs []xlsxXf

	fills = make([]xlsxFill, 1)
	fills[0] = xlsxFill{
		PatternFill: xlsxPatternFill{
			PatternType: "solid",
			FgColor:     xlsxColor{RGB: "FF000000"},
			BgColor:     xlsxColor{RGB: "00FF0000"}}}
	cellXfs = make([]xlsxXf, 1)
	cellXfs[0] = xlsxXf{ApplyFill: true, FillId: 0}

	xStyles = &xlsxStyles{Fills: fills, CellXfs: cellXfs}

	cell = &Cell{Value: "123", styleIndex: 0, styles: xStyles}
	style = cell.GetStyle()
	fill := style.Fill
	c.Assert(fill.PatternType, Equals, "solid")
	c.Assert(fill.BgColor, Equals, "00FF0000")
	c.Assert(fill.FgColor, Equals, "FF000000")
}

// Test that GetStyle correctly converts the xlsxStyle.Borders.
func (l *LibSuite) TestGetStyleWithBorders(c *C) {
	var cell *Cell
	var style *Style
	var xStyles *xlsxStyles
	var borders []xlsxBorder
	var cellXfs []xlsxXf

	borders = make([]xlsxBorder, 1)
	borders[0] = xlsxBorder{
		Left:   xlsxLine{Style: "thin"},
		Right:  xlsxLine{Style: "thin"},
		Top:    xlsxLine{Style: "thin"},
		Bottom: xlsxLine{Style: "thin"}}

	cellXfs = make([]xlsxXf, 1)
	cellXfs[0] = xlsxXf{ApplyBorder: true, BorderId: 0}

	xStyles = &xlsxStyles{Borders: borders, CellXfs: cellXfs}

	cell = &Cell{Value: "123", styleIndex: 0, styles: xStyles}
	style = cell.GetStyle()
	border := style.Border
	c.Assert(border.Left, Equals, "thin")
	c.Assert(border.Right, Equals, "thin")
	c.Assert(border.Top, Equals, "thin")
	c.Assert(border.Bottom, Equals, "thin")
}

// Test that we can correctly extract a reference table from the
// sharedStrings.xml file embedded in the XLSX file and return a
// reference table of string values from it.
func (l *LibSuite) TestReadSharedStringsFromZipFile(c *C) {
	var xlsxFile *File
	var err error
	xlsxFile, err = OpenFile("testfile.xlsx")
	c.Assert(err, IsNil)
	c.Assert(xlsxFile.referenceTable, NotNil)
}

// Helper function used to test contents of a given xlsxXf against
// expectations.
func testXf(c *C, result, expected *xlsxXf) {
	c.Assert(result.ApplyAlignment, Equals, expected.ApplyAlignment)
	c.Assert(result.ApplyBorder, Equals, expected.ApplyBorder)
	c.Assert(result.ApplyFont, Equals, expected.ApplyFont)
	c.Assert(result.ApplyFill, Equals, expected.ApplyFill)
	c.Assert(result.ApplyProtection, Equals, expected.ApplyProtection)
	c.Assert(result.BorderId, Equals, expected.BorderId)
	c.Assert(result.FillId, Equals, expected.FillId)
	c.Assert(result.FontId, Equals, expected.FontId)
	c.Assert(result.NumFmtId, Equals, expected.NumFmtId)
}

// We can correctly extract a style table from the style.xml file
// embedded in the XLSX file and return a styles struct from it.
func (l *LibSuite) TestReadStylesFromZipFile(c *C) {
	var xlsxFile *File
	var err error
	var fontCount, fillCount, borderCount, cellStyleXfCount, cellXfCount int
	var font xlsxFont
	var fill xlsxFill
	var border xlsxBorder
	var xf xlsxXf

	xlsxFile, err = OpenFile("testfile.xlsx")
	c.Assert(err, IsNil)
	c.Assert(xlsxFile.styles, NotNil)

	fontCount = len(xlsxFile.styles.Fonts)
	c.Assert(fontCount, Equals, 4)

	font = xlsxFile.styles.Fonts[0]
	c.Assert(font.Sz.Val, Equals, "11")
	c.Assert(font.Name.Val, Equals, "Calibri")

	fillCount = len(xlsxFile.styles.Fills)
	c.Assert(fillCount, Equals, 3)

	fill = xlsxFile.styles.Fills[2]
	c.Assert(fill.PatternFill.PatternType, Equals, "solid")

	borderCount = len(xlsxFile.styles.Borders)
	c.Assert(borderCount, Equals, 2)

	border = xlsxFile.styles.Borders[1]
	c.Assert(border.Left.Style, Equals, "thin")
	c.Assert(border.Right.Style, Equals, "thin")
	c.Assert(border.Top.Style, Equals, "thin")
	c.Assert(border.Bottom.Style, Equals, "thin")

	cellStyleXfCount = len(xlsxFile.styles.CellStyleXfs)
	c.Assert(cellStyleXfCount, Equals, 20)

	xf = xlsxFile.styles.CellStyleXfs[0]
	expectedXf := &xlsxXf{
		ApplyAlignment:  true,
		ApplyBorder:     true,
		ApplyFont:       true,
		ApplyFill:       false,
		ApplyProtection: true,
		BorderId:        0,
		FillId:          0,
		FontId:          0,
		NumFmtId:        164}
	testXf(c, &xf, expectedXf)

	cellXfCount = len(xlsxFile.styles.CellXfs)
	c.Assert(cellXfCount, Equals, 3)

	xf = xlsxFile.styles.CellXfs[0]
	expectedXf = &xlsxXf{
		ApplyAlignment:  false,
		ApplyBorder:     false,
		ApplyFont:       false,
		ApplyFill:       false,
		ApplyProtection: false,
		BorderId:        0,
		FillId:          0,
		FontId:          0,
		NumFmtId:        164}
	testXf(c, &xf, expectedXf)
}

// We can correctly extract a map of relationship Ids to the worksheet files in
// which they are contained from the XLSX file.
func (l *LibSuite) TestReadWorkbookRelationsFromZipFile(c *C) {
	var xlsxFile *File
	var err error

	xlsxFile, err = OpenFile("testfile.xlsx")
	c.Assert(err, IsNil)
	sheetCount := len(xlsxFile.Sheet)
	c.Assert(sheetCount, Equals, 3)
}

// which they are contained from the XLSX file, even when the
// worksheet files have arbitrary, non-numeric names.
func (l *LibSuite) TestReadWorkbookRelationsFromZipFileWithFunnyNames(c *C) {
	var xlsxFile *File
	var err error

	xlsxFile, err = OpenFile("testrels.xlsx")
	c.Assert(err, IsNil)
	sheetCount := len(xlsxFile.Sheet)
	c.Assert(sheetCount, Equals, 2)
	bob := xlsxFile.Sheet["Bob"]
	row1 := bob.Rows[0]
	cell1 := row1.Cells[0]
	c.Assert(cell1.String(), Equals, "I am Bob")
}

func (l *LibSuite) TestGetStyleFromZipFile(c *C) {
	var xlsxFile *File
	var err error

	xlsxFile, err = OpenFile("testfile.xlsx")
	c.Assert(err, IsNil)
	sheetCount := len(xlsxFile.Sheet)
	c.Assert(sheetCount, Equals, 3)

	tabelle1 := xlsxFile.Sheet["Tabelle1"]

	row0 := tabelle1.Rows[0]
	cellFoo := row0.Cells[0]
	c.Assert(cellFoo.String(), Equals, "Foo")
	c.Assert(cellFoo.GetStyle().Fill.BgColor, Equals, "FF33CCCC")

	row1 := tabelle1.Rows[1]
	cellQuuk := row1.Cells[1]
	c.Assert(cellQuuk.String(), Equals, "Quuk")
	c.Assert(cellQuuk.GetStyle().Border.Left, Equals, "thin")

	cellBar := row0.Cells[1]
	c.Assert(cellBar.String(), Equals, "Bar")
	c.Assert(cellBar.GetStyle().Fill.BgColor, Equals, "")
}

func (l *LibSuite) TestLettersToNumeric(c *C) {
	cases := map[string]int{"A": 0, "G": 6, "z": 25, "AA": 26, "Az": 51,
		"BA": 52, "Bz": 77, "ZA": 26*26 + 0, "ZZ": 26*26 + 25,
		"AAA": 26*26 + 26 + 0, "AMI": 1022}
	for input, ans := range cases {
		output := lettersToNumeric(input)
		c.Assert(output, Equals, ans)
	}
}

func (l *LibSuite) TestLetterOnlyMapFunction(c *C) {
	var input string = "ABC123"
	var output string = strings.Map(letterOnlyMapF, input)
	c.Assert(output, Equals, "ABC")
	input = "abc123"
	output = strings.Map(letterOnlyMapF, input)
	c.Assert(output, Equals, "ABC")
}

func (l *LibSuite) TestIntOnlyMapFunction(c *C) {
	var input string = "ABC123"
	var output string = strings.Map(intOnlyMapF, input)
	c.Assert(output, Equals, "123")
}

func (l *LibSuite) TestGetCoordsFromCellIDString(c *C) {
	var cellIDString string = "A3"
	var x, y int
	var err error
	x, y, err = getCoordsFromCellIDString(cellIDString)
	c.Assert(err, IsNil)
	c.Assert(x, Equals, 0)
	c.Assert(y, Equals, 2)
}

func (l *LibSuite) TestGetMaxMinFromDimensionRef(c *C) {
	var dimensionRef string = "A1:B2"
	var minx, miny, maxx, maxy int
	var err error
	minx, miny, maxx, maxy, err = getMaxMinFromDimensionRef(dimensionRef)
	c.Assert(err, IsNil)
	c.Assert(minx, Equals, 0)
	c.Assert(miny, Equals, 0)
	c.Assert(maxx, Equals, 1)
	c.Assert(maxy, Equals, 1)
}

func (l *LibSuite) TestCalculateMaxMinFromWorksheet(c *C) {
	var sheetxml = bytes.NewBufferString(`
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main"
           xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
           xmlns:mv="urn:schemas-microsoft-com:mac:vml"
           xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
           xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
           xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr customHeight="1" defaultColWidth="14.43" defaultRowHeight="15.75"/>
  <sheetData>
    <row r="1">
      <c t="s" s="1" r="A1">
        <v>0</v>
      </c>
      <c t="s" s="1" r="B1">
        <v>1</v>
      </c>
    </row>
    <row r="2">
      <c t="s" s="1" r="A2">
        <v>2</v>
      </c>
      <c t="s" s="1" r="B2">
        <v>3</v>
      </c>
    </row>
  </sheetData>
  <drawing r:id="rId1"/>
</worksheet>`)
	worksheet := new(xlsxWorksheet)
	err := xml.NewDecoder(sheetxml).Decode(worksheet)
	c.Assert(err, IsNil)
	minx, miny, maxx, maxy, err := calculateMaxMinFromWorksheet(worksheet)
	c.Assert(err, IsNil)
	c.Assert(minx, Equals, 0)
	c.Assert(miny, Equals, 0)
	c.Assert(maxx, Equals, 1)
	c.Assert(maxy, Equals, 1)
}

func (l *LibSuite) TestGetRangeFromString(c *C) {
	var rangeString string
	var lower, upper int
	var err error
	rangeString = "1:3"
	lower, upper, err = getRangeFromString(rangeString)
	c.Assert(err, IsNil)
	c.Assert(lower, Equals, 1)
	c.Assert(upper, Equals, 3)
}

func (l *LibSuite) TestMakeRowFromSpan(c *C) {
	var rangeString string
	var row *Row
	var length int
	rangeString = "1:3"
	row = makeRowFromSpan(rangeString)
	length = len(row.Cells)
	c.Assert(length, Equals, 3)
	rangeString = "5:7" // Note - we ignore lower bound!
	row = makeRowFromSpan(rangeString)
	length = len(row.Cells)
	c.Assert(length, Equals, 7)
	rangeString = "1:1"
	row = makeRowFromSpan(rangeString)
	length = len(row.Cells)
	c.Assert(length, Equals, 1)
}

func (l *LibSuite) TestReadRowsFromSheet(c *C) {
	var sharedstringsXML = bytes.NewBufferString(`
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
  <si>
    <t>Foo</t>
  </si>
  <si>
    <t>Bar</t>
  </si>
  <si>
    <t xml:space="preserve">Baz </t>
  </si>
  <si>
    <t>Quuk</t>
  </si>
</sst>`)
	var sheetxml = bytes.NewBufferString(`
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:B2"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <selection activeCell="C2" sqref="C2"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
  <sheetData>
    <row r="1" spans="1:2">
      <c r="A1" t="s">
        <v>0</v>
      </c>
      <c r="B1" t="s">
        <v>1</v>
      </c>
    </row>
    <row r="2" spans="1:2">
      <c r="A2" t="s">
        <v>2</v>
      </c>
      <c r="B2" t="s">
        <v>3</v>
      </c>
    </row>
  </sheetData>
  <pageMargins left="0.7" right="0.7"
               top="0.78740157499999996"
               bottom="0.78740157499999996"
               header="0.3"
               footer="0.3"/>
</worksheet>`)
	worksheet := new(xlsxWorksheet)
	err := xml.NewDecoder(sheetxml).Decode(worksheet)
	c.Assert(err, IsNil)
	sst := new(xlsxSST)
	err = xml.NewDecoder(sharedstringsXML).Decode(sst)
	c.Assert(err, IsNil)
	file := new(File)
	file.referenceTable = MakeSharedStringRefTable(sst)
	rows, maxCols, maxRows := readRowsFromSheet(worksheet, file)
	c.Assert(maxRows, Equals, 2)
	c.Assert(maxCols, Equals, 2)
	row := rows[0]
	c.Assert(len(row.Cells), Equals, 2)
	cell1 := row.Cells[0]
	c.Assert(cell1.String(), Equals, "Foo")
	cell2 := row.Cells[1]
	c.Assert(cell2.String(), Equals, "Bar")
}

func (l *LibSuite) TestReadRowsFromSheetWithLeadingEmptyRows(c *C) {
	var sharedstringsXML = bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t>ABC</t></si><si><t>DEF</t></si></sst>`)
	var sheetxml = bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <dimension ref="A4:A5"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <selection activeCell="A2" sqref="A2"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr baseColWidth="10" defaultRowHeight="15" x14ac:dyDescent="0"/>
  <sheetData>
    <row r="4" spans="1:1">
      <c r="A4" t="s">
        <v>0</v>
      </c>
    </row>
    <row r="5" spans="1:1">
      <c r="A5" t="s">
        <v>1</v>
      </c>
    </row>
  </sheetData>
  <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
  <pageSetup paperSize="9" orientation="portrait" horizontalDpi="4294967292" verticalDpi="4294967292"/>
  <extLst>
    <ext uri="{64002731-A6B0-56B0-2670-7721B7C09600}" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main">
      <mx:PLV Mode="0" OnePage="0" WScale="0"/>
    </ext>
  </extLst>
</worksheet>
`)
	worksheet := new(xlsxWorksheet)
	err := xml.NewDecoder(sheetxml).Decode(worksheet)
	c.Assert(err, IsNil)
	sst := new(xlsxSST)
	err = xml.NewDecoder(sharedstringsXML).Decode(sst)
	c.Assert(err, IsNil)

	file := new(File)
	file.referenceTable = MakeSharedStringRefTable(sst)
	_, maxCols, maxRows := readRowsFromSheet(worksheet, file)
	c.Assert(maxRows, Equals, 2)
	c.Assert(maxCols, Equals, 1)
}

func (l *LibSuite) TestReadRowsFromSheetWithEmptyCells(c *C) {
	var sharedstringsXML = bytes.NewBufferString(`
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="8" uniqueCount="5">
  <si>
    <t>Bob</t>
  </si>
  <si>
    <t>Alice</t>
  </si>
  <si>
    <t>Sue</t>
  </si>
  <si>
    <t>Yes</t>
  </si>
  <si>
    <t>No</t>
  </si>
</sst>
`)
	var sheetxml = bytes.NewBufferString(`
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:C3"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="D3" sqref="D3"/></sheetView></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
<sheetData>
  <row r="1" spans="1:3">
    <c r="A1" t="s">
      <v>
        0
      </v>
    </c>
    <c r="B1" t="s">
      <v>
        1
      </v>
    </c>
    <c r="C1" t="s">
      <v>
        2
      </v>
    </c>
  </row>
  <row r="2" spans="1:3">
    <c r="A2" t="s">
      <v>
        3
      </v>
    </c>
    <c r="B2" t="s">
      <v>
        4
      </v>
    </c>
    <c r="C2" t="s">
      <v>
        3
      </v>
    </c>
  </row>
  <row r="3" spans="1:3">
    <c r="A3" t="s">
      <v>
        4
      </v>
    </c>
    <c r="C3" t="s">
      <v>
        3
      </v>
    </c>
  </row>
</sheetData>
<pageMargins left="0.7" right="0.7" top="0.78740157499999996" bottom="0.78740157499999996" header="0.3" footer="0.3"/>
</worksheet>

`)
	worksheet := new(xlsxWorksheet)
	err := xml.NewDecoder(sheetxml).Decode(worksheet)
	c.Assert(err, IsNil)
	sst := new(xlsxSST)
	err = xml.NewDecoder(sharedstringsXML).Decode(sst)
	c.Assert(err, IsNil)
	file := new(File)
	file.referenceTable = MakeSharedStringRefTable(sst)
	rows, maxCols, maxRows := readRowsFromSheet(worksheet, file)
	c.Assert(maxRows, Equals, 3)
	c.Assert(maxCols, Equals, 3)

	row := rows[2]
	c.Assert(len(row.Cells), Equals, 3)

	cell1 := row.Cells[0]
	c.Assert(cell1.String(), Equals, "No")

	cell2 := row.Cells[1]
	c.Assert(cell2.String(), Equals, "")

	cell3 := row.Cells[2]
	c.Assert(cell3.String(), Equals, "Yes")
}

func (l *LibSuite) TestReadRowsFromSheetWithTrailingEmptyCells(c *C) {
	var row *Row
	var cell1, cell2, cell3, cell4 *Cell
	var sharedstringsXML = bytes.NewBufferString(`
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4"><si><t>A</t></si><si><t>B</t></si><si><t>C</t></si><si><t>D</t></si></sst>`)
	var sheetxml = bytes.NewBufferString(`
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:D8"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A7" sqref="A7"/></sheetView></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="15"/><sheetData><row r="1" spans="1:4"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c></row><row r="2" spans="1:4"><c r="A2"><v>1</v></c></row><row r="3" spans="1:4"><c r="B3"><v>1</v></c></row><row r="4" spans="1:4"><c r="C4"><v>1</v></c></row><row r="5" spans="1:4"><c r="D5"><v>1</v></c></row><row r="6" spans="1:4"><c r="C6"><v>1</v></c></row><row r="7" spans="1:4"><c r="B7"><v>1</v></c></row><row r="8" spans="1:4"><c r="A8"><v>1</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.78740157499999996" bottom="0.78740157499999996" header="0.3" footer="0.3"/></worksheet>
`)
	worksheet := new(xlsxWorksheet)
	err := xml.NewDecoder(sheetxml).Decode(worksheet)
	c.Assert(err, IsNil)

	sst := new(xlsxSST)
	err = xml.NewDecoder(sharedstringsXML).Decode(sst)
	c.Assert(err, IsNil)

	file := new(File)
	file.referenceTable = MakeSharedStringRefTable(sst)
	rows, maxCol, maxRow := readRowsFromSheet(worksheet, file)
	c.Assert(maxCol, Equals, 4)
	c.Assert(maxRow, Equals, 8)

	row = rows[0]
	c.Assert(len(row.Cells), Equals, 4)

	cell1 = row.Cells[0]
	c.Assert(cell1.String(), Equals, "A")

	cell2 = row.Cells[1]
	c.Assert(cell2.String(), Equals, "B")

	cell3 = row.Cells[2]
	c.Assert(cell3.String(), Equals, "C")

	cell4 = row.Cells[3]
	c.Assert(cell4.String(), Equals, "D")

	row = rows[1]
	c.Assert(len(row.Cells), Equals, 4)

	cell1 = row.Cells[0]
	c.Assert(cell1.String(), Equals, "1")

	cell2 = row.Cells[1]
	c.Assert(cell2.String(), Equals, "")

	cell3 = row.Cells[2]
	c.Assert(cell3.String(), Equals, "")

	cell4 = row.Cells[3]
	c.Assert(cell4.String(), Equals, "")
}

func (l *LibSuite) TestReadRowsFromSheetWithMultipleSpans(c *C) {
	var sharedstringsXML = bytes.NewBufferString(`
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
  <si>
    <t>Foo</t>
  </si>
  <si>
    <t>Bar</t>
  </si>
  <si>
    <t xml:space="preserve">Baz </t>
  </si>
  <si>
    <t>Quuk</t>
  </si>
</sst>`)
	var sheetxml = bytes.NewBufferString(`
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:D2"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <selection activeCell="C2" sqref="C2"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
  <sheetData>
    <row r="1" spans="1:2 3:4">
      <c r="A1" t="s">
        <v>0</v>
      </c>
      <c r="B1" t="s">
        <v>1</v>
      </c>
      <c r="C1" t="s">
        <v>0</v>
      </c>
      <c r="D1" t="s">
        <v>1</v>
      </c>
    </row>
    <row r="2" spans="1:2 3:4">
      <c r="A2" t="s">
        <v>2</v>
      </c>
      <c r="B2" t="s">
        <v>3</v>
      </c>
      <c r="C2" t="s">
        <v>2</v>
      </c>
      <c r="D2" t="s">
        <v>3</v>
      </c>
    </row>
  </sheetData>
  <pageMargins left="0.7" right="0.7"
               top="0.78740157499999996"
               bottom="0.78740157499999996"
               header="0.3"
               footer="0.3"/>
</worksheet>`)
	worksheet := new(xlsxWorksheet)
	err := xml.NewDecoder(sheetxml).Decode(worksheet)
	c.Assert(err, IsNil)
	sst := new(xlsxSST)
	err = xml.NewDecoder(sharedstringsXML).Decode(sst)
	c.Assert(err, IsNil)
	file := new(File)
	file.referenceTable = MakeSharedStringRefTable(sst)
	rows, maxCols, maxRows := readRowsFromSheet(worksheet, file)
	c.Assert(maxRows, Equals, 2)
	c.Assert(maxCols, Equals, 4)
	row := rows[0]
	c.Assert(len(row.Cells), Equals, 4)
	cell1 := row.Cells[0]
	c.Assert(cell1.String(), Equals, "Foo")
	cell2 := row.Cells[1]
	c.Assert(cell2.String(), Equals, "Bar")
	cell3 := row.Cells[2]
	c.Assert(cell3.String(), Equals, "Foo")
	cell4 := row.Cells[3]
	c.Assert(cell4.String(), Equals, "Bar")

}
