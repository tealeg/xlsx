package xlsx

import (
	"bytes"
	"encoding/xml"
	"strconv"
	"strings"
	"testing"
)

// Test we can correctly open a XSLX file and return a xlsx.File
// struct.
func TestOpenFile(t *testing.T) {
	var xlsxFile *File
	var error error

	files := []string{
		"testfile.xlsx",
		"issue18.xlsx",
	}
	for _, f := range files {
		xlsxFile, error = OpenFile(f)
		if error != nil {
			t.Error(error.Error())
			return
		}
		if xlsxFile == nil {
			t.Error("OpenFile returned nil FileInterface without generating an os.Error")
			return
		}
	}

}

// Test that when we open a real XLSX file we create xlsx.Sheet
// objects for the sheets inside the file and that these sheets are
// themselves correct.
func TestCreateSheet(t *testing.T) {
	var xlsxFile *File
	var error error
	var sheet *Sheet
	var row *Row
	xlsxFile, error = OpenFile("testfile.xlsx")
	if error != nil {
		t.Error(error.Error())
		return
	}
	if xlsxFile == nil {
		t.Error("OpenFile returned a nil File pointer but did not generate an error.")
		return
	}
	sheetLen := len(xlsxFile.Sheets)
	if sheetLen == 0 {
		t.Error("Expected len(xlsxFile.Sheets) > 0, but got ", sheetLen)
		return
	}
	sheet = xlsxFile.Sheets[0]
	rowLen := len(sheet.Rows)
	if rowLen != 2 {
		t.Error("Expected len(sheet.Rows) == 2, but got ", rowLen)
		return
	}
	row = sheet.Rows[0]
	if len(row.Cells) != 2 {
		t.Error("Expected len(row.Cells) == 2")
		return
	}
	cell := row.Cells[0]
	cellstring := cell.String()
	if cellstring != "Foo" {
		t.Error("Expected cell.String() == 'Foo', got ", cellstring)
	}
}

// Test that GetStyle correctly converts the xlsxStyle.Fonts.
func TestGetStyleWithFonts(t *testing.T) {
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

	cell = &Cell{Value: "123", styleIndex: 1, styles: xStyles}
	style = cell.GetStyle()
	if style == nil {
		t.Error("No style returned by cell.GetStyle()")
	}
	if style.Font.Size != 10 {
		t.Error("Expected style.Font.Size == 10, but got ", style.Font.Size)
	}
	if style.Font.Name != "Calibra" {
		t.Error("Expected style.Font.Name == 'Calibra', but got ", style.Font.Name)
	}
}

// Test that GetStyle correctly converts the xlsxStyle.Fills.
func TestGetStyleWithFills(t *testing.T) {
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

	cell = &Cell{Value: "123", styleIndex: 1, styles: xStyles}
	style = cell.GetStyle()
	fill := style.Fill
	if fill.PatternType != "solid" {
		t.Error("Expected fill.PatternType == 'solid', but got ",
			fill.PatternType)
	}
	if fill.BgColor != "00FF0000" {
		t.Error("Expected fill.BgColor == '00FF0000', but got ",
			fill.BgColor)
	}
	if fill.FgColor != "FF000000" {
		t.Error("Expected fill.FgColor == 'FF000000', but got ",
			fill.FgColor)
	}
}

// Test that GetStyle correctly converts the xlsxStyle.Borders.
func TestGetStyleWithBorders(t *testing.T) {
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

	cell = &Cell{Value: "123", styleIndex: 1, styles: xStyles}
	style = cell.GetStyle()
	border := style.Border
	if border.Left != "thin" {
		t.Error("Expected border.Left == 'thin', but got ",
			border.Left)
	}
	if border.Right != "thin" {
		t.Error("Expected border.Right == 'thin', but got ",
			border.Right)
	}
	if border.Top != "thin" {
		t.Error("Expected border.Top == 'thin', but got ",
			border.Top)
	}
	if border.Bottom != "thin" {
		t.Error("Expected border.Bottom == 'thin', but got ",
			border.Bottom)
	}
}

// Test that we can correctly extract a reference table from the
// sharedStrings.xml file embedded in the XLSX file and return a
// reference table of string values from it.
func TestReadSharedStringsFromZipFile(t *testing.T) {
	var xlsxFile *File
	var error error
	xlsxFile, error = OpenFile("testfile.xlsx")
	if error != nil {
		t.Error(error.Error())
		return
	}
	if xlsxFile.referenceTable == nil {
		t.Error("expected non nil xlsxFile.referenceTable")
		return
	}
}

func testXf(t *testing.T, result, expected *xlsxXf) {
	if result.ApplyAlignment != expected.ApplyAlignment {
		t.Error("Expected result.ApplyAlignment == ", expected.ApplyAlignment,
			", got", result.ApplyAlignment)
		return
	}
	if result.ApplyBorder != expected.ApplyBorder {
		t.Error("Expected result.ApplyBorder == ", expected.ApplyBorder,
			", got ", result.ApplyBorder)
		return
	}
	if result.ApplyFont != expected.ApplyFont {
		t.Error("Expect result.ApplyFont == ", expected.ApplyFont,
			", got ", result.ApplyFont)
		return
	}
	if result.ApplyFill != expected.ApplyFill {
		t.Error("Expected result.ApplyFill == ", expected.ApplyFill,
			", got ", result.ApplyFill)
		return
	}
	if result.ApplyProtection != expected.ApplyProtection {
		t.Error("Expexcted result.ApplyProtection == ", expected.ApplyProtection,
			", got ", result.ApplyProtection)
		return
	}
	if result.BorderId != expected.BorderId {
		t.Error("Expected BorderId == ", expected.BorderId,
			". got ", result.BorderId)
		return
	}
	if result.FillId != expected.FillId {
		t.Error("Expected result.FillId == ", expected.FillId,
			", got ", result.FillId)
		return
	}
	if result.FontId != expected.FontId {
		t.Error("Expected result.FontId == ", expected.FontId,
			", got ", result.FontId)
		return
	}
	if result.NumFmtId != expected.NumFmtId {
		t.Error("Expected result.NumFmtId == ", expected.NumFmtId,
			", got ", result.NumFmtId)
		return
	}
}

// We can correctly extract a style table from the style.xml file
// embedded in the XLSX file and return a styles struct from it.
func TestReadStylesFromZipFile(t *testing.T) {
	var xlsxFile *File
	var error error
	var fontCount, fillCount, borderCount, cellStyleXfCount, cellXfCount int
	var font xlsxFont
	var fill xlsxFill
	var border xlsxBorder
	var xf xlsxXf

	xlsxFile, error = OpenFile("testfile.xlsx")
	if error != nil {
		t.Error(error.Error())
		return
	}
	if xlsxFile.styles == nil {
		t.Error("expected non nil xlsxFile.styles")
		return
	}
	fontCount = len(xlsxFile.styles.Fonts)
	if fontCount != 4 {
		t.Error("expected exactly 4 xslxFonts, got ", fontCount)
		return
	}
	font = xlsxFile.styles.Fonts[0]
	if font.Sz.Val != "11" {
		t.Error("expected font.Sz.Val == 11, got ", font.Sz.Val)
		return
	}
	if font.Name.Val != "Calibri" {
		t.Error("expected font.Name.Val == 'Calibri', got ", font.Name.Val)
		return
	}
	fillCount = len(xlsxFile.styles.Fills)
	if fillCount != 3 {
		t.Error("Expected exactly 3 xlsxFills, got ", fillCount)
		return
	}
	fill = xlsxFile.styles.Fills[2]
	if fill.PatternFill.PatternType != "solid" {
		t.Error("Expected PatternFill.PatternType == 'solid', but got ",
			fill.PatternFill.PatternType)
		return
	}
	borderCount = len(xlsxFile.styles.Borders)
	if borderCount != 2 {
		t.Error("Expected exactly 2 xlsxBorders, got ", borderCount)
		return
	}
	border = xlsxFile.styles.Borders[1]
	if border.Left.Style != "thin" {
		t.Error("Expected border.Left.Style == 'thin', got ", border.Left.Style)
		return
	}
	if border.Right.Style != "thin" {
		t.Error("Expected border.Right.Style == 'thin', got ", border.Right.Style)
		return
	}
	if border.Top.Style != "thin" {
		t.Error("Expected border.Top.Style == 'thin', got ", border.Top.Style)
		return
	}
	if border.Bottom.Style != "thin" {
		t.Error("Expected border.Bottom.Style == 'thin', got ", border.Bottom.Style)
		return
	}
	cellStyleXfCount = len(xlsxFile.styles.CellStyleXfs)
	if cellStyleXfCount != 20 {
		t.Error("Expected excactly 20 cellStyleXfs, got ", cellStyleXfCount)
		return
	}
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
	testXf(t, &xf, expectedXf)

	cellXfCount = len(xlsxFile.styles.CellXfs)
	if cellXfCount != 3 {
		t.Error("Expected excactly 3 cellXfs, got ", cellXfCount)
		return
	}
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
	testXf(t, &xf, expectedXf)
}

func TestLettersToNumeric(t *testing.T) {
	cases := map[string]int{"A": 0, "G": 6, "z": 25, "AA": 26, "Az": 51,
		"BA": 52, "Bz": 77, "ZA": 26*26 + 0, "ZZ": 26*26 + 25,
		"AAA": 26*26 + 26 + 0, "AMI": 1022}
	for input, ans := range cases {
		output := lettersToNumeric(input)
		if output != ans {
			t.Error("Expected output '"+input+"' == ", ans,
				"but got ", strconv.Itoa(output))
		}
	}
}

func TestLetterOnlyMapFunction(t *testing.T) {
	var input string = "ABC123"
	var output string = strings.Map(letterOnlyMapF, input)
	if output != "ABC" {
		t.Error("Expected output == 'ABC' but got ", output)
	}
	input = "abc123"
	output = strings.Map(letterOnlyMapF, input)
	if output != "ABC" {
		t.Error("Expected output == 'ABC' but got ", output)
	}
}

func TestIntOnlyMapFunction(t *testing.T) {
	var input string = "ABC123"
	var output string = strings.Map(intOnlyMapF, input)
	if output != "123" {
		t.Error("Expected output == '123' but got ", output)
	}
}

func TestGetCoordsFromCellIDString(t *testing.T) {
	var cellIDString string = "A3"
	var x, y int
	var error error
	x, y, error = getCoordsFromCellIDString(cellIDString)
	if error != nil {
		t.Error(error)
	}
	if x != 0 {
		t.Error("Expected x == 0, but got ", strconv.Itoa(x))
	}
	if y != 2 {
		t.Error("Expected y == 2, but got ", strconv.Itoa(y))
	}
}

func TestGetMaxMinFromDimensionRef(t *testing.T) {
	var dimensionRef string = "A1:B2"
	var minx, miny, maxx, maxy int
	var err error
	minx, miny, maxx, maxy, err = getMaxMinFromDimensionRef(dimensionRef)
	if err != nil {
		t.Error(err)
	}
	if minx != 0 {
		t.Error("Expected minx == 0, but got ", strconv.Itoa(minx))
	}
	if miny != 0 {
		t.Error("Expected miny == 0, but got ", strconv.Itoa(miny))
	}
	if maxx != 1 {
		t.Error("Expected maxx == 0, but got ", strconv.Itoa(maxx))
	}
	if maxy != 1 {
		t.Error("Expected maxy == 0, but got ", strconv.Itoa(maxy))
	}

}

func TestGetRangeFromString(t *testing.T) {
	var rangeString string
	var lower, upper int
	var error error
	rangeString = "1:3"
	lower, upper, error = getRangeFromString(rangeString)
	if error != nil {
		t.Error(error)
	}
	if lower != 1 {
		t.Error("Expected lower bound == 1, but got ", strconv.Itoa(lower))
	}
	if upper != 3 {
		t.Error("Expected upper bound == 3, but got ", strconv.Itoa(upper))
	}
}

func TestMakeRowFromSpan(t *testing.T) {
	var rangeString string
	var row *Row
	var length int
	rangeString = "1:3"
	row = makeRowFromSpan(rangeString)
	length = len(row.Cells)
	if length != 3 {
		t.Error("Expected a row with 3 cells, but got ", strconv.Itoa(length))
	}
	rangeString = "5:7" // Note - we ignore lower bound!
	row = makeRowFromSpan(rangeString)
	length = len(row.Cells)
	if length != 7 {
		t.Error("Expected a row with 7 cells, but got ", strconv.Itoa(length))
	}
	rangeString = "1:1"
	row = makeRowFromSpan(rangeString)
	length = len(row.Cells)
	if length != 1 {
		t.Error("Expected a row with 1 cells, but got ", strconv.Itoa(length))
	}
}

func TestReadRowsFromSheet(t *testing.T) {
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
	error := xml.NewDecoder(sheetxml).Decode(worksheet)
	if error != nil {
		t.Error(error.Error())
		return
	}
	sst := new(xlsxSST)
	error = xml.NewDecoder(sharedstringsXML).Decode(sst)
	if error != nil {
		t.Error(error.Error())
		return
	}
	file := new(File)
	file.referenceTable = MakeSharedStringRefTable(sst)
	rows, maxCols, maxRows := readRowsFromSheet(worksheet, file)
	if maxRows != 2 {
		t.Error("Expected maxRows == 2")
	}
	if maxCols != 2 {
		t.Error("Expected maxCols == 2")
	}
	row := rows[0]
	if len(row.Cells) != 2 {
		t.Error("Expected len(row.Cells) == 2, got ", strconv.Itoa(len(row.Cells)))
	}
	cell1 := row.Cells[0]
	if cell1.String() != "Foo" {
		t.Error("Expected cell1.String() == 'Foo', got ", cell1.String())
	}
	cell2 := row.Cells[1]
	if cell2.String() != "Bar" {
		t.Error("Expected cell2.String() == 'Bar', got ", cell2.String())
	}

}

func TestReadRowsFromSheetWithEmptyCells(t *testing.T) {
	var sharedstringsXML = bytes.NewBufferString(`
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="8" uniqueCount="5"><si><t>Bob</t></si><si><t>Alice</t></si><si><t>Sue</t></si><si><t>Yes</t></si><si><t>No</t></si></sst>`)
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
	error := xml.NewDecoder(sheetxml).Decode(worksheet)
	if error != nil {
		t.Error(error.Error())
		return
	}
	sst := new(xlsxSST)
	error = xml.NewDecoder(sharedstringsXML).Decode(sst)
	if error != nil {
		t.Error(error.Error())
		return
	}
	file := new(File)
	file.referenceTable = MakeSharedStringRefTable(sst)
	rows, maxCols, maxRows := readRowsFromSheet(worksheet, file)
	if maxRows != 3 {
		t.Error("Expected maxRows == 3, got ", strconv.Itoa(len(rows)))
	}
	if maxCols != 3 {
		t.Error("Expected maxCols == 3, got ", strconv.Itoa(maxCols))
	}
	row := rows[2]
	if len(row.Cells) != 3 {
		t.Error("Expected len(row.Cells) == 3, got ", strconv.Itoa(len(row.Cells)))
	}
	cell1 := row.Cells[0]
	if cell1.String() != "No" {
		t.Error("Expected cell1.String() == 'No', got ", cell1.String())
	}
	cell2 := row.Cells[1]
	if cell2.String() != "" {
		t.Error("Expected cell2.String() == '', got ", cell2.String())
	}
	cell3 := row.Cells[2]
	if cell3.String() != "Yes" {
		t.Error("Expected cell3.String() == 'Yes', got ", cell3.String())
	}

}

func TestReadRowsFromSheetWithTrailingEmptyCells(t *testing.T) {
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
	error := xml.NewDecoder(sheetxml).Decode(worksheet)
	if error != nil {
		t.Error(error.Error())
		return
	}
	sst := new(xlsxSST)
	error = xml.NewDecoder(sharedstringsXML).Decode(sst)
	if error != nil {
		t.Error(error.Error())
		return
	}
	file := new(File)
	file.referenceTable = MakeSharedStringRefTable(sst)
	rows, maxCol, maxRow := readRowsFromSheet(worksheet, file)
	if maxCol != 4 {
		t.Error("Expected maxCol == 4, got ", strconv.Itoa(maxCol))

	}
	if maxRow != 8 {
		t.Error("Expected maxRow == 8, got ", strconv.Itoa(maxRow))

	}

	row = rows[0]
	if len(row.Cells) != 4 {
		t.Error("Expected len(row.Cells) == 4, got ", strconv.Itoa(len(row.Cells)))
	}
	cell1 = row.Cells[0]
	if cell1.String() != "A" {
		t.Error("Expected cell1.String() == 'A', got ", cell1.String())
	}
	cell2 = row.Cells[1]
	if cell2.String() != "B" {
		t.Error("Expected cell2.String() == 'B', got ", cell2.String())
	}
	cell3 = row.Cells[2]
	if cell3.String() != "C" {
		t.Error("Expected cell3.String() == 'C', got ", cell3.String())
	}
	cell4 = row.Cells[3]
	if cell4.String() != "D" {
		t.Error("Expected cell4.String() == 'D', got ", cell4.String())
	}

	row = rows[1]
	if len(row.Cells) != 4 {
		t.Error("Expected len(row.Cells) == 4, got ", strconv.Itoa(len(row.Cells)))
	}
	cell1 = row.Cells[0]
	if cell1.String() != "1" {
		t.Error("Expected cell1.String() == '1', got ", cell1.String())
	}
	cell2 = row.Cells[1]
	if cell2.String() != "" {
		t.Error("Expected cell2.String() == '', got ", cell2.String())
	}
	cell3 = row.Cells[2]
	if cell3.String() != "" {
		t.Error("Expected cell3.String() == '', got ", cell3.String())
	}
	cell4 = row.Cells[3]
	if cell4.String() != "" {
		t.Error("Expected cell4.String() == '', got ", cell4.String())
	}

}
