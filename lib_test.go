package xlsx

import (
	"bytes"
	"encoding/xml"
	"os"
	"strings"
	"testing"

	. "gopkg.in/check.v1"
)

type LibSuite struct{}

var _ = Suite(&LibSuite{})

// Attempting to open a file without workbook.xml.rels returns an error.
func (l *LibSuite) TestReadZipReaderWithFileWithNoWorkbookRels(c *C) {
	_, err := OpenFile("./testdocs/badfile_noWorkbookRels.xlsx")
	c.Assert(err, NotNil)
	c.Assert(err.Error(), Equals, "xl/_rels/workbook.xml.rels not found in input xlsx.")
}

// Attempting to open a file with no worksheets returns an error.
func (l *LibSuite) TestReadZipReaderWithFileWithNoWorksheets(c *C) {
	_, err := OpenFile("./testdocs/badfile_noWorksheets.xlsx")
	c.Assert(err, NotNil)
	c.Assert(err.Error(), Equals, "Input xlsx contains no worksheets.")
}

// Attempt to read data from a file with inlined string sheet data.
func (l *LibSuite) TestReadWithInlineStrings(c *C) {
	var xlsxFile *File
	var err error

	xlsxFile, err = OpenFile("./testdocs/inlineStrings.xlsx")
	c.Assert(err, IsNil)
	sheet := xlsxFile.Sheets[0]
	r1 := sheet.Rows[0]
	c1 := r1.Cells[1]

	val, err := c1.FormattedValue()
	if err != nil {
		c.Error(err)
		return
	}
	if val == "" {
		c.Error("Expected a string value")
		return
	}
	c.Assert(val, Equals, "HL Retail - North America - Activity by Day - MTD")
}

// which they are contained from the XLSX file, even when the
// worksheet files have arbitrary, non-numeric names.
func (l *LibSuite) TestReadWorkbookRelationsFromZipFileWithFunnyNames(c *C) {
	var xlsxFile *File
	var err error

	xlsxFile, err = OpenFile("./testdocs/testrels.xlsx")
	c.Assert(err, IsNil)
	bob := xlsxFile.Sheet["Bob"]
	row1 := bob.Rows[0]
	cell1 := row1.Cells[0]
	if val, err := cell1.FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "I am Bob")
	}
}

// We can marshal WorkBookRels to an xml file
func (l *LibSuite) TestWorkBookRelsMarshal(c *C) {
	var rels WorkBookRels = make(WorkBookRels)
	rels["rId1"] = "worksheets/sheet.xml"
	expectedXML := `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"></Relationship><Relationship Id="rId2" Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"></Relationship><Relationship Id="rId3" Target="theme/theme1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"></Relationship><Relationship Id="rId4" Target="styles.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"></Relationship></Relationships>`
	xRels := rels.MakeXLSXWorkbookRels()

	output := bytes.NewBufferString(xml.Header)
	body, err := xml.Marshal(xRels)
	c.Assert(err, IsNil)
	c.Assert(body, NotNil)
	_, err = output.Write(body)
	c.Assert(err, IsNil)
	c.Assert(output.String(), Equals, expectedXML)
}

// Excel column codes are a special form of base26 that doesn't allow
// zeros, except in the least significant part of the code.  Test we
// can smoosh the numbers in a normal base26 representation (presented
// as a slice of integers) down to this form.
func (l *LibSuite) TestSmooshBase26Slice(c *C) {
	input := []int{20, 0, 1}
	expected := []int{19, 26, 1}
	c.Assert(smooshBase26Slice(input), DeepEquals, expected)
}

// formatColumnName converts slices of base26 integers to alphabetical
// column names.  Note that the least signifcant character has a
// different numeric offset (Yuck!)
func (l *LibSuite) TestFormatColumnName(c *C) {
	c.Assert(formatColumnName([]int{0}), Equals, "A")
	c.Assert(formatColumnName([]int{25}), Equals, "Z")
	c.Assert(formatColumnName([]int{1, 25}), Equals, "AZ")
	c.Assert(formatColumnName([]int{26, 25}), Equals, "ZZ")
	c.Assert(formatColumnName([]int{26, 26, 25}), Equals, "ZZZ")
}

// getLargestDenominator returns the largest power of a provided value
// that can fit within a given value.
func (l *LibSuite) TestGetLargestDenominator(c *C) {
	d, p := getLargestDenominator(0, 1, 2, 0)
	c.Assert(d, Equals, 1)
	c.Assert(p, Equals, 0)
	d, p = getLargestDenominator(1, 1, 2, 0)
	c.Assert(d, Equals, 1)
	c.Assert(p, Equals, 0)
	d, p = getLargestDenominator(2, 1, 2, 0)
	c.Assert(d, Equals, 2)
	c.Assert(p, Equals, 1)
	d, p = getLargestDenominator(4, 1, 2, 0)
	c.Assert(d, Equals, 4)
	c.Assert(p, Equals, 2)
	d, p = getLargestDenominator(8, 1, 2, 0)
	c.Assert(d, Equals, 8)
	c.Assert(p, Equals, 3)
	d, p = getLargestDenominator(9, 1, 2, 0)
	c.Assert(d, Equals, 8)
	c.Assert(p, Equals, 3)
	d, p = getLargestDenominator(15, 1, 2, 0)
	c.Assert(d, Equals, 8)
	c.Assert(p, Equals, 3)
	d, p = getLargestDenominator(16, 1, 2, 0)
	c.Assert(d, Equals, 16)
	c.Assert(p, Equals, 4)
}

func (l *LibSuite) TestLettersToNumeric(c *C) {
	cases := map[string]int{"A": 0, "G": 6, "z": 25, "AA": 26, "Az": 51,
		"BA": 52, "BZ": 77, "ZA": 26*26 + 0, "ZZ": 26*26 + 25,
		"AAA": 26*26 + 26 + 0, "AMI": 1022}
	for input, ans := range cases {
		output := ColLettersToIndex(input)
		c.Assert(output, Equals, ans)
	}
}

func (l *LibSuite) TestNumericToLetters(c *C) {
	cases := map[string]int{
		"A":  0,
		"G":  6,
		"Z":  25,
		"AA": 26,
		"AZ": 51,
		"BA": 52,
		"BZ": 77, "ZA": 26 * 26, "ZB": 26*26 + 1,
		"ZZ":  26*26 + 25,
		"AAA": 26*26 + 26 + 0, "AMI": 1022}
	for ans, input := range cases {
		output := ColIndexToLetters(input)
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
	x, y, err = GetCoordsFromCellIDString(cellIDString)
	c.Assert(err, IsNil)
	c.Assert(x, Equals, 0)
	c.Assert(y, Equals, 2)
}

func (l *LibSuite) TestGetCellIDStringFromCoords(c *C) {
	c.Assert(GetCellIDStringFromCoords(0, 0), Equals, "A1")
	c.Assert(GetCellIDStringFromCoords(2, 2), Equals, "C3")
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
	var sheet *Sheet
	sheet = new(Sheet)
	rangeString = "1:3"
	row = makeRowFromSpan(rangeString, sheet)
	length = len(row.Cells)
	c.Assert(length, Equals, 3)
	c.Assert(row.Sheet, Equals, sheet)
	rangeString = "5:7" // Note - we ignore lower bound!
	row = makeRowFromSpan(rangeString, sheet)
	length = len(row.Cells)
	c.Assert(length, Equals, 7)
	c.Assert(row.Sheet, Equals, sheet)
	rangeString = "1:1"
	row = makeRowFromSpan(rangeString, sheet)
	length = len(row.Cells)
	c.Assert(length, Equals, 1)
	c.Assert(row.Sheet, Equals, sheet)
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
	  <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
  <sheetData>
    <row r="1" spans="1:2" ht="123.45" customHeight="1">
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
	sheet := new(Sheet)
	rows, cols, maxCols, maxRows := readRowsFromSheet(worksheet, file, sheet, NoRowLimit)
	c.Assert(maxRows, Equals, 2)
	c.Assert(maxCols, Equals, 2)
	row := rows[0]
	c.Assert(row.Sheet, Equals, sheet)
	c.Assert(len(row.Cells), Equals, 2)
	c.Assert(row.Height, Equals, 123.45)
	c.Assert(row.isCustom, Equals, true)
	cell1 := row.Cells[0]
	c.Assert(cell1.Value, Equals, "Foo")
	cell2 := row.Cells[1]
	c.Assert(cell2.Value, Equals, "Bar")
	col := cols[0]
	c.Assert(col.Min, Equals, 0)
	c.Assert(col.Max, Equals, 0)
	c.Assert(col.Hidden, Equals, false)
	c.Assert(len(worksheet.SheetViews.SheetView), Equals, 1)
	sheetView := worksheet.SheetViews.SheetView[0]
	c.Assert(sheetView.Pane, NotNil)
	pane := sheetView.Pane
	c.Assert(pane.XSplit, Equals, 0.0)
	c.Assert(pane.YSplit, Equals, 1.0)
}

func (l *LibSuite) TestReadRowsFromSheetWithMergeCells(c *C) {
	var sharedstringsXML = bytes.NewBufferString(`
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si>
    <t>Value A</t>
  </si>
  <si>
    <t>Value B</t>
  </si>
  <si>
    <t>Value C</t>
  </si>
</sst>
`)
	var sheetxml = bytes.NewBufferString(`
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr customHeight="1" defaultColWidth="17.29" defaultRowHeight="15.0"/>
  <cols>
    <col customWidth="1" min="1" max="6" width="14.43"/>
  </cols>
  <sheetData>
    <row r="1" ht="15.75" customHeight="1">
      <c r="A1" s="1" t="s">
        <v>0</v>
      </c>
    </row>
    <row r="2" ht="15.75" customHeight="1">
      <c r="A2" s="1" t="s">
        <v>1</v>
      </c>
      <c r="B2" s="1" t="s">
        <v>2</v>
      </c>
    </row>
  </sheetData>
  <mergeCells count="1">
    <mergeCell ref="A1:B1"/>
  </mergeCells>
  <drawing r:id="rId1"/>
</worksheet>`)
	worksheet := new(xlsxWorksheet)
	err := xml.NewDecoder(sheetxml).Decode(worksheet)
	c.Assert(err, IsNil)
	sst := new(xlsxSST)
	err = xml.NewDecoder(sharedstringsXML).Decode(sst)
	c.Assert(err, IsNil)
	file := new(File)
	file.referenceTable = MakeSharedStringRefTable(sst)
	sheet := new(Sheet)
	rows, _, _, _ := readRowsFromSheet(worksheet, file, sheet, NoRowLimit)
	row := rows[0] //
	cell1 := row.Cells[0]
	c.Assert(cell1.HMerge, Equals, 1)
	c.Assert(cell1.VMerge, Equals, 0)
}

// An invalid value in the "r" attribute in a <row> was causing a panic
// in readRowsFromSheet. This test is a copy of TestReadRowsFromSheet,
// with the important difference of the value 1048576 below in <row r="1048576", which is
// higher than the number of rows in the sheet. That number itself isn't significant;
// it just happens to be the value found to trigger the error in a user's file.
func (l *LibSuite) TestReadRowsFromSheetBadR(c *C) {
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
	  <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
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
    <row r="1048576" spans="1:2">
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

	sheet := new(Sheet)
	// Discarding all return values; this test is a regression for
	// a panic due to an "index out of range."
	readRowsFromSheet(worksheet, file, sheet, NoRowLimit)
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
	sheet := new(Sheet)
	rows, _, maxCols, maxRows := readRowsFromSheet(worksheet, file, sheet, NoRowLimit)
	c.Assert(maxRows, Equals, 5)
	c.Assert(maxCols, Equals, 1)

	c.Assert(len(rows[0].Cells), Equals, 0)
	c.Assert(len(rows[1].Cells), Equals, 0)
	c.Assert(len(rows[2].Cells), Equals, 0)
	c.Assert(len(rows[3].Cells), Equals, 1)
	if val, err := rows[3].Cells[0].FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "ABC")
	}
	c.Assert(len(rows[4].Cells), Equals, 1)
	if val, err := rows[4].Cells[0].FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "DEF")
	}
}

func (l *LibSuite) TestReadRowsFromSheetWithLeadingEmptyCols(c *C) {
	var sharedstringsXML = bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t>ABC</t></si><si><t>DEF</t></si></sst>`)
	var sheetxml = bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <dimension ref="C1:D2"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <selection activeCell="A2" sqref="A2"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr baseColWidth="10" defaultRowHeight="15" x14ac:dyDescent="0"/>
  <cols>
  	<col min="3" max="3" width="17" customWidth="1"/>
  	<col min="4" max="4" width="18" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1" spans="3:4">
      <c r="C1" t="s"><v>0</v></c>
      <c r="D1" t="s"><v>1</v></c>
    </row>
    <row r="2" spans="3:4">
      <c r="C2" t="s"><v>0</v></c>
      <c r="D2" t="s"><v>1</v></c>
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
	sheet := new(Sheet)
	rows, cols, maxCols, maxRows := readRowsFromSheet(worksheet, file, sheet, NoRowLimit)
	c.Assert(maxRows, Equals, 2)
	c.Assert(maxCols, Equals, 4)

	c.Assert(len(rows[0].Cells), Equals, 4)
	if val, err := rows[0].Cells[0].FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "")
	}
	if val, err := rows[0].Cells[1].FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "")
	}
	if val, err := rows[0].Cells[2].FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "ABC")
	}
	if val, err := rows[0].Cells[3].FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "DEF")
	}
	c.Assert(len(rows[1].Cells), Equals, 4)
	if val, err := rows[1].Cells[0].FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "")
	}
	if val, err := rows[1].Cells[1].FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "")
	}
	if val, err := rows[1].Cells[2].FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "ABC")
	}
	if val, err := rows[1].Cells[3].FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "DEF")
	}

	c.Assert(len(cols), Equals, 4)
	c.Assert(cols[0].Width, Equals, 0.0)
	c.Assert(cols[1].Width, Equals, 0.0)
	c.Assert(cols[2].Width, Equals, 17.0)
	c.Assert(cols[3].Width, Equals, 18.0)
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
	sheet := new(Sheet)
	rows, cols, maxCols, maxRows := readRowsFromSheet(worksheet, file, sheet, NoRowLimit)
	c.Assert(maxRows, Equals, 3)
	c.Assert(maxCols, Equals, 3)

	row := rows[2]
	c.Assert(row.Sheet, Equals, sheet)
	c.Assert(len(row.Cells), Equals, 3)

	cell1 := row.Cells[0]
	c.Assert(cell1.Value, Equals, "No")

	cell2 := row.Cells[1]
	c.Assert(cell2.Value, Equals, "")

	cell3 := row.Cells[2]
	c.Assert(cell3.Value, Equals, "Yes")

	col := cols[0]
	c.Assert(col.Min, Equals, 0)
	c.Assert(col.Max, Equals, 0)
	c.Assert(col.Hidden, Equals, false)
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
	sheet := new(Sheet)
	rows, _, maxCol, maxRow := readRowsFromSheet(worksheet, file, sheet, NoRowLimit)
	c.Assert(maxCol, Equals, 4)
	c.Assert(maxRow, Equals, 8)

	row = rows[0]
	c.Assert(row.Sheet, Equals, sheet)
	c.Assert(len(row.Cells), Equals, 4)

	cell1 = row.Cells[0]
	c.Assert(cell1.Value, Equals, "A")

	cell2 = row.Cells[1]
	c.Assert(cell2.Value, Equals, "B")

	cell3 = row.Cells[2]
	c.Assert(cell3.Value, Equals, "C")

	cell4 = row.Cells[3]
	c.Assert(cell4.Value, Equals, "D")

	row = rows[1]
	c.Assert(row.Sheet, Equals, sheet)
	c.Assert(len(row.Cells), Equals, 4)

	cell1 = row.Cells[0]
	c.Assert(cell1.Value, Equals, "1")

	cell2 = row.Cells[1]
	c.Assert(cell2.Value, Equals, "")

	cell3 = row.Cells[2]
	c.Assert(cell3.Value, Equals, "")

	cell4 = row.Cells[3]
	c.Assert(cell4.Value, Equals, "")
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
	sheet := new(Sheet)
	rows, _, maxCols, maxRows := readRowsFromSheet(worksheet, file, sheet, NoRowLimit)
	c.Assert(maxRows, Equals, 2)
	c.Assert(maxCols, Equals, 4)
	row := rows[0]
	c.Assert(row.Sheet, Equals, sheet)
	c.Assert(len(row.Cells), Equals, 4)
	cell1 := row.Cells[0]
	c.Assert(cell1.Value, Equals, "Foo")
	cell2 := row.Cells[1]
	c.Assert(cell2.Value, Equals, "Bar")
	cell3 := row.Cells[2]
	c.Assert(cell3.Value, Equals, "Foo")
	cell4 := row.Cells[3]
	c.Assert(cell4.Value, Equals, "Bar")

}

func (l *LibSuite) TestReadRowsFromSheetWithMultipleTypes(c *C) {
	var sharedstringsXML = bytes.NewBufferString(`
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
  <si>
    <t>Hello World</t>
  </si>
</sst>`)
	var sheetxml = bytes.NewBufferString(`
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:F1"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <selection activeCell="C1" sqref="C1"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
  <sheetData>
    <row r="1" spans="1:6">
      <c r="A1" t="s">
        <v>0</v>
      </c>
      <c r="B1">
        <v>12345</v>
      </c>
      <c r="C1">
        <v>1.024</v>
      </c>
      <c r="D1" t="b">
        <v>1</v>
      </c>
      <c r="E1">
      	<f>10+20</f>
        <v>30</v>
      </c>
      <c r="F1" t="e">
      	<f>10/0</f>
        <v>#DIV/0!</v>
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
	sheet := new(Sheet)
	rows, _, maxCols, maxRows := readRowsFromSheet(worksheet, file, sheet, NoRowLimit)
	c.Assert(maxRows, Equals, 1)
	c.Assert(maxCols, Equals, 6)
	row := rows[0]
	c.Assert(row.Sheet, Equals, sheet)
	c.Assert(len(row.Cells), Equals, 6)

	cell1 := row.Cells[0]
	c.Assert(cell1.Type(), Equals, CellTypeString)
	if val, err := cell1.FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "Hello World")
	}

	cell2 := row.Cells[1]
	c.Assert(cell2.Type(), Equals, CellTypeNumeric)
	intValue, _ := cell2.Int()
	c.Assert(intValue, Equals, 12345)

	cell3 := row.Cells[2]
	c.Assert(cell3.Type(), Equals, CellTypeNumeric)
	float, _ := cell3.Float()
	c.Assert(float, Equals, 1.024)

	cell4 := row.Cells[3]
	c.Assert(cell4.Type(), Equals, CellTypeBool)
	c.Assert(cell4.Bool(), Equals, true)

	cell5 := row.Cells[4]
	c.Assert(cell5.Type(), Equals, CellTypeNumeric)
	c.Assert(cell5.Formula(), Equals, "10+20")
	c.Assert(cell5.Value, Equals, "30")

	cell6 := row.Cells[5]
	c.Assert(cell6.Type(), Equals, CellTypeError)
	c.Assert(cell6.Formula(), Equals, "10/0")
	c.Assert(cell6.Value, Equals, "#DIV/0!")
}

func (l *LibSuite) TestReadRowsFromSheetWithHiddenColumn(c *C) {
	var sharedstringsXML = bytes.NewBufferString(`
		<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
		    <si><t>This is a test.</t></si>
		    <si><t>This should be invisible.</t></si>
		</sst>`)
	var sheetxml = bytes.NewBufferString(`
		<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<worksheet xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main"
		    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"
		    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
			<sheetViews><sheetView workbookViewId="0"/>
			</sheetViews>
			<sheetFormatPr customHeight="1" defaultColWidth="14.43" defaultRowHeight="15.75"/>
			<cols>
				<col hidden="1" max="2" min="2"/>
			</cols>
		    <sheetData>
		        <row r="1">
		            <c r="A1" s="1" t="s"><v>0</v></c>
		            <c r="B1" s="1" t="s"><v>1</v></c>
		        </row>
		    </sheetData><drawing r:id="rId1"/></worksheet>`)
	worksheet := new(xlsxWorksheet)
	err := xml.NewDecoder(sheetxml).Decode(worksheet)
	c.Assert(err, IsNil)
	sst := new(xlsxSST)
	err = xml.NewDecoder(sharedstringsXML).Decode(sst)
	c.Assert(err, IsNil)
	file := new(File)
	file.referenceTable = MakeSharedStringRefTable(sst)
	sheet := new(Sheet)
	rows, _, maxCols, maxRows := readRowsFromSheet(worksheet, file, sheet, NoRowLimit)
	c.Assert(maxRows, Equals, 1)
	c.Assert(maxCols, Equals, 2)
	row := rows[0]
	c.Assert(row.Sheet, Equals, sheet)
	c.Assert(len(row.Cells), Equals, 2)

	cell1 := row.Cells[0]
	c.Assert(cell1.Type(), Equals, CellTypeString)
	if val, err := cell1.FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "This is a test.")
	}
	c.Assert(cell1.Hidden, Equals, false)

	cell2 := row.Cells[1]
	c.Assert(cell2.Type(), Equals, CellTypeString)
	if val, err := cell2.FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "This should be invisible.")
	}
	c.Assert(cell2.Hidden, Equals, true)
}

// When converting the xlsxRow to a Row we create a as many cells as we find.
func (l *LibSuite) TestReadRowFromRaw(c *C) {
	var rawRow xlsxRow
	var cell xlsxC
	var row *Row

	rawRow = xlsxRow{}
	cell = xlsxC{R: "A1"}
	cell = xlsxC{R: "A2"}
	rawRow.C = append(rawRow.C, cell)
	sheet := new(Sheet)
	row = makeRowFromRaw(rawRow, sheet)
	c.Assert(row, NotNil)
	c.Assert(row.Cells, HasLen, 1)
	c.Assert(row.Sheet, Equals, sheet)
}

// When a cell claims it is at a position greater than its ordinal
// position in the file we make up the missing cells.
func (l *LibSuite) TestReadRowFromRawWithMissingCells(c *C) {
	var rawRow xlsxRow
	var cell xlsxC
	var row *Row

	rawRow = xlsxRow{}
	cell = xlsxC{R: "A1"}
	rawRow.C = append(rawRow.C, cell)
	cell = xlsxC{R: "E1"}
	rawRow.C = append(rawRow.C, cell)
	sheet := new(Sheet)
	row = makeRowFromRaw(rawRow, sheet)
	c.Assert(row, NotNil)
	c.Assert(row.Cells, HasLen, 5)
	c.Assert(row.Sheet, Equals, sheet)
}

// We can cope with missing coordinate references
func (l *LibSuite) TestReadRowFromRawWithPartialCoordinates(c *C) {
	var rawRow xlsxRow
	var cell xlsxC
	var row *Row

	rawRow = xlsxRow{}
	cell = xlsxC{R: "A1"}
	rawRow.C = append(rawRow.C, cell)
	cell = xlsxC{}
	rawRow.C = append(rawRow.C, cell)
	cell = xlsxC{R: "Z:1"}
	rawRow.C = append(rawRow.C, cell)
	cell = xlsxC{}
	rawRow.C = append(rawRow.C, cell)
	sheet := new(Sheet)
	row = makeRowFromRaw(rawRow, sheet)
	c.Assert(row, NotNil)
	c.Assert(row.Cells, HasLen, 27)
	c.Assert(row.Sheet, Equals, sheet)
}

func (l *LibSuite) TestSharedFormulas(c *C) {
	var sheetxml = bytes.NewBufferString(`
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:C2"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <selection activeCell="C1" sqref="C1"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
  <sheetData>
    <row r="1" spans="1:3">
      <c r="A1">
        <v>1</v>
      </c>
      <c r="B1">
        <v>2</v>
      </c>
      <c r="C1">
        <v>3</v>
      </c>
    </row>
    <row r="2" spans="1:3">
      <c r="A2">
        <v>2</v>
		<f t="shared" ref="A2:C2" si="0">2*A1</f>
      </c>
      <c r="B2">
        <v>4</v>
		<f t="shared" si="0"/>
      </c>
      <c r="C2">
        <v>6</v>
		<f t="shared" si="0"/>
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

	file := new(File)
	sheet := new(Sheet)
	rows, _, maxCols, maxRows := readRowsFromSheet(worksheet, file, sheet, NoRowLimit)
	c.Assert(maxCols, Equals, 3)
	c.Assert(maxRows, Equals, 2)

	row := rows[1]
	c.Assert(row.Cells[1].Formula(), Equals, "2*B1")
	c.Assert(row.Cells[2].Formula(), Equals, "2*C1")
}

// Test shared formulas that have absolute references ($) in them
func (l *LibSuite) TestSharedFormulasWithAbsoluteReferences(c *C) {
	formulas := []string{
		"A1",
		"$A1",
		"A$1",
		"$A$1",
		"A1+B1",
		"$A1+B1",
		"$A$1+B1",
		"A1+$B1",
		"A1+B$1",
		"A1+$B$1",
		"$A$1+$B$1",
		`IF(C23>=E$12,"Q4",IF(C23>=$D$12,"Q3",IF(C23>=C$12,"Q2","Q1")))`,
		`SUM(D44:H44)*IM_A_DEFINED_NAME`,
		`IM_A_DEFINED_NAME+SUM(D44:H44)*IM_A_DEFINED_NAME_ALSO`,
		`SUM(D44:H44)*IM_A_DEFINED_NAME+A1`,
		"AA1",
		"$AA1",
		"AA$1",
		"$AA$1",
	}

	expected := []string{
		"B2",
		"$A2",
		"B$1",
		"$A$1",
		"B2+C2",
		"$A2+C2",
		"$A$1+C2",
		"B2+$B2",
		"B2+C$1",
		"B2+$B$1",
		"$A$1+$B$1",
		`IF(D24>=F$12,"Q4",IF(D24>=$D$12,"Q3",IF(D24>=D$12,"Q2","Q1")))`,
		`SUM(E45:I45)*IM_A_DEFINED_NAME`,
		`IM_A_DEFINED_NAME+SUM(E45:I45)*IM_A_DEFINED_NAME_ALSO`,
		`SUM(E45:I45)*IM_A_DEFINED_NAME+B2`,
		"AB2",
		"$AA2",
		"AB$1",
		"$AA$1",
	}

	anchorCell := "C4"

	sharedFormulas := map[int]sharedFormula{}
	x, y, _ := GetCoordsFromCellIDString(anchorCell)
	for i, formula := range formulas {
		res := formula
		sharedFormulas[i] = sharedFormula{x, y, res}
	}

	for i, formula := range formulas {
		testCell := xlsxC{
			R: "D5",
			F: &xlsxF{
				Content: formula,
				T:       "shared",
				Si:      i,
			},
		}

		c.Assert(formulaForCell(testCell, sharedFormulas), Equals, expected[i])
	}
}

// Avoid panic when cell.F.T is "e" (for error)
func (l *LibSuite) TestFormulaForCellPanic(c *C) {
	cell := xlsxC{R: "A1"}
	// This line would panic before the fix.
	sharedFormulas := make(map[int]sharedFormula)

	// Not really an important test; getting here without a
	// panic is the real win.
	c.Assert(formulaForCell(cell, sharedFormulas), Equals, "")
}

func (l *LibSuite) TestRowNotOverwrittenWhenFollowedByEmptyRow(c *C) {
	sheetXML := bytes.NewBufferString(`
	<?xml version="1.0" encoding="UTF-8"?>
	<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
		<sheetViews>
			<sheetView workbookViewId="0" />
		</sheetViews>
		<sheetFormatPr customHeight="1" defaultColWidth="14.43" defaultRowHeight="15.75" />
		<sheetData>
			<row r="2">
				<c r="A2" t="str">
					<f t="shared" ref="A2" si="1">RANDBETWEEN(1,100)</f>
					<v>66</v>
				</c>
			</row>
			<row r="3">
				<c r="A3" t="str">
					<f t="shared" ref="A3" si="2">RANDBETWEEN(1,100)</f>
					<v>30</v>
				</c>
			</row>
			<row r="4">
				<c r="A4" t="str">
					<f t="shared" ref="A4" si="3">RANDBETWEEN(1,100)</f>
					<v>75</v>
				</c>
			</row>
			<row r="7">
				<c r="A7" s="1" t="str">
					<f t="shared" ref="A7" si="4">A4/A2</f>
					<v>1.14</v>
				</c>
			</row>
		</sheetData>
		<drawing r:id="rId1" />
	</worksheet>
	`)

	sharedstringsXML := bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`)

	worksheet := new(xlsxWorksheet)
	err := xml.NewDecoder(sheetXML).Decode(worksheet)
	c.Assert(err, IsNil)

	sst := new(xlsxSST)
	err = xml.NewDecoder(sharedstringsXML).Decode(sst)
	c.Assert(err, IsNil)

	file := new(File)
	file.referenceTable = MakeSharedStringRefTable(sst)

	sheet := new(Sheet)
	rows, _, _, _ := readRowsFromSheet(worksheet, file, sheet, NoRowLimit)
	cells := rows[3].Cells

	c.Assert(cells, HasLen, 1)
	c.Assert(cells[0].Value, Equals, "75")
}

// This was a specific issue raised by a user.
func (l *LibSuite) TestRoundTripFileWithNoSheetCols(c *C) {
	originalXlFile, err := OpenFile("testdocs/original.xlsx")
	c.Assert(err, IsNil)
	err = originalXlFile.Save("testdocs/after_write.xlsx")
	c.Assert(err, IsNil)
	_, err = OpenFile("testdocs/after_write.xlsx")
	c.Assert(err, IsNil)
	err = os.Remove("testdocs/after_write.xlsx")
	c.Assert(err, IsNil)
}

func (l *LibSuite) TestReadRestEmptyRowsFromSheet(c *C) {
	originalXlFile, err := OpenFile("testdocs/empty_rows_in_the_rest.xlsx")
	c.Assert(err, IsNil)
	for _, sheet := range originalXlFile.Sheets {
		for _, row := range sheet.Rows {
			if row == nil {
				c.Errorf("Row should not be nil")
			}
		}
	}
}

// See issue #362
// An XSLX file with an invalid sheet name (xl/worksheets.xml) caused an exception
func TestFuzzCrashers(t *testing.T) {
	var crashers = []string{
		"PK\x03\x04\x14\x00\x00\b\b\x00D\xae\fC\xf4\xeb\xcaY=\x01" +
			"\x00\x00g\x05\x00\x00\x13\x00\x00\x00[Content_T" +
			"ypes].xml͔\xd1K\xc30\x10\xc6\xdf\xf7W" +
			"\x94\xbcJ\x9bm\x82\x88\xb4ۃ\xe0\xa3\x0e\x9c\xcf\x12\x93\xdb\x1a\xda" +
			"&\xe1.\xce\xed\xbf\xf7R7A\x11\\q\xa8/\r%\xf7}\xbf" +
			"\xef\x8ek\xcb\xf9\xb6k\xb3\r Y\xef*1)\xc6\"\x03\xa7\xbd" +
			"\xb1n]\x89\x87\xe5M~)\xe6\xb3Q\xb9\xdc\x05\xa0\x8ck\x1dU" +
			"\xa2\x8e1\\II\xba\x86NQ\xe1\x038\xbeYy\xecT\xe4W" +
			"\\ˠt\xa3\xd6 \xa7\xe3\xf1\x85\xd4\xdeEp1\x8f\xc9C\xcc" +
			"\xca;ơ5\x90-\x14\xc6[\xd5A%\xe4#BK\xb2HO" +
			"\x91]\xbf\t\x12\xb3\x12*\x84\xd6j\x159\x9f\xdc8\xf3\x89\x96\xef" +
			"II\xd9\xd7Pm\x03\x9dq\x81\x90_\x93\x8c\xd7\v\xf4\x81$\x1b" +
			"\x17\xa9n\x10ίVV\x03{<w,)`\xcbJ\x03&\x0f" +
			"l\t\x18-\x1c\xc7\xd6\x1ea8\xfc\xd0kR\x1fIܶ\xfb\xd1" +
			"\xbexl\x9e\xbco\x12\xf57\xc6\xcc`\xaa\x15\x82\xb9\x8fȻD" +
			"?\x1e5\x05\x04e\xa8\x06\x88\xdc\xc1\a\xefor\xa4\xd6{\x1d\xc9" +
			"\xfe8?q\x96w\xff\x819\xa6\xff$\xc7\xe4\x8frPܵp" +
			"\xf2\xc5\xe8M\x8f\x98\xc0\xe1c8\xe9R\xf2Ytʺ=\u007fT" +
			"\xca\xfe\xc79{\x05PK\x03\x04\x14\x00\x00\b\b\x00D\xae\fC" +
			"f\xaa\x82\xb7\xe0\x00\x00\x00;\x02\x00\x00\v\x00\x00\x00_rel" +
			"s/.rels\xad\x92\xcfJ\x031\x10\x87\xef}\x8a\x90{" +
			"w\xb6\x15Dd\xb3\xbd\x88ЛH}\x80\x98\xcc\xfea7\x990" +
			"\x19u}{\x83\bZ\xa9\xa5\a\x8fI~\xf3\xcd7C\x9a\xdd\x12" +
			"f\xf5\x8a\x9cG\x8aFo\xaaZ+\x8c\x8e\xfc\x18{\xa3\x9f\x0e\xf7" +
			"\xeb\x1b\xbdkW\xcd#\xceVJ$\x0fcʪ\xd4\xc4l\xf4 " +
			"\x92n\x01\xb2\x1b0\xd8\\Q\xc2X^:\xe2`\xa5\x1c\xb9\x87d" +
			"\xddd{\x84m]_\x03\xffd\xe8\xf6\x88\xa9\xf6\xdeh\xde\xfb\x8d" +
			"V\x87\xf7\x84\x97\xb0\xa9\xebF\x87w\xe4^\x02F9\xd1\xe2W\xa2" +
			"\x90-\xf7(F/3\xbc\x11O\xcfDSU\xa0\x1aN\xbbl/" +
			"w\xf9{N\b(\xd6[\xb1\xe0\x88q\x9d\xb8T\xb3\x8c\x98\xbfu" +
			"<\xb9\x87r\x9d?\x13焮\xfes9\xb8\bF\x8f\xfe\xbc\x92" +
			"M\xe9\xcbh\xd5\xc0\xd1'h?\x00PK\x03\x04\x14\x00\x00\b\b" +
			"\x00D\xae\fC\x17ϯ\xa7\xbc\x00\x00\x005\x01\x00\x00\x10\x00\x00" +
			"\x00docProps/app.xml\x9d\x8f\xb1" +
			"j\x031\x10\x05{\u007f\x85Po\xeb\xe2\xc2\x04\xa3\x93\t$\xee\x02" +
			"..酴g\v\xa4]\xa1ݘ\xf3\xdfG!\x10\xa7v9" +
			"\f\f\xef\xd9\xc3R\xb2\xbaB\xe3D8\xea\xa7͠\x15`\xa0\x98" +
			"\xf0<\xea\x8f\xe9\xb8~֊\xc5c\xf4\x99\x10F}\x03\xd6\a\xb7" +
			"\xb2\xa7F\x15\x9a$`\xd5\vȣ\xbe\x88Խ1\x1c.P<" +
			"o\xba\xc6nfj\xc5K\xc7v64\xcf)\xc0+\x85\xaf\x02(" +
			"f;\f;\x03\x8b\x00F\x88\xeb\xfa\x17Կ\xc5\xfdU\x1e\x8dF" +
			"\n?\xfb\xf8s\xba\xd5\xdesv\"\xf1yJ\x05\xdc`\xcd\x1d\xec" +
			"K\xad9\x05/\xfd\xbc{O\xa1\x11\xd3,\xeam\t\x90\xad\xf9/" +
			"\xad\xb9\x1fv\xdfPK\x03\x04\x14\x00\x00\b\b\x00D\xae\fC\x17" +
			"qy\xdb:\x01\x00\x00x\x02\x00\x00\x11\x00\x00\x00docPr" +
			"ops/core.xml\x8d\x92_O\xc3 \x14\xc5" +
			"\xdf\xfd\x14\r\xef-\xa5]\x96\x85\xb4]\xa2f\xbe\xb8\xc4\xc4\x1a\x8d" +
			"o\x04\xee:b\xa1\x04pݾ\xbdm\xb7\xe1\xd4=\xf8ƽ\xe7" +
			"\xdc\x1f\x87?\xc5r\xaf\xdah\a\xd6\xc9N\x97\x88$)\x8a@\xf3" +
			"NHݔ\xe8\xa5^\xc5\v\x149ϴ`m\xa7\xa1D\ap" +
			"hY\xdd\x14\xdcP\xdeYx\xb2\x9d\x01\xeb%\xb8h\x00iG\xb9" +
			")\xd1\xd6{C1v|\v\x8a\xb9dp\xe8A\xdctV1?" +
			"\x94\xb6\xc1\x86\xf1\x0f\xd6\x00\xce\xd2t\x8e\x15x&\x98gx\x04\xc6" +
			"&\x10\xd1\t)x@\x9aO\xdbN\x00\xc11\xb4\xa0@{\x87I" +
			"B\xf0\xb7WI\u007f0pu\xe2,^\xb8=X定'%" +
			"8\xf7N\x06W\xdf\xf7I\x9fO\xbe!?\xc1o\xeb\xc7\xe7\xe9\xa8" +
			"\xb1\xd4\xe3Uq@UqBSn\x81y\x10\xd1\x00\xa0\xc7`g" +
			"\xe55\xbf\xbb\xafW\xa8\xcaRB\xe2t\x1eg\x8b\x9a\xcc(\xc9i" +
			"\x9a\xbf\x17\xf8\xd7\xfc\b<\xae;[\xd5\xc0Zx\x18=\xa15\xbe" +
			"G˜_\x0f/\xb7\x91 n\x0f\xc1\xf5W\t\xe1ԩ\xf7\xff" +
			"t3\x9ag\x17\xe9\u0380i\u007f\v;9~\xa3*\x9d6\r\xe5" +
			"T\xfd\xfc,\xd5\x17PK\x03\x04\x14\x00\x00\b\b\x00D\xae\fC" +
			"(\xba\xe5Ҧ\x00\x00\x00\xec\x00\x00\x00\x14\x00\x00\x00xl/s" +
			"haredStrings.xmle\xce\xc1\x8a" +
			"\xc20\x10\xc6\xf1\xbbO\x11\xe6\xbeMw\x11\x11IRX\xa1\xf7\x85" +
			"\xf5\x01B;\xda`3\xa9\x99\xc9\xe2\xfa\xf4V\x04\x05=\xfe\u007f\x03" +
			"\x1fc\x9as\x1c\xd5\x1ff\x0e\x89,|V5(\xa4.\xf5\x81\x0e" +
			"\x16v\xbf\xed\xc7\x1a\x14\x8b\xa7ޏ\x89\xd0\xc2?24na\x98" +
			"Eu\xa9\x90XX\x82*\x14N\x05\xb7\x8f\x9eG\x89-\f\"\xd3" +
			"Fk\xee\x06\x8c\x9e\xab4!͗}\xca\xd1˜\xf9\xa0y\xca" +
			"\xe8{\x1e\x10%\x8e\xfa\xab\xaeW:\xfa@\xe0\f\agĵ)" +
			"\x19-\xce\xe8[\xde\xe9\xdb\xe7w\xba\xbc\xd2O)ǧ\xe9\xf9]" +
			"w\x05PK\x03\x04\x14\x00\x00\b\b\x00\xcf,\rC\x0ep\x99\x04" +
			"\f\x04\x00\x00\x96\x1f\x00\x00\r\x00\x1c\x00xl/style" +
			"s.xmlUT\t\x00\x035\xaa\tR5\xaa\tRux" +
			"\v\x00\x01\x04\xe8\x03\x00\x00\x04\xe8\x03\x00\x00\xedYQo\xdb*\x14" +
			"~\xef\xaf@~\xbf\xb3\x13\xa7^|\x95tjs\x97\xabI\xd3T" +
			"m\x9dt\xa5\xab\xfb@ll\xa3a\xb00ْ\xfd\xfa\x81\xb1\xb1" +
			"\x9d6\x90\xaa/ծ\x1dU\x82\xc3\xc7\xc7\xc7\xe1p\\\xf0\xea\xdd" +
			"\xa1$\xe0;\xe25ft\xed\xcd\xde\x04\x1e@4a)\xa6\xf9\xda" +
			"\xfb\xfa\xb0\xfdc\xe9\x81Z@\x9aB\xc2(Z{GT{\xefn" +
			"\xaeV\xb58\x12\xf4\xa5@H\x00\xc9@\xeb\xb5W\bQ\xfd\xe9\xfb" +
			"uR\xa0\x12\xd6oX\x85\xa8l\xc9\x18/\xa1\x90U\x9e\xfbu\xc5" +
			"\x11Lkթ$\xfe<\b\"\xbf\x84\x98z7W\x00\xac\xe8\xbe" +
			"ܖ\xa2\x06\t\xdbS!\x954Vc\a\x9ag\xc3R\xa9\xe1\xef" +
			"\xf7\x9f\xde\u007f\xbe\xfd\xe8\x01\xdd\xf6!\x95\xf0h\xe1\xf9\r\x8f\xdf\x12" +
			"5\x95\x8cўr\xd1Q*\xab.\xcaJ\xfd\x13|\x87D2\xcc" +
			"4AcL\x18a\x1c\xf0|\xb7\xf6\xb6۠y\x06\xad\x14\x96H" +
			"w\xda@\x82w\x1c\x0f\xda2Xbrԭ\xf3!c\x01y-" +
			"\x9d\xa5\xc7\xea\x1aV~\xaf匬\xa7\a\xbe\xe5\x18\x923\xc3\x06" +
			"\xbf\x17\xbb.\xb5\xeb\x89\t1\xeb\x19\x9a\xf5\x94V\xc3VA!\x10" +
			"\xa7[i\x03m\xf9\xe1Xɨ\xa12~\a\xec\xa6υ\xdds" +
			"\x0e\x8f\xb3\xf9\xf5\v\x18jFp\xeau8\xd5-ߌ\xe3\xecn" +
			"\xb9\xdd\xf6\x8e\x91\x88\xdd\x18\x11\x86\x1b\xf9\f\\\xe7\x0f\x06;ե" +
			"K\xdao;\xc6S\xb9\xc7;\xcf\xcd;\xcfi;H1\xcc\x19\x85" +
			"\xe4/\xf6C&\x81\f\x92\x1ay\xc6\xf8\xb5\xeaLfX\x822\xd1" +
			"\x8b\xe08/\x06U\xc1\xaa\xbe\xb2cB\xb0\xb2\xafw\xa4ƋZ" +
			"\xc1\xcb\xe5\x80&!\xad=QȄr\xa2\xedL\x9b\x14z\xa6E" +
			"\xab>\xd3h\x99BW\xd6NO\x10!_\x14\xc5?Y\xef\xf9\xa0" +
			"s\xfd!\x03\xb0\xaa\xc8\xf1\x96\xe0\x9c\x96H5\n\xbe\x97\x13m\xac" +
			"w\r\xcfȴe'\x98{\xce\x04JD\x93\xb8\xb5Y\x8f\xae\xf2" +
			"\xa1L\xe3j\xf5\xbb\xa2\xec\xda\x16\xc7)\xd3L\nv*@\xc18" +
			"\xfe)\xf1j7\xe6\x88\".\xf7*\xc04m\x14J\x82\xba\xe0\x98" +
			"~{`[,\xcc\xda\bt\x10\x9f\x99\x80Z\x8b\x04\xc9\x17\x8a\xc0" +
			"\x89\xa2Ю\xf4\xc0\x0f\x0e\xab\a\t\xeb\xfa\xf4\xfe\xac\xcc4@\x81" +
			"S9\x8e\xa1%,\xf9\x86\xd2vr\xc6݇\xec\xbc\aێ#" +
			"\x17\x0emV\x1f\xb6@\x97\x13gC'\x9a\xe5\x9cd=-k>" +
			"\xc9z\x86\xac`\x925ɚdM\xb2&Y\xbf\xb1\xacѫz" +
			"\x11\xbeR]\xb3W\xaak\xf1Ju\xcd_\xa7\xae\xf8D\xd6\xca\x1f" +
			"\x9eK\xccAepF\t-G\x94\xcb\xe4\x0f-/ݶ\xea\x98" +
			"\x02\x0e\xd9p\aO\xe7\x15\xa77\xe7\x937\x9fw\x80~\xe4\xcc\xd9" +
			"\xff&4uF\x18&\x83&;\x98|\x10u\xf9\xc04\x81\xdd\x1e" +
			"\x13\x81i\xeb\x91d_Kaw\xdaf\x86T\x17~k\uf4fa" +
			"\xb7%\xbd\x9b|\x1bWh\xe7ڰ\xb2\x84\x1d\xd5\xec\xda\xce\x15]" +
			"\xc0\x05\xfe\r\xfe3|\x91\x9do\xe1\xe0\xdbs\x8ehr4to" +
			"\xedto/\xa3\x1b)\\\xda)\xaf\xed\x94\xf7\x88'2\xf2\f[" +
			"\xdc]\x98\xf7+\xae\x03@\xdd8\xd6\xed@*Z\x0f(\xdd\fl" +
			"\xeaV-\xdfm\xac7\xe4\xa7\x00\xfdX\x01\x0e\x86 P\u007fV\x80" +
			"\x82848D:\x18T\xb3\x15\xb0t\xf9!\b\x96.\x80\x82X" +
			"\x87p0,]\f\xaa\xd9\n\xd8\x04\xea\xe7\xd0`e\x88\xe5cu" +
			"T\x1c\x87a\x149\x16kt\xd3\xfdH\xe4Ʊ\x16Q\x14\x04\x8e" +
			"!\x1c\xb3P\xfd\x1d\x1a\x94\x8a\x97\xac\xa63\xe4\x9cA{QL:" +
			"\"ꂠu8ʹq\xc6_6\x9e\xbf\x9a\n`]\v\xf5" +
			"ı5\xe4\x1c\x8b\xa5!V\r\x8e\xa8\xd6\x10\v@\x05\xbd\x95\xe1" +
			"\xd1\a\x9e\xa7f\xe1HbN@\x1c;\x00j\xebXwV\x14\xd9" +
			"]\x1d\xa9\x9f5\x1e\x1c\xdb?\f\xe3\xd8\nP\fV\x91a\xe8\x00" +
			"\xa8\x14\xe4\x048D*\x99\x0e@\x18\xf6\xffn=z\x99\xcawo" +
			"[^\xf9\xfd\a\xf5\x9b\xab_PK\x03\x04\x14\x00\x00\b\b\x00\xe0" +
			"\x15\xf7D\b\xc40\xf9\xbe\x01\x00\x00}\x03\x00\x00\x0f\x00\x1c\x00x" +
			"l/workbook.xmlUT\t\x00\x03\x03" +
			"\x06\xcfS\x03\x06\xcfSux\v\x00\x01\x04\xe8\x03\x00\x00\x04\xe8\x03" +
			"\x00\x00\x8dRMO\xdb@\x10\xbd\xe7W\xac\xf6N\xfc\x91\x88\x86\xc8" +
			"\x0e\xaa@\b\x0e\x94\xaaP8\x8f\xd7\xe3x\x9b\xf5\xae\xb5;I\xa0" +
			"\xbf\xbec;.)\xaa\"n;_o\xdf{3\xd9\xe5kc\xc4" +
			"\x0e}\xd0\xce\xe62\x99\xc6R\xa0U\xae\xd4v\x9d˟O7g" +
			"\v)\x02\x81-\xc18\x8b\xb9|\xc3 /W\x93l\xef\xfc\xa6p" +
			"n#xކ\\\xd6D\xed2\x8a\x82\xaa\xb1\x810u-Z\xae" +
			"T\xce7@\x1c\xfau\x14Z\x8fP\x86\x1a\x91\x1a\x13\xa5q|\x1e" +
			"5\xa0\xad\x1c\x10\x96\xfe3\x18\xae\xaa\xb4\xc2k\xa7\xb6\rZ\x1a@" +
			"<\x1a f\x1fj\xdd\x06\xb9\x9a\b\x91U\xda\xe0\xf3\xa0I@\xdb" +
			"~\x83\x86\x99_\x81Q2\xea\xeb#\xf9\xef^\x14\xa06\xdb\xf6\x86" +
			"\arY\x81\t\xc8rk\xb7\u007f(~\xa1\"\xd6\x05\xc6HQ\x02" +
			"ar\x11\xcfǖ\x8f(\x8e\xb8\x99?\x1b\xf2]\xeeY\xe3>t" +
			"\xd1Q_\x97\x13\xc0\x8d;|\x82\"\x97lu\xa5}\xa0\xc7Β" +
			">\xec~\xbeu^\xffv\x96\xc0<*\xef\x8c\xc9%\xf9\xed\x81U" +
			"\xdfɳ\xe18\xc9BI\xab\x8f\xed\x04ŏΖ\\ο\xf0" +
			"\n\xf7ږ\x8c\x8dz]\xf3W\x8b\xe4\"\x1ds/\xba\xa4\x9a7" +
			"\u007f>[\xccy\x17/}\xb2g\xf3\xf6\xfe\x1e\x84E\xff(\xcb\xfa" +
			"U\x8e\"\xfb@\xd8\xdeif\x88\xc6`\xd2\xf1\xe3\xec]\xc9\xf0\xfd" +
			"\x15\x11\x17w:\xe8\xc20C\xbf\xd4\\\xf0we:\xc0\xff\x1f$" +
			"=\x02IO\x80\xccN\x81̎@f'@\xe6\xa3\xd0wi\x19" +
			";\xab\xf8N4\xa1\xe7\xa9+\xb7\xb5l`\x12\xb3?\x1e\xab{W" +
			"2\xd0WVw\xa8\xff=\xa2C|\x8d\x86\x80\x1d\x9c\xc6q\x9ct" +
			"\xe8Y4\x9e\xc3j\xf2\aPK\x03\x04\x14\x00\x00\b\b\x00\xb2\x04" +
			"4C\xa0J\x80\x9e\x84\x03\x00\x00\x8f\b\x00\x00\x18\x00\x1c\x00xl" +
			"/worksheets/sheet1.x" +
			"mlUT\t\x00\x03\xb0|;R\xb0|;Rux\v\x00\x01" +
			"\x04\xe8\x03\x00\x00\x04\xe8\x03\x00\x00\xc5V\xc1r\xdb6\x10\xbd\xfb+" +
			"0\xb8W\x94\x14+V<\x923\x89]5\x9dq\"O\xe543" +
			"\xbdA\xc4R\xc4\x18\x04X\x00\x94b\u007f}w\x01\x92\xa2\x155\xd3" +
			"[u\x11\xb1\x8bžݷx\xe4\xe2\xfd\xf7J\xb3=8\xaf\xac" +
			"Y\xf2\xc9h\xcc\x19\x98\xdcJevK\xfe\xf5q\xf5˜3\x1f" +
			"\x84\x91B[\x03K\xfe\f\x9e\xbf\xbf\xb9X\x1c\xac{\xf2%@`" +
			"x\x80\xf1K^\x86P_g\x99\xcfK\xa8\x84\x1f\xd9\x1a\fz\n" +
			"\xeb*\x11p\xe9v\x99\xaf\x1d\b\x19\x83*\x9dM\xc7\xe3\xb7Y%" +
			"\x94\xe1\xe9\x84k\xf7_ΰE\xa1r\xb8\xb3yS\x81\t\xe9\x10" +
			"\aZ\x04\x84\xefKU{~s\xc1\xd8\"&yp\xacP:\x80" +
			"\xfbl%\"/\x84\xf6\x10ݸ\xa1\x16;\xd8@\xf8Z\xc7M\xe1" +
			"\xd1>\xa0\xa1ۓ\xc53\xb2\xf6\x90\xb8\x90\n\xf3Q\x8f\x98\x83b" +
			"\xc9?L\xae?N\xdb}q۟\n\x0e\xbe=\xbb7\xb0\xdcj" +
			"\xeb~\x97K\xfe\xf6\x923\t\x85ht\xf8\xcd)yK\xf6%\x0f" +
			"\xae\x01Μڕ\b\xe0\x1e\x8a\xd0\x01`\xbe\xb4\x87\x15\x96\xddh" +
			"\xe1_\x19)\xfa^\x19\xf0]4\x19\xd7M\xd0h\xdb<W[\xab" +
			"_y\xfe\xb0\a\xcc\xf5\tێ\f\x0f\x1d\u007f\x81\xb3\xbd!\x88\xed" +
			"\x064\xe4\x01do\xb25\x01\xba\x05\xad\xa9Z\xce\xf6Xϒ\x1b" +
			"\xa2BsvPF\xdaÃ\xb3\x01\xa3\xe2\xe4\xb4\x18i,\xb6\xd6" +
			">Q\xf9T8\x8eӋ\xb5\xd5&\x17\x1a\xdb;\x19\x0f\xd7_\xe2" +
			"a\xa7V\"\xe2^<\xdb&\xb60ySc\xa9\xb5\x11&\xd1 " +
			"\xf0o\x0f\t\x1fR1X\xb7ikA\xe3\xdaցE\xff\x1d\x89" +
			"\xebX\xeb\xf9\xa5$\x03\xba{\x1e\x13\x8b\xab8z8#-y\xd8" +
			"\xcfO@|!\xae\x19\x1f\xc4u\x1b\xa3\ti\xeff\x01\x1fi\n" +
			"\xb4\xa8=5\xb7\xedR\xa9\xa4\x84c\xd3*\xf1\x9d\xea\x9c\xce\xf0Q" +
			"\xd1-\xa4+\xf7L\r\x1bS\xabe(\xc9=\x9a]]\xce\xdf\xcd" +
			"\xaf\xe6\xb3\xcb\xe9l\xd2\ri\x97+\xe1\xbd\x13A\xb4\x99\x9d=\x9c" +
			"ɜ7>\xd8*\xa1=1v\x85\x9d\xc7\x18K\xbe\x1c\xbd\xe3\xcc" +
			"\xa6i\xbb\x87=\xe8\x88\xd0\x11➢\x9c\xd641>\x16\x82a" +
			"\xbew\xa2{\u007f3^d\xfb~w\x96\xbf\x0e\xfc\x98\x02\xc7\xe7\x02" +
			"'\xe7\x02\x17\x19\x16\xfa\u007f\x94<=-y\xfa\xefȧ?-9" +
			"\x05N\xcf\x05\xbe\xf9i\xc9\xed\xecu\xa4/j\xa7LX\xd7Q\v" +
			"Y\x89\x97\x1eE\xfc(\x1f\xbb\xa3t\x9cZP\t\xbb\x8b_Z\xa7" +
			"^\xac\tBߢ\xc0\x82\x1b4\x12\xdf\x12A\xe5?:\xd2$\x92" +
			"\xa2~\x16n\xa70\xb7\x8eZ6\x1e]\xb5\xea\xd6>\xe3e\x8cO" +
			"\xf3+\x9c\xf4\xad\r\xc8\xc0q]F\x91\xa2\xf5l2\x99\x8fg\xfd" +
			"\x8f\xb3¢Μu\x1dSc\tMͶZ\xe4O\x1f\x8c\xfc" +
			"V\xaa\x00G\xb2Q\x15nmEo\fO\x1af\xc8fk\x05i" +
			"D\xa5\x13\x03\xe9-\x94\xf3\x81D\xe8KSmcV\x9e\xde\x10\xfd" +
			"\xc5o\xd7\xdfڋ9\xec\xd9]\xad\x96\xfc\r\xe9\x19\x9a0\x9dH" +
			"\xe2X[\x17\x9cP\x81Ti\ak\x17\vE\x015\x8f%\x985" +
			"\xf6\x95\x1c5\xb8\x8dzA\xd48p~\xa0\x97\x8d\x87\xd5)\xa6\x16" +
			"+\xba\x1e\x88spwI\x9f\xfc\x0fd\xf5\x88R\xabR\x97W\xb1" +
			"\xa1L\xaa\xa2@\"M\x88\xe7\xf7\xa1\xbdy-\xe5\xaf\xfb\xe3\x85h" +
			"\xa7\xcfJ\x99^(\xdd4\x9e\x1ap\x9d\x12\f6\x1c\r\x8bl\x88" +
			"\x01\xbf$\xb2\xfeS\xe2\xe6\xe2\x1fPK\x03\x04\x14\x00\x00\b\b\x00" +
			"\xb3\x044C\xccJ\xae2\x0e\x03\x00\x00\x99\x06\x00\x00\x18\x00\x1c\x00" +
			"xl/worksheets/sheet2" +
			".xmlUT\t\x00\x03\xb1|;R\xb1|;Rux\v" +
			"\x00\x01\x04\xe8\x03\x00\x00\x04\xe8\x03\x00\x00\x8dUMO\x1b1\x10\xbd" +
			"\xf3+,\xdf\xcb&\x94@@,\bAS*\x01A\r\x14\xa97" +
			"g=\x9b\xb5\xf0z\xb6\xb6\x97\x00\xbf\xbec\xef'\xd0Cs\x89=" +
			"c\xcf\xc7{o\xbc'g/\xa5f\xcf`\x9dB\x93\xf2\xe9\xee\x84" +
			"30\x19Je6)\u007f\xb8_|\x99s\xe6\xbc0Rh4\x90" +
			"\xf2Wp\xfc\xect\xe7d\x8b\xf6\xc9\x15\x00\x9eQ\x00\xe3R^x" +
			"_\x1d'\x89\xcb\n(\x85\xdb\xc5\n\fyr\xb4\xa5\U00034d5b" +
			"\xc4U\x16\x84\x8c\x97J\x9d\xecM&\aI)\x94\xe1M\x84c\xfb" +
			"?10\xcfU\x06\x97\x98\xd5%\x18\xdf\x04\xb1\xa0\x85\xa7\xf2]\xa1" +
			"*\xc7Ow\x18;\x89I\xee,˕\xf6`oPR\xe5\xb9\xd0" +
			"\x0e\xa2\x9b\x0eTb\x03+\xf0\x0fU<\xe4\xef\xf1\x8e\fݙ$" +
			"\xc6H\xda q#\x15\xe5\v\x181\vy\xcaϧ\xed\xa1x\xe6" +
			"\x97\x82\xadk\x03\xf7\x06\x96\xa1F\xfbC\xa6\xfc`\x9f3\t\xb9\xa8" +
			"\xb5\xffn\x95\xbc\b\xf6\x94{[\x03gVm\n\xca~\r\xb9\xef" +
			"\xb23W\xe0vA=\xd7Z\xb8w\xc6p\xfbZ\x19p\xdd\xed`" +
			"\\\xd6^\x93m\xf5Z\xaeQ\xbf\xf3\xfc\xc4-\xe5\xba\"̉\xde" +
			"\xb1\xe37X\xec\r^\xacW\xa0!\xf3 \xfbd\x1e\xabP\xd1\x05" +
			"h\x1d{e\xcf\xd4P\xcaM Bs\xb6UF\xe2\xf6\u03a2\xa7" +
			"kQ7\xed\xbd \x8a5\xe2S\xe8?tNbzC,W\x99" +
			"\xd0\x04\xeet2\xde\xdf\xc6`\x1f\xad\x81\x86k\xf1\x8auİ\xf1" +
			"6\xc8\x06lc\x9d\x81\x04A\u007f\xcf0\xd47\xec۴\x95\bb" +
			"m\xfb\xa0\xae\xff\xbc\xa3\xadg7$\x19\x91\xdd\x13\xd9и\x88\xc2" +
			"#\x85\xb4\xec\x11\xa0W\x10\b\xa3\xbaf|t\xaf;\x18M\xc4{" +
			"'\x06Z\x06\x19hQ\xb9\x11\xba\x85\x92\x12\x06\xd0J\xf1\x12\xfa\xdc" +
			"\x9b\xd1R\x85\x19\f\x03\xf7\x1a\x00\x9b\x04\xa8\xa5/\x82{wv\xb8" +
			"??\x9a\x1f\xceg\xfb{\xb3N}I\x97\xab\xa9\xf7Rx\xd18" +
			"*\xab\x8c_Vq*XA\n\xa0q\x1e\xb4\xb4\x19t\xf4\xd1B" +
			"3\xd1\t\xa3@\xab\xde\xd0x\xa1/h\xd4\xc0\x8e:\xa0\xf7«" +
			"쳣MN\x1c\xde\b\xbbQ\x94[GaOv\x0f[\xa9\xb7" +
			"k\"&\xae\xe6\x87\xd4\xf5\x1a\xbd\xc7r\xd8\x17Q\xb1a?\x9bN" +
			"\xe7\x93Y\xff\xe3,G\xd2\xdc?]Cjj\xa1\xae\xd8Z\x8b\xec" +
			"\xe9\xdc\xc8\xc7B\xf9~\xb0YF\n\xb9\xc02\xbc\x1d.\xe8\xd9\x04" +
			"\x1bV*@A\xb8K+Fs\x98+\xeb|\x10\xe4m]\xaec" +
			"V\u07bc\x15\xbd\b\xda\xfdcK\xd2\x18\xb3\xcbJ\xa5\xfck\xd06" +
			"\x99(\x9dh\x06\xa5B\xeb\xadP>(t\x03K\x1b\x1b\xa5a2" +
			"\xf7\x05\x98%\xe1\x1a\x1c\x15ؕz\xa3\xaa\x8fH\n\xa3٩\x1d" +
			",>\xd6\xd4\xd6J\xae\xbb\xc09\xd8\xcbF\xab\xee\x13Y}E\r" +
			"T\rʋ\b(\x93*ωH\xe3c\xfc\xfejo^J\xf9" +
			"\xedy\xd0l+o\x94\xb2y]\xba\x99\xfah\xa0}\x93`t`" +
			"0\x9c$\xe3\x1a蛒\xf4\x1f\x95ӝ\xbfPK\x03\x04\x14\x00" +
			"\x00\b\b\x00D\xae\fC\xb6w\xb3\xfa\xf6\x02\x00\x00;\x06\x00\x00" +
			"\x18\x00\x00\x00xl/worksheets/sh" +
			"eet3.xml\x8dTMS\xdb0\x10\xbd\xf7Wht" +
			"o\x9cP\x02\x81\x89\xc30\xa4)\x9d\xa1\x84i\xa0\xcc\xf4\xa6X\xeb" +
			"X\x83\xacu%\x99\x10~}W\xb2\xe3\x18\xe8\xa19\x90\xd5.\xfb" +
			"\xf5\xde\xdbL/^J͞\xc1:\x85&\xe5\xa3\xc1\x9030\x19" +
			"Je6)\u007f\xb8_|\x9ep\xe6\xbc0Rh4\x90\xf2\x1d8" +
			"~1\xfb4ݢ}r\x05\x80gT\xc0\xb8\x94\x17\xdeW\xe7I" +
			"\xe2\xb2\x02J\xe1\x06X\x81\xa1H\x8e\xb6\x14\x9e\x9ev\x93\xb8ʂ" +
			"\x901\xa9\xd4\xc9\xd1px\x92\x94B\x19\xdeT8\xb7\xffS\x03\xf3" +
			"\\e0Ǭ.\xc1\xf8\xa6\x88\x05-<\x8d\xef\nU9>\x9b" +
			"\xc6\x0ew\x96\xe5J{\xb0?P\xd2ع\xd0\x0e(V\x89\r\xac" +
			"\xc0?T1\xee\xef\xf1\x8e\x1c\xfbp2\x9b&m\xf2l*\x15u" +
			"\b\xa80\vy\xca/G!\x1c\xa3\xbf\x14l]\xcff\x19j\xb4" +
			"\xdfe\xcaO\x8e9\x93\x90\x8bZ\xfboVɫ\xe0O\xb9\xb75" +
			"pfզ\xa0~7\x90\xfb}?\xe6\n\xdc.h\xbbZ\v\xf7" +
			"\xc6\x19\xb2o\x94\x01\xb7\xcf\x0e\xcee\xed5\xf9V\xbbr\x8d\xfaM" +
			"\xe4'n\xa9\xd75\xa1KD\xf6\x03\xbf\xc1b\xe7\xf0b\xbd\x02\r" +
			"\x99\a\xd95\xf3X\x85\x89\xae@\xeb\xb8#{\xa6\x85Rn\x02\xe4" +
			"\x9a\xb3\xad2\x12\xb7w\x16=\xa5E\x85\xb4y\x81\xfe5\xe2S\xd8" +
			"?lN\xb2yE,W\x99\xd0\x04\xe7h\xd8\u007f\xdf\xc6b\xef\xbd" +
			"\x01\xf8\x1b\xb1\xc3:b\xd8D\t\xd48`@]\xd0\xd73\x1c\x06" +
			";\xbc\xdb~\x95\bzl\x17\xa0u\xff\xf4xJ:r\xfa\xf6\x9e" +
			"\xb4E\x14\x14)\xa0\xe5\x8aໆ@\x0fM1\xe6\xfb\x8c\xfd\u007f" +
			"ͦį\x8b\u007f\x03\xd3ZT\xae\a`\xa1\xa4\x84\x03.\xa5x" +
			"\t\xab\x1c\x8d\xc9T\xe1\xa0\xc2\xf5\xec\x02&À\xa6\xf4E\b\x0f" +
			"ƧǓ\xb3\xc9\xe9d||4n\x06nz\xc4\xc6s\xe1\x05" +
			"\xb9*\xab\x8c_VQ٬ n\xe9$\x0f*\xd9\x1c\x14\xf2\xde" +
			"C\xfa\xdeS^\xa0U\xafh\xbc\xd0Wt.`{\x83\xd3\xcd{" +
			"\x95}\f$͑\xfc\x10v\xa3\xa8\xb1\x8ez\x1d\x0eN[\x05\xb7" +
			"6\xc1\x1e\xad\xc9)m\xbaF\xef\xb1<\xbc\x8b(\xc4\xf0\x1e\x8fF" +
			"\x93\xe1\xb8\xfbp\x96#I韡\xa4;κbk-\xb2\xa7" +
			"K#\x1f\v\xe5\xbb\xf3d\x19\x91\u007f\x85e\xb8|\x174j\x82\x0f" +
			"+\x15@ \xa0\xa5\x15\xbd\xdbʕu>\x88\xec\xb6.ױ%" +
			"o.\xbe\xa3\xba}?\xb6\xac\xf4њW*\xe5_\x82^\xc9E" +
			"\xedD#\xfe\n\xad\xb7B\xf9 \xbe\r,mܒ\x0e\xc4\xdc\x17" +
			"`\x96\x84h\bT`Wꕦ>#\xee{\xf7P;X\xbc" +
			"\x9f\xa9\x9d\x95Bw\x81m\xb0\xf3F\x91\xee\x03M\xddD\x84S\x83" +
			"\xef\"Bɤ\xcas\xe2\xcf\xf8X\xbc\xcb\xeb\xdcK)\xbf>\x1f" +
			"\x14:\x9b\xa2\x94\xcd/\x05\xa9\xaeg\x93\xd9Tlܝ\xddoF" +
			"\xcf\xeew\u007f\xf6\x17PK\x01\x02\x14\x00\x14\x00\x00\b\b\x00D\xae" +
			"\fC\xf4\xeb\xcaY=\x01\x00\x00g\x05\x00\x00\x13\x00\x00\x00\x00\x00" +
			"\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00[Content" +
			"_Types].xmlPK\x01\x02\x14\x00\x14\x00\x00" +
			"\b\b\x00D\xae\fCf\xaa\x82\xb7\xe0\x00\x00\x00;\x02\x00\x00\v" +
			"\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00n\x01\x00\x00_re" +
			"ls/.relsPK\x01\x02\x14\x00\x14\x00\x00\b\b\x00" +
			"D\xae\fC\x17ϯ\xa7\xbc\x00\x00\x005\x01\x00\x00\x10\x00\x00\x00" +
			"\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00w\x02\x00\x00docPro" +
			"ps/app.xmlPK\x01\x02\x14\x00\x14\x00\x00\b" +
			"\b\x00D\xae\fC\x17qy\xdb:\x01\x00\x00x\x02\x00\x00\x11\x00" +
			"\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00a\x03\x00\x00xl/w" +
			"orksheets.xmlPK\x01\x02\x14\x00\x14" +
			"\x00\x00\b\b\x00D\xae\fC(\xba\xe5Ҧ\x00\x00\x00\xec\x00\x00" +
			"\x00\x14\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\xca\x04\x00\x00x" +
			"l/sharedStrings.xmlP" +
			"K\x01\x02\x1e\x03\x14\x00\x00\b\b\x00\xcf,\rC\x0ep\x99\x04\f" +
			"\x04\x00\x00\x96\x1f\x00\x00\r\x00\x18\x00\x00\x00\x00\x00\x01\x00\x00\x00\xb6" +
			"\x81\xa2\x05\x00\x00xl/styles.xmlUT" +
			"\x05\x00\x035\xaa\tRux\v\x00\x01\x04\xe8\x03\x00\x00\x04\xe8\x03" +
			"\x00\x00PK\x01\x02\x1e\x03\x14\x00\x00\b\b\x00\xe0\x15\xf7D\b\xc4" +
			"0\xf9\xbe\x01\x00\x00}\x03\x00\x00\x0f\x00\x18\x00\x00\x00\x00\x00\x01\x00" +
			"\x00\x00\xb6\x81\xf5\t\x00\x00xl/workbook." +
			"xmlUT\x05\x00\x03\x03\x06\xcfSux\v\x00\x01\x04\xe8\x03" +
			"\x00\x00\x04\xe8\x03\x00\x00PK\x01\x02\x1e\x03\x14\x00\x00\b\b\x00\xb2" +
			"\x044C\xa0J\x80\x9e\x84\x03\x00\x00\x8f\b\x00\x00\x18\x00\x18\x00\x00" +
			"\x00\x00\x00\x01\x00\x00\x00\xb6\x81\xfc\v\x00\x00xl/work" +
			"sheets/sheet1.xmlUT\x05" +
			"\x00\x03\xb0|;Rux\v\x00\x01\x04\xe8\x03\x00\x00\x04\xe8\x03\x00" +
			"\x00PK\x01\x02\x1e\x03\x14\x00\x00\b\b\x00\xb3\x044C\xccJ\xae" +
			"2\x0e\x03\x00\x00\x99\x06\x00\x00\x18\x00\x18\x00\x00\x00\x00\x00\x01\x00\x00" +
			"\x00\xb6\x81\xd2\x0f\x00\x00xl/worksheets" +
			"/sheet2.xmlUT\x05\x00\x03\xb1|;R" +
			"ux\v\x00\x01\x04\xe8\x03\x00\x00\x04\xe8\x03\x00\x00PK\x01\x02D" +
			"\xae\fC\xb6w\xb3\xfa\xf6\x02\x00\x00;\x06\x00\x00\x18\x00\x00\x00\x00" +
			"\x00\x00\x00\x00\x00\x00\x00\x00\x00PK\x05\x06\x00\x00\x00\x00\n\x00\n" +
			"\x00\xe3\x02\x00\x00^\x16\x00\x00\x00\x00",
	}

	for _, f := range crashers {
		_, err := OpenBinary([]byte(f))
		if err == nil {
			t.Fatal("Expected a well formed error from opening this file")
		}
	}
}
