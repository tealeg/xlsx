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
	xlsxFile, error = OpenFile("testfile.xlsx")
	if error != nil {
		t.Error(error.Error())
		return
	}
	if xlsxFile == nil {
		t.Error("OpenFile returned nil FileInterface without generating an os.Error")
		return
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

func TestLettersToNumeric(t *testing.T) {
	var input string
	var output int

	input = "A"
	output = lettersToNumeric(input)
	if output != 0 {
		t.Error("Expected output 'A' == 0, but got ", strconv.Itoa(output))
	}
	input = "z"
	output = lettersToNumeric(input)
	if output != 25 {
		t.Error("Expected output 'z' == 25, but got ", strconv.Itoa(output))
	}
	input = "AA"
	output = lettersToNumeric(input)
	if output != 26 {
		t.Error("Expected output 'AA' == 26, but got ", strconv.Itoa(output))
	}
	input = "Az"
	output = lettersToNumeric(input)
	if output != 51 {
		t.Error("Expected output 'Az' == 51, but got ", strconv.Itoa(output))
	}
	input = "BA"
	output = lettersToNumeric(input)
	if output != 52 {
		t.Error("Expected output 'BA' == 52, but got ", strconv.Itoa(output))
	}
	input = "Bz"
	output = lettersToNumeric(input)
	if output != 77 {
		t.Error("Expected output 'Bz' == 77, but got ", strconv.Itoa(output))
	}
	input = "AAA"
	output = lettersToNumeric(input)
	if output != 676 {
		t.Error("Expected output 'AAA' == 676, but got ", strconv.Itoa(output))
	}

}

func TestPositionalLetterMultiplier(t *testing.T) {
	var output int
	output = positionalLetterMultiplier(1, 0)
	if output != 1 {
		t.Error("Expected positionalLetterMultiplier(1, 0) == 1, got ", output)
	}
	output = positionalLetterMultiplier(2, 0)
	if output != 26 {
		t.Error("Expected positionalLetterMultiplier(2, 0) == 26, got ", output)
	}
	output = positionalLetterMultiplier(2, 1)
	if output != 1 {
		t.Error("Expected positionalLetterMultiplier(2, 1) == 1, got ", output)
	}
	output = positionalLetterMultiplier(3, 0)
	if output != 676 {
		t.Error("Expected positionalLetterMultiplier(3, 0) == 676, got ", output)
	}
	output = positionalLetterMultiplier(3, 1)
	if output != 26 {
		t.Error("Expected positionalLetterMultiplier(3, 1) == 26, got ", output)
	}
	output = positionalLetterMultiplier(3, 2)
	if output != 1 {
		t.Error("Expected positionalLetterMultiplier(3, 2) == 1, got ", output)
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
	reftable := MakeSharedStringRefTable(sst)
	rows, maxCols, maxRows := readRowsFromSheet(worksheet, reftable)
	if maxRows != 2 {
		t.Error("Expected maxRows == 2")
	}
	if maxCols != 22 {
		t.Error("Expected maxCols == 22")
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
	reftable := MakeSharedStringRefTable(sst)
	rows, _, _ := readRowsFromSheet(worksheet, reftable)
	if len(rows) != 3 {
		t.Error("Expected len(rows) == 3, got ", strconv.Itoa(len(rows)))
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
	reftable := MakeSharedStringRefTable(sst)
	rows, maxCol, maxRow  := readRowsFromSheet(worksheet, reftable)
	if len(rows) != 8 {
		t.Error("Expected len(rows) == 8, got ", strconv.Itoa(len(rows)))
	}
	if maxCol != 22 {
		t.Error("Expected maxCol == 22, got ", strconv.Itoa(maxCol))

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
