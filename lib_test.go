package xlsx


import (
	"bytes"
	"os"
	"strconv"
	"strings"
	"testing"
	"xml"
)


// Test we can correctly open a XSLX file and return a xlsx.File
// struct.
func TestOpenFile(t *testing.T) {
	var xlsxFile *File
	var error os.Error
	xlsxFile, error = OpenFile("testfile.xlsx")
	if error != nil {
		t.Error(error.String())
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
	var error os.Error
	var sheet *Sheet
	var row *Row
	xlsxFile, error = OpenFile("testfile.xlsx")
	if error != nil {
		t.Error(error.String())
		return
	}
	if xlsxFile == nil {
		t.Error("OpenFile returned a nil File pointer but did not generate an error.")
		return
	}
	if len(xlsxFile.Sheets) == 0 {
		t.Error("Expected len(xlsxFile.Sheets) > 0")
		return
	}
	sheet = xlsxFile.Sheets[0]
	if len(sheet.Rows) != 2 {
		t.Error("Expected len(sheet.Rows) == 2")
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
	var error os.Error
	xlsxFile, error = OpenFile("testfile.xlsx")
	if error != nil {
		t.Error(error.String())
		return
	}
	if xlsxFile.referenceTable == nil {
		t.Error("expected non nil xlsxFile.referenceTable")
		return
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
	worksheet := new(XLSXWorksheet)
	error := xml.Unmarshal(sheetxml, worksheet)
	if error != nil {
		t.Error(error.String())
		return
	}
	sst := new(XLSXSST)
	error = xml.Unmarshal(sharedstringsXML, sst)
	if error != nil {
		t.Error(error.String())
		return
	}
	reftable := MakeSharedStringRefTable(sst)
	rows := readRowsFromSheet(worksheet, reftable)
	if len(rows) != 2 {
		t.Error("Expected len(rows) == 2")
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
	var error os.Error
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
	var error os.Error
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


// func TestReadRowsFromSheetWithEmptyCells(t *testing.T) {
// 	var sharedstringsXML = bytes.NewBufferString(`
// <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
// <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="8" uniqueCount="5"><si><t>Bob</t></si><si><t>Alice</t></si><si><t>Sue</t></si><si><t>Yes</t></si><si><t>No</t></si></sst>`)
// 	var sheetxml = bytes.NewBufferString(`
// <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
// <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:C3"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="D3" sqref="D3"/></sheetView></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
// <sheetData>
//   <row r="1" spans="1:3">
//     <c r="A1" t="s">
//       <v>
//         0
//       </v>
//     </c>
//     <c r="B1" t="s">
//       <v>
//         1
//       </v>
//     </c>
//     <c r="C1" t="s">
//       <v>
//         2
//       </v>
//     </c>
//   </row>
//   <row r="2" spans="1:3">
//     <c r="A2" t="s">
//       <v>
//         3
//       </v>
//     </c>
//     <c r="B2" t="s">
//       <v>
//         4
//       </v>
//     </c>
//     <c r="C2" t="s">
//       <v>
//         3
//       </v>
//     </c>
//   </row>
//   <row r="3" spans="1:3">
//     <c r="A3" t="s">
//       <v>
//         4
//       </v>
//     </c>
//     <c r="C3" t="s">
//       <v>
//         3
//       </v>
//     </c>
//   </row>
// </sheetData>
// <pageMargins left="0.7" right="0.7" top="0.78740157499999996" bottom="0.78740157499999996" header="0.3" footer="0.3"/>
// </worksheet>

// `)
// 	worksheet := new(XLSXWorksheet)
// 	error := xml.Unmarshal(sheetxml, worksheet)
// 	if error != nil {
// 		t.Error(error.String())
// 		return
// 	}
// 	sst := new(XLSXSST)
// 	error = xml.Unmarshal(sharedstringsXML, sst)
// 	if error != nil {
// 		t.Error(error.String())
// 		return
// 	}
// 	reftable := MakeSharedStringRefTable(sst)
// 	rows := readRowsFromSheet(worksheet, reftable)
// 	if len(rows) != 3 {
// 		t.Error("Expected len(rows) == 3")
// 	}
	
// }
