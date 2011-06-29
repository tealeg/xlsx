package xlsx


import (
	"bytes"
	"os"
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
