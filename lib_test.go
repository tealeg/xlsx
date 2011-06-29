package xlsx

import (
	"bytes"
	// "archive/zip"
	"fmt"
	"os"
	"testing"
	"xml"
)

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

func TestExtractSheets(t *testing.T) {
	var xlsxFile *File
	var sheets []*Sheet
	xlsxFile, _ = OpenFile("testfile.xlsx")
	sheets = xlsxFile.Sheets
	if len(sheets) == 0 {
		t.Error("No sheets read from XLSX file")
		return
	}
	fmt.Printf("%v\n", len(sheets))
}


func TestMakeSharedStringRefTable(t *testing.T) {
	var sharedstringsXML = bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4"><si><t>Foo</t></si><si><t>Bar</t></si><si><t xml:space="preserve">Baz </t></si><si><t>Quuk</t></si></sst>`)
	sst := new(XLSXSST)
	error := xml.Unmarshal(sharedstringsXML, sst)
	if error != nil {
		t.Error(error.String())
		return
	}
	reftable := MakeSharedStringRefTable(sst)
	if len(reftable) == 0 {
		t.Error("Reftable is zero length.")
		return
	}
	if reftable[0] != "Foo" {
		t.Error("RefTable lookup failed, expected reftable[0] == 'Foo'")
	}
	if reftable[1] != "Bar" {
		t.Error("RefTable lookup failed, expected reftable[1] == 'Bar'")
	}

}


func TestResolveSharedString(t *testing.T) {
	var sharedstringsXML = bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4"><si><t>Foo</t></si><si><t>Bar</t></si><si><t xml:space="preserve">Baz </t></si><si><t>Quuk</t></si></sst>`)
	sst := new(XLSXSST)
	error := xml.Unmarshal(sharedstringsXML, sst)
	if error != nil {
		t.Error(error.String())
		return
	}
	reftable := MakeSharedStringRefTable(sst)
	if ResolveSharedString(reftable, 0) != "Foo" {
		t.Error("Expected ResolveSharedString(reftable, 0) == 'Foo'")
	}
}


func TestUnmarshallSharedStrings(t *testing.T) {
	var sharedstringsXML = bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4"><si><t>Foo</t></si><si><t>Bar</t></si><si><t xml:space="preserve">Baz </t></si><si><t>Quuk</t></si></sst>`)
	sst := new(XLSXSST)
	error := xml.Unmarshal(sharedstringsXML, sst)
	if error != nil {
		t.Error(error.String())
		return
	}
	if sst.Count != "4" {
		t.Error(`sst.Count != "4"`)
	}
	if sst.UniqueCount != "4" {
		t.Error(`sst.UniqueCount != 4`)
	}
	if len(sst.SI) == 0 {
		t.Error("Expected 4 sst.SI but found none")
	}
	si := sst.SI[0]
	if si.T.Data != "Foo" {
		t.Error("Expected s.T.Data == 'Foo'")
	}

}


func TestUnmarshallSheet(t *testing.T) {
	var sheetxml = bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:B2"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="C2" sqref="C2"/></sheetView></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="15"/><sheetData><row r="1" spans="1:2"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row><row r="2" spans="1:2"><c r="A2" t="s"><v>2</v></c><c r="B2" t="s"><v>3</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.78740157499999996" bottom="0.78740157499999996" header="0.3" footer="0.3"/></worksheet>`)
	worksheet := new(XLSXWorksheet)
	error := xml.Unmarshal(sheetxml, worksheet)
	if error != nil {
		t.Error(error.String())
		return
	}
	if worksheet.Dimension.Ref != "A1:B2" {
		t.Error("Expected worksheet.Dimension.Ref == 'A1:B2'")
	}
	if len(worksheet.SheetViews.SheetView) == 0 {
		t.Error("Expected len(worksheet.SheetViews.SheetView) == 1")
	}
	sheetview := worksheet.SheetViews.SheetView[0]
	if sheetview.TabSelected != "1" {
		t.Error("Expected sheetview.TabSelected == '1'")
	}
	if sheetview.WorkbookViewID != "0" {
		t.Error("Expected sheetview.WorkbookViewID == '0'")
	}
	if sheetview.Selection.ActiveCell != "C2" {
		t.Error("Expeceted sheetview.Selection.ActiveCell == 'C2'")
	}
	if sheetview.Selection.SQRef != "C2" {
		t.Error("Expected sheetview.Selection.SQRef == 'C2'")
	}
	if worksheet.SheetFormatPr.BaseColWidth != "10" {
		t.Error("Expected worksheet.SheetFormatPr.BaseColWidth == '10'")
	}
	if worksheet.SheetFormatPr.DefaultRowHeight != "15" {
		t.Error("Expected worksheet.SheetFormatPr.DefaultRowHeight == '15'")
	}
	if len(worksheet.SheetData.Row) == 0 {
		t.Error("Expected len(worksheet.SheetData.Row) == '2'")
	}
	row := worksheet.SheetData.Row[0]
	if row.R != "1" {
		t.Error("Expected row.r == '1'")
	}
	if row.Spans != "1:2" {
		t.Error("Expected row.Spans == '1:2'")
	}
	if len(row.C) != 2 {
		t.Error("Expected len(row.C) == 2")
	}
	c := row.C[0]
	if c.R != "A1" {
		t.Error("Expected c.R == 'A1'")
	}
	if c.T != "s" {
		t.Error("Expected c.T == 's'")
	}
	if c.V.Data != "0" {
		t.Error("Expected c.V.Data == '0'")
	}

}

func TestCreateXSLXSheetStruct(t *testing.T) {
	var xlsxFile *File
	var error os.Error
	var sheet *Sheet
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
	if len(sheet.Cells) == 0 {
		t.Error("Expected len(sheet.Cells) == 4")
	}
}


func TestUnmarshallXML(t *testing.T) {
	var error os.Error
	var buf = bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4506"/><workbookPr defaultThemeVersion="124226"/><bookViews><workbookView xWindow="120" yWindow="75" windowWidth="15135" windowHeight="7620"/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId2"/><sheet name="Sheet3" sheetId="3" r:id="rId3"/></sheets><definedNames><definedName name="monitors" localSheetId="0">Sheet1!$A$1533</definedName></definedNames><calcPr calcId="125725"/></workbook>`)
	var workbook *XLSXWorkbook
	workbook = new(XLSXWorkbook)
	error = xml.Unmarshal(buf, workbook)
	if error != nil {
		t.Error(error.String())
		return
	}
	if workbook.FileVersion.AppName != "xl" {
		t.Error("Expected FileVersion.AppName == 'xl')")
	}
	if workbook.FileVersion.LastEdited != "4" {
		t.Error("Expected FileVersion.LastEdited == '4'")
	}
	if workbook.FileVersion.LowestEdited != "4" {
		t.Error("Expected FileVersion.LowestEdited == '4'")
	}
	if workbook.FileVersion.RupBuild != "4506" {
		t.Error("Expected FileVersion.RupBuild == '4506'")
	}
	if workbook.WorkbookPr.DefaultThemeVersion != "124226" {
		t.Error("Expected workbook.WorkbookPr.DefaultThemeVersion == '124226'")
	}
	if len(workbook.BookViews.WorkBookView) == 0 {
		t.Error("Expected len(workbook.BookViews.WorkBookView) == 0")
	}
	workBookView := workbook.BookViews.WorkBookView[0]
	if workBookView.XWindow != "120" {
		t.Error("Expected workBookView.XWindow == '120'")
	}
	if workBookView.YWindow != "75" {
		t.Error("Expected workBookView.YWindow == '75'")
	}
	if workBookView.WindowWidth != "15135" {
		t.Error("Expected workBookView.WindowWidth == '15135'")
	}
	if workBookView.WindowHeight != "7620" {
		t.Error("Expected workBookView.WindowHeight == '7620'")
	}
	if len(workbook.Sheets.Sheet) == 0 {
		t.Error("Expected len(workbook.Sheets.Sheet) == 0")
	}
	sheet := workbook.Sheets.Sheet[0]
	if sheet.Id != "rId1" {
		t.Error("Expected sheet.Id == 'rID1'")
	}
	if sheet.Name != "Sheet1" {
		t.Error("Expected sheet.Name == 'Sheet1'")
	}
	if sheet.SheetId != "1" {
		t.Error("Expected sheet.SheetId == '1'")
	}
	if len(workbook.DefinedNames.DefinedName) == 0 {
		t.Error("Expected len(workbook.DefinedNames.DefinedName) == 0")
	}
	dname := workbook.DefinedNames.DefinedName[0]
	if dname.Data != "Sheet1!$A$1533" {
		t.Error("dname.Data == 'Sheet1!$A$1533'")
	}
	if dname.LocalSheetID != "0" {
		t.Error("dname.LocalSheetID == '0'")
	}
	if dname.Name != "monitors" {
		t.Error("Expected dname.Name == 'monitors'")
	}
	if workbook.CalcPr.CalcId != "125725" {
		t.Error("workbook.CalcPr.CalcId != '125725'")
	}
}
