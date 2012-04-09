package xlsx

import (
	"bytes"
	"encoding/xml"
	"testing"
)

// Test we can succesfully unmarshal the workbook.xml file from within
// an XLSX file and return a XLSXWorkbook struct (and associated
// children).
func TestUnmarshallWorkbookXML(t *testing.T) {
	var error error
	var buf = bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4506"/><workbookPr defaultThemeVersion="124226"/><bookViews><workbookView xWindow="120" yWindow="75" windowWidth="15135" windowHeight="7620"/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId2"/><sheet name="Sheet3" sheetId="3" r:id="rId3"/></sheets><definedNames><definedName name="monitors" localSheetId="0">Sheet1!$A$1533</definedName></definedNames><calcPr calcId="125725"/></workbook>`)
	var workbook *XLSXWorkbook
	workbook = new(XLSXWorkbook)
	error = xml.NewDecoder(buf).Decode(workbook)
	if error != nil {
		t.Error(error.Error())
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
