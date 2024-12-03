package xlsx

import (
	"bytes"
	"encoding/xml"
	"testing"

	qt "github.com/frankban/quicktest"
)

// Test we can succesfully unmarshal the workbook.xml file from within
// an XLSX file and return a xlsxWorkbook struct (and associated
// children).
func TestUnmarshallWorkbookXML(t *testing.T) {
	c := qt.New(t)
	var buf = bytes.NewBufferString(
		`<?xml version="1.0"
        encoding="UTF-8"
        standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <fileVersion appName="xl"
                       lastEdited="4"
                       lowestEdited="4"
                       rupBuild="4506"/>
          <workbookPr defaultThemeVersion="124226" date1904="true"/>
          <bookViews>
            <workbookView xWindow="120"
                          yWindow="75"
                          windowWidth="15135"
                          windowHeight="7620"/>
          </bookViews>
          <sheets>
            <sheet name="Sheet1"
                   sheetId="1"
                   r:id="rId1"
                   state="visible"/>
            <sheet name="Sheet2"
                   sheetId="2"
                   r:id="rId2"
                   state="hidden"/>
            <sheet name="Sheet3"
                   sheetId="3"
                   r:id="rId3"
                   state="veryHidden"/>
          </sheets>
          <definedNames>
            <definedName name="monitors" comment="this is the comment"
                         description="give cells a name"
                         localSheetId="0">Sheet1!$A$1533</definedName>
            <definedName name="global" comment="this is the comment"
                         description="a global defined name">Sheet1!$A$1533</definedName>
          </definedNames>
          <calcPr calcId="125725"/>
          </workbook>`)
	workbook := new(xlsxWorkbook)
	err := xml.NewDecoder(buf).Decode(workbook)
	c.Assert(err, qt.IsNil)
	c.Assert(workbook.FileVersion.AppName, qt.Equals, "xl")
	c.Assert(workbook.FileVersion.LastEdited, qt.Equals, "4")
	c.Assert(workbook.FileVersion.LowestEdited, qt.Equals, "4")
	c.Assert(workbook.FileVersion.RupBuild, qt.Equals, "4506")
	c.Assert(workbook.WorkbookPr.DefaultThemeVersion, qt.Equals, "124226")
	c.Assert(workbook.WorkbookPr.Date1904, qt.Equals, true)
	c.Assert(workbook.BookViews.WorkBookView, qt.HasLen, 1)
	workBookView := workbook.BookViews.WorkBookView[0]
	c.Assert(workBookView.XWindow, qt.Equals, "120")
	c.Assert(workBookView.YWindow, qt.Equals, "75")
	c.Assert(workBookView.WindowWidth, qt.Equals, 15135)
	c.Assert(workBookView.WindowHeight, qt.Equals, 7620)
	c.Assert(workbook.Sheets.Sheet, qt.HasLen, 3)
	sheet := workbook.Sheets.Sheet[0]
	c.Assert(sheet.Id, qt.Equals, "rId1")
	c.Assert(sheet.Name, qt.Equals, "Sheet1")
	c.Assert(sheet.SheetId, qt.Equals, "1")
	c.Assert(sheet.State, qt.Equals, "visible")
	c.Assert(workbook.DefinedNames.DefinedName, qt.HasLen, 2)
	dname := workbook.DefinedNames.DefinedName[0]
	c.Assert(dname.Data, qt.Equals, "Sheet1!$A$1533")
	c.Assert(*dname.LocalSheetID, qt.Equals, 0)
	c.Assert(dname.Name, qt.Equals, "monitors")
	c.Assert(dname.Comment, qt.Equals, "this is the comment")
	c.Assert(dname.Description, qt.Equals, "give cells a name")
	c.Assert(workbook.CalcPr.CalcId, qt.Equals, "125725")
	dname2 := workbook.DefinedNames.DefinedName[1]
	c.Assert(dname2.Data, qt.Equals, "Sheet1!$A$1533")
	c.Assert(dname2.LocalSheetID, qt.Equals, (*int)(nil))
	c.Assert(dname2.Name, qt.Equals, "global")
	c.Assert(dname2.Comment, qt.Equals, "this is the comment")
	c.Assert(dname2.Description, qt.Equals, "a global defined name")
	c.Assert(workbook.CalcPr.CalcId, qt.Equals, "125725")
}

// Test we can marshall a Workbook to xml
func TestMarshallWorkbook(t *testing.T) {
	c := qt.New(t)
	workbook := new(xlsxWorkbook)
	workbook.FileVersion = xlsxFileVersion{}
	workbook.FileVersion.AppName = "xlsx"
	workbook.WorkbookPr = xlsxWorkbookPr{BackupFile: false}
	workbook.BookViews = xlsxBookViews{}
	workbook.BookViews.WorkBookView = make([]xlsxWorkBookView, 1)
	workbook.BookViews.WorkBookView[0] = xlsxWorkBookView{}
	workbook.Sheets = xlsxSheets{}
	workbook.Sheets.Sheet = make([]xlsxSheet, 1)
	workbook.Sheets.Sheet[0] = xlsxSheet{Name: "sheet1", SheetId: "1", Id: "rId2"}

	body, err := xml.Marshal(workbook)
	c.Assert(err, qt.IsNil)
	expectedWorkbook := `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fileVersion appName="xlsx"></fileVersion><workbookPr date1904="false"></workbookPr><workbookProtection></workbookProtection><bookViews><workbookView></workbookView></bookViews><sheets><sheet name="sheet1" sheetId="1" xmlns:relationships="http://schemas.openxmlformats.org/officeDocument/2006/relationships" relationships:id="rId2"></sheet></sheets><definedNames></definedNames><calcPr></calcPr></workbook>`
	c.Assert(string(body), qt.Equals, expectedWorkbook)
}
