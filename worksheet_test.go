package xlsx

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"testing"
)

// Test we can succesfully unmarshal the sheetN.xml files within and
// XLSX file into an XLSXWorksheet struct (and it's related children).
func TestUnmarshallWorksheet(t *testing.T) {
	var sheetxml = bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:B2"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="C2" sqref="C2"/></sheetView></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="15"/><sheetData><row r="1" spans="1:2"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row><row r="2" spans="1:2"><c r="A2" t="s"><v>2</v></c><c r="B2" t="s"><v>3</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.78740157499999996" bottom="0.78740157499999996" header="0.3" footer="0.3"/></worksheet>`)
	worksheet := new(XLSXWorksheet)
	error := xml.NewDecoder(sheetxml).Decode(worksheet)
	if error != nil {
		t.Error(error.String())
		return
	}
	if worksheet.Dimension.Ref != "A1:B2" {
		t.Error(fmt.Sprintf("Expected worksheet.Dimension.Ref == 'A1:B2', got %s", worksheet.Dimension.Ref))
	}
	if len(worksheet.SheetViews.SheetView) == 0 {
		t.Error(fmt.Sprintf("Expected len(worksheet.SheetViews.SheetView) == 1, got %d", len(worksheet.SheetViews.SheetView)))
	}
	sheetview := worksheet.SheetViews.SheetView[0]
	if sheetview.TabSelected != "1" {
		t.Error(fmt.Sprintf("Expected sheetview.TabSelected == '1', got %s", sheetview.TabSelected))
	}
	if sheetview.WorkbookViewID != "0" {
		t.Error(fmt.Sprintf("Expected sheetview.WorkbookViewID == '0', got %s", sheetview.WorkbookViewID))
	}
	if sheetview.Selection.ActiveCell != "C2" {
		t.Error(fmt.Sprintf("Expeceted sheetview.Selection.ActiveCell == 'C2', got %s", sheetview.Selection.ActiveCell))
	}
	if sheetview.Selection.SQRef != "C2" {
		t.Error(fmt.Sprintf("Expected sheetview.Selection.SQRef == 'C2', got %s", sheetview.Selection.SQRef))
	}
	if worksheet.SheetFormatPr.BaseColWidth != "10" {
		t.Error(fmt.Sprintf("Expected worksheet.SheetFormatPr.BaseColWidth == '10', got %s", worksheet.SheetFormatPr.BaseColWidth))
	}
	if worksheet.SheetFormatPr.DefaultRowHeight != "15" {
		t.Error(fmt.Sprintf("Expected worksheet.SheetFormatPr.DefaultRowHeight == '15', got %s", worksheet.SheetFormatPr.DefaultRowHeight))
	}
	if len(worksheet.SheetData.Row) == 0 {
		t.Error(fmt.Sprintf("Expected len(worksheet.SheetData.Row) == '2', got %d", worksheet.SheetData.Row))
	}
	row := worksheet.SheetData.Row[0]
	if row.R != "1" {
		t.Error(fmt.Sprintf("Expected row.r == '1', got %s", row.R))
	}
	if row.Spans != "1:2" {
		t.Error(fmt.Sprintf("Expected row.Spans == '1:2', got %s", row.Spans))
	}
	if len(row.C) != 2 {
		t.Error(fmt.Sprintf("Expected len(row.C) == 2, got %s", row.C))
	}
	c := row.C[0]
	if c.R != "A1" {
		t.Error(fmt.Sprintf("Expected c.R == 'A1' got %s", c.R))
	}
	if c.T != "s" {
		t.Error(fmt.Sprintf("Expected c.T == 's' got %s", c.T))
	}
	if c.V.Data != "0" {
		t.Error(fmt.Sprintf("Expected c.V.Data == '0', got %s", c.V.Data))
	}

}
