package xlsx

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"testing"
)

// Test we can succesfully unmarshal the sheetN.xml files within and
// XLSX file into an xlsxWorksheet struct (and it's related children).
func TestUnmarshallWorksheet(t *testing.T) {
	var sheetxml = bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetPr filterMode="false">
    <pageSetUpPr fitToPage="false"/>
  </sheetPr>
  <dimension ref="A1:B2"/>
  <sheetViews>
    <sheetView colorId="64" defaultGridColor="true" rightToLeft="false" showFormulas="false" showGridLines="true" showOutlineSymbols="true" showRowColHeaders="true" showZeros="true" tabSelected="true" topLeftCell="A1" view="normal" windowProtection="false" workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">
      <selection activeCell="B2" activeCellId="0" pane="topLeft" sqref="B2"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15">
  </sheetFormatPr>
  <cols>
    <col collapsed="false" hidden="false" max="1025" min="1" style="0" width="10.5748987854251"/>
  </cols>
  <sheetData>
    <row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="14.9" outlineLevel="0" r="1">
      <c r="A1" s="1" t="s">
        <v>0</v>
      </c>
      <c r="B1" s="0" t="s">
        <v>1</v>
      </c>
    </row>
    <row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="14.9" outlineLevel="0" r="2">
      <c r="A2" s="0" t="s">
        <v>2</v>
      </c>
      <c r="B2" s="2" t="s">
        <v>3</v>
      </c>
    </row>
  </sheetData>
  <printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>
  <pageMargins left="0.7" right="0.7" top="0.7875" bottom="0.7875" header="0.511805555555555" footer="0.511805555555555"/>
  <pageSetup blackAndWhite="false" cellComments="none" copies="1" draft="false" firstPageNumber="0" fitToHeight="1" fitToWidth="1" horizontalDpi="300" orientation="portrait" pageOrder="downThenOver" paperSize="9" scale="100" useFirstPageNumber="false" usePrinterDefaults="false" verticalDpi="300"/>
  <headerFooter differentFirst="false" differentOddEven="false">
    <oddHeader>
    </oddHeader>
    <oddFooter>
    </oddFooter>
  </headerFooter>
</worksheet>`)
	worksheet := new(xlsxWorksheet)
	error := xml.NewDecoder(sheetxml).Decode(worksheet)
	if error != nil {
		t.Error(error.Error())
		return
	}
	if worksheet.Dimension.Ref != "A1:B2" {
		t.Error(fmt.Sprintf("Expected worksheet.Dimension.Ref == 'A1:B2', got %s", worksheet.Dimension.Ref))
	}
	if len(worksheet.SheetData.Row) == 0 {
		t.Error(fmt.Sprintf("Expected len(worksheet.SheetData.Row) == '2', got %d", worksheet.SheetData.Row))
	}
	row := worksheet.SheetData.Row[0]
	if row.R != 1 {
		t.Error(fmt.Sprintf("Expected row.r == 1, got %d", row.R))
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
	if c.V != "0" {
		t.Error(fmt.Sprintf("Expected c.V.Data == '0', got %s", c.V))
	}

}
