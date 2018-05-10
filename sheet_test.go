package xlsx

import (
	"bytes"
	"encoding/xml"

	. "gopkg.in/check.v1"
)

type SheetSuite struct{}

var _ = Suite(&SheetSuite{})

// Test we can add a Row to a Sheet
func (s *SheetSuite) TestAddRow(c *C) {
	var f *File
	f = NewFile()
	sheet, _ := f.AddSheet("MySheet")
	row := sheet.AddRow()
	c.Assert(row, NotNil)
	c.Assert(len(sheet.Rows), Equals, 1)
}

// Test we can get row by index from  Sheet
func (s *SheetSuite) TestGetRowByIndex(c *C) {
	var f *File
	f = NewFile()
	sheet, _ := f.AddSheet("MySheet")
	row := sheet.Row(10)
	c.Assert(row, NotNil)
	c.Assert(len(sheet.Rows), Equals, 10)
	row = sheet.Row(2)
	c.Assert(row, NotNil)
	c.Assert(len(sheet.Rows), Equals, 10)
}

func (s *SheetSuite) TestMakeXLSXSheetFromRows(c *C) {
	file := NewFile()
	sheet, _ := file.AddSheet("Sheet1")
	row := sheet.AddRow()
	cell := row.AddCell()
	cell.Value = "A cell!"
	refTable := NewSharedStringRefTable()
	styles := newXlsxStyleSheet(nil)
	xSheet := sheet.makeXLSXSheet(refTable, styles)
	c.Assert(xSheet.Dimension.Ref, Equals, "A1")
	c.Assert(xSheet.SheetData.Row, HasLen, 1)
	xRow := xSheet.SheetData.Row[0]
	c.Assert(xRow.R, Equals, 1)
	c.Assert(xRow.Spans, Equals, "")
	c.Assert(xRow.C, HasLen, 1)
	xC := xRow.C[0]
	c.Assert(xC.R, Equals, "A1")
	c.Assert(xC.S, Equals, 0)
	c.Assert(xC.T, Equals, "s") // Shared string type
	c.Assert(xC.V, Equals, "0") // reference to shared string
	xSST := refTable.makeXLSXSST()
	c.Assert(xSST.Count, Equals, 1)
	c.Assert(xSST.UniqueCount, Equals, 1)
	c.Assert(xSST.SI, HasLen, 1)
	xSI := xSST.SI[0]
	c.Assert(xSI.T, Equals, "A cell!")
}

// Test if the NumFmts assigned properly according the FormatCode in cell.
func (s *SheetSuite) TestMakeXLSXSheetWithNumFormats(c *C) {
	file := NewFile()
	sheet, _ := file.AddSheet("Sheet1")
	row := sheet.AddRow()

	cell1 := row.AddCell()
	cell1.Value = "A cell!"
	cell1.NumFmt = "general"

	cell2 := row.AddCell()
	cell2.Value = "37947.7500001"
	cell2.NumFmt = "0"

	cell3 := row.AddCell()
	cell3.Value = "37947.7500001"
	cell3.NumFmt = "mm-dd-yy"

	cell4 := row.AddCell()
	cell4.Value = "37947.7500001"
	cell4.NumFmt = "hh:mm:ss"

	refTable := NewSharedStringRefTable()
	styles := newXlsxStyleSheet(nil)
	worksheet := sheet.makeXLSXSheet(refTable, styles)

	c.Assert(styles.CellStyleXfs, IsNil)

	c.Assert(styles.CellXfs.Count, Equals, 4)
	c.Assert(styles.CellXfs.Xf[0].NumFmtId, Equals, 0)
	c.Assert(styles.CellXfs.Xf[1].NumFmtId, Equals, 1)
	c.Assert(styles.CellXfs.Xf[2].NumFmtId, Equals, 14)
	c.Assert(styles.CellXfs.Xf[3].NumFmtId, Equals, 164)
	c.Assert(styles.NumFmts.Count, Equals, 1)
	c.Assert(styles.NumFmts.NumFmt[0].NumFmtId, Equals, 164)
	c.Assert(styles.NumFmts.NumFmt[0].FormatCode, Equals, "hh:mm:ss")

	// Finally we check that the cell points to the right CellXf /
	// CellStyleXf.
	c.Assert(worksheet.SheetData.Row[0].C[0].S, Equals, 0)
	c.Assert(worksheet.SheetData.Row[0].C[1].S, Equals, 1)
}

// When we create the xlsxSheet we also populate the xlsxStyles struct
// with style information.
func (s *SheetSuite) TestMakeXLSXSheetAlsoPopulatesXLSXSTyles(c *C) {
	file := NewFile()
	sheet, _ := file.AddSheet("Sheet1")
	row := sheet.AddRow()

	cell1 := row.AddCell()
	cell1.Value = "A cell!"
	style1 := NewStyle()
	style1.Font = *NewFont(10, "Verdana")
	style1.Fill = *NewFill("solid", "FFFFFFFF", "00000000")
	style1.Border = *NewBorder("none", "thin", "none", "thin")
	cell1.SetStyle(style1)

	// We need a second style to check that Xfs are populated correctly.
	cell2 := row.AddCell()
	cell2.Value = "Another cell!"
	style2 := NewStyle()
	style2.Font = *NewFont(10, "Verdana")
	style2.Fill = *NewFill("solid", "FFFFFFFF", "00000000")
	style2.Border = *NewBorder("none", "thin", "none", "thin")
	cell2.SetStyle(style2)

	refTable := NewSharedStringRefTable()
	styles := newXlsxStyleSheet(nil)
	worksheet := sheet.makeXLSXSheet(refTable, styles)

	c.Assert(styles.Fonts.Count, Equals, 2)
	c.Assert(styles.Fonts.Font[0].Sz.Val, Equals, "12")
	c.Assert(styles.Fonts.Font[0].Name.Val, Equals, "Verdana")
	c.Assert(styles.Fonts.Font[1].Sz.Val, Equals, "10")
	c.Assert(styles.Fonts.Font[1].Name.Val, Equals, "Verdana")

	c.Assert(styles.Fills.Count, Equals, 3)
	c.Assert(styles.Fills.Fill[0].PatternFill.PatternType, Equals, "none")
	c.Assert(styles.Fills.Fill[0].PatternFill.FgColor.RGB, Equals, "FFFFFFFF")
	c.Assert(styles.Fills.Fill[0].PatternFill.BgColor.RGB, Equals, "00000000")

	c.Assert(styles.Borders.Count, Equals, 2)
	c.Assert(styles.Borders.Border[1].Left.Style, Equals, "none")
	c.Assert(styles.Borders.Border[1].Right.Style, Equals, "thin")
	c.Assert(styles.Borders.Border[1].Top.Style, Equals, "none")
	c.Assert(styles.Borders.Border[1].Bottom.Style, Equals, "thin")

	c.Assert(styles.CellStyleXfs, IsNil)

	c.Assert(styles.CellXfs.Count, Equals, 2)
	c.Assert(styles.CellXfs.Xf[0].FontId, Equals, 0)
	c.Assert(styles.CellXfs.Xf[0].FillId, Equals, 0)
	c.Assert(styles.CellXfs.Xf[0].BorderId, Equals, 0)

	// Finally we check that the cell points to the right CellXf /
	// CellStyleXf.
	c.Assert(worksheet.SheetData.Row[0].C[0].S, Equals, 1)
	c.Assert(worksheet.SheetData.Row[0].C[1].S, Equals, 1)
}

// If the column width is not customised, the xslxCol.CustomWidth field is set to 0.
func (s *SheetSuite) TestMakeXLSXSheetDefaultsCustomColWidth(c *C) {
	file := NewFile()
	sheet, _ := file.AddSheet("Sheet1")
	row := sheet.AddRow()
	cell1 := row.AddCell()
	cell1.Value = "A cell!"

	refTable := NewSharedStringRefTable()
	styles := newXlsxStyleSheet(nil)
	worksheet := sheet.makeXLSXSheet(refTable, styles)
	c.Assert(worksheet.Cols.Col[0].CustomWidth, Equals, false)
}

// If the column width is customised, the xslxCol.CustomWidth field is set to 1.
func (s *SheetSuite) TestMakeXLSXSheetSetsCustomColWidth(c *C) {
	file := NewFile()
	sheet, _ := file.AddSheet("Sheet1")
	row := sheet.AddRow()
	cell1 := row.AddCell()
	cell1.Value = "A cell!"
	err := sheet.SetColWidth(0, 1, 10.5)
	c.Assert(err, IsNil)

	refTable := NewSharedStringRefTable()
	styles := newXlsxStyleSheet(nil)
	worksheet := sheet.makeXLSXSheet(refTable, styles)
	c.Assert(worksheet.Cols.Col[1].CustomWidth, Equals, true)
}

func (s *SheetSuite) TestMarshalSheet(c *C) {
	file := NewFile()
	sheet, _ := file.AddSheet("Sheet1")
	row := sheet.AddRow()
	cell := row.AddCell()
	cell.Value = "A cell!"
	refTable := NewSharedStringRefTable()
	styles := newXlsxStyleSheet(nil)
	xSheet := sheet.makeXLSXSheet(refTable, styles)

	output := bytes.NewBufferString(xml.Header)
	body, err := xml.Marshal(xSheet)
	c.Assert(err, IsNil)
	c.Assert(body, NotNil)
	_, err = output.Write(body)
	c.Assert(err, IsNil)

	expectedXLSXSheet := `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr filterMode="false"><pageSetUpPr fitToPage="false"></pageSetUpPr></sheetPr><dimension ref="A1"></dimension><sheetViews><sheetView windowProtection="false" showFormulas="false" showGridLines="true" showRowColHeaders="true" showZeros="true" rightToLeft="false" tabSelected="true" showOutlineSymbols="true" defaultGridColor="true" view="normal" topLeftCell="A1" colorId="64" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100" workbookViewId="0"><selection pane="topLeft" activeCell="A1" activeCellId="0" sqref="A1"></selection></sheetView></sheetViews><sheetFormatPr defaultRowHeight="12.85"></sheetFormatPr><cols><col collapsed="false" hidden="false" max="1" min="1" style="0" width="9.5"></col></cols><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData><printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"></printOptions><pageMargins left="0.7875" right="0.7875" top="1.05277777777778" bottom="1.05277777777778" header="0.7875" footer="0.7875"></pageMargins><pageSetup paperSize="9" scale="100" firstPageNumber="1" fitToWidth="1" fitToHeight="1" pageOrder="downThenOver" orientation="portrait" usePrinterDefaults="false" blackAndWhite="false" draft="false" cellComments="none" useFirstPageNumber="true" horizontalDpi="300" verticalDpi="300" copies="1"></pageSetup><headerFooter differentFirst="false" differentOddEven="false"><oddHeader>&amp;C&amp;&#34;Times New Roman,Regular&#34;&amp;12&amp;A</oddHeader><oddFooter>&amp;C&amp;&#34;Times New Roman,Regular&#34;&amp;12Page &amp;P</oddFooter></headerFooter></worksheet>`

	c.Assert(output.String(), Equals, expectedXLSXSheet)
}

func (s *SheetSuite) TestMarshalSheetWithMultipleCells(c *C) {
	file := NewFile()
	sheet, _ := file.AddSheet("Sheet1")
	row := sheet.AddRow()
	cell := row.AddCell()
	cell.Value = "A cell (with value 1)!"
	cell = row.AddCell()
	cell.Value = "A cell (with value 2)!"
	refTable := NewSharedStringRefTable()
	styles := newXlsxStyleSheet(nil)
	xSheet := sheet.makeXLSXSheet(refTable, styles)

	output := bytes.NewBufferString(xml.Header)
	body, err := xml.Marshal(xSheet)
	c.Assert(err, IsNil)
	c.Assert(body, NotNil)
	_, err = output.Write(body)
	c.Assert(err, IsNil)

	expectedXLSXSheet := `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr filterMode="false"><pageSetUpPr fitToPage="false"></pageSetUpPr></sheetPr><dimension ref="A1:B1"></dimension><sheetViews><sheetView windowProtection="false" showFormulas="false" showGridLines="true" showRowColHeaders="true" showZeros="true" rightToLeft="false" tabSelected="true" showOutlineSymbols="true" defaultGridColor="true" view="normal" topLeftCell="A1" colorId="64" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100" workbookViewId="0"><selection pane="topLeft" activeCell="A1" activeCellId="0" sqref="A1"></selection></sheetView></sheetViews><sheetFormatPr defaultRowHeight="12.85"></sheetFormatPr><cols><col collapsed="false" hidden="false" max="1" min="1" style="0" width="9.5"></col><col collapsed="false" hidden="false" max="2" min="2" style="0" width="9.5"></col></cols><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData><printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"></printOptions><pageMargins left="0.7875" right="0.7875" top="1.05277777777778" bottom="1.05277777777778" header="0.7875" footer="0.7875"></pageMargins><pageSetup paperSize="9" scale="100" firstPageNumber="1" fitToWidth="1" fitToHeight="1" pageOrder="downThenOver" orientation="portrait" usePrinterDefaults="false" blackAndWhite="false" draft="false" cellComments="none" useFirstPageNumber="true" horizontalDpi="300" verticalDpi="300" copies="1"></pageSetup><headerFooter differentFirst="false" differentOddEven="false"><oddHeader>&amp;C&amp;&#34;Times New Roman,Regular&#34;&amp;12&amp;A</oddHeader><oddFooter>&amp;C&amp;&#34;Times New Roman,Regular&#34;&amp;12Page &amp;P</oddFooter></headerFooter></worksheet>`
	c.Assert(output.String(), Equals, expectedXLSXSheet)
}

func (s *SheetSuite) TestSetColWidth(c *C) {
	file := NewFile()
	sheet, _ := file.AddSheet("Sheet1")
	_ = sheet.SetColWidth(0, 0, 10.5)
	_ = sheet.SetColWidth(1, 5, 11)

	c.Assert(sheet.Cols[0].Width, Equals, 10.5)
	c.Assert(sheet.Cols[0].Max, Equals, 1)
	c.Assert(sheet.Cols[0].Min, Equals, 1)
	c.Assert(sheet.Cols[1].Width, Equals, float64(11))
	c.Assert(sheet.Cols[1].Max, Equals, 6)
	c.Assert(sheet.Cols[1].Min, Equals, 2)
}

func (s *SheetSuite) TestSetRowHeightCM(c *C) {
	file := NewFile()
	sheet, _ := file.AddSheet("Sheet1")
	row := sheet.AddRow()
	row.SetHeightCM(1.5)
	c.Assert(row.Height, Equals, 42.51968505)
}

func (s *SheetSuite) TestAlignment(c *C) {
	leftalign := *DefaultAlignment()
	leftalign.Horizontal = "left"
	centerHalign := *DefaultAlignment()
	centerHalign.Horizontal = "center"
	rightalign := *DefaultAlignment()
	rightalign.Horizontal = "right"

	file := NewFile()
	sheet, _ := file.AddSheet("Sheet1")

	style := NewStyle()

	hrow := sheet.AddRow()

	// Horizontals
	cell := hrow.AddCell()
	cell.Value = "left"
	style.Alignment = leftalign
	style.ApplyAlignment = true
	cell.SetStyle(style)

	style = NewStyle()
	cell = hrow.AddCell()
	cell.Value = "centerH"
	style.Alignment = centerHalign
	style.ApplyAlignment = true
	cell.SetStyle(style)

	style = NewStyle()
	cell = hrow.AddCell()
	cell.Value = "right"
	style.Alignment = rightalign
	style.ApplyAlignment = true
	cell.SetStyle(style)

	// Verticals
	topalign := *DefaultAlignment()
	topalign.Vertical = "top"
	centerValign := *DefaultAlignment()
	centerValign.Vertical = "center"
	bottomalign := *DefaultAlignment()
	bottomalign.Vertical = "bottom"

	style = NewStyle()
	vrow := sheet.AddRow()
	cell = vrow.AddCell()
	cell.Value = "top"
	style.Alignment = topalign
	style.ApplyAlignment = true
	cell.SetStyle(style)

	style = NewStyle()
	cell = vrow.AddCell()
	cell.Value = "centerV"
	style.Alignment = centerValign
	style.ApplyAlignment = true
	cell.SetStyle(style)

	style = NewStyle()
	cell = vrow.AddCell()
	cell.Value = "bottom"
	style.Alignment = bottomalign
	style.ApplyAlignment = true
	cell.SetStyle(style)

	parts, err := file.MarshallParts()
	c.Assert(err, IsNil)
	obtained := parts["xl/styles.xml"]

	shouldbe := `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="12"/><name val="Verdana"/><family val="0"/><charset val="0"/></font></fonts><fills count="2"><fill><patternFill patternType="none"><fgColor rgb="FFFFFFFF"/><bgColor rgb="00000000"/></patternFill></fill><fill><patternFill patternType="lightGray"/></fill></fills><borders count="1"><border><left style="none"></left><right style="none"></right><top style="none"></top><bottom style="none"></bottom></border></borders><cellXfs count="8"><xf applyAlignment="0" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="0" fillId="0" fontId="0" numFmtId="0"><alignment horizontal="general" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf><xf applyAlignment="0" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="0" fillId="0" fontId="0" numFmtId="0"><alignment horizontal="general" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf><xf applyAlignment="1" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="0" fillId="0" fontId="0" numFmtId="0"><alignment horizontal="left" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf><xf applyAlignment="1" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="0" fillId="0" fontId="0" numFmtId="0"><alignment horizontal="center" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf><xf applyAlignment="1" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="0" fillId="0" fontId="0" numFmtId="0"><alignment horizontal="right" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf><xf applyAlignment="1" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="0" fillId="0" fontId="0" numFmtId="0"><alignment horizontal="general" indent="0" shrinkToFit="0" textRotation="0" vertical="top" wrapText="0"/></xf><xf applyAlignment="1" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="0" fillId="0" fontId="0" numFmtId="0"><alignment horizontal="general" indent="0" shrinkToFit="0" textRotation="0" vertical="center" wrapText="0"/></xf><xf applyAlignment="1" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="0" fillId="0" fontId="0" numFmtId="0"><alignment horizontal="general" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf></cellXfs></styleSheet>`

	expected := bytes.NewBufferString(shouldbe)

	c.Assert(obtained, Equals, expected.String())
}

func (s *SheetSuite) TestBorder(c *C) {
	file := NewFile()
	sheet, _ := file.AddSheet("Sheet1")
	row := sheet.AddRow()

	cell1 := row.AddCell()
	cell1.Value = "A cell!"
	style1 := NewStyle()
	style1.Border = *NewBorder("thin", "thin", "thin", "thin")
	style1.ApplyBorder = true
	cell1.SetStyle(style1)

	refTable := NewSharedStringRefTable()
	styles := newXlsxStyleSheet(nil)
	worksheet := sheet.makeXLSXSheet(refTable, styles)

	c.Assert(styles.Borders.Border[1].Left.Style, Equals, "thin")
	c.Assert(styles.Borders.Border[1].Right.Style, Equals, "thin")
	c.Assert(styles.Borders.Border[1].Top.Style, Equals, "thin")
	c.Assert(styles.Borders.Border[1].Bottom.Style, Equals, "thin")

	c.Assert(worksheet.SheetData.Row[0].C[0].S, Equals, 1)
}

func (s *SheetSuite) TestOutlineLevels(c *C) {
	file := NewFile()
	sheet, _ := file.AddSheet("Sheet1")

	r1 := sheet.AddRow()
	c11 := r1.AddCell()
	c11.Value = "A1"
	c12 := r1.AddCell()
	c12.Value = "B1"

	r2 := sheet.AddRow()
	c21 := r2.AddCell()
	c21.Value = "A2"
	c22 := r2.AddCell()
	c22.Value = "B2"

	r3 := sheet.AddRow()
	c31 := r3.AddCell()
	c31.Value = "A3"
	c32 := r3.AddCell()
	c32.Value = "B3"

	// Add some groups
	r1.OutlineLevel = 1
	r2.OutlineLevel = 2
	sheet.Col(0).OutlineLevel = 1

	refTable := NewSharedStringRefTable()
	styles := newXlsxStyleSheet(nil)
	worksheet := sheet.makeXLSXSheet(refTable, styles)

	c.Assert(worksheet.SheetFormatPr.OutlineLevelCol, Equals, uint8(1))
	c.Assert(worksheet.SheetFormatPr.OutlineLevelRow, Equals, uint8(2))

	c.Assert(worksheet.Cols.Col[0].OutlineLevel, Equals, uint8(1))
	c.Assert(worksheet.Cols.Col[1].OutlineLevel, Equals, uint8(0))
	c.Assert(worksheet.SheetData.Row[0].OutlineLevel, Equals, uint8(1))
	c.Assert(worksheet.SheetData.Row[1].OutlineLevel, Equals, uint8(2))
	c.Assert(worksheet.SheetData.Row[2].OutlineLevel, Equals, uint8(0))
}

func (s *SheetSuite) TestAutoFilter(c *C) {
	file := NewFile()
	sheet, _ := file.AddSheet("Sheet1")

	r1 := sheet.AddRow()
	r1.AddCell()
	r1.AddCell()
	r1.AddCell()

	r2 := sheet.AddRow()
	r2.AddCell()
	r2.AddCell()
	r2.AddCell()

	r3 := sheet.AddRow()
	r3.AddCell()
	r3.AddCell()
	r3.AddCell()

	// Define a filter area
	sheet.AutoFilter = &AutoFilter{TopLeftCell: "B2", BottomRightCell: "C3"}

	refTable := NewSharedStringRefTable()
	styles := newXlsxStyleSheet(nil)
	worksheet := sheet.makeXLSXSheet(refTable, styles)

	c.Assert(worksheet.AutoFilter, NotNil)
	c.Assert(worksheet.AutoFilter.Ref, Equals, "B2:C3")
}
