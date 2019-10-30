package xlsx

import (
	"bytes"
	"encoding/xml"
	"testing"

	qt "github.com/frankban/quicktest"
	. "gopkg.in/check.v1"
)

type SheetSuite struct{}

var _ = Suite(&SheetSuite{})

// Test we can add a Row to a Sheet
func (s *SheetSuite) TestAddRow(c *C) {
	// Create a file with three rows.
	var f *File
	f = NewFile()
	sheet, _ := f.AddSheet("MySheet")
	row0 := sheet.AddRow()
	cell0 := row0.AddCell()
	cell0.Value = "Row 0"
	c.Assert(row0, NotNil)
	row1 := sheet.AddRow()
	cell1 := row1.AddCell()
	cell1.Value = "Row 1"
	row2 := sheet.AddRow()
	cell2 := row2.AddCell()
	cell2.Value = "Row 2"
	// Check the file
	expected := []string{"Row 0", "Row 1", "Row 2"}
	c.Assert(len(sheet.Rows), Equals, len(expected))
	for i, row := range sheet.Rows {
		c.Assert(row.Cells[0].Value, Equals, expected[i])
	}

	// Insert a row in the middle
	row1pt5, err := sheet.AddRowAtIndex(2)
	c.Assert(err, IsNil)
	cell1pt5 := row1pt5.AddCell()
	cell1pt5.Value = "Row 1.5"

	expected = []string{"Row 0", "Row 1", "Row 1.5", "Row 2"}
	c.Assert(len(sheet.Rows), Equals, len(expected))
	for i, row := range sheet.Rows {
		c.Assert(row.Cells[0].Value, Equals, expected[i])
	}

	// Insert a row at the beginning
	rowNewStart, err := sheet.AddRowAtIndex(0)
	c.Assert(err, IsNil)
	cellNewStart := rowNewStart.AddCell()
	cellNewStart.Value = "Row -1"
	// Insert a row at one index past the end, this is the same as AddRow().
	row2pt5, err := sheet.AddRowAtIndex(5)
	c.Assert(err, IsNil)
	cell2pt5 := row2pt5.AddCell()
	cell2pt5.Value = "Row 2.5"

	expected = []string{"Row -1", "Row 0", "Row 1", "Row 1.5", "Row 2", "Row 2.5"}
	c.Assert(len(sheet.Rows), Equals, len(expected))
	for i, row := range sheet.Rows {
		c.Assert(row.Cells[0].Value, Equals, expected[i])
	}

	// Negative and out of range indicies should fail for insert
	_, err = sheet.AddRowAtIndex(-1)
	c.Assert(err, NotNil)
	// Since we allow inserting into the position that does not yet exist, it has to be 1 greater
	// than you would think in order to fail.
	_, err = sheet.AddRowAtIndex(7)
	c.Assert(err, NotNil)

	// Negative and out of range indicies should fail for remove
	err = sheet.RemoveRowAtIndex(-1)
	c.Assert(err, NotNil)
	err = sheet.RemoveRowAtIndex(6)
	c.Assert(err, NotNil)

	// Remove from the beginning, the end, and the middle.
	err = sheet.RemoveRowAtIndex(0)
	c.Assert(err, IsNil)
	err = sheet.RemoveRowAtIndex(4)
	c.Assert(err, IsNil)
	err = sheet.RemoveRowAtIndex(2)
	c.Assert(err, IsNil)

	expected = []string{"Row 0", "Row 1", "Row 2"}
	c.Assert(len(sheet.Rows), Equals, len(expected))
	for i, row := range sheet.Rows {
		c.Assert(row.Cells[0].Value, Equals, expected[i])
	}
}

// Test we can get row by index from  Sheet
func (s *SheetSuite) TestGetRowByIndex(c *C) {
	var f *File
	f = NewFile()
	sheet, _ := f.AddSheet("MySheet")
	row := sheet.Row(10)
	c.Assert(row, NotNil)
	c.Assert(len(sheet.Rows), Equals, 11)
	row = sheet.Row(2)
	c.Assert(row, NotNil)
	c.Assert(len(sheet.Rows), Equals, 11)
}

func (s *SheetSuite) TestMakeXLSXSheetFromRows(c *C) {
	file := NewFile()
	sheet, _ := file.AddSheet("Sheet1")
	row := sheet.AddRow()
	cell := row.AddCell()
	cell.Value = "A cell!"
	refTable := NewSharedStringRefTable()
	styles := newXlsxStyleSheet(nil)
	xSheet := sheet.makeXLSXSheet(refTable, styles, nil)
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
func TestMakeXLSXSheetWithNumFormats(t *testing.T) {
	c := qt.New(t)
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
	worksheet := sheet.makeXLSXSheet(refTable, styles, nil)

	c.Assert(styles.CellStyleXfs, qt.IsNil)

	c.Assert(styles.CellXfs.Count, qt.Equals, 4)
	c.Assert(styles.CellXfs.Xf[0].NumFmtId, qt.Equals, 0)
	c.Assert(styles.CellXfs.Xf[1].NumFmtId, qt.Equals, 1)
	c.Assert(styles.CellXfs.Xf[2].NumFmtId, qt.Equals, 14)
	c.Assert(styles.CellXfs.Xf[3].NumFmtId, qt.Equals, 164)
	c.Assert(styles.NumFmts.Count, qt.Equals, 1)
	c.Assert(styles.NumFmts.NumFmt[0].NumFmtId, qt.Equals, 164)
	c.Assert(styles.NumFmts.NumFmt[0].FormatCode, qt.Equals, "hh:mm:ss")

	// Finally we check that the cell points to the right CellXf /
	// CellStyleXf.
	c.Assert(worksheet.SheetData.Row[0].C[0].S, qt.Equals, 0)
	c.Assert(worksheet.SheetData.Row[0].C[1].S, qt.Equals, 1)
}

// When we create the xlsxSheet we also populate the xlsxStyles struct
// with style information.
func TestMakeXLSXSheetAlsoPopulatesXLSXSTyles(t *testing.T) {
	c := qt.New(t)
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
	worksheet := sheet.makeXLSXSheet(refTable, styles, nil)

	c.Assert(styles.Fonts.Count, qt.Equals, 1)
	c.Assert(styles.Fonts.Font[0].Sz.Val, qt.Equals, "10")
	c.Assert(styles.Fonts.Font[0].Name.Val, qt.Equals, "Verdana")

	c.Assert(styles.Fills.Count, qt.Equals, 2)
	c.Assert(styles.Fills.Fill[0].PatternFill.PatternType, qt.Equals, "solid")
	c.Assert(styles.Fills.Fill[0].PatternFill.FgColor.RGB, qt.Equals, "FFFFFFFF")
	c.Assert(styles.Fills.Fill[0].PatternFill.BgColor.RGB, qt.Equals, "00000000")

	c.Assert(styles.Borders.Count, qt.Equals, 1)
	c.Assert(styles.Borders.Border[0].Left.Style, qt.Equals, "none")
	c.Assert(styles.Borders.Border[0].Right.Style, qt.Equals, "thin")
	c.Assert(styles.Borders.Border[0].Top.Style, qt.Equals, "none")
	c.Assert(styles.Borders.Border[0].Bottom.Style, qt.Equals, "thin")

	c.Assert(styles.CellStyleXfs, qt.IsNil)

	c.Assert(styles.CellXfs.Count, qt.Equals, 1)
	c.Assert(styles.CellXfs.Xf[0].FontId, qt.Equals, 0)
	c.Assert(styles.CellXfs.Xf[0].FillId, qt.Equals, 0)
	c.Assert(styles.CellXfs.Xf[0].BorderId, qt.Equals, 0)

	// Finally we check that the cell points to the right CellXf /
	// CellStyleXf.
	c.Assert(worksheet.SheetData.Row[0].C[0].S, qt.Equals, 0)
	c.Assert(worksheet.SheetData.Row[0].C[1].S, qt.Equals, 0)
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
	worksheet := sheet.makeXLSXSheet(refTable, styles, nil)
	c.Assert(worksheet.Cols, IsNil)
}

// If the column width is customised, the xslxCol.CustomWidth field is set to 1.
func TestMakeXLSXSheetSetsCustomColWidth(t *testing.T) {
	c := qt.New(t)
	file := NewFile()
	sheet, _ := file.AddSheet("Sheet1")
	row := sheet.AddRow()
	cell1 := row.AddCell()
	cell1.Value = "A cell!"
	sheet.SetColWidth(1, 1, 10.5)

	refTable := NewSharedStringRefTable()
	styles := newXlsxStyleSheet(nil)
	worksheet := sheet.makeXLSXSheet(refTable, styles, nil)
	c.Assert(worksheet.Cols.Col[0].CustomWidth, qt.Equals, true)
}

func (s *SheetSuite) TestMarshalSheet(c *C) {
	file := NewFile()
	sheet, _ := file.AddSheet("Sheet1")
	row := sheet.AddRow()
	cell := row.AddCell()
	cell.Value = "A cell!"
	refTable := NewSharedStringRefTable()
	styles := newXlsxStyleSheet(nil)
	xSheet := sheet.makeXLSXSheet(refTable, styles, nil)

	output := bytes.NewBufferString(xml.Header)
	body, err := xml.Marshal(xSheet)
	c.Assert(err, IsNil)
	c.Assert(body, NotNil)
	_, err = output.Write(body)
	c.Assert(err, IsNil)

	expectedXLSXSheet := `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr filterMode="false"><pageSetUpPr fitToPage="false"></pageSetUpPr></sheetPr><dimension ref="A1"></dimension><sheetViews><sheetView windowProtection="false" showFormulas="false" showGridLines="true" showRowColHeaders="true" showZeros="true" rightToLeft="false" tabSelected="true" showOutlineSymbols="true" defaultGridColor="true" view="normal" topLeftCell="A1" colorId="64" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100" workbookViewId="0"><selection pane="topLeft" activeCell="A1" activeCellId="0" sqref="A1"></selection></sheetView></sheetViews><sheetFormatPr defaultRowHeight="12.85"></sheetFormatPr><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData><printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"></printOptions><pageMargins left="0.7875" right="0.7875" top="1.05277777777778" bottom="1.05277777777778" header="0.7875" footer="0.7875"></pageMargins><pageSetup paperSize="9" scale="100" firstPageNumber="1" fitToWidth="1" fitToHeight="1" pageOrder="downThenOver" orientation="portrait" usePrinterDefaults="false" blackAndWhite="false" draft="false" cellComments="none" useFirstPageNumber="true" horizontalDpi="300" verticalDpi="300" copies="1"></pageSetup><headerFooter differentFirst="false" differentOddEven="false"><oddHeader>&amp;C&amp;&#34;Times New Roman,Regular&#34;&amp;12&amp;A</oddHeader><oddFooter>&amp;C&amp;&#34;Times New Roman,Regular&#34;&amp;12Page &amp;P</oddFooter></headerFooter></worksheet>`

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
	xSheet := sheet.makeXLSXSheet(refTable, styles, nil)

	output := bytes.NewBufferString(xml.Header)
	body, err := xml.Marshal(xSheet)
	c.Assert(err, IsNil)
	c.Assert(body, NotNil)
	_, err = output.Write(body)
	c.Assert(err, IsNil)

	expectedXLSXSheet := `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr filterMode="false"><pageSetUpPr fitToPage="false"></pageSetUpPr></sheetPr><dimension ref="A1:B1"></dimension><sheetViews><sheetView windowProtection="false" showFormulas="false" showGridLines="true" showRowColHeaders="true" showZeros="true" rightToLeft="false" tabSelected="true" showOutlineSymbols="true" defaultGridColor="true" view="normal" topLeftCell="A1" colorId="64" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100" workbookViewId="0"><selection pane="topLeft" activeCell="A1" activeCellId="0" sqref="A1"></selection></sheetView></sheetViews><sheetFormatPr defaultRowHeight="12.85"></sheetFormatPr><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData><printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"></printOptions><pageMargins left="0.7875" right="0.7875" top="1.05277777777778" bottom="1.05277777777778" header="0.7875" footer="0.7875"></pageMargins><pageSetup paperSize="9" scale="100" firstPageNumber="1" fitToWidth="1" fitToHeight="1" pageOrder="downThenOver" orientation="portrait" usePrinterDefaults="false" blackAndWhite="false" draft="false" cellComments="none" useFirstPageNumber="true" horizontalDpi="300" verticalDpi="300" copies="1"></pageSetup><headerFooter differentFirst="false" differentOddEven="false"><oddHeader>&amp;C&amp;&#34;Times New Roman,Regular&#34;&amp;12&amp;A</oddHeader><oddFooter>&amp;C&amp;&#34;Times New Roman,Regular&#34;&amp;12Page &amp;P</oddFooter></headerFooter></worksheet>`
	c.Assert(output.String(), Equals, expectedXLSXSheet)
}

func TestSetColWidth(t *testing.T) {
	c := qt.New(t)
	file := NewFile()
	sheet, _ := file.AddSheet("Sheet1")
	sheet.SetColWidth(1, 1, 10.5)
	sheet.SetColWidth(2, 6, 11)
	c.Assert(sheet.Cols.FindColByIndex(1).Width, qt.Equals, 10.5)
	c.Assert(sheet.Cols.FindColByIndex(1).Max, qt.Equals, 1)
	c.Assert(sheet.Cols.FindColByIndex(1).Min, qt.Equals, 1)
	c.Assert(sheet.Cols.FindColByIndex(2).Width, qt.Equals, float64(11))
	c.Assert(sheet.Cols.FindColByIndex(2).Max, qt.Equals, 6)
	c.Assert(sheet.Cols.FindColByIndex(2).Min, qt.Equals, 2)
}

func TestSetDataValidation(t *testing.T) {
	c := qt.New(t)
	file := NewFile()
	sheet, _ := file.AddSheet("Sheet1")

	dd := NewDataValidation(0, 0, 10, 0, true)
	err := dd.SetDropList([]string{"a1", "a2", "a3"})
	c.Assert(err, qt.IsNil)

	sheet.AddDataValidation(dd)
	c.Assert(sheet.DataValidations, qt.HasLen, 1)
	c.Assert(sheet.DataValidations[0], qt.Equals, dd)
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
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Arial"/><family val="2"/><color theme="1" /><scheme val="minor"/></font><font><sz val="12"/><name val="Verdana"/><family val="0"/><charset val="0"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="lightGray"/></fill></fills><borders count="2"><border><left/><right/><top/><bottom/></border><border><left style="none"></left><right style="none"></right><top style="none"></top><bottom style="none"></bottom></border></borders><cellStyleXfs count="1"><xf applyAlignment="0" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="0" fillId="0" fontId="0" numFmtId="0"><alignment horizontal="general" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf></cellStyleXfs><cellXfs count="7"><xf applyAlignment="0" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="0" fillId="0" fontId="0" numFmtId="0"><alignment horizontal="general" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf><xf applyAlignment="1" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="1" fillId="0" fontId="1" numFmtId="0"><alignment horizontal="left" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf><xf applyAlignment="1" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="1" fillId="0" fontId="1" numFmtId="0"><alignment horizontal="center" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf><xf applyAlignment="1" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="1" fillId="0" fontId="1" numFmtId="0"><alignment horizontal="right" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf><xf applyAlignment="1" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="1" fillId="0" fontId="1" numFmtId="0"><alignment horizontal="general" indent="0" shrinkToFit="0" textRotation="0" vertical="top" wrapText="0"/></xf><xf applyAlignment="1" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="1" fillId="0" fontId="1" numFmtId="0"><alignment horizontal="general" indent="0" shrinkToFit="0" textRotation="0" vertical="center" wrapText="0"/></xf><xf applyAlignment="1" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="1" fillId="0" fontId="1" numFmtId="0"><alignment horizontal="general" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf></cellXfs></styleSheet>`

	expected := bytes.NewBufferString(shouldbe)
	c.Assert(obtained, Equals, expected.String())
}

func TestBorder(t *testing.T) {
	c := qt.New(t)
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
	worksheet := sheet.makeXLSXSheet(refTable, styles, nil)

	c.Assert(styles.Borders.Border[0].Left.Style, qt.Equals, "thin")
	c.Assert(styles.Borders.Border[0].Right.Style, qt.Equals, "thin")
	c.Assert(styles.Borders.Border[0].Top.Style, qt.Equals, "thin")
	c.Assert(styles.Borders.Border[0].Bottom.Style, qt.Equals, "thin")

	c.Assert(worksheet.SheetData.Row[0].C[0].S, qt.Equals, 0)
}

func TestOutlineLevels(t *testing.T) {
	c := qt.New(t)
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
	sheet.SetOutlineLevel(1, 1, 1)

	refTable := NewSharedStringRefTable()
	styles := newXlsxStyleSheet(nil)
	worksheet := sheet.makeXLSXSheet(refTable, styles, nil)

	c.Assert(worksheet.SheetFormatPr.OutlineLevelCol, qt.Equals, uint8(1))
	c.Assert(worksheet.SheetFormatPr.OutlineLevelRow, qt.Equals, uint8(2))

	c.Assert(worksheet.Cols.Col[0].OutlineLevel, qt.Equals, uint8(1))
	c.Assert(worksheet.SheetData.Row[0].OutlineLevel, qt.Equals, uint8(1))
	c.Assert(worksheet.SheetData.Row[1].OutlineLevel, qt.Equals, uint8(2))
	c.Assert(worksheet.SheetData.Row[2].OutlineLevel, qt.Equals, uint8(0))
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
	worksheet := sheet.makeXLSXSheet(refTable, styles, nil)

	c.Assert(worksheet.AutoFilter, NotNil)
	c.Assert(worksheet.AutoFilter.Ref, Equals, "B2:C3")
}
