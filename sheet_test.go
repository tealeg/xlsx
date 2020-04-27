package xlsx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"io/ioutil"
	"path/filepath"
	"strings"
	"testing"

	qt "github.com/frankban/quicktest"
)

func TestSheet(t *testing.T) {
	c := qt.New(t)
	// Test we can add a Row to a Sheet
	csRunO(c, "TestAddAndRemoveRow", func(c *qt.C, option FileOption) {
		setUp := func() (*Sheet, error) {
			var f *File
			f = NewFile(option)
			sheet, err := f.AddSheet("MySheet")
			if err != nil {
				return nil, err
			}

			row0 := sheet.AddRow()
			cell0 := row0.AddCell()
			cell0.Value = "Row 0"
			c.Assert(row0, qt.Not(qt.IsNil))
			row1 := sheet.AddRow()
			cell1 := row1.AddCell()
			cell1.Value = "Row 1"
			row2 := sheet.AddRow()
			cell2 := row2.AddCell()
			cell2.Value = "Row 2"
			return sheet, nil
		}

		assertRow := func(c *qt.C, sheet *Sheet, expected []string) {
			c.Assert(sheet.MaxRow, qt.Equals, len(expected))
			sheet.ForEachRow(func(row *Row) error {
				c.Assert(row.GetCell(0).Value, qt.Equals, expected[row.num])
				return nil
			})

		}

		c.Run("AddRow", func(c *qt.C) {
			c.Run("SimpleAddRow", func(c *qt.C) {
				sheet, err := setUp()
				c.Assert(err, qt.Equals, nil)
				// Check the file
				assertRow(c, sheet, []string{"Row 0", "Row 1", "Row 2"})
			})

			c.Run("InsertARowInTheMiddle", func(c *qt.C) {
				sheet, err := setUp()
				c.Assert(err, qt.Equals, nil)
				row1pt5, err := sheet.AddRowAtIndex(2)
				c.Assert(err, qt.IsNil)
				cell1pt5 := row1pt5.AddCell()
				cell1pt5.Value = "Row 1.5"

				assertRow(c, sheet, []string{"Row 0", "Row 1", "Row 1.5", "Row 2"})
			})

			c.Run("InsertARowAtBeginning", func(c *qt.C) {
				sheet, err := setUp()
				c.Assert(err, qt.Equals, nil)
				assertRow(c, sheet, []string{"Row 0", "Row 1", "Row 2"})
				rowNewStart, err := sheet.AddRowAtIndex(0)
				c.Assert(err, qt.IsNil)
				cellNewStart := rowNewStart.AddCell()
				cellNewStart.Value = "Row -1"
				assertRow(c, sheet, []string{"Row -1", "Row 0", "Row 1", "Row 2"})

			})

			c.Run("InsertARowAtTheEnd", func(c *qt.C) {
				// Insert a row at one index past the end,
				// this is the same as AddRow().
				sheet, err := setUp()
				c.Assert(err, qt.Equals, nil)

				row2pt5, err := sheet.AddRowAtIndex(3)
				c.Assert(err, qt.IsNil)
				cell2pt5 := row2pt5.AddCell()
				cell2pt5.Value = "Row 2.5"

				assertRow(c, sheet, []string{"Row 0", "Row 1", "Row 2", "Row 2.5"})
			})

			c.Run("NegativeIndexFails", func(c *qt.C) {
				// TODO: why do we accept an int type instead of a uint?
				sheet, err := setUp()
				c.Assert(err, qt.Equals, nil)

				_, err = sheet.AddRowAtIndex(-1)
				c.Assert(err, qt.Not(qt.IsNil))
			})

			c.Run("BeyondMaxRowPlusOne", func(c *qt.C) {
				// TODO: The behaviour right now is that we
				// allow you to AddRowAtIndex up to an
				// including in one place after the final row,
				// this seems arbitrary and wrong.
				sheet, err := setUp()
				c.Assert(err, qt.Equals, nil)

				_, err = sheet.AddRowAtIndex(7)
				c.Assert(err, qt.Not(qt.IsNil))
			})
		})
		c.Run("RemoveRow", func(c *qt.C) {
			c.Run("NegativeIndex", func(c *qt.C) {
				// Negative and out of range indicies should fail for remove
				sheet, err := setUp()
				c.Assert(err, qt.Equals, nil)

				err = sheet.RemoveRowAtIndex(-1)
				c.Assert(err, qt.Not(qt.IsNil))
			})
			c.Run("IndexOutOfRange", func(c *qt.C) {
				sheet, err := setUp()
				c.Assert(err, qt.Equals, nil)

				err = sheet.RemoveRowAtIndex(6)
				c.Assert(err, qt.Not(qt.IsNil))

			})

			c.Run("RemoveFromTheBeginning", func(c *qt.C) {
				sheet, err := setUp()
				c.Assert(err, qt.Equals, nil)

				err = sheet.RemoveRowAtIndex(0)
				c.Assert(err, qt.IsNil)
				assertRow(c, sheet, []string{"Row 1", "Row 2"})
			})

			c.Run("RemoveFromTheEnd", func(c *qt.C) {
				sheet, err := setUp()
				c.Assert(err, qt.Equals, nil)

				err = sheet.RemoveRowAtIndex(2)
				c.Assert(err, qt.IsNil)
				assertRow(c, sheet, []string{"Row 0", "Row 1"})
			})

			c.Run("RemoveFromTheMiddle", func(c *qt.C) {
				sheet, err := setUp()
				c.Assert(err, qt.Equals, nil)

				err = sheet.RemoveRowAtIndex(1)
				c.Assert(err, qt.IsNil)
				assertRow(c, sheet, []string{"Row 0", "Row 2"})
			})

		})
	})

	// Test we can get row by index from  Sheet
	csRunO(c, "TestGetRowByIndex", func(c *qt.C, option FileOption) {
		var f *File
		f = NewFile()
		sheet, _ := f.AddSheet("MySheet")
		row, err := sheet.Row(10)
		c.Assert(err, qt.Equals, nil)
		c.Assert(row, qt.Not(qt.IsNil))
		c.Assert(sheet.MaxRow, qt.Equals, 11)
		row, err = sheet.Row(2)
		c.Assert(err, qt.Equals, nil)
		c.Assert(row, qt.Not(qt.IsNil))
		c.Assert(sheet.MaxRow, qt.Equals, 11)
	})

	csRunO(c, "TestMakeXLSXSheetFromRows", func(c *qt.C, option FileOption) {
		file := NewFile(option)
		sheet, _ := file.AddSheet("Sheet1")
		row := sheet.AddRow()
		cell := row.AddCell()
		cell.Value = "A cell!"

		var buf bytes.Buffer

		refTable := NewSharedStringRefTable()
		styles := newXlsxStyleSheet(nil)
		err := sheet.MarshalSheet(&buf, refTable, styles, nil)
		c.Assert(err, qt.Equals, nil)
		var xSheet xlsxWorksheet
		err = xml.Unmarshal(buf.Bytes(), &xSheet)
		c.Assert(err, qt.Equals, nil)

		c.Assert(xSheet.Dimension.Ref, qt.Equals, "A1")
		c.Assert(len(xSheet.SheetData.Row), qt.Equals, 1)
		xRow := xSheet.SheetData.Row[0]
		c.Assert(xRow.R, qt.Equals, 1)
		c.Assert(xRow.Spans, qt.Equals, "")
		c.Assert(len(xRow.C), qt.Equals, 1)
		xC := xRow.C[0]
		c.Assert(xC.R, qt.Equals, "A1")
		c.Assert(xC.S, qt.Equals, 0)
		c.Assert(xC.T, qt.Equals, "s") // Shared string type
		c.Assert(xC.V, qt.Equals, "0") // reference to shared string
		xSST := refTable.makeXLSXSST()
		c.Assert(xSST.Count, qt.Equals, 1)
		c.Assert(xSST.UniqueCount, qt.Equals, 1)
		c.Assert(len(xSST.SI), qt.Equals, 1)
		xSI := xSST.SI[0]
		c.Assert(xSI.T.Text, qt.Equals, "A cell!")
		c.Assert(xSI.R, qt.HasLen, 0)
	})

	csRunO(c, "TestMarshalSheetFromRows", func(c *qt.C, option FileOption) {
		file := NewFile(option)
		sheet, _ := file.AddSheet("Sheet1")
		row := sheet.AddRow()
		cell := row.AddCell()
		cell.Value = "A cell!"
		refTable := NewSharedStringRefTable()
		styles := newXlsxStyleSheet(nil)
		var output bytes.Buffer
		err := sheet.MarshalSheet(&output, refTable, styles, nil)
		c.Assert(err, qt.IsNil)

		var xSheet xlsxWorksheet
		err = xml.Unmarshal(output.Bytes(), &xSheet)
		c.Assert(xSheet.Dimension.Ref, qt.Equals, "A1")
		c.Assert(len(xSheet.SheetData.Row), qt.Equals, 1)
		xRow := xSheet.SheetData.Row[0]
		c.Assert(xRow.R, qt.Equals, 1)
		c.Assert(xRow.Spans, qt.Equals, "")
		c.Assert(len(xRow.C), qt.Equals, 1)
		xC := xRow.C[0]
		c.Assert(xC.R, qt.Equals, "A1")
		c.Assert(xC.S, qt.Equals, 0)
		c.Assert(xC.T, qt.Equals, "s") // Shared string type
		c.Assert(xC.V, qt.Equals, "0") // reference to shared string
		xSST := refTable.makeXLSXSST()
		c.Assert(xSST.Count, qt.Equals, 1)
		c.Assert(xSST.UniqueCount, qt.Equals, 1)
		c.Assert(len(xSST.SI), qt.Equals, 1)
		xSI := xSST.SI[0]
		c.Assert(xSI.T.Text, qt.Equals, "A cell!")
		c.Assert(xSI.R, qt.HasLen, 0)
	})

	// If the column width is not customised, the xslxCol.CustomWidth field is set to 0.
	csRunO(c, "TestMarshalSheetDefaultsCustomColWidth", func(c *qt.C, option FileOption) {
		file := NewFile(option)
		sheet, _ := file.AddSheet("Sheet1")
		row := sheet.AddRow()
		cell1 := row.AddCell()
		cell1.Value = "A cell!"

		refTable := NewSharedStringRefTable()
		styles := newXlsxStyleSheet(nil)

		var output bytes.Buffer
		err := sheet.MarshalSheet(&output, refTable, styles, nil)
		c.Assert(err, qt.IsNil)

		var result xlsxWorksheet
		err = xml.Unmarshal(output.Bytes(), &result)
		c.Assert(result.Cols, qt.IsNil)
	})

	csRunO(c, "TestMarshalSheetWithColStyle", func(qc *qt.C, option FileOption) {
		file := NewFile(option)
		sheet, _ := file.AddSheet("Sheet1")
		row := sheet.AddRow()
		cell1 := row.AddCell()
		cell1.Value = "A cell!"

		colStyle := NewStyle()
		colStyle.Fill.FgColor = "EEEEEE00"
		colStyle.Fill.PatternType = "solid"
		colStyle.ApplyFill = true
		col := NewColForRange(10, 11)
		col.SetStyle(colStyle)
		sheet.Cols.Add(col)

		var buf bytes.Buffer

		refTable := NewSharedStringRefTable()
		styles := newXlsxStyleSheet(nil)
		err := sheet.MarshalSheet(&buf, refTable, styles, nil)
		c.Assert(err, qt.Equals, nil)
		var xSheet xlsxWorksheet
		err = xml.Unmarshal(buf.Bytes(), &xSheet)
		c.Assert(err, qt.Equals, nil)
		c.Assert(xSheet.Cols, qt.Not(qt.IsNil))
		c.Assert(*xSheet.Cols.Col[0].Style, qt.Equals, 0)
		c.Assert(styles.getStyle(0), qt.DeepEquals, colStyle)
	})

	csRunO(c, "TestMarshalSheet", func(c *qt.C, option FileOption) {
		file := NewFile(option)
		sheet, _ := file.AddSheet("Sheet1")
		row := sheet.AddRow()
		cell := row.AddCell()
		cell.Value = "A cell!"
		refTable := NewSharedStringRefTable()
		styles := newXlsxStyleSheet(nil)
		var output strings.Builder
		err := sheet.MarshalSheet(&output, refTable, styles, nil)
		c.Assert(err, qt.IsNil)

		expectedXLSXSheet := `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr filterMode="false"><pageSetUpPr fitToPage="false"/></sheetPr><dimension ref="A1"/><sheetViews><sheetView windowProtection="false" showFormulas="false" showGridLines="true" showRowColHeaders="true" showZeros="true" rightToLeft="false" tabSelected="true" showOutlineSymbols="true" defaultGridColor="true" view="normal" topLeftCell="A1" colorId="64" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100" workbookViewId="0"><selection pane="topLeft" activeCell="A1" activeCellId="0" sqref="A1"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="12.85"/><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>`

		c.Assert(output.String(), qt.Equals, expectedXLSXSheet)
	})

	csRunO(c, "TestMarshalSheetWithMultipleCells", func(c *qt.C, option FileOption) {
		file := NewFile(option)
		sheet, _ := file.AddSheet("Sheet1")
		row := sheet.AddRow()
		cell := row.AddCell()
		cell.Value = "A cell (with value 1)!"
		cell = row.AddCell()
		cell.Value = "A cell (with value 2)!"
		var buf bytes.Buffer

		refTable := NewSharedStringRefTable()
		styles := newXlsxStyleSheet(nil)
		err := sheet.MarshalSheet(&buf, refTable, styles, nil)
		c.Assert(err, qt.Equals, nil)

		expectedXLSXSheet := `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr filterMode="false"><pageSetUpPr fitToPage="false"/></sheetPr><dimension ref="A1:B1"/><sheetViews><sheetView windowProtection="false" showFormulas="false" showGridLines="true" showRowColHeaders="true" showZeros="true" rightToLeft="false" tabSelected="true" showOutlineSymbols="true" defaultGridColor="true" view="normal" topLeftCell="A1" colorId="64" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100" workbookViewId="0"><selection pane="topLeft" activeCell="A1" activeCellId="0" sqref="A1"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="12.85"/><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData></worksheet>`
		c.Assert(buf.String(), qt.Equals, expectedXLSXSheet)
	})
	csRunO(c, "TestSetRowHeightCM", func(c *qt.C, option FileOption) {
		file := NewFile(option)
		sheet, _ := file.AddSheet("Sheet1")
		row := sheet.AddRow()
		row.SetHeightCM(1.5)
		c.Assert(row.GetHeight(), qt.Equals, 42.51968505)
	})

	csRunO(c, "TestAlignment", func(c *qt.C, option FileOption) {
		leftalign := *DefaultAlignment()
		leftalign.Horizontal = "left"
		centerHalign := *DefaultAlignment()
		centerHalign.Horizontal = "center"
		rightalign := *DefaultAlignment()
		rightalign.Horizontal = "right"

		file := NewFile(option)
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

		dir := c.Mkdir()
		path := filepath.Join(dir, "test.xlsx")
		err := file.Save(path)
		c.Assert(err, qt.IsNil)
		r, err := zip.OpenReader(path)
		c.Assert(err, qt.IsNil)
		defer r.Close()

		var obtained []byte
		for _, f := range r.File {
			if f.Name == "xl/styles.xml" {
				rc, err := f.Open()
				c.Assert(err, qt.Equals, nil)
				obtained, err = ioutil.ReadAll(rc)
				c.Assert(err, qt.Equals, nil)
			}
		}

		shouldbe := `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Arial"/><family val="2"/><color theme="1" /><scheme val="minor"/></font><font><sz val="12"/><name val="Verdana"/><family val="0"/><charset val="0"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="lightGray"/></fill></fills><borders count="2"><border><left/><right/><top/><bottom/></border><border><left style="none"></left><right style="none"></right><top style="none"></top><bottom style="none"></bottom></border></borders><cellStyleXfs count="1"><xf applyAlignment="0" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="0" fillId="0" fontId="0" numFmtId="0"><alignment horizontal="general" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf></cellStyleXfs><cellXfs count="7"><xf applyAlignment="0" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="0" fillId="0" fontId="0" numFmtId="0"><alignment horizontal="general" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf><xf applyAlignment="1" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="1" fillId="0" fontId="1" numFmtId="0"><alignment horizontal="left" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf><xf applyAlignment="1" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="1" fillId="0" fontId="1" numFmtId="0"><alignment horizontal="center" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf><xf applyAlignment="1" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="1" fillId="0" fontId="1" numFmtId="0"><alignment horizontal="right" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf><xf applyAlignment="1" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="1" fillId="0" fontId="1" numFmtId="0"><alignment horizontal="general" indent="0" shrinkToFit="0" textRotation="0" vertical="top" wrapText="0"/></xf><xf applyAlignment="1" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="1" fillId="0" fontId="1" numFmtId="0"><alignment horizontal="general" indent="0" shrinkToFit="0" textRotation="0" vertical="center" wrapText="0"/></xf><xf applyAlignment="1" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="1" fillId="0" fontId="1" numFmtId="0"><alignment horizontal="general" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf></cellXfs></styleSheet>`

		expected := bytes.NewBufferString(shouldbe)
		c.Assert(string(obtained), qt.Equals, expected.String())
	})

	csRunO(c, "TestAutoFilter", func(c *qt.C, option FileOption) {
		file := NewFile(option)
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

		var buf bytes.Buffer

		refTable := NewSharedStringRefTable()
		styles := newXlsxStyleSheet(nil)
		err := sheet.MarshalSheet(&buf, refTable, styles, nil)
		c.Assert(err, qt.Equals, nil)
		var xSheet xlsxWorksheet
		err = xml.Unmarshal(buf.Bytes(), &xSheet)
		c.Assert(err, qt.Equals, nil)

		c.Assert(xSheet.AutoFilter, qt.Not(qt.IsNil))
		c.Assert(xSheet.AutoFilter.Ref, qt.Equals, "B2:C3")
	})

}

func TestMakeXLSXSheet(t *testing.T) {
	c := qt.New(t)

	// Test if the NumFmts assigned properly according the FormatCode in cell.
	csRunO(c, "SheetWithNumFormats", func(c *qt.C, option FileOption) {
		file := NewFile(option)
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

		var buf bytes.Buffer

		refTable := NewSharedStringRefTable()
		styles := newXlsxStyleSheet(nil)
		err := sheet.MarshalSheet(&buf, refTable, styles, nil)
		c.Assert(err, qt.Equals, nil)
		var xSheet xlsxWorksheet
		err = xml.Unmarshal(buf.Bytes(), &xSheet)
		c.Assert(err, qt.Equals, nil)

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
		c.Assert(xSheet.SheetData.Row[0].C[0].S, qt.Equals, 0)
		c.Assert(xSheet.SheetData.Row[0].C[1].S, qt.Equals, 1)
	})

	// When we create the xlsxSheet we also populate the xlsxStyles struct
	// with style information.
	csRunO(c, "PopulateXLSXSTyles", func(c *qt.C, option FileOption) {
		file := NewFile(option)
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
		var buf bytes.Buffer

		refTable := NewSharedStringRefTable()
		styles := newXlsxStyleSheet(nil)
		err := sheet.MarshalSheet(&buf, refTable, styles, nil)
		c.Assert(err, qt.Equals, nil)
		var xSheet xlsxWorksheet
		err = xml.Unmarshal(buf.Bytes(), &xSheet)
		c.Assert(err, qt.Equals, nil)

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
		c.Assert(xSheet.SheetData.Row[0].C[0].S, qt.Equals, 0)
		c.Assert(xSheet.SheetData.Row[0].C[1].S, qt.Equals, 0)
	})

	// If the column width is customised, the xslxCol.CustomWidth field is set to 1.
	csRunO(c, "SetCustomColWidth", func(c *qt.C, option FileOption) {
		file := NewFile(option)
		sheet, _ := file.AddSheet("Sheet1")
		row := sheet.AddRow()
		cell1 := row.AddCell()
		cell1.Value = "A cell!"
		sheet.SetColWidth(1, 1, 10.5)
		var buf bytes.Buffer

		refTable := NewSharedStringRefTable()
		styles := newXlsxStyleSheet(nil)
		err := sheet.MarshalSheet(&buf, refTable, styles, nil)
		c.Assert(err, qt.Equals, nil)
		var xSheet xlsxWorksheet
		err = xml.Unmarshal(buf.Bytes(), &xSheet)
		c.Assert(err, qt.Equals, nil)
		c.Assert(*xSheet.Cols.Col[0].CustomWidth, qt.Equals, true)
	})

	csRunO(c, "SetColWidth", func(c *qt.C, option FileOption) {
		file := NewFile(option)
		sheet, _ := file.AddSheet("Sheet1")
		sheet.SetColWidth(1, 1, 10.5)
		sheet.SetColWidth(2, 6, 11)
		c.Assert(*sheet.Cols.FindColByIndex(1).Width, qt.Equals, 10.5)
		c.Assert(sheet.Cols.FindColByIndex(1).Max, qt.Equals, 1)
		c.Assert(sheet.Cols.FindColByIndex(1).Min, qt.Equals, 1)
		c.Assert(*sheet.Cols.FindColByIndex(2).Width, qt.Equals, float64(11))
		c.Assert(sheet.Cols.FindColByIndex(2).Max, qt.Equals, 6)
		c.Assert(sheet.Cols.FindColByIndex(2).Min, qt.Equals, 2)
	})

	csRunO(c, "SetDataValidation", func(c *qt.C, option FileOption) {
		file := NewFile(option)
		sheet, _ := file.AddSheet("Sheet1")

		dd := NewDataValidation(0, 0, 10, 0, true)
		err := dd.SetDropList([]string{"a1", "a2", "a3"})
		c.Assert(err, qt.IsNil)

		sheet.AddDataValidation(dd)
		c.Assert(sheet.DataValidations, qt.HasLen, 1)
		c.Assert(sheet.DataValidations[0], qt.Equals, dd)
	})

	csRunO(c, "Border", func(c *qt.C, option FileOption) {
		file := NewFile(option)
		sheet, _ := file.AddSheet("Sheet1")
		row := sheet.AddRow()

		cell1 := row.AddCell()
		cell1.Value = "A cell!"
		style1 := NewStyle()
		style1.Border = *NewBorder("thin", "thin", "thin", "thin")
		style1.ApplyBorder = true
		cell1.SetStyle(style1)
		var buf bytes.Buffer

		refTable := NewSharedStringRefTable()
		styles := newXlsxStyleSheet(nil)
		err := sheet.MarshalSheet(&buf, refTable, styles, nil)
		c.Assert(err, qt.Equals, nil)
		var xSheet xlsxWorksheet
		err = xml.Unmarshal(buf.Bytes(), &xSheet)
		c.Assert(err, qt.Equals, nil)

		c.Assert(styles.Borders.Border[0].Left.Style, qt.Equals, "thin")
		c.Assert(styles.Borders.Border[0].Right.Style, qt.Equals, "thin")
		c.Assert(styles.Borders.Border[0].Top.Style, qt.Equals, "thin")
		c.Assert(styles.Borders.Border[0].Bottom.Style, qt.Equals, "thin")

		c.Assert(xSheet.SheetData.Row[0].C[0].S, qt.Equals, 0)
	})

	csRunO(c, "OutlineLevels", func(c *qt.C, option FileOption) {
		file := NewFile(option)
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
		r1.SetOutlineLevel(1)
		r2.SetOutlineLevel(2)
		sheet.SetOutlineLevel(1, 1, 1)

		var buf bytes.Buffer

		refTable := NewSharedStringRefTable()
		styles := newXlsxStyleSheet(nil)
		err := sheet.MarshalSheet(&buf, refTable, styles, nil)
		c.Assert(err, qt.Equals, nil)
		var xSheet xlsxWorksheet
		err = xml.Unmarshal(buf.Bytes(), &xSheet)
		c.Assert(err, qt.Equals, nil)

		c.Assert(xSheet.SheetFormatPr.OutlineLevelCol, qt.Equals, uint8(1))
		c.Assert(xSheet.SheetFormatPr.OutlineLevelRow, qt.Equals, uint8(2))

		c.Assert(*xSheet.Cols.Col[0].OutlineLevel, qt.Equals, uint8(1))
		c.Assert(xSheet.SheetData.Row[0].OutlineLevel, qt.Equals, uint8(1))
		c.Assert(xSheet.SheetData.Row[1].OutlineLevel, qt.Equals, uint8(2))
		c.Assert(xSheet.SheetData.Row[2].OutlineLevel, qt.Equals, uint8(0))
	})
}
