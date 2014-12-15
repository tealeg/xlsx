package xlsx

import (
	"fmt"
	"strconv"
)

// Sheet is a high level structure intended to provide user access to
// the contents of a particular sheet within an XLSX file.
type Sheet struct {
	Name   string
	File   File
	Rows   []*Row
	Cols   []*Col
	MaxRow int
	MaxCol int
	Hidden bool
}

// Add a new Row to a Sheet
func (s *Sheet) AddRow() *Row {
	row := &Row{Sheet: s}
	s.Rows = append(s.Rows, row)
	if len(s.Rows) > s.MaxRow {
		s.MaxRow = len(s.Rows)
	}
	return row
}

// Make sure we always have as many Cols as we do cells.
func (s *Sheet) maybeAddCol(cellCount int) {
	if cellCount > s.MaxCol {
		col := &Col{
			Min:       cellCount,
			Max:       cellCount,
			Hidden:    false,
			Collapsed: false,
			// Style:     0,
			Width: ColWidth}
		s.Cols = append(s.Cols, col)
		s.MaxCol = cellCount
	}
}

// Get a Cell by passing it's cartesian coordinates (zero based) as
// row and column integer indexes.
//
// For example:
//
//    cell := sheet.Cell(0,0)
//
// ... would set the variable "cell" to contain a Cell struct
// containing the data from the field "A1" on the spreadsheet.
func (sh *Sheet) Cell(row, col int) *Cell {

	if len(sh.Rows) > row && sh.Rows[row] != nil && len(sh.Rows[row].Cells) > col {
		return sh.Rows[row].Cells[col]
	}
	return new(Cell)
}

// Dump sheet to it's XML representation, intended for internal use only
func (s *Sheet) makeXLSXSheet(refTable *RefTable, styles *xlsxStyleSheet) *xlsxWorksheet {
	worksheet := newXlsxWorksheet()
	xSheet := xlsxSheetData{}
	maxRow := 0
	maxCell := 0
	for r, row := range s.Rows {
		if r > maxRow {
			maxRow = r
		}
		xRow := xlsxRow{}
		xRow.R = r + 1
		for c, cell := range row.Cells {
			style := cell.GetStyle()
			xNumFmt, xFont, xFill, xBorder, xCellStyleXf, xCellXf := style.makeXLSXStyleElements()
			fontId := styles.addFont(xFont)
			fillId := styles.addFill(xFill)
			borderId := styles.addBorder(xBorder)
			styles.addNumFmt(xNumFmt)
			xCellStyleXf.FontId = fontId
			xCellStyleXf.FillId = fillId
			xCellStyleXf.BorderId = borderId
			xCellStyleXf.NumFmtId = xNumFmt.NumFmtId
			xCellXf.FontId = fontId
			xCellXf.FillId = fillId
			xCellXf.BorderId = borderId
			xCellXf.NumFmtId = xNumFmt.NumFmtId
			styles.addCellStyleXf(xCellStyleXf)
			XfId := styles.addCellXf(xCellXf)
			if c > maxCell {
				maxCell = c
			}
			xC := xlsxC{}
			xC.R = fmt.Sprintf("%s%d", numericToLetters(c), r+1)
			xC.V = strconv.Itoa(refTable.AddString(cell.Value))
			xC.T = "s" // Hardcode string type, for now.
			xC.S = XfId
			xRow.C = append(xRow.C, xC)
		}
		xSheet.Row = append(xSheet.Row, xRow)
	}

	worksheet.Cols = xlsxCols{Col: []xlsxCol{}}
	for _, col := range s.Cols {
		if col.Width == 0 {
			col.Width = ColWidth
		}
		worksheet.Cols.Col = append(worksheet.Cols.Col,
			xlsxCol{Min: col.Min,
				Max:       col.Max,
				Hidden:    col.Hidden,
				Width:     col.Width,
				Collapsed: col.Collapsed,
				// Style:     col.Style
			})
	}
	worksheet.SheetData = xSheet
	dimension := xlsxDimension{}
	dimension.Ref = fmt.Sprintf("A1:%s%d",
		numericToLetters(maxCell), maxRow+1)
	if dimension.Ref == "A1:A1" {
		dimension.Ref = "A1"
	}
	worksheet.Dimension = dimension
	return worksheet
}
