package xlsx

import (
	"fmt"
	"strconv"
)

// Sheet is a high level structure intended to provide user access to
// the contents of a particular sheet within an XLSX file.
type Sheet struct {
	Name   string
	Rows   []*Row
	MaxRow int
	MaxCol int
}

// Add a new Row to a Sheet
func (s *Sheet) AddRow() *Row {
	row := &Row{}
	s.Rows = append(s.Rows, row)
	if len(s.Rows) > s.MaxRow {
		s.MaxRow = len(s.Rows)
	}
	return row
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
func (s *Sheet) makeXLSXSheet(refTable *RefTable, styles *xlsxStyles) *xlsxWorksheet {
	worksheet := &xlsxWorksheet{}
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
			xFont, xFill, xBorder, xCellStyleXf, xCellXf := style.makeXLSXStyleElements()
			fontId := styles.addFont(xFont)
			fillId := styles.addFill(xFill)
			borderId := styles.addBorder(xBorder)
			xCellStyleXf.FontId = fontId
			xCellStyleXf.FillId = fillId
			xCellStyleXf.BorderId = borderId
			xCellXf.FontId = fontId
			xCellXf.FillId = fillId
			xCellXf.BorderId = borderId
			styles.addCellStyleXf(xCellStyleXf)
			styles.addCellXf(xCellXf)
			if c > maxCell {
				maxCell = c
			}
			xC := xlsxC{}
			xC.R = fmt.Sprintf("%s%d", numericToLetters(c), r+1)
			xC.V = strconv.Itoa(refTable.AddString(cell.Value))
			xC.T = "s" // Hardcode string type, for now.
			xRow.C = append(xRow.C, xC)
		}
		xSheet.Row = append(xSheet.Row, xRow)
	}
	worksheet.SheetData = xSheet
	dimension := xlsxDimension{}
	dimension.Ref = fmt.Sprintf("A1:%s%d",
		numericToLetters(maxCell), maxRow+1)
	worksheet.Dimension = dimension
	return worksheet
}
