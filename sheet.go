package xlsx

import (
	"fmt"
	"strconv"
)

// Sheet is a high level structure intended to provide user access to
// the contents of a particular sheet within an XLSX file.
type Sheet struct {
	Name       string
	File       *File
	Rows       []*Row
	Cols       []*Col
	MaxRow     int
	MaxCol     int
	Hidden     bool
	SheetViews []SheetView
}

type SheetView struct {
	Pane *Pane
}

type Pane struct {
	XSplit      int
	YSplit      int
	TopLeftCell string
	ActivePane  string
	State       string // Either "split" or "frozen"
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

//Set the width of a single column or multipel columns.
func (s *Sheet) SetColWidth(startcol, endcol int, width float64) error {
	if startcol > endcol {
		return fmt.Errorf("Could not set width for range %d-%d: startcol must be less than endcol.", startcol, endcol)
	}
	col := &Col{
		Min:       startcol + 1,
		Max:       endcol + 1,
		Hidden:    false,
		Collapsed: false,
		// Style:     0,
		Width: width}
	s.Cols = append(s.Cols, col)
	if endcol+1 > s.MaxCol {
		s.MaxCol = endcol + 1
	}
	return nil
}

// Dump sheet to it's XML representation, intended for internal use only
func (s *Sheet) makeXLSXSheet(refTable *RefTable, styles *xlsxStyleSheet) *xlsxWorksheet {
	worksheet := newXlsxWorksheet()
	xSheet := xlsxSheetData{}
	maxRow := 0
	maxCell := 0
	XfId := 0
	for r, row := range s.Rows {
		if r > maxRow {
			maxRow = r
		}
		xRow := xlsxRow{}
		xRow.R = r + 1
		for c, cell := range row.Cells {
			style := cell.GetStyle()
			if style != nil {
				xFont, xFill, xBorder, xCellStyleXf, xCellXf := style.makeXLSXStyleElements()
				fontId := styles.addFont(xFont)
				fillId := styles.addFill(xFill)
				borderId := styles.addBorder(xBorder)
				xCellStyleXf.FontId = fontId
				xCellStyleXf.FillId = fillId
				xCellStyleXf.BorderId = borderId
				xCellStyleXf.NumFmtId = 0 // General
				xCellXf.FontId = fontId
				xCellXf.FillId = fillId
				xCellXf.BorderId = borderId
				xCellXf.NumFmtId = 0 // General
				styles.addCellStyleXf(xCellStyleXf)
				XfId = styles.addCellXf(xCellXf)
			}
			if c > maxCell {
				maxCell = c
			}
			xC := xlsxC{}
			xC.R = fmt.Sprintf("%s%d", numericToLetters(c), r+1)
			switch cell.cellType {
			case CellTypeString:
				xC.V = strconv.Itoa(refTable.AddString(cell.Value))
				xC.T = "s"
				xC.S = XfId
			case CellTypeBool:
				xC.V = cell.Value
				xC.T = "b"
				xC.S = XfId
			case CellTypeNumeric:
				xC.V = cell.Value
				xC.S = XfId
			case CellTypeFormula:
				xC.V = cell.Value
				xC.F = &xlsxF{Content: cell.formula}
				xC.S = XfId
			case CellTypeError:
				xC.V = cell.Value
				xC.F = &xlsxF{Content: cell.formula}
				xC.T = "e"
				xC.S = XfId
			}
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
