package xlsx

import (
	"fmt"
	"strconv"
)

// Sheet is a high level structure intended to provide user access to
// the contents of a particular sheet within an XLSX file.
type Sheet struct {
	Name	string
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

// Dump sheet to it's XML representation
func (s *Sheet) makeXLSXSheet(refTable *RefTable) *xlsxWorksheet {
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
			if c > maxCell {
				maxCell = c
			}
			xC := xlsxC{}
			xC.R = fmt.Sprintf("%s%d", numericToLetters(c), r + 1)
			xC.V = strconv.Itoa(refTable.AddString(cell.Value))
			xC.T = "s" // Hardcode string type, for now.
			xRow.C = append(xRow.C, xC)
		}
		xSheet.Row = append(xSheet.Row, xRow)
	}
	worksheet.SheetData = xSheet
	dimension := xlsxDimension{}
	dimension.Ref = fmt.Sprintf("A1:%s%d",
		numericToLetters(maxCell), maxRow + 1)
	worksheet.Dimension = dimension
	return worksheet
}
