package xlsx

import (
	"encoding/xml"
	"fmt"
)

// Sheet is a high level structure intended to provide user access to
// the contents of a particular sheet within an XLSX file.
type Sheet struct {
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
func (s *Sheet) makeXLSXSheet() ([]byte, error) {
	xSheet := xlsxSheetData{}
	for r, row := range s.Rows {
		xRow := xlsxRow{}
		xRow.R = r + 1
		for c, cell := range row.Cells {
			xC := xlsxC{}
			xC.R = fmt.Sprintf("%s%d", numericToLetters(c), r + 1)
			xC.V = cell.Value
			xC.T = "s" // Hardcode string type, for now.
			xRow.C = append(xRow.C, xC)
		}
		xSheet.Row = append(xSheet.Row, xRow)
	}
	return xml.MarshalIndent(xSheet, "  ", "  ")
}
