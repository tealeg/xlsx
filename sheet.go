package xlsx

import (
	"fmt"
	"strconv"
)

// Sheet is a high level structure intended to provide user access to
// the contents of a particular sheet within an XLSX file.
type Sheet struct {
	Name        string
	File        *File
	Rows        []*Row
	Cols        []*Col
	MaxRow      int
	MaxCol      int
	Hidden      bool
	Selected    bool
	SheetViews  []SheetView
	SheetFormat SheetFormat
}

type SheetView struct {
	Pane *Pane
}

type Pane struct {
	XSplit      float64
	YSplit      float64
	TopLeftCell string
	ActivePane  string
	State       string // Either "split" or "frozen"
}

type SheetFormat struct {
	DefaultColWidth  float64
	DefaultRowHeight float64
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
			style:     NewStyle(),
			Min:       cellCount,
			Max:       cellCount,
			Hidden:    false,
			Collapsed: false}
		s.Cols = append(s.Cols, col)
		s.MaxCol = cellCount
	}
}

// Make sure we always have as many Cols as we do cells.
func (s *Sheet) Col(idx int) *Col {
	s.maybeAddCol(idx + 1)
	return s.Cols[idx]
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

	// If the user requests a row beyond what we have, then extend.
	for len(sh.Rows) <= row {
		sh.AddRow()
	}

	r := sh.Rows[row]
	for len(r.Cells) <= col {
		r.AddCell()
	}

	return r.Cells[col]
}

//Set the width of a single column or multiple columns.
func (s *Sheet) SetColWidth(startcol, endcol int, width float64) error {
	if startcol > endcol {
		return fmt.Errorf("Could not set width for range %d-%d: startcol must be less than endcol.", startcol, endcol)
	}
	col := &Col{
		style:     NewStyle(),
		Min:       startcol + 1,
		Max:       endcol + 1,
		Hidden:    false,
		Collapsed: false,
		Width:     width}
	s.Cols = append(s.Cols, col)
	if endcol+1 > s.MaxCol {
		s.MaxCol = endcol + 1
	}
	return nil
}

// When merging cells, the cell may be the 'original' or the 'covered'.
// First, figure out which cells are merge starting points. Then create
// the necessary cells underlying the merge area.
// Then go through all the underlying cells and apply the appropriate
// border, based on the original cell.
func (s *Sheet) handleMerged() {
	merged := make(map[string]*Cell)

	for r, row := range s.Rows {
		for c, cell := range row.Cells {
			if cell.HMerge > 0 || cell.VMerge > 0 {
				coord := fmt.Sprintf("%s%d", numericToLetters(c), r+1)
				merged[coord] = cell
			}
		}
	}

	// This loop iterates over all cells that should be merged and applies the correct
	// borders to them depending on their position. If any cells required by the merge
	// are missing, they will be allocated by s.Cell().
	for key, cell := range merged {
		mainstyle := cell.GetStyle()

		top := mainstyle.Border.Top
		left := mainstyle.Border.Left
		right := mainstyle.Border.Right
		bottom := mainstyle.Border.Bottom

		// When merging cells, the upper left cell does not maintain
		// the original borders
		mainstyle.Border.Top = ""
		mainstyle.Border.Left = ""
		mainstyle.Border.Right = ""
		mainstyle.Border.Bottom = ""

		maincol, mainrow, _ := getCoordsFromCellIDString(key)
		for rownum := 0; rownum <= cell.VMerge; rownum++ {
			for colnum := 0; colnum <= cell.HMerge; colnum++ {
				tmpcell := s.Cell(mainrow+rownum, maincol+colnum)
				style := tmpcell.GetStyle()
				style.ApplyBorder = true

				if rownum == 0 {
					style.Border.Top = top
				}

				if rownum == (cell.VMerge) {
					style.Border.Bottom = bottom
				}

				if colnum == 0 {
					style.Border.Left = left
				}

				if colnum == (cell.HMerge) {
					style.Border.Right = right
				}
			}
		}
	}
}

// Dump sheet to its XML representation, intended for internal use only
func (s *Sheet) makeXLSXSheet(refTable *RefTable, styles *xlsxStyleSheet) *xlsxWorksheet {
	worksheet := newXlsxWorksheet()
	xSheet := xlsxSheetData{}
	maxRow := 0
	maxCell := 0

	// Scan through the sheet and see if there are any merged cells. If there
	// are, we may need to extend the size of the sheet. There needs to be
	// phantom cells underlying the area covered by the merged cell
	s.handleMerged()

	if s.Selected {
		worksheet.SheetViews.SheetView[0].TabSelected = true
	}

	if s.SheetFormat.DefaultRowHeight != 0 {
		worksheet.SheetFormatPr.DefaultRowHeight = s.SheetFormat.DefaultRowHeight
	}
	worksheet.SheetFormatPr.DefaultColWidth = s.SheetFormat.DefaultColWidth

	colsXfIdList := make([]int, len(s.Cols))
	worksheet.Cols = &xlsxCols{Col: []xlsxCol{}}
	for c, col := range s.Cols {
		XfId := 0
		if col.Min == 0 {
			col.Min = 1
		}
		if col.Max == 0 {
			col.Max = 1
		}
		style := col.GetStyle()
		//col's style always not nil
		if style != nil {
			xNumFmt := styles.newNumFmt(col.numFmt)
			XfId = handleStyleForXLSX(style, xNumFmt.NumFmtId, styles)
		}
		colsXfIdList[c] = XfId

		var customWidth int
		if col.Width == 0 {
			col.Width = ColWidth
		} else {
			customWidth = 1
		}
		worksheet.Cols.Col = append(worksheet.Cols.Col,
			xlsxCol{Min: col.Min,
				Max:         col.Max,
				Hidden:      col.Hidden,
				Width:       col.Width,
				CustomWidth: customWidth,
				Collapsed:   col.Collapsed,
				Style:       XfId,
			})
	}

	for r, row := range s.Rows {
		if r > maxRow {
			maxRow = r
		}
		xRow := xlsxRow{}
		xRow.R = r + 1
		if row.isCustom {
			xRow.CustomHeight = true
			xRow.Ht = fmt.Sprintf("%g", row.Height)
		}
		for c, cell := range row.Cells {
			XfId := colsXfIdList[c]

			// generate NumFmtId and add new NumFmt
			xNumFmt := styles.newNumFmt(cell.NumFmt)

			style := cell.style
			if style != nil {
				XfId = handleStyleForXLSX(style, xNumFmt.NumFmtId, styles)
			} else if len(cell.NumFmt) > 0 && s.Cols[c].numFmt != cell.NumFmt {
				XfId = handleNumFmtIdForXLSX(xNumFmt.NumFmtId, styles)
			}

			if c > maxCell {
				maxCell = c
			}
			xC := xlsxC{}
			xC.R = fmt.Sprintf("%s%d", numericToLetters(c), r+1)
			switch cell.cellType {
			case CellTypeString:
				if len(cell.Value) > 0 {
					xC.V = strconv.Itoa(refTable.AddString(cell.Value))
				}
				xC.T = "s"
				xC.S = XfId
			case CellTypeBool:
				xC.V = cell.Value
				xC.T = "b"
				xC.S = XfId
			case CellTypeNumeric:
				xC.V = cell.Value
				xC.S = XfId
			case CellTypeDate:
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
			case CellTypeGeneral:
				xC.V = cell.Value
				xC.S = XfId
			}

			xRow.C = append(xRow.C, xC)

			if cell.HMerge > 0 || cell.VMerge > 0 {
				// r == rownum, c == colnum
				mc := xlsxMergeCell{}
				start := fmt.Sprintf("%s%d", numericToLetters(c), r+1)
				endcol := c + cell.HMerge
				endrow := r + cell.VMerge + 1
				end := fmt.Sprintf("%s%d", numericToLetters(endcol), endrow)
				mc.Ref = start + ":" + end
				if worksheet.MergeCells == nil {
					worksheet.MergeCells = &xlsxMergeCells{}
				}
				worksheet.MergeCells.Cells = append(worksheet.MergeCells.Cells, mc)
			}
		}
		xSheet.Row = append(xSheet.Row, xRow)
	}

	if worksheet.MergeCells != nil {
		worksheet.MergeCells.Count = len(worksheet.MergeCells.Cells)
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

func handleStyleForXLSX(style *Style, NumFmtId int, styles *xlsxStyleSheet) (XfId int) {
	xFont, xFill, xBorder, xCellXf := style.makeXLSXStyleElements()
	fontId := styles.addFont(xFont)
	fillId := styles.addFill(xFill)

	// HACK - adding light grey fill, as in OO and Google
	greyfill := xlsxFill{}
	greyfill.PatternFill.PatternType = "lightGray"
	styles.addFill(greyfill)

	borderId := styles.addBorder(xBorder)
	xCellXf.FontId = fontId
	xCellXf.FillId = fillId
	xCellXf.BorderId = borderId
	xCellXf.NumFmtId = NumFmtId
	// apply the numFmtId when it is not the default cellxf
	if xCellXf.NumFmtId > 0 {
		xCellXf.ApplyNumberFormat = true
	}

	xCellXf.Alignment.Horizontal = style.Alignment.Horizontal
	xCellXf.Alignment.Indent = style.Alignment.Indent
	xCellXf.Alignment.ShrinkToFit = style.Alignment.ShrinkToFit
	xCellXf.Alignment.TextRotation = style.Alignment.TextRotation
	xCellXf.Alignment.Vertical = style.Alignment.Vertical
	xCellXf.Alignment.WrapText = style.Alignment.WrapText

	XfId = styles.addCellXf(xCellXf)
	return
}

func handleNumFmtIdForXLSX(NumFmtId int, styles *xlsxStyleSheet) (XfId int) {
	xCellXf := makeXLSXCellElement()
	xCellXf.NumFmtId = NumFmtId
	if xCellXf.NumFmtId > 0 {
		xCellXf.ApplyNumberFormat = true
	}
	XfId = styles.addCellXf(xCellXf)
	return
}
