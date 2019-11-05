package xlsx

import (
	"encoding/xml"
	"errors"
	"fmt"
	"strconv"
)

// Sheet is a high level structure intended to provide user access to
// the contents of a particular sheet within an XLSX file.
type Sheet struct {
	Name            string
	File            *File
	Rows            []*Row
	Cols            *ColStore
	MaxRow          int
	MaxCol          int
	Hidden          bool
	Selected        bool
	SheetViews      []SheetView
	SheetFormat     SheetFormat
	AutoFilter      *AutoFilter
	Relations       []Relation
	DataValidations []*xlsxDataValidation
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
	OutlineLevelCol  uint8
	OutlineLevelRow  uint8
}

type AutoFilter struct {
	TopLeftCell     string
	BottomRightCell string
}

type Relation struct {
	Type       RelationshipType
	Target     string
	TargetMode RelationshipTargetMode
}

func (s *Sheet) makeXLSXSheetRelations() *xlsxWorksheetRels {
	relSheet := xlsxWorksheetRels{XMLName: xml.Name{Local: "Relationships"}, Relationships: []xlsxWorksheetRelation{}}
	for id, rel := range s.Relations {
		xRel := xlsxWorksheetRelation{Id: "rId" + strconv.Itoa(id+1), Type: rel.Type, Target: rel.Target, TargetMode: rel.TargetMode}
		relSheet.Relationships = append(relSheet.Relationships, xRel)
	}
	if len(relSheet.Relationships) == 0 {
		return nil
	}
	return &relSheet
}

func (s *Sheet) addRelation(relType RelationshipType, target string, targetMode RelationshipTargetMode) {
	newRel := Relation{Type: relType, Target: target, TargetMode: targetMode}
	for _, rel := range s.Relations {
		if rel == newRel {
			return
		}
	}
	s.Relations = append(s.Relations, newRel)
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

// Add a new Row to a Sheet at a specific index
func (s *Sheet) AddRowAtIndex(index int) (*Row, error) {
	if index < 0 || index > len(s.Rows) {
		return nil, errors.New("AddRowAtIndex: index out of bounds")
	}
	row := &Row{Sheet: s}
	s.Rows = append(s.Rows, nil)

	if index < len(s.Rows) {
		copy(s.Rows[index+1:], s.Rows[index:])
	}
	s.Rows[index] = row
	if len(s.Rows) > s.MaxRow {
		s.MaxRow = len(s.Rows)
	}
	return row, nil
}

// Add a DataValidation to a range of cells
func (s *Sheet) AddDataValidation(dv *xlsxDataValidation) {
	s.DataValidations = append(s.DataValidations, dv)
}

// Removes a row at a specific index
func (s *Sheet) RemoveRowAtIndex(index int) error {
	if index < 0 || index >= len(s.Rows) {
		return errors.New("RemoveRowAtIndex: index out of bounds")
	}
	s.Rows = append(s.Rows[:index], s.Rows[index+1:]...)
	return nil
}

// Make sure we always have as many Rows as we do cells.
func (s *Sheet) maybeAddRow(rowCount int) {
	if rowCount > s.MaxRow {
		loopCnt := rowCount - s.MaxRow
		for i := 0; i < loopCnt; i++ {

			row := &Row{Sheet: s}
			s.Rows = append(s.Rows, row)
		}
		s.MaxRow = rowCount
	}
}

// Make sure we always have as many Rows as we do cells.
func (s *Sheet) Row(idx int) *Row {
	s.maybeAddRow(idx + 1)
	return s.Rows[idx]
}

// Return the Col that applies to this Column index, or return nil if no such Col exists
func (s *Sheet) Col(idx int) *Col {
	if s.Cols == nil {
		panic("trying to use uninitialised ColStore")
	}
	return s.Cols.FindColByIndex(idx + 1)
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
func (s *Sheet) Cell(row, col int) *Cell {

	// If the user requests a row beyond what we have, then extend.
	for len(s.Rows) <= row {
		s.AddRow()
	}

	r := s.Rows[row]
	for len(r.Cells) <= col {
		r.AddCell()
	}

	return r.Cells[col]
}

//Set the parameters of a column.  Parameters are passed as a pointer
//to a Col structure which you much construct yourself.
func (s *Sheet) SetColParameters(col *Col) {
	if s.Cols == nil {
		panic("trying to use uninitialised ColStore")
	}
	s.Cols.Add(col)
}

func (s *Sheet) setCol(min, max int, setter func(col *Col)) {
	if s.Cols == nil {
		panic("trying to use uninitialised ColStore")
	}

	cols := s.Cols.getOrMakeColsForRange(s.Cols.Root, min, max)

	for _, col := range cols {
		switch {
		case col.Min < min && col.Max > max:
			// The column completely envelops the range,
			// so we'll split it into three parts and only
			// set the width on the part within the range.
			// The ColStore will do most of this work for
			// us, we just need to create the new Col
			// based on the old one.
			newCol := col.copyToRange(min, max)
			setter(newCol)
			s.Cols.Add(newCol)
		case col.Min < min:
			// If this column crosses the minimum boundary
			// of the range we must split it and only
			// apply the change within the range.  Again,
			// we can lean on the ColStore to deal with
			// the rest we just need to make the new
			// Col.
			newCol := col.copyToRange(min, col.Max)
			setter(newCol)
			s.Cols.Add(newCol)
		case col.Max > max:
			// Likewise if a col definition crosses the
			// maximum boundary of the range, it must also
			// be split
			newCol := col.copyToRange(col.Min, max)
			setter(newCol)
			s.Cols.Add(newCol)
		default:
			newCol := col.copyToRange(min, max)
			setter(newCol)
			s.Cols.Add(newCol)

		}
	}
	return
}

// Set the width of a range of columns.
func (s *Sheet) SetColWidth(min, max int, width float64) {
	s.setCol(min, max, func(col *Col) {
		col.SetWidth(width)
	})
}

// Set the outline level for a range of columns.
func (s *Sheet) SetOutlineLevel(minCol, maxCol int, outlineLevel uint8) {
	s.setCol(minCol, maxCol, func(col *Col) {
		col.SetOutlineLevel(outlineLevel)
	})
}

// Set the type for a range of columns.
func (s *Sheet) SetType(minCol, maxCol int, cellType CellType) {
	s.setCol(minCol, maxCol, func(col *Col) {
		col.SetType(cellType)
	})

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
				coord := GetCellIDStringFromCoords(c, r)
				merged[coord] = cell
			}
		}
	}

	// This loop iterates over all cells that should be merged and applies the correct
	// borders to them depending on their position. If any cells required by the merge
	// are missing, they will be allocated by s.Cell().
	for key, cell := range merged {

		maincol, mainrow, _ := GetCoordsFromCellIDString(key)
		for rownum := 0; rownum <= cell.VMerge; rownum++ {
			for colnum := 0; colnum <= cell.HMerge; colnum++ {
				// make cell
				s.Cell(mainrow+rownum, maincol+colnum)

			}
		}
	}
}

func (s *Sheet) makeSheetView(worksheet *xlsxWorksheet) {
	for index, sheetView := range s.SheetViews {
		if sheetView.Pane != nil {
			worksheet.SheetViews.SheetView[index].Pane = &xlsxPane{
				XSplit:      sheetView.Pane.XSplit,
				YSplit:      sheetView.Pane.YSplit,
				TopLeftCell: sheetView.Pane.TopLeftCell,
				ActivePane:  sheetView.Pane.ActivePane,
				State:       sheetView.Pane.State,
			}

		}
	}
	if s.Selected {
		worksheet.SheetViews.SheetView[0].TabSelected = true
	}

}

func (s *Sheet) makeSheetFormatPr(worksheet *xlsxWorksheet) {
	if s.SheetFormat.DefaultRowHeight != 0 {
		worksheet.SheetFormatPr.DefaultRowHeight = s.SheetFormat.DefaultRowHeight
	}
	worksheet.SheetFormatPr.DefaultColWidth = s.SheetFormat.DefaultColWidth
}

//
func (s *Sheet) makeCols(worksheet *xlsxWorksheet, styles *xlsxStyleSheet) (maxLevelCol uint8) {
	maxLevelCol = 0
	if s.Cols == nil {
		panic("trying to use uninitialised ColStore")
	}
	s.Cols.ForEach(
		func(c int, col *Col) {
			XfId := 0
			style := col.GetStyle()

			hasNumFmt := len(col.numFmt) > 0
			if style == nil && hasNumFmt {
				style = NewStyle()
			}

			if hasNumFmt {
				xNumFmt := styles.newNumFmt(col.numFmt)
				XfId = handleStyleForXLSX(style, xNumFmt.NumFmtId, styles)
			}
			col.outXfID = XfId

			// When the cols content is empty, the cols flag is not output in the xml file.
			if worksheet.Cols == nil {
				worksheet.Cols = &xlsxCols{Col: []xlsxCol{}}
			}
			worksheet.Cols.Col = append(worksheet.Cols.Col,
				xlsxCol{
					Min:          col.Min,
					Max:          col.Max,
					Hidden:       col.Hidden,
					Width:        col.Width,
					CustomWidth:  col.CustomWidth,
					Collapsed:    col.Collapsed,
					OutlineLevel: col.OutlineLevel,
					Style:        XfId,
					BestFit:      col.BestFit,
					Phonetic:     col.Phonetic,
				})

			if col.OutlineLevel > maxLevelCol {
				maxLevelCol = col.OutlineLevel
			}
		})

	return maxLevelCol
}

func (s *Sheet) makeRows(worksheet *xlsxWorksheet, styles *xlsxStyleSheet, refTable *RefTable, relations *xlsxWorksheetRels, maxLevelCol uint8) {
	maxRow := 0
	maxCell := 0
	var maxLevelRow uint8
	xSheet := xlsxSheetData{}

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
		xRow.OutlineLevel = row.OutlineLevel
		if row.OutlineLevel > maxLevelRow {
			maxLevelRow = row.OutlineLevel
		}
		for c, cell := range row.Cells {
			var XfId int

			col := s.Col(c)
			if col != nil {
				XfId = col.outXfID
			}

			// generate NumFmtId and add new NumFmt
			xNumFmt := styles.newNumFmt(cell.NumFmt)

			style := cell.style
			switch {
			case style != nil:
				XfId = handleStyleForXLSX(style, xNumFmt.NumFmtId, styles)
			case len(cell.NumFmt) == 0:
				// Do nothing
			case col == nil:
				XfId = handleNumFmtIdForXLSX(xNumFmt.NumFmtId, styles)
			case !compareFormatString(col.numFmt, cell.NumFmt):
				XfId = handleNumFmtIdForXLSX(xNumFmt.NumFmtId, styles)
			}

			if c > maxCell {
				maxCell = c
			}
			xC := xlsxC{
				S: XfId,
				R: GetCellIDStringFromCoords(c, r),
			}
			if cell.formula != "" {
				xC.F = &xlsxF{Content: cell.formula}
			}
			switch cell.cellType {
			case CellTypeInline:
				// Inline strings are turned into shared strings since they are more efficient.
				// This is what Excel does as well.
				fallthrough
			case CellTypeString:
				if len(cell.Value) > 0 {
					xC.V = strconv.Itoa(refTable.AddString(cell.Value))
				}
				xC.T = "s"
			case CellTypeNumeric:
				// Numeric is the default, so the type can be left blank
				xC.V = cell.Value
			case CellTypeBool:
				xC.V = cell.Value
				xC.T = "b"
			case CellTypeError:
				xC.V = cell.Value
				xC.T = "e"
			case CellTypeDate:
				xC.V = cell.Value
				xC.T = "d"
			case CellTypeStringFormula:
				xC.V = cell.Value
				xC.T = "str"
			default:
				panic(errors.New("unknown cell type cannot be marshaled"))
			}

			xRow.C = append(xRow.C, xC)
			if nil != cell.DataValidation {
				if nil == worksheet.DataValidations {
					worksheet.DataValidations = &xlsxDataValidations{}
				}
				cell.DataValidation.Sqref = xC.R
				worksheet.DataValidations.DataValidation = append(worksheet.DataValidations.DataValidation, cell.DataValidation)
				worksheet.DataValidations.Count = len(worksheet.DataValidations.DataValidation)
			}

			if cell.Hyperlink != (Hyperlink{}) {
				if worksheet.Hyperlinks == nil {
					worksheet.Hyperlinks = &xlsxHyperlinks{HyperLinks: []xlsxHyperlink{}}
				}

				var relId string
				for _, rel := range relations.Relationships {
					if rel.Target == cell.Hyperlink.Link {
						relId = rel.Id
					}
				}

				if relId != "" {

					xlsxLink := xlsxHyperlink{
						RelationshipId: relId,
						Reference:      xC.R,
						DisplayString:  cell.Hyperlink.DisplayString,
						Tooltip:        cell.Hyperlink.Tooltip}
					worksheet.Hyperlinks.HyperLinks = append(worksheet.Hyperlinks.HyperLinks, xlsxLink)
				}
			}

			if cell.HMerge > 0 || cell.VMerge > 0 {
				// r == rownum, c == colnum
				mc := xlsxMergeCell{}
				start := fmt.Sprintf("%s%d", ColIndexToLetters(c), r+1)
				endcol := c + cell.HMerge
				endrow := r + cell.VMerge + 1
				end := fmt.Sprintf("%s%d", ColIndexToLetters(endcol), endrow)
				mc.Ref = start + ":" + end
				if worksheet.MergeCells == nil {
					worksheet.MergeCells = &xlsxMergeCells{}
				}
				worksheet.MergeCells.Cells = append(worksheet.MergeCells.Cells, mc)
				worksheet.MergeCells.addCell(mc)
			}
		}
		xSheet.Row = append(xSheet.Row, xRow)
	}
	// Update sheet format with the freshly determined max levels
	s.SheetFormat.OutlineLevelCol = maxLevelCol
	s.SheetFormat.OutlineLevelRow = maxLevelRow
	// .. and then also apply this to the xml worksheet
	worksheet.SheetFormatPr.OutlineLevelCol = s.SheetFormat.OutlineLevelCol
	worksheet.SheetFormatPr.OutlineLevelRow = s.SheetFormat.OutlineLevelRow
	if worksheet.MergeCells != nil {
		worksheet.MergeCells.Count = len(worksheet.MergeCells.Cells)
	}

	if s.AutoFilter != nil {
		worksheet.AutoFilter = &xlsxAutoFilter{Ref: fmt.Sprintf("%v:%v", s.AutoFilter.TopLeftCell, s.AutoFilter.BottomRightCell)}
	}

	worksheet.SheetData = xSheet
	dimension := xlsxDimension{}
	dimension.Ref = "A1:" + GetCellIDStringFromCoords(maxCell, maxRow)
	if dimension.Ref == "A1:A1" {
		dimension.Ref = "A1"
	}
	worksheet.Dimension = dimension
}

func (s *Sheet) makeDataValidations(worksheet *xlsxWorksheet) {
	if len(s.DataValidations) > 0 {
		if worksheet.DataValidations == nil {
			worksheet.DataValidations = &xlsxDataValidations{}
		}
		for _, dv := range s.DataValidations {
			worksheet.DataValidations.DataValidation = append(worksheet.DataValidations.DataValidation, dv)
		}
		worksheet.DataValidations.Count = len(worksheet.DataValidations.DataValidation)
	}
}

// Dump sheet to its XML representation, intended for internal use only
func (s *Sheet) makeXLSXSheet(refTable *RefTable, styles *xlsxStyleSheet, relations *xlsxWorksheetRels) *xlsxWorksheet {
	worksheet := newXlsxWorksheet()

	// Scan through the sheet and see if there are any merged cells. If there
	// are, we may need to extend the size of the sheet. There needs to be
	// phantom cells underlying the area covered by the merged cell
	s.handleMerged()

	s.makeSheetView(worksheet)
	s.makeSheetFormatPr(worksheet)
	maxLevelCol := s.makeCols(worksheet, styles)
	s.makeDataValidations(worksheet)
	s.makeRows(worksheet, styles, refTable, relations, maxLevelCol)

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
