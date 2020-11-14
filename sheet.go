package xlsx

import (
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"strconv"
	"strings"

	"github.com/shabbyrobe/xmlwriter"
)

// Sheet is a high level structure intended to provide user access to
// the contents of a particular sheet within an XLSX file.
type Sheet struct {
	Name            string
	File            *File
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
	cellStore       CellStore
	currentRow      *Row
}

// NewSheet constructs a Sheet with the default CellStore and returns
// a pointer to it.
func NewSheet(name string) (*Sheet, error) {
	return NewSheetWithCellStore(name, NewMemoryCellStore)
}

// NewSheetWithCellStore constructs a Sheet, backed by a CellStore,
// for which you must provide the constructor function.
func NewSheetWithCellStore(name string, constructor CellStoreConstructor) (*Sheet, error) {
	sheet := &Sheet{
		Name: name,
		Cols: &ColStore{},
	}
	var err error
	sheet.cellStore, err = constructor()
	if err != nil {
		return nil, fmt.Errorf("NewSheetWithCellStore: %w", err)
	}
	return sheet, err

}

// Remove Sheet's dependant resources - if you are done with operations on a sheet this should be called to clear down the Sheet's persistent cache.  Note: if you call this, all further read operaton on the sheet will fail - including any attempt to save the file, or dump it's contents to a byte stream.  Therefore only call this *after* you've saved your changes, of when you're done reading a sheet in a file you don't plan to persist. 
func (s *Sheet) Close() {
	s.cellStore.Close()
	s.cellStore = nil
}

func (s *Sheet) getState() string {
	if s.Hidden {
		return "hidden"
	}
	return "visible"
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

func (s *Sheet) setCurrentRow(r *Row) {
	if r != nil && r == s.currentRow {
		return
	}
	if s.currentRow != nil && s.currentRow.isCustom {
		err := s.cellStore.WriteRow(s.currentRow)
		if err != nil {
			panic(err)
		}
	}
	s.currentRow = r
}

// rowVisitorFlags contains flags that can be set by a RowVisitorOption to affect the behaviour of sheet.ForEachRow
type rowVisitorFlags struct {
	skipEmptyRows bool
}

// RowVisitorOption defines the call signature of functions that can be passed as options to the Sheet.ForEachRow function to affect its behaviour.
type RowVisitorOption func(flags *rowVisitorFlags)

// SkipEmptyRows can be passed to the Sheet.ForEachRow function to
// cause it to skip over empty Rows.
func SkipEmptyRows(flags *rowVisitorFlags) {
	flags.skipEmptyRows = true
}

// A RowVisitor function should be provided by the user when calling
// Sheet.ForEachRow, it will be called once for every Row visited.
type RowVisitor func(r *Row) error

func (s *Sheet) mustBeOpen() {
	if s.cellStore == nil {
		panic("Attempt to iterate over sheet with no cellstore. Perhaps you called Close() on this sheet?")
	}
}

func (s *Sheet) ForEachRow(rv RowVisitor, options ...RowVisitorOption) error {
	s.mustBeOpen()
	flags := &rowVisitorFlags{}
	for _, opt := range options {
		opt(flags)
	}
	if s.currentRow != nil {
		err := s.cellStore.WriteRow(s.currentRow)
		if err != nil {
			return err
		}
	}
	for i := 0; i < s.MaxRow; i++ {
		r, err := s.cellStore.ReadRow(makeRowKey(s, i), s)
		if err != nil {
			if _, ok := err.(*RowNotFoundError); !ok {
				return err

			}
			if flags.skipEmptyRows {
				continue
			}
			r = s.cellStore.MakeRow(s)
			r.num = i
		}
		if r.cellStoreRow.CellCount() == 0 && flags.skipEmptyRows {
			continue
		}
		r.Sheet = s
		s.setCurrentRow(r)
		err = rv(r)
		if err != nil {
			return err
		}
	}
	return nil
}

// Add a new Row to a Sheet
func (s *Sheet) AddRow() *Row {
	s.mustBeOpen()
	// NOTE - this is not safe to use concurrently
	if s.currentRow != nil {
		s.cellStore.WriteRow(s.currentRow)
	}
	row := s.cellStore.MakeRow(s)
	row.num = s.MaxRow
	s.MaxRow++
	s.setCurrentRow(row)
	return row
}

func makeRowKey(s *Sheet, i int) string {
	return fmt.Sprintf("%s:%06d", s.Name, i)
}

// Add a new Row to a Sheet at a specific index
func (s *Sheet) AddRowAtIndex(index int) (*Row, error) {
	s.mustBeOpen()
	if index < 0 || index > s.MaxRow {
		return nil, errors.New("AddRowAtIndex: index out of bounds")
	}

	if s.currentRow != nil {
		s.cellStore.WriteRow(s.currentRow)
	}

	// We move rows in reverse order to avoid overwriting anyting
	for i := (s.MaxRow - 1); i >= index; i-- {
		nRow, err := s.cellStore.ReadRow(makeRowKey(s, i), s)
		if err != nil {
			continue
		}
		nRow.Sheet = s
		s.setCurrentRow(nRow)
		s.cellStore.MoveRow(nRow, i+1)
	}
	row := s.cellStore.MakeRow(s)
	row.num = index
	s.setCurrentRow(row)
	err := s.cellStore.WriteRow(row)
	if err != nil {
		return nil, err
	}
	s.MaxRow++
	return row, nil
}

// Add a DataValidation to a range of cells
func (s *Sheet) AddDataValidation(dv *xlsxDataValidation) {
	s.mustBeOpen()
	s.DataValidations = append(s.DataValidations, dv)
}

// Removes a row at a specific index
func (s *Sheet) RemoveRowAtIndex(index int) error {
	s.mustBeOpen()
	if index < 0 || index >= s.MaxRow {
		return fmt.Errorf("Cannot remove row: index out of range: %d", index)
	}
	if s.currentRow != nil {
		s.setCurrentRow(nil)
	}
	err := s.cellStore.RemoveRow(makeRowKey(s, index))
	if err != nil {
		return err
	}
	for i := index + 1; i < s.MaxRow; i++ {
		nRow, err := s.cellStore.ReadRow(makeRowKey(s, i), s)
		if err != nil {
			continue
		}
		nRow.Sheet = s
		s.cellStore.MoveRow(nRow, i-1)
	}
	s.MaxRow--
	return nil
}

// Make sure we always have as many Rows as we do cells.
func (s *Sheet) maybeAddRow(rowCount int) {
	s.mustBeOpen()
	if rowCount > s.MaxRow {
		loopCnt := rowCount - s.MaxRow
		for i := 0; i < loopCnt; i++ {
			row := s.cellStore.MakeRow(s)
			row.num = i
			s.setCurrentRow(row)
		}
		s.MaxRow = rowCount
	}
}

// Make sure we always have as many Rows as we do cells.
func (s *Sheet) Row(idx int) (*Row, error) {
	s.mustBeOpen()
	s.maybeAddRow(idx + 1)
	if s.currentRow != nil && idx == s.currentRow.num {
		return s.currentRow, nil
	}
	r, err := s.cellStore.ReadRow(makeRowKey(s, idx), s)
	if err != nil {
		if _, ok := err.(*RowNotFoundError); !ok {
			return nil, err
		}
	}
	if r == nil {
		r = s.cellStore.MakeRow(s)
		r.num = idx
	} else {
		r.Sheet = s
	}
	s.setCurrentRow(r)
	return r, nil
}

// Return the Col that applies to this Column index, or return nil if no such Col exists
func (s *Sheet) Col(idx int) *Col {
	s.mustBeOpen()
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
func (s *Sheet) Cell(row, col int) (*Cell, error) {
	s.mustBeOpen()
	// If the user requests a row beyond what we have, then extend.
	for s.MaxRow <= row {
		s.AddRow()
	}

	r, err := s.Row(row)
	if err != nil {
		return nil, err
	}
	cell := r.GetCell(col)
	cell.Row = r
	return cell, err
}

//Set the parameters of a column.  Parameters are passed as a pointer
//to a Col structure which you much construct yourself.
func (s *Sheet) SetColParameters(col *Col) {
	s.mustBeOpen()
	if s.Cols == nil {
		panic("trying to use uninitialised ColStore")
	}
	s.Cols.Add(col)
}

func (s *Sheet) setCol(min, max int, setter func(col *Col)) {
	s.mustBeOpen()
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
	s.mustBeOpen()
	s.setCol(min, max, func(col *Col) {
		col.SetWidth(width)
	})
}

// This can be use as the default scale function for the autowidth.
// It works well with the default font sizes.
func DefaultAutoWidth(s string) float64 {
	return (float64(strings.Count(s, "")) + 3.0 ) * 1.2
}

// Tries to guess the best width for a column, based on the largest
// cell content. A scale function needs to be provided.
func (s *Sheet) SetColAutoWidth(colIndex int, width func (string) float64) error {
	s.mustBeOpen()
	largestWidth := 0.0
	rowVisitor := func (r *Row) error {
		cell := r.GetCell(colIndex)
		value, err := cell.FormattedValue()
		if err != nil {
			return err
		}

		if width(value) > largestWidth {
			largestWidth = width(value)
		}
		return nil
	}
	err := s.ForEachRow(rowVisitor)

	if err != nil {
		return err
	}

	s.SetColWidth(colIndex, colIndex, largestWidth)

	return nil
}

// Set the outline level for a range of columns.
func (s *Sheet) SetOutlineLevel(minCol, maxCol int, outlineLevel uint8) {
	s.mustBeOpen()
	s.setCol(minCol, maxCol, func(col *Col) {
		col.SetOutlineLevel(outlineLevel)
	})
}

// Set the type for a range of columns.
func (s *Sheet) SetType(minCol, maxCol int, cellType CellType) {
	s.mustBeOpen()
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

	s.ForEachRow(func(row *Row) error {
		return row.ForEachCell(func(cell *Cell) error {
			if cell.HMerge > 0 || cell.VMerge > 0 {
				coord := GetCellIDStringFromCoords(cell.num, row.num)
				merged[coord] = cell
			}
			return nil
		}, SkipEmptyCells)

	}, SkipEmptyRows)

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
	s.mustBeOpen()
	maxLevelCol = 0
	if s.Cols == nil {
		panic("trying to use uninitialised ColStore")
	}
	s.Cols.ForEach(
		func(c int, col *Col) {
			XfId := 0
			style := col.GetStyle()

			hasNumFmt := len(col.numFmt) > 0
			if hasNumFmt {
				if style == nil {
					style = NewStyle()
				}

				xNumFmt := styles.newNumFmt(col.numFmt)
				XfId = handleStyleForXLSX(style, xNumFmt.NumFmtId, styles)
			} else {
				if style != nil {
					XfId = handleStyleForXLSX(style, 0, styles)
				}
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
					Style:        &XfId,
					BestFit:      col.BestFit,
					Phonetic:     col.Phonetic,
				})

			if col.OutlineLevel != nil && *col.OutlineLevel > maxLevelCol {
				maxLevelCol = *col.OutlineLevel
			}
		})

	return maxLevelCol
}

func (s *Sheet) prepSheetForMarshalling(maxLevelCol uint8) {
	s.SheetFormat.OutlineLevelCol = maxLevelCol
}

func (s *Sheet) prepWorksheetFromRows(worksheet *xlsxWorksheet, relations *xlsxWorksheetRels) error {
	s.mustBeOpen()
	var maxCell, maxRow int

	prepRow := func(row *Row) error {
		if row.num > maxRow {
			maxRow = row.num
		}

		prepCell := func(cell *Cell) error {
			if cell.num > maxCell {
				maxCell = cell.num
			}
			cellID := GetCellIDStringFromCoords(cell.num, row.num)
			if nil != cell.DataValidation {
				if nil == worksheet.DataValidations {
					worksheet.DataValidations = &xlsxDataValidations{}
				}
				cell.DataValidation.Sqref = cellID
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
						Reference:      cellID,
						DisplayString:  cell.Hyperlink.DisplayString,
						Tooltip:        cell.Hyperlink.Tooltip}
					worksheet.Hyperlinks.HyperLinks = append(worksheet.Hyperlinks.HyperLinks, xlsxLink)
				}
			}

			if cell.HMerge > 0 || cell.VMerge > 0 {
				mc := xlsxMergeCell{}
				start := fmt.Sprintf("%s%d", ColIndexToLetters(cell.num), row.num+1)
				endcol := cell.num + cell.HMerge
				endrow := row.num + cell.VMerge + 1
				end := fmt.Sprintf("%s%d", ColIndexToLetters(endcol), endrow)
				mc.Ref = start + ":" + end
				if worksheet.MergeCells == nil {
					worksheet.MergeCells = &xlsxMergeCells{}
				}
				worksheet.MergeCells.Cells = append(worksheet.MergeCells.Cells, mc)
				worksheet.MergeCells.addCell(mc)
			}
			return nil
		}

		return row.ForEachCell(prepCell, SkipEmptyCells)
	}

	err := s.ForEachRow(prepRow, SkipEmptyRows)
	if err != nil {
		return err
	}
	worksheet.SheetFormatPr.OutlineLevelCol = s.SheetFormat.OutlineLevelCol
	worksheet.SheetFormatPr.OutlineLevelRow = s.SheetFormat.OutlineLevelRow
	if worksheet.MergeCells != nil {
		worksheet.MergeCells.Count = len(worksheet.MergeCells.Cells)
	}

	if s.AutoFilter != nil {
		worksheet.AutoFilter = &xlsxAutoFilter{Ref: fmt.Sprintf("%v:%v", s.AutoFilter.TopLeftCell, s.AutoFilter.BottomRightCell)}
	}

	dimension := xlsxDimension{}
	dimension.Ref = "A1:" + GetCellIDStringFromCoords(maxCell, maxRow)
	if dimension.Ref == "A1:A1" {
		dimension.Ref = "A1"
	}
	worksheet.Dimension = dimension
	return nil
}

func (s *Sheet) makeRows(worksheet *xlsxWorksheet, styles *xlsxStyleSheet, refTable *RefTable, relations *xlsxWorksheetRels, maxLevelCol uint8) error {
	s.mustBeOpen()
	maxRow := 0
	maxCell := 0
	var maxLevelRow uint8
	xSheet := xlsxSheetData{}
	makeR := func(row *Row) error {
		r := row.num
		if r > maxRow {
			maxRow = r
		}
		xRow := xlsxRow{}
		xRow.R = r + 1
		if row.isCustom {
			xRow.CustomHeight = true
			xRow.Ht = fmt.Sprintf("%g", row.GetHeight())
		}
		xRow.OutlineLevel = row.GetOutlineLevel()
		if xRow.OutlineLevel > maxLevelRow {
			maxLevelRow = xRow.OutlineLevel
		}
		makeC := func(cell *Cell) error {
			var XfId int

			c := cell.num
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
				} else if len(cell.RichText) > 0 {
					xC.V = strconv.Itoa(refTable.AddRichText(cell.RichText))
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
			return nil
		}
		err := row.ForEachCell(makeC, SkipEmptyCells)
		if err != nil {
			return err
		}
		xSheet.Row = append(xSheet.Row, xRow)
		return nil
	}

	err := s.ForEachRow(makeR, SkipEmptyRows)
	if err != nil {
		return err
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
	return nil
}

func (s *Sheet) makeDataValidations(worksheet *xlsxWorksheet) {
	s.mustBeOpen()
	if len(s.DataValidations) > 0 {
		if worksheet.DataValidations == nil {
			worksheet.DataValidations = &xlsxDataValidations{}
		}
		worksheet.DataValidations.DataValidation = append(worksheet.DataValidations.DataValidation, s.DataValidations...)
		worksheet.DataValidations.Count = len(worksheet.DataValidations.DataValidation)
	}
}

func (s *Sheet) MarshalSheet(w io.Writer, refTable *RefTable, styles *xlsxStyleSheet, relations *xlsxWorksheetRels) error {
	worksheet := newXlsxWorksheet()

	s.handleMerged()
	s.makeSheetView(worksheet)
	s.makeSheetFormatPr(worksheet)
	maxLevelCol := s.makeCols(worksheet, styles)
	s.makeDataValidations(worksheet)
	s.prepSheetForMarshalling(maxLevelCol)
	err := s.prepWorksheetFromRows(worksheet, relations)
	if err != nil {
		return err
	}
	xw := xmlwriter.Open(w)

	err = xw.StartDoc(xmlwriter.Doc{})
	if err != nil {
		return err
	}
	err = worksheet.WriteXML(xw, s, styles, refTable)
	if err != nil {
		return err
	}

	return xw.EndAllFlush()
}

// Dump sheet to its XML representation, intended for internal use only
func (s *Sheet) makeXLSXSheet(refTable *RefTable, styles *xlsxStyleSheet, relations *xlsxWorksheetRels) *xlsxWorksheet {
	s.mustBeOpen()
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
