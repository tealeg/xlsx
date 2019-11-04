package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"io"
	"strconv"
)

type StreamFile struct {
	xlsxFile               *File
	sheetXmlPrefix         []string
	sheetXmlSuffix         []string
	zipWriter              *zip.Writer
	currentSheet           *streamSheet
	styleIds               [][]int
	styleIdMap             map[StreamStyle]int
	streamingCellMetadatas map[int]*StreamingCellMetadata
	sheetStreamStyles      map[int]cellStreamStyle
	sheetDefaultCellType   map[int]defaultCellType
	err                    error
}

type streamSheet struct {
	// sheetIndex is the XLSX sheet index, which starts at 1
	index int
	// The number of rows that have been written to the sheet so far
	rowCount int
	// The number of columns in the sheet
	columnCount int
	// The writer to write to this sheet's file in the XLSX Zip file
	writer     io.Writer
	styleIds   []int
	mergeCells []string
}

var (
	NoCurrentSheetError      = errors.New("no Current Sheet")
	WrongNumberOfRowsError   = errors.New("invalid number of cells passed to Write. All calls to Write on the same sheet must have the same number of cells")
	AlreadyOnLastSheetError  = errors.New("NextSheet() called, but already on last sheet")
	UnsupportedCellTypeError = errors.New("the given cell type is not supported")
)

// Write will write a row of cells to the current sheet. Every call to Write on the same sheet must contain the
// same number of cells as the header provided when the sheet was created or an error will be returned. This function
// will always trigger a flush on success. Currently the only supported data type is string data.
func (sf *StreamFile) Write(cells []string) error {
	if sf.err != nil {
		return sf.err
	}
	err := sf.write(cells)
	if err != nil {
		sf.err = err
		return err
	}
	return sf.zipWriter.Flush()
}

// WriteWithColumnDefaultMetadata will write a row of cells to the current sheet. Every call to WriteWithColumnDefaultMetadata
// on the same sheet must contain the same number of cells as the header provided when the sheet was created or
// an error will be returned. This function will always trigger a flush on success. Each cell will be encoded with the
// default StreamingCellMetadata of the column that it belongs to. However, if the cell data string cannot be
// parsed into the cell type in StreamingCellMetadata, we fall back on encoding the cell as a string and giving it a default
// string style
func (sf *StreamFile) WriteWithColumnDefaultMetadata(cells []string) error {
	if sf.err != nil {
		return sf.err
	}
	err := sf.writeWithColumnDefaultMetadata(cells)
	if err != nil {
		sf.err = err
		return err
	}
	return sf.zipWriter.Flush()
}

// WriteS will write a row of cells to the current sheet. Every call to WriteS on the same sheet must
// contain the same number of cells as the number of columns provided when the sheet was created or an error
// will be returned. This function will always trigger a flush on success. WriteS supports all data types
// and styles that are supported by StreamCell.
func (sf *StreamFile) WriteS(cells []StreamCell) error {
	if sf.err != nil {
		return sf.err
	}
	err := sf.writeS(cells)
	if err != nil {
		sf.err = err
		return err
	}
	return sf.zipWriter.Flush()
}

func (sf *StreamFile) WriteAll(records [][]string) error {
	if sf.err != nil {
		return sf.err
	}
	for _, row := range records {
		err := sf.write(row)
		if err != nil {
			sf.err = err
			return err
		}
	}
	return sf.zipWriter.Flush()
}

// WriteAllS will write all the rows provided in records. All rows must have the same number of cells as
// the number of columns given when creating the sheet. This function will always trigger a flush on success.
// WriteAllS supports all data types and styles that are supported by StreamCell.
func (sf *StreamFile) WriteAllS(records [][]StreamCell) error {
	if sf.err != nil {
		return sf.err
	}
	for _, row := range records {
		err := sf.writeS(row)
		if err != nil {
			sf.err = err
			return err
		}
	}
	return sf.zipWriter.Flush()
}

func (sf *StreamFile) AddMergeCells(startRowIdx, startColumnIdx, endRowIdx, endColumnIdx int) {
	start := GetCellIDStringFromCoords(startColumnIdx, startRowIdx)
	end := GetCellIDStringFromCoords(endColumnIdx, endRowIdx)
	ref := start + cellRangeChar + end
	sf.currentSheet.mergeCells = append(sf.currentSheet.mergeCells, ref)
}

func (sf *StreamFile) write(cells []string) error {
	if sf.currentSheet == nil {
		return NoCurrentSheetError
	}
	cellCount := len(cells)
	if cellCount != sf.currentSheet.columnCount {
		if sf.currentSheet.columnCount != 0 {
			return WrongNumberOfRowsError
		}
		sf.currentSheet.columnCount = cellCount
	}

	sf.currentSheet.rowCount++
	if err := sf.currentSheet.write(`<row r="` + strconv.Itoa(sf.currentSheet.rowCount) + `">`); err != nil {
		return err
	}
	for colIndex, cellData := range cells {
		// documentation for the c.t (cell.Type) attribute:
		// b (Boolean): Cell containing a boolean.
		// d (Date): Cell contains a date in the ISO 8601 format.
		// e (Error): Cell containing an error.
		// inlineStr (Inline String): Cell containing an (inline) rich string, i.e., one not in the shared string table.
		// If this cell type is used, then the cell value is in the is element rather than the v element in the cell (c element).
		// n (Number): Cell containing a number.
		// s (Shared String): Cell containing a shared string.
		// str (String): Cell containing a formula string.
		cellCoordinate := GetCellIDStringFromCoords(colIndex, sf.currentSheet.rowCount-1)
		cellType := "inlineStr"
		cellOpen := `<c r="` + cellCoordinate + `" t="` + cellType + `"`
		// Add in the style id if the cell isn't using the default style
		if colIndex < len(sf.currentSheet.styleIds) && sf.currentSheet.styleIds[colIndex] != 0 {
			cellOpen += ` s="` + strconv.Itoa(sf.currentSheet.styleIds[colIndex]) + `"`
		}
		cellOpen += `><is><t>`
		cellClose := `</t></is></c>`

		if err := sf.currentSheet.write(cellOpen); err != nil {
			return err
		}
		if err := xml.EscapeText(sf.currentSheet.writer, []byte(cellData)); err != nil {
			return err
		}
		if err := sf.currentSheet.write(cellClose); err != nil {
			return err
		}
	}
	if err := sf.currentSheet.write(`</row>`); err != nil {
		return err
	}
	return sf.zipWriter.Flush()
}

func (sf *StreamFile) writeWithColumnDefaultMetadata(cells []string) error {

	if sf.currentSheet == nil {
		return NoCurrentSheetError
	}

	sheetIndex := sf.currentSheet.index - 1
	currentSheet := sf.xlsxFile.Sheets[sheetIndex]

	var streamCells []StreamCell

	if currentSheet.Cols == nil {
		panic("trying to use uninitialised ColStore")
	}

	if len(cells) != sf.currentSheet.columnCount {
		if sf.currentSheet.columnCount != 0 {
			return WrongNumberOfRowsError

		}
		sf.currentSheet.columnCount = len(cells)
	}

	cSS := sf.sheetStreamStyles[sheetIndex]
	cDCT := sf.sheetDefaultCellType[sheetIndex]
	for ci, c := range cells {
		// TODO: Legacy code paths like `StreamFileBuilder.AddSheet` could
		// leave style empty and if cell data cannot be parsed into cell type then
		// we need a sensible default StreamStyle to fall back to
		style := StreamStyleDefaultString
		// Because `cellData` could be anything we need to attempt to
		// parse into the default cell type and if parsing fails fall back
		// to some sensible default
		cellType := CellTypeInline
		if dct, ok := cDCT[ci]; ok {
			defaultType := dct
			cellType = defaultType.fallbackTo(cells[ci], CellTypeString)
			if ss, ok := cSS[ci]; ok {
				// TODO: Again `CellType` could be nil if sheet was created through
				// legacy code path so, like style, hardcoding for now
				if defaultType != nil && *defaultType == cellType {
					style = ss
				}
			}

		}

		streamCells = append(
			streamCells,
			NewStreamCell(
				c,
				style,
				cellType,
			))

	}
	return sf.writeS(streamCells)
}

func (sf *StreamFile) writeS(cells []StreamCell) error {
	if sf.currentSheet == nil {
		return NoCurrentSheetError
	}
	if len(cells) != sf.currentSheet.columnCount {
		if sf.currentSheet.columnCount != 0 {
			return WrongNumberOfRowsError
		}
		sf.currentSheet.columnCount = len(cells)
	}

	sf.currentSheet.rowCount++
	// Write the row opening
	if err := sf.currentSheet.write(`<row r="` + strconv.Itoa(sf.currentSheet.rowCount) + `">`); err != nil {
		return err
	}

	// Add cells one by one
	for colIndex, cell := range cells {

		xlsxCell, err := sf.getXlsxCell(cell, colIndex)
		if err != nil {
			return err
		}

		marshaledCell, err := xml.Marshal(xlsxCell)
		if err != nil {
			return nil
		}
		// Write the cell
		if _, err := sf.currentSheet.writer.Write(marshaledCell); err != nil {
			return err
		}

	}
	// Write the row ending
	if err := sf.currentSheet.write(`</row>`); err != nil {
		return err
	}
	return sf.zipWriter.Flush()
}

func (sf *StreamFile) getXlsxCell(cell StreamCell, colIndex int) (xlsxC, error) {
	// Get the cell reference (location)
	cellCoordinate := GetCellIDStringFromCoords(colIndex, sf.currentSheet.rowCount-1)

	var cellStyleId int

	if cell.cellStyle != (StreamStyle{}) {
		if idx, ok := sf.styleIdMap[cell.cellStyle]; ok {
			cellStyleId = idx
		} else {
			return xlsxC{}, errors.New("trying to make use of a style that has not been added")
		}
	}

	return makeXlsxCell(cell.cellType, cellCoordinate, cellStyleId, cell.cellData)
}

func makeXlsxCell(cellType CellType, cellCoordinate string, cellStyleId int, cellData string) (xlsxC, error) {
	// documentation for the c.t (cell.Type) attribute:
	// b (Boolean): Cell containing a boolean.
	// d (Date): Cell contains a date in the ISO 8601 format.
	// e (Error): Cell containing an error.
	// inlineStr (Inline String): Cell containing an (inline) rich string, i.e., one not in the shared string table.
	// If this cell type is used, then the cell value is in the is element rather than the v element in the cell (c element).
	// n (Number): Cell containing a number.
	// s (Shared String): Cell containing a shared string.
	// str (String): Cell containing a formula string.
	switch cellType {
	case CellTypeBool:
		return xlsxC{XMLName: xml.Name{Local: "c"}, R: cellCoordinate, S: cellStyleId, T: "b", V: cellData}, nil
	// Dates are better represented using CellTyleNumeric and the date formatting
	//case CellTypeDate:
	//return xlsxC{XMLName: xml.Name{Local: "c"}, R: cellCoordinate, S: cellStyleId, T: "d", V: cellData}, nil
	case CellTypeError:
		return xlsxC{XMLName: xml.Name{Local: "c"}, R: cellCoordinate, S: cellStyleId, T: "e", V: cellData}, nil
	case CellTypeInline:
		return xlsxC{XMLName: xml.Name{Local: "c"}, R: cellCoordinate, S: cellStyleId, T: "inlineStr", Is: &xlsxSI{T: cellData}}, nil
	case CellTypeNumeric:
		return xlsxC{XMLName: xml.Name{Local: "c"}, R: cellCoordinate, S: cellStyleId, T: "n", V: cellData}, nil
	case CellTypeString:
		// TODO Currently shared strings are types as inline strings
		return xlsxC{XMLName: xml.Name{Local: "c"}, R: cellCoordinate, S: cellStyleId, T: "inlineStr", Is: &xlsxSI{T: cellData}}, nil
	// TODO currently not supported
	// case CellTypeStringFormula:
	// return xlsxC{}, UnsupportedCellTypeError
	default:
		return xlsxC{}, UnsupportedCellTypeError
	}
}

// Error reports any error that has occurred during a previous Write or Flush.
func (sf *StreamFile) Error() error {
	return sf.err
}

func (sf *StreamFile) Flush() {
	if sf.err != nil {
		sf.err = sf.zipWriter.Flush()
	}
}

// NextSheet will switch to the next sheet. Sheets are selected in the same order they were added.
// Once you leave a sheet, you cannot return to it.
func (sf *StreamFile) NextSheet() error {
	if sf.err != nil {
		return sf.err
	}
	var sheetIndex int
	if sf.currentSheet != nil {
		if sf.currentSheet.index >= len(sf.xlsxFile.Sheets) {
			sf.err = AlreadyOnLastSheetError
			return AlreadyOnLastSheetError
		}
		if err := sf.writeSheetEnd(); err != nil {
			sf.currentSheet = nil
			sf.err = err
			return err
		}
		sheetIndex = sf.currentSheet.index
	}
	sheetIndex++
	sf.currentSheet = &streamSheet{
		index:       sheetIndex,
		columnCount: sf.xlsxFile.Sheets[sheetIndex-1].MaxCol,
		styleIds:    sf.styleIds[sheetIndex-1],
		rowCount:    len(sf.xlsxFile.Sheets[sheetIndex-1].Rows),
	}
	sheetPath := sheetFilePathPrefix + strconv.Itoa(sf.currentSheet.index) + sheetFilePathSuffix
	fileWriter, err := sf.zipWriter.Create(sheetPath)
	if err != nil {
		sf.err = err
		return err
	}
	sf.currentSheet.writer = fileWriter

	if err := sf.writeSheetStart(); err != nil {
		sf.err = err
		return err
	}
	return nil
}

// Close closes the Stream File.
// Any sheets that have not yet been written to will have an empty sheet created for them.
func (sf *StreamFile) Close() error {
	if sf.err != nil {
		return sf.err
	}
	// If there are sheets that have not been written yet, call NextSheet() which will add files to the zip for them.
	// XLSX readers may error if the sheets registered in the metadata are not present in the file.
	if sf.currentSheet != nil {
		for sf.currentSheet.index < len(sf.xlsxFile.Sheets) {
			if err := sf.NextSheet(); err != nil {
				sf.err = err
				return err
			}
		}
		// Write the end of the last sheet.
		if err := sf.writeSheetEnd(); err != nil {
			sf.err = err
			return err
		}
	}
	err := sf.zipWriter.Close()
	if err != nil {
		sf.err = err
	}
	return err
}

// writeSheetStart will write the start of the Sheet's XML
func (sf *StreamFile) writeSheetStart() error {
	if sf.currentSheet == nil {
		return NoCurrentSheetError
	}
	return sf.currentSheet.write(sf.sheetXmlPrefix[sf.currentSheet.index-1])
}

// writeSheetEnd will write the end of the Sheet's XML
func (sf *StreamFile) writeSheetEnd() error {
	if sf.currentSheet == nil {
		return NoCurrentSheetError
	}
	if err := sf.currentSheet.write(endSheetDataTag); err != nil {
		return err
	}

	if len(sf.currentSheet.mergeCells) > 0 {
		mergeCellData := "<mergeCells count=\"" + strconv.Itoa(len(sf.currentSheet.mergeCells)) + "\">"
		for _, ref := range sf.currentSheet.mergeCells {
			mergeCellData += "<mergeCell ref=\"" + ref + "\"/>"
		}
		if err := sf.currentSheet.write(mergeCellData + "</mergeCells>"); err != nil {
			return err
		}
	}

	return sf.currentSheet.write(sf.sheetXmlSuffix[sf.currentSheet.index-1])
}

func (ss *streamSheet) write(data string) error {
	_, err := ss.writer.Write([]byte(data))
	return err
}
