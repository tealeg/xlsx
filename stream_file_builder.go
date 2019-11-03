// Authors: Ryan Hollis (ryanh@)

// The purpose of StreamFileBuilder and StreamFile is to allow streamed writing of XLSX files.
// Directions:
// 1. Create a StreamFileBuilder with NewStreamFileBuilder() or NewStreamFileBuilderForPath().
// 2. Add the sheets and their first row of data by calling AddSheet().
// 3. Call Build() to get a StreamFile. Once built, all functions on the builder will return an error.
// 4. Write to the StreamFile with Write(). Writes begin on the first sheet. New rows are always written and flushed
// to the io. All rows written to the same sheet must have the same number of cells as the header provided when the sheet
// was created or an error will be returned.
// 5. Call NextSheet() to proceed to the next sheet. Once NextSheet() is called, the previous sheet can not be edited.
// 6. Call Close() to finish.

// Directions for using custom styles and different data types:
// 1. Create a StreamFileBuilder with NewStreamFileBuilder() or NewStreamFileBuilderForPath().
// 2. Use MakeStyle() to create the styles you want yo use in your document. Keep a list of these styles.
// 3. Add all the styles you created by using AddStreamStyle() or AddStreamStyleList().
// 4. Add the sheets and their column styles of data by calling AddSheetS().
// 5. Call Build() to get a StreamFile. Once built, all functions on the builder will return an error.
// 6. Write to the StreamFile with WriteS(). Writes begin on the first sheet. New rows are always written and flushed
// to the io. All rows written to the same sheet must have the same number of cells as the number of column styles
// provided when adding the sheet with AddSheetS() or an error will be returned.
// 5. Call NextSheet() to proceed to the next sheet. Once NextSheet() is called, the previous sheet can not be edited.
// 6. Call Close() to finish.

// Future work suggestions:
// The current default style uses fonts that are not on Macs by default so opening the XLSX files in Numbers causes a
// pop up that says there are missing fonts. The font could be changed to something that is usually found on Mac and PC.
// Extend support for Formulas and Shared Strings.

package xlsx

import (
	"archive/zip"
	"errors"
	"io"
	"os"
	"strconv"
	"strings"
)

type cellStreamStyle map[int]StreamStyle
type defaultCellType map[int]*CellType

type StreamFileBuilder struct {
	built                                   bool
	firstSheetAdded                         bool
	customStylesAdded                       bool
	xlsxFile                                *File
	zipWriter                               *zip.Writer
	cellTypeToStyleIds                      map[CellType]int
	maxStyleId                              int
	styleIds                                [][]int
	customStreamStyles                      map[StreamStyle]struct{}
	customNumFormats                        map[int]xlsxNumFmt
	styleIdMap                              map[StreamStyle]int
	streamingCellMetadatas                  map[int]*StreamingCellMetadata
	sheetStreamStyles                       map[int]cellStreamStyle
	sheetDefaultCellType                    map[int]defaultCellType
	defaultColumnStreamingCellMetadataAdded bool
}

const (
	sheetFilePathPrefix = "xl/worksheets/sheet"
	sheetFilePathSuffix = ".xml"
	endSheetDataTag     = "</sheetData>"
	dimensionTag        = `<dimension ref="%s"></dimension>`
	// This is the index of the max style that this library will insert into XLSX sheets by default.
	// This allows us to predict what the style id of styles that we add will be.
	// TestXlsxStyleBehavior tests that this behavior continues to be what we expect.
	initMaxStyleId = 1
)

var BuiltStreamFileBuilderError = errors.New("StreamFileBuilder has already been built, functions may no longer be used")

// NewStreamFileBuilder creates an StreamFileBuilder that will write to the the provided io.writer
func NewStreamFileBuilder(writer io.Writer) *StreamFileBuilder {
	return &StreamFileBuilder{
		zipWriter:              zip.NewWriter(writer),
		xlsxFile:               NewFile(),
		cellTypeToStyleIds:     make(map[CellType]int),
		maxStyleId:             initMaxStyleId,
		customStreamStyles:     make(map[StreamStyle]struct{}),
		customNumFormats:       make(map[int]xlsxNumFmt),
		styleIdMap:             make(map[StreamStyle]int),
		streamingCellMetadatas: make(map[int]*StreamingCellMetadata),
		sheetStreamStyles:      make(map[int]cellStreamStyle),
		sheetDefaultCellType:   make(map[int]defaultCellType),
	}
}

// NewStreamFileBuilderForPath takes the name of an XLSX file and returns a builder for it.
// The file will be created if it does not exist, or truncated if it does.
func NewStreamFileBuilderForPath(path string) (*StreamFileBuilder, error) {
	file, err := os.Create(path)
	if err != nil {
		return nil, err
	}
	return NewStreamFileBuilder(file), nil
}

// AddSheet will add sheets with the given name.  Sheet names must be unique, or an
// error will be thrown.
func (sb *StreamFileBuilder) AddSheet(name string, cellTypes []*CellType) error {
	return sb.addSheet(name, cellTypes, false)
}

// AddSheetWithAutoFilters will add sheets with the given name.
// Sheet names must be unique, or an error will be thrown.
// AddSheetWithAutoFilters will also add autoFilters for every column in the sheet.
func (sb *StreamFileBuilder) AddSheetWithAutoFilters(name string, cellTypes []*CellType) error {
	return sb.addSheet(name, cellTypes, true)
}

func (sb *StreamFileBuilder) addSheet(name string, cellTypes []*CellType, addAutofilters bool) error {
	if sb.built {
		return BuiltStreamFileBuilderError
	}

	sheet, err := sb.xlsxFile.AddSheet(name)
	if addAutofilters {
		lastColCoordinate := GetCellIDStringFromCoords(len(cellTypes)-1, 0)
		sheet.AutoFilter = &AutoFilter{"A1", lastColCoordinate}
	}

	if err != nil {
		// Set built on error so that all subsequent calls to the builder will also fail.
		sb.built = true
		return err
	}
	sb.styleIds = append(sb.styleIds, []int{})

	for i, cellType := range cellTypes {
		var cellStyleIndex int
		var ok bool
		if cellType != nil {
			// The cell type is one of the attributes of a Style.
			// Since it is the only attribute of Style that we use, we can assume that cell types
			// map one to one with Styles and their Style ID.
			// If a new cell type is used, a new style gets created with an increased id, if an existing cell type is
			// used, the pre-existing style will also be used.
			cellStyleIndex, ok = sb.cellTypeToStyleIds[*cellType]
			if !ok {
				sb.maxStyleId++
				cellStyleIndex = sb.maxStyleId
				sb.cellTypeToStyleIds[*cellType] = sb.maxStyleId
			}
			sheet.SetType(i+1, i+1, *cellType)

		}
		sb.styleIds[len(sb.styleIds)-1] = append(sb.styleIds[len(sb.styleIds)-1], cellStyleIndex)
	}
	return nil
}

func (sb *StreamFileBuilder) AddSheetWithDefaultColumnMetadata(name string, columnsDefaultStreamingCellMetadata []*StreamingCellMetadata) error {
	if sb.built {
		return BuiltStreamFileBuilderError
	}
	_, err := sb.xlsxFile.AddSheet(name)
	if err != nil {
		// Set built on error so that all subsequent calls to the builder will also fail.
		sb.built = true
		return err
	}
	sb.styleIds = append(sb.styleIds, []int{})
	sheetIndex := len(sb.xlsxFile.Sheets) - 1

	cSS := make(cellStreamStyle)
	dCT := make(defaultCellType)
	for i, streamingCellMetadata := range columnsDefaultStreamingCellMetadata {
		var cellStyleIndex int
		var ok bool
		if streamingCellMetadata != nil {
			// Exact same logic as `AddSheet` to ensure compatibility as much as possible
			// with the `AddSheet` + `StreamFile.Write` code path
			cellStyleIndex, ok = sb.cellTypeToStyleIds[streamingCellMetadata.cellType]
			if !ok {
				sb.maxStyleId++
				cellStyleIndex = sb.maxStyleId
				sb.cellTypeToStyleIds[streamingCellMetadata.cellType] = sb.maxStyleId
			}

			// Add streamStyle and set default cell metadata on col
			sb.customStreamStyles[streamingCellMetadata.streamStyle] = struct{}{}
			sb.streamingCellMetadatas[i+1] = streamingCellMetadata
			cSS[i] = streamingCellMetadata.streamStyle
			dCT[i] = streamingCellMetadata.cellType.Ptr()
		}
		sb.styleIds[len(sb.styleIds)-1] = append(sb.styleIds[len(sb.styleIds)-1], cellStyleIndex)
	}
	// Add fall back streamStyle
	sb.customStreamStyles[StreamStyleDefaultString] = struct{}{}
	// Toggle to true to ensure `styleIdMap` is constructed from `customStreamStyles` on `Build`
	sb.customStylesAdded = true
	// Hack to ensure the `dimension` tag on each `worksheet` xml is stripped. Otherwise only the first
	// row of each worksheet will be read back rather than all rows
	sb.defaultColumnStreamingCellMetadataAdded = true
	sb.sheetStreamStyles[sheetIndex] = cSS
	sb.sheetDefaultCellType[sheetIndex] = dCT
	return nil
}

// AddSheetS will add a sheet with the given name and column styles. The number of column styles given
// is the number of columns that will be created, and thus the number of cells each row has to have.
// columnStyles[0] becomes the style of the first column, columnStyles[1] the style of the second column etc.
// All the styles in columnStyles have to have been added or an error will be returned.
// Sheet names must be unique, or an error will be returned.
func (sb *StreamFileBuilder) AddSheetS(name string, columnStyles []StreamStyle) error {
	if sb.built {
		return BuiltStreamFileBuilderError
	}
	sheet, err := sb.xlsxFile.AddSheet(name)
	if err != nil {
		// Set built on error so that all subsequent calls to the builder will also fail.
		sb.built = true
		return err
	}
	// To make sure no new styles can be added after adding a sheet
	sb.firstSheetAdded = true

	// Check if all styles that will be used for columns have been created
	for _, colStyle := range columnStyles {
		if _, ok := sb.customStreamStyles[colStyle]; !ok {
			return errors.New("trying to make use of a style that has not been added")
		}
	}

	// Is needed for stream file to work but is not needed for streaming with styles
	sb.styleIds = append(sb.styleIds, []int{})

	if sheet.Cols == nil {
		panic("trying to use uninitialised ColStore")
	}

	cSS := make(map[int]StreamStyle)
	// Set default column styles based on the cel styles in the first row
	// Set the default column width to 11. This makes enough places for the
	// default date style cells to display the dates correctly
	for i, colStyle := range columnStyles {
		colNum := i + 1
		cSS[colNum] = colStyle
		sheet.SetColWidth(colNum, colNum, 11)
	}
	sheetIndex := len(sb.xlsxFile.Sheets) - 1
	sb.sheetStreamStyles[sheetIndex] = cSS
	return nil
}

// AddValidation will add a validation to a sheet.
func (sb *StreamFileBuilder) AddValidation(sheetIndex int, validation *xlsxDataValidation) {
	sheet := sb.xlsxFile.Sheets[sheetIndex]
	sheet.AddDataValidation(validation)
}

// Build begins streaming the XLSX file to the io, by writing all the XLSX metadata. It creates a StreamFile struct
// that can be used to write the rows to the sheets.
func (sb *StreamFileBuilder) Build() (*StreamFile, error) {
	if sb.built {
		return nil, BuiltStreamFileBuilderError
	}
	sb.built = true

	parts, err := sb.xlsxFile.MarshallParts()
	if err != nil {
		return nil, err
	}

	if sb.customStylesAdded {
		parts["xl/styles.xml"], err = sb.marshalStyles()
		if err != nil {
			return nil, err
		}
	}

	es := &StreamFile{
		zipWriter:              sb.zipWriter,
		xlsxFile:               sb.xlsxFile,
		sheetXmlPrefix:         make([]string, len(sb.xlsxFile.Sheets)),
		sheetXmlSuffix:         make([]string, len(sb.xlsxFile.Sheets)),
		styleIds:               sb.styleIds,
		styleIdMap:             sb.styleIdMap,
		streamingCellMetadatas: sb.streamingCellMetadatas,
		sheetStreamStyles:      sb.sheetStreamStyles,
		sheetDefaultCellType:   sb.sheetDefaultCellType,
	}
	for path, data := range parts {
		// If the part is a sheet, don't write it yet. We only want to write the XLSX metadata files, since at this
		// point the sheets are still empty. The sheet files will be written later as their rows come in.
		if strings.HasPrefix(path, sheetFilePathPrefix) {
			// sb.default ColumnStreamingCellMetadataAdded is a hack because neither the `AddSheet` nor `AddSheetS` codepaths
			// actually encode a valid worksheet dimension. `AddSheet` encodes an empty one: "" and `AddSheetS` encodes
			// an effectively empty one: "A1". `AddSheetWithDefaultColumnMetadata` uses logic from both paths which results
			// in an effectively invalid dimension being encoded which, upon read, results in only reading in the header of
			// a given worksheet and non of the rows that follow
			if err := sb.processEmptySheetXML(es, path, data, !sb.customStylesAdded || sb.defaultColumnStreamingCellMetadataAdded); err != nil {
				return nil, err
			}
			continue
		}
		metadataFile, err := sb.zipWriter.Create(path)
		if err != nil {
			return nil, err
		}
		_, err = metadataFile.Write([]byte(data))
		if err != nil {
			return nil, err
		}
	}

	if err := es.NextSheet(); err != nil {
		return nil, err
	}
	return es, nil
}

func (sb *StreamFileBuilder) AddNewNumberFormat(formatCode string) int {
	if sb.xlsxFile.styles == nil {
		sb.xlsxFile.styles = newXlsxStyleSheet(sb.xlsxFile.theme)
	}
	numFmt := sb.xlsxFile.styles.newNumFmt(formatCode)
	sb.customNumFormats[numFmt.NumFmtId] = numFmt
	return numFmt.NumFmtId
}

func (sb *StreamFileBuilder) marshalStyles() (string, error) {

	for streamStyle := range sb.customStreamStyles {
		XfId := handleStyleForXLSX(streamStyle.style, streamStyle.xNumFmtId, sb.xlsxFile.styles)
		sb.styleIdMap[streamStyle] = XfId
		if xNumFmt, ok := sb.customNumFormats[streamStyle.xNumFmtId]; ok {
			sb.xlsxFile.styles.addNumFmt(xNumFmt)
		}
	}

	styleSheetXMLString, err := sb.xlsxFile.styles.Marshal()
	if err != nil {
		return "", err
	}
	return styleSheetXMLString, nil
}

// AddStreamStyle adds a new style to the style sheet.
// Only Styles that have been added through this function will be usable.
// This function cannot be used after AddSheetS or Build has been called, and if it is
// called after AddSheetS or Buildit will return an error.
func (sb *StreamFileBuilder) AddStreamStyle(streamStyle StreamStyle) error {
	if sb.firstSheetAdded {
		return errors.New("at least one sheet has been added, cannot add new styles anymore")
	}
	if sb.built {
		return errors.New("file has been build, cannot add new styles anymore")
	}
	sb.customStreamStyles[streamStyle] = struct{}{}
	sb.customStylesAdded = true
	return nil
}

// AddStreamStyleList adds a list of new styles to the style sheet.
// Only Styles that have been added through either this function or AddStreamStyle will be usable.
// This function cannot be used after AddSheetS and Build has been called, and if it is
// called after AddSheetS and Build it will return an error.
func (sb *StreamFileBuilder) AddStreamStyleList(streamStyles []StreamStyle) error {
	for _, streamStyle := range streamStyles {
		err := sb.AddStreamStyle(streamStyle)
		if err != nil {
			return err
		}
	}
	return nil
}

// processEmptySheetXML will take in the path and XML data of an empty sheet, and will save the beginning and end of the
// XML file so that these can be written at the right time.
func (sb *StreamFileBuilder) processEmptySheetXML(sf *StreamFile, path, data string, removeDimensionTagFlag bool) error {
	// Get the sheet index from the path
	sheetIndex, err := getSheetIndex(sf, path)
	if err != nil {
		return err
	}

	// Remove the Dimension tag. Since more rows are going to be written to the sheet, it will be wrong.
	// It is valid to for a sheet to be missing a Dimension tag, but it is not valid for it to be wrong.
	if removeDimensionTagFlag {
		data = removeDimensionTag(data)
	}

	// Split the sheet at the end of its SheetData tag so that more rows can be added inside.
	prefix, suffix, err := splitSheetIntoPrefixAndSuffix(data)
	if err != nil {
		return err
	}
	sf.sheetXmlPrefix[sheetIndex] = prefix
	sf.sheetXmlSuffix[sheetIndex] = suffix
	return nil
}

// getSheetIndex parses the path to the XLSX sheet data and returns the index
// The files that store the data for each sheet must have the format:
// xl/worksheets/sheet123.xml
// where 123 is the index of the sheet. This file path format is part of the XLSX file standard.
func getSheetIndex(sf *StreamFile, path string) (int, error) {
	indexString := path[len(sheetFilePathPrefix) : len(path)-len(sheetFilePathSuffix)]
	sheetXLSXIndex, err := strconv.Atoi(indexString)
	if err != nil {
		return -1, errors.New("unexpected sheet file name from xlsx package")
	}
	if sheetXLSXIndex < 1 || len(sf.sheetXmlPrefix) < sheetXLSXIndex ||
		len(sf.sheetXmlSuffix) < sheetXLSXIndex || len(sf.xlsxFile.Sheets) < sheetXLSXIndex {
		return -1, errors.New("unexpected sheet index")
	}
	sheetArrayIndex := sheetXLSXIndex - 1
	return sheetArrayIndex, nil
}

// removeDimensionTag will return the passed in XLSX Spreadsheet XML with the dimension tag removed.
// data is the XML data for the sheet
// sheet is the Sheet struct that the XML was created from.
func removeDimensionTag(data string) string {
	start := strings.Index(data, "<dimension")
	end := strings.Index(data, "</dimension>") + 12
	return data[0:start] + data[end:]
}

// splitSheetIntoPrefixAndSuffix will split the provided XML sheet into a prefix and a suffix so that
// more spreadsheet rows can be inserted in between.
func splitSheetIntoPrefixAndSuffix(data string) (string, string, error) {
	// Split the sheet at the end of its SheetData tag so that more rows can be added inside.
	sheetParts := strings.Split(data, endSheetDataTag)
	if len(sheetParts) != 2 {
		return "", "", errors.New("unexpected Sheet XML: SheetData close tag not found")
	}
	return sheetParts[0], sheetParts[1], nil
}
