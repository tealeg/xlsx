package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"strconv"
	"strings"
)

// XLSXReaderError is the standard error type for otherwise undefined
// errors in the XSLX reading process.
type XLSXReaderError struct {
	Err string
}

// String() returns a string value from an XLSXReaderError struct in
// order that it might comply with the os.Error interface.
func (e *XLSXReaderError) Error() string {
	return e.Err
}

// Cell is a high level structure intended to provide user access to
// the contents of Cell within an xlsx.Row.
type Cell struct {
	Value      string
	styleIndex int
	styles     *xlsxStyles
}

// CellInterface defines the public API of the Cell.
type CellInterface interface {
	String() string
}

func (c *Cell) String() string {
	return c.Value
}

// TODO: TestMe!
func (c *Cell) GetStyle() *Style {
	style := &Style{}

	if c.styleIndex > 0 && c.styleIndex <= len(c.styles.CellXfs) {
		xf := c.styles.CellXfs[c.styleIndex-1]
		if xf.ApplyBorder {
			var border Border
			border.Left = c.styles.Borders[xf.BorderId].Left.Style
			border.Right = c.styles.Borders[xf.BorderId].Right.Style
			border.Top = c.styles.Borders[xf.BorderId].Top.Style
			border.Bottom = c.styles.Borders[xf.BorderId].Bottom.Style
			style.Border = border
		}
		if xf.ApplyFill {
			var fill Fill
			fill.PatternType = c.styles.Fills[xf.FillId].PatternFill.PatternType
			fill.BgColor = c.styles.Fills[xf.FillId].PatternFill.BgColor.RGB
			fill.FgColor = c.styles.Fills[xf.FillId].PatternFill.FgColor.RGB
			style.Fill = fill
		}
		if xf.ApplyFont {
			font := c.styles.Fonts[xf.FontId]
			style.Font = Font{}
			style.Font.Size, _ = strconv.Atoi(font.Sz.Val)
			style.Font.Name = font.Name.Val
			style.Font.Family, _ = strconv.Atoi(font.Family.Val)
			style.Font.Charset, _ = strconv.Atoi(font.Charset.Val)
		}
	}
	return style
}

// Row is a high level structure indended to provide user access to a
// row within a xlsx.Sheet.  An xlsx.Row contains a slice of xlsx.Cell.
type Row struct {
	Cells []*Cell
}

// Sheet is a high level structure intended to provide user access to
// the contents of a particular sheet within an XLSX file.
type Sheet struct {
	Rows   []*Row
	MaxRow int
	MaxCol int
}

// Style is a high level structure intended to provide user access to
// the contents of Style within an XLSX file.
type Style struct {
	Border Border
	Fill   Fill
	Font   Font
}

// Border is a high level structure intended to provide user access to
// the contents of Border Style within an Sheet.
type Border struct {
	Left   string
	Right  string
	Top    string
	Bottom string
}

// Fill is a high level structure intended to provide user access to
// the contents of background and foreground color index within an Sheet.
type Fill struct {
	PatternType string
	BgColor     string
	FgColor     string
}

type Font struct {
	Size    int
	Name    string
	Family  int
	Charset int
}

// File is a high level structure providing a slice of Sheet structs
// to the user.
type File struct {
	worksheets     map[string]*zip.File
	referenceTable []string
	styles         *xlsxStyles
	Sheets         []*Sheet          // sheet access by index
	Sheet          map[string]*Sheet // sheet access by name
}

// getRangeFromString is an internal helper function that converts
// XLSX internal range syntax to a pair of integers.  For example,
// the range string "1:3" yield the upper and lower intergers 1 and 3.
func getRangeFromString(rangeString string) (lower int, upper int, error error) {
	var parts []string
	parts = strings.SplitN(rangeString, ":", 2)
	if parts[0] == "" {
		error = errors.New(fmt.Sprintf("Invalid range '%s'\n", rangeString))
	}
	if parts[1] == "" {
		error = errors.New(fmt.Sprintf("Invalid range '%s'\n", rangeString))
	}
	lower, error = strconv.Atoi(parts[0])
	if error != nil {
		error = errors.New(fmt.Sprintf("Invalid range (not integer in lower bound) %s\n", rangeString))
	}
	upper, error = strconv.Atoi(parts[1])
	if error != nil {
		error = errors.New(fmt.Sprintf("Invalid range (not integer in upper bound) %s\n", rangeString))
	}
	return lower, upper, error
}

// lettersToNumeric is used to convert a character based column
// reference to a zero based numeric column identifier.
func lettersToNumeric(letters string) int {
	sum, mul, n := 0, 1, 0
	for i := len(letters) - 1; i >= 0; i, mul, n = i-1, mul*26, 1 {
		c := letters[i]
		switch {
		case 'A' <= c && c <= 'Z':
			n += int(c - 'A')
		case 'a' <= c && c <= 'z':
			n += int(c - 'a')
		}
		sum += n * mul
	}
	return sum
}

// letterOnlyMapF is used in conjunction with strings.Map to return
// only the characters A-Z and a-z in a string
func letterOnlyMapF(rune rune) rune {
	switch {
	case 'A' <= rune && rune <= 'Z':
		return rune
	case 'a' <= rune && rune <= 'z':
		return rune - 32
	}
	return -1
}

// intOnlyMapF is used in conjunction with strings.Map to return only
// the numeric portions of a string.
func intOnlyMapF(rune rune) rune {
	if rune >= 48 && rune < 58 {
		return rune
	}
	return -1
}

// getCoordsFromCellIDString returns the zero based cartesian
// coordinates from a cell name in Excel format, e.g. the cellIDString
// "A1" returns 0, 0 and the "B3" return 1, 2.
func getCoordsFromCellIDString(cellIDString string) (x, y int, error error) {
	var letterPart string = strings.Map(letterOnlyMapF, cellIDString)
	y, error = strconv.Atoi(strings.Map(intOnlyMapF, cellIDString))
	if error != nil {
		return x, y, error
	}
	y -= 1 // Zero based
	x = lettersToNumeric(letterPart)
	return x, y, error
}

// getMaxMinFromDimensionRef return the zero based cartesian maximum
// and minimum coordinates from the dimension reference embedded in a
// XLSX worksheet.  For example, the dimension reference "A1:B2"
// returns "0,0", "1,1".
func getMaxMinFromDimensionRef(ref string) (minx, miny, maxx, maxy int, err error) {
	var parts []string
	parts = strings.Split(ref, ":")
	minx, miny, err = getCoordsFromCellIDString(parts[0])
	if err != nil {
		return -1, -1, -1, -1, err
	}
	if len(parts) == 1 {
		maxx, maxy = minx, miny
		return
	}
	maxx, maxy, err = getCoordsFromCellIDString(parts[1])
	if err != nil {
		return -1, -1, -1, -1, err
	}
	return
}

// makeRowFromSpan will, when given a span expressed as a string,
// return an empty Row large enough to encompass that span and
// populate it with empty cells.  All rows start from cell 1 -
// regardless of the lower bound of the span.
func makeRowFromSpan(spans string) *Row {
	var error error
	var upper int
	var row *Row
	var cell *Cell

	row = new(Row)
	_, upper, error = getRangeFromString(spans)
	if error != nil {
		panic(error)
	}
	error = nil
	row.Cells = make([]*Cell, upper)
	for i := 0; i < upper; i++ {
		cell = new(Cell)
		cell.Value = ""
		row.Cells[i] = cell
	}
	return row
}

// get the max column
// return the cells of columns
func makeRowFromRaw(rawrow xlsxRow) *Row {
	var upper int
	var row *Row
	var cell *Cell

	row = new(Row)
	upper = -1

	for _, rawcell := range rawrow.C {
		x, _, error := getCoordsFromCellIDString(rawcell.R)
		if error != nil {
			panic(fmt.Sprintf("Invalid Cell Coord, %s\n", rawcell.R))
		}
		if x > upper {
			upper = x
		}
	}
	upper++

	row.Cells = make([]*Cell, upper)
	for i := 0; i < upper; i++ {
		cell = new(Cell)
		cell.Value = ""
		row.Cells[i] = cell
	}
	return row
}

// getValueFromCellData attempts to extract a valid value, usable in CSV form from the raw cell value.
// Note - this is not actually general enough - we should support retaining tabs and newlines.
func getValueFromCellData(rawcell xlsxC, reftable []string) string {
	var value string = ""
	var data string = rawcell.V
	if len(data) > 0 {
		vval := strings.Trim(data, " \t\n\r")
		if rawcell.T == "s" {
			ref, error := strconv.Atoi(vval)
			if error != nil {
				panic(error)
			}
			value = reftable[ref]
		} else {
			value = vval
		}
	}
	return value
}

// readRowsFromSheet is an internal helper function that extracts the
// rows from a XSLXWorksheet, poulates them with Cells and resolves
// the value references from the reference table and stores them in
func readRowsFromSheet(Worksheet *xlsxWorksheet, file *File) ([]*Row, int, int) {
	var rows []*Row
	var row *Row
	var minCol, maxCol, minRow, maxRow, colCount, rowCount int
	var reftable []string
	var err error

	if len(Worksheet.SheetData.Row) == 0 {
		return nil, 0, 0
	}
	reftable = file.referenceTable
	minCol, minRow, maxCol, maxRow, err = getMaxMinFromDimensionRef(Worksheet.Dimension.Ref)
	if err != nil {
		panic(err.Error())
	}
	rowCount = (maxRow - minRow) + 1
	colCount = (maxCol - minCol) + 1
	rows = make([]*Row, rowCount)
	for rowIndex := 0; rowIndex < rowCount; rowIndex++ {
		rawrow := Worksheet.SheetData.Row[rowIndex]
		// range is not empty
		if len(rawrow.Spans) != 0 {
			row = makeRowFromSpan(rawrow.Spans)
		} else {
			row = makeRowFromRaw(rawrow)
		}
		for _, rawcell := range rawrow.C {
			x, _, _ := getCoordsFromCellIDString(rawcell.R)
			if x < len(row.Cells) {
				row.Cells[x].Value = getValueFromCellData(rawcell, reftable)
				row.Cells[x].styleIndex = rawcell.S
				row.Cells[x].styles = file.styles
			}
		}
		rows[rowIndex] = row
	}
	return rows, colCount, rowCount
}

type indexedSheet struct {
	Index int
	Sheet *Sheet
	Error error
}

// readSheetFromFile is the logic of converting a xlsxSheet struct
// into a Sheet struct.  This work can be done in parallel and so
// readSheetsFromZipFile will spawn an instance of this function per
// sheet and get the results back on the provided channel.
func readSheetFromFile(sc chan *indexedSheet, index int, rsheet xlsxSheet, fi *File, sheetXMLMap map[string]string) {
	result := &indexedSheet{Index: index, Sheet: nil, Error: nil}
	worksheet, error := getWorksheetFromSheet(rsheet, fi.worksheets, sheetXMLMap)
	if error != nil {
		result.Error = error
		sc <- result
		return
	}
	sheet := new(Sheet)
	sheet.Rows, sheet.MaxCol, sheet.MaxRow = readRowsFromSheet(worksheet, fi)
	result.Sheet = sheet
	sc <- result
}

// readSheetsFromZipFile is an internal helper function that loops
// over the Worksheets defined in the XSLXWorkbook and loads them into
// Sheet objects stored in the Sheets slice of a xlsx.File struct.
func readSheetsFromZipFile(f *zip.File, file *File, sheetXMLMap map[string]string) ([]*Sheet, []string, error) {
	var workbook *xlsxWorkbook
	var error error
	var rc io.ReadCloser
	var decoder *xml.Decoder
	var sheetCount int
	workbook = new(xlsxWorkbook)
	rc, error = f.Open()
	if error != nil {
		return nil, nil, error
	}
	decoder = xml.NewDecoder(rc)
	error = decoder.Decode(workbook)
	if error != nil {
		return nil, nil, error
	}
	sheetCount = len(workbook.Sheets.Sheet)
	sheets := make([]*Sheet, sheetCount)
	names := make([]string, sheetCount)
	sheetChan := make(chan *indexedSheet, sheetCount)
	for i, rawsheet := range workbook.Sheets.Sheet {
		go readSheetFromFile(sheetChan, i, rawsheet, file, sheetXMLMap)
	}
	for j := 0; j < sheetCount; j++ {
		sheet := <-sheetChan
		if sheet.Error != nil {
			return nil, nil, sheet.Error
		}
		sheets[sheet.Index] = sheet.Sheet
		names[sheet.Index] = workbook.Sheets.Sheet[sheet.Index].Name
	}
	return sheets, names, nil
}

// readSharedStringsFromZipFile() is an internal helper function to
// extract a reference table from the sharedStrings.xml file within
// the XLSX zip file.
func readSharedStringsFromZipFile(f *zip.File) ([]string, error) {
	var sst *xlsxSST
	var error error
	var rc io.ReadCloser
	var decoder *xml.Decoder
	var reftable []string
	rc, error = f.Open()
	if error != nil {
		return nil, error
	}
	sst = new(xlsxSST)
	decoder = xml.NewDecoder(rc)
	error = decoder.Decode(sst)
	if error != nil {
		return nil, error
	}
	reftable = MakeSharedStringRefTable(sst)
	return reftable, nil
}

// readStylesFromZipFile() is an internal helper function to
// extract a style table from the style.xml file within
// the XLSX zip file.
func readStylesFromZipFile(f *zip.File) (*xlsxStyles, error) {
	var style *xlsxStyles
	var error error
	var rc io.ReadCloser
	var decoder *xml.Decoder
	rc, error = f.Open()
	if error != nil {
		return nil, error
	}
	style = new(xlsxStyles)
	decoder = xml.NewDecoder(rc)
	error = decoder.Decode(style)
	if error != nil {
		return nil, error
	}
	return style, nil
}

func readWorkbookRelationsFromZipFile(workbook_rels *zip.File) (sheetXMLMap map[string]string) {
	sheetXMLMap = make(map[string]string)
	var wb_relationships *xlsxWorkbookRels
	var error error
	var rc io.ReadCloser
	var decoder *xml.Decoder
	rc, error = workbook_rels.Open()
	if error != nil {
		return
	}
	decoder = xml.NewDecoder(rc)
	wb_relationships = new(xlsxWorkbookRels)
	error = decoder.Decode(wb_relationships)
	if error != nil {
		return
	}
	for _, rel := range wb_relationships.Relationships {
		if strings.HasSuffix(rel.Target, ".xml") && strings.HasPrefix(rel.Target, "worksheets/") {
			sheetXMLMap[rel.Id] = strings.Replace(rel.Target[len("worksheets/"):], ".xml", "", 1)
		}
	}
	return
}

// OpenFile() take the name of an XLSX file and returns a populated
// xlsx.File struct for it.
func OpenFile(filename string) (*File, error) {
	var f *zip.ReadCloser
	f, err := zip.OpenReader(filename)
	if err != nil {
		return nil, err
	}
	return ReadZip(f)
}

func ReadZip(f *zip.ReadCloser) (*File, error) {
	var error error
	var file *File
	var v *zip.File
	var workbook *zip.File
	var workbook_rels *zip.File
	var styles *zip.File
	var sharedStrings *zip.File
	var reftable []string
	var worksheets map[string]*zip.File
	var sheetMap map[string]*Sheet

	file = new(File)
	worksheets = make(map[string]*zip.File, len(f.File))
	for _, v = range f.File {
		switch v.Name {
		case "xl/sharedStrings.xml":
			sharedStrings = v
		case "xl/workbook.xml":
			workbook = v
		case "xl/_rels/workbook.xml.rels":
			workbook_rels = v
		case "xl/styles.xml":
			styles = v
		default:
			if len(v.Name) > 12 {
				if v.Name[0:13] == "xl/worksheets" {
					worksheets[v.Name[14:len(v.Name)-4]] = v
				}
			}
		}
	}
	sheetXMLMap := readWorkbookRelationsFromZipFile(workbook_rels)

	file.worksheets = worksheets
	reftable, error = readSharedStringsFromZipFile(sharedStrings)
	if error != nil {
		return nil, error
	}
	if reftable == nil {
		error := new(XLSXReaderError)
		error.Err = "No valid sharedStrings.xml found in XLSX file"
		return nil, error
	}
	file.referenceTable = reftable
	style, error := readStylesFromZipFile(styles)
	if error != nil {
		return nil, error
	}
	file.styles = style
	sheets, names, error := readSheetsFromZipFile(workbook, file, sheetXMLMap)
	if error != nil {
		return nil, error
	}
	if sheets == nil {
		error := new(XLSXReaderError)
		error.Err = "No sheets found in XLSX File"
		return nil, error
	}
	file.Sheets = sheets
	sheetMap = make(map[string]*Sheet, len(names))
	for i := 0; i < len(names); i++ {
		sheetMap[names[i]] = sheets[i]
	}
	file.Sheet = sheetMap
	f.Close()
	return file, nil
}
