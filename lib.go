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

func (c *Cell) GetStyle() *Style {
	style := new(Style)
	if c.styleIndex > 0 && c.styleIndex < len(c.styles.CellXfs) {
		xf := c.styles.CellXfs[c.styleIndex]
		if xf.ApplyBorder != "0" {
			var border Border
			border.Left = c.styles.Borders[xf.BorderId].Left.Style
			border.Right = c.styles.Borders[xf.BorderId].Right.Style
			border.Top = c.styles.Borders[xf.BorderId].Top.Style
			border.Bottom = c.styles.Borders[xf.BorderId].Bottom.Style
			style.Boders = border
		}
		if xf.ApplyFill != "0" {
			var fill Fill
			fill.BgColorIndex = c.styles.Fills[xf.FillId].BgColorIndex
			style.Fills = fill
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
	Boders Border
	Fills  Fill
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
	BgColorIndex string
	FgColorIndex string
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
	var maxCol int
	var maxRow int
	var reftable []string

	reftable = file.referenceTable
	maxCol = 0
	maxRow = 0
	for _, rawrow := range Worksheet.SheetData.Row {
		for _, rawcell := range rawrow.C {
			x, y, error := getCoordsFromCellIDString(rawcell.R)
			if error != nil {
				panic(fmt.Sprintf("Invalid Cell Coord, %s\n", rawcell.R))
			}
			if x > maxCol {
				maxCol = x
			}
			if y > maxRow {
				maxRow = y
			}
		}
	}
	maxCol += 1
	maxRow += 1
	rows = make([]*Row, maxRow)
	for _, rawrow := range Worksheet.SheetData.Row {
		// range is not empty
		if len(rawrow.Spans) != 0 {
			row = makeRowFromSpan(rawrow.Spans)
		} else {
			row = makeRowFromRaw(rawrow)
		}
		rowno := 0
		for _, rawcell := range rawrow.C {
			x, y, _ := getCoordsFromCellIDString(rawcell.R)
			if y != 0 && rowno == 0 {
				rowno = y
			}
			if x < len(row.Cells) {
				row.Cells[x].Value = getValueFromCellData(rawcell, reftable)
				row.Cells[x].styleIndex = rawcell.S
				row.Cells[x].styles = file.styles
			}
		}
		rows[rowno] = row
	}
	return rows, maxCol, maxRow
}

// readSheetsFromZipFile is an internal helper function that loops
// over the Worksheets defined in the XSLXWorkbook and loads them into
// Sheet objects stored in the Sheets slice of a xlsx.File struct.
func readSheetsFromZipFile(f *zip.File, file *File) ([]*Sheet, []string, error) {
	var workbook *xlsxWorkbook
	var error error
	var rc io.ReadCloser
	var decoder *xml.Decoder
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
	sheets := make([]*Sheet, len(workbook.Sheets.Sheet))
	names := make([]string, len(workbook.Sheets.Sheet))
	for i, rawsheet := range workbook.Sheets.Sheet {
		worksheet, error := getWorksheetFromSheet(rawsheet, file.worksheets)
		if error != nil {
			return nil, nil, error
		}
		sheet := new(Sheet)
		sheet.Rows, sheet.MaxCol, sheet.MaxRow = readRowsFromSheet(worksheet, file)
		sheets[i] = sheet
		names[i] = rawsheet.Name
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
	sheets, names, error := readSheetsFromZipFile(workbook, file)
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
