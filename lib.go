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


// Get the largestDenominator that is a multiple of a basedDenominator
// and fits at least once into a given numerator.
func getLargestDenominator(numerator, multiple, baseDenominator, power int) (int, int) {
	if numerator / multiple == 0 {
		return 1, power
	}
	next, nextPower := getLargestDenominator(
		numerator, multiple * baseDenominator, baseDenominator, power + 1)
	if next > multiple {
		return next, nextPower
	}
	return multiple, power
}

// Convers a list of numbers representing a column into a alphabetic
// representation, as used in the spreadsheet.
func formatColumnName(colId []int) string {
	lastPart := len(colId) - 1

	result := ""
	for n, part := range(colId) {
		if n == lastPart {
			// The least significant number is in the
			// range 0-25, all other numbers are 1-26,
			// hence we use a differente offset for the
			// last part.
			result += string(part + 65)
		} else {
			// Don't output leading 0s, as there is no
			// representation of 0 in this format.
			if part > 0 {
				result += string(part + 64)
			}
		}
	}
	return result
}

func smooshBase26Slice(b26 []int) []int {
	// Smoosh values together, eliminating 0s from all but the
	// least significant part.
	lastButOnePart := len(b26) - 2
	for i := lastButOnePart; i > 0; i-- {
		part := b26[i]
		if part == 0 {
			greaterPart := b26[i-1]
			if greaterPart > 0 {
				b26[i-1] = greaterPart - 1
				b26[i] = 26
			}
		}
	}
	return b26
}

func intToBase26(x int) (parts []int) {
	// Excel column codes are pure evil - in essence they're just
	// base26, but they don't represent the number 0.
	b26Denominator, _ := getLargestDenominator(x, 1, 26, 0)

	// This loop terminates because integer division of 1 / 26
	// returns 0.
	for d := b26Denominator; d > 0; d = d / 26 {
		value := x / d
		remainder := x % d
		parts = append(parts, value)
		x = remainder
	}
	return parts
}

// numericToLetters is used to convert a zero based, numeric column
// indentifier into a character code.
func numericToLetters(colRef int) string {
	parts := intToBase26(colRef)
	return formatColumnName(smooshBase26Slice(parts))
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

// getCellIDStringFromCoords returns the Excel format cell name that
// represents a pair of zero based cartesian coordinates.
func getCellIDStringFromCoords(x, y int) string {
	letterPart := numericToLetters(x);
	numericPart := y + 1
	return fmt.Sprintf("%s%d", letterPart, numericPart)
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

// makeRowFromRaw returns the Row representation of the xlsxRow.
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

// getValueFromCellData attempts to extract a valid value, usable in
// CSV form from the raw cell value.  Note - this is not actually
// general enough - we should support retaining tabs and newlines.
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
	var insertRowIndex, insertColIndex int

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
	insertRowIndex = minRow
	for rowIndex := 0; rowIndex < len(Worksheet.SheetData.Row); rowIndex++ {
		rawrow := Worksheet.SheetData.Row[rowIndex]
		// Some spreadsheets will omit blank rows from the
		// stored data
		for rawrow.R > (insertRowIndex + 1) {
			// Put an empty Row into the array
			rows[insertRowIndex-minRow] = new(Row)
			insertRowIndex++
		}
		// range is not empty
		if len(rawrow.Spans) != 0 {
			row = makeRowFromSpan(rawrow.Spans)
		} else {
			row = makeRowFromRaw(rawrow)
		}

		insertColIndex = minCol
		for _, rawcell := range rawrow.C {
			x, _, _ := getCoordsFromCellIDString(rawcell.R)

			// Some spreadsheets will omit blank cells
			// from the data.
			for x > insertColIndex {
				// Put an empty Cell into the array
				row.Cells[insertColIndex-minCol] = new(Cell)
				insertColIndex++
			}
			cellX := insertColIndex - minCol
			row.Cells[cellX].Value = getValueFromCellData(rawcell, reftable)
			row.Cells[cellX].styleIndex = rawcell.S
			row.Cells[cellX].styles = file.styles
			insertColIndex++
		}
		rows[insertRowIndex-minRow] = row
		insertRowIndex++
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

// readWorkbookRelationsFromZipFile is an internal helper function to
// extract a map of relationship ID strings to the name of the
// worksheet.xml file they refer to.  The resulting map can be used to
// reliably derefence the worksheets in the XLSX file.
func readWorkbookRelationsFromZipFile(workbookRels *zip.File) (map[string]string, error) {
	var sheetXMLMap map[string]string
	var wbRelationships *xlsxWorkbookRels
	var rc io.ReadCloser
	var decoder *xml.Decoder
	var err error

	rc, err = workbookRels.Open()
	if err != nil {
		return nil, err
	}
	decoder = xml.NewDecoder(rc)
	wbRelationships = new(xlsxWorkbookRels)
	err = decoder.Decode(wbRelationships)
	if err != nil {
		return nil, err
	}
	sheetXMLMap = make(map[string]string)
	for _, rel := range wbRelationships.Relationships {
		if strings.HasSuffix(rel.Target, ".xml") && strings.HasPrefix(rel.Target, "worksheets/") {
			sheetXMLMap[rel.Id] = strings.Replace(rel.Target[len("worksheets/"):], ".xml", "", 1)
		}
	}
	return sheetXMLMap, nil
}


// ReadZip() takes a pointer to a zip.ReadCloser and returns a
// xlsx.File struct populated with its contents.  In most cases
// ReadZip is not used directly, but is called internally by OpenFile.
func ReadZip(f *zip.ReadCloser) (*File, error) {
	defer f.Close()
	return ReadZipReader(&f.Reader)
}

// ReadZipReader() can be used to read an XLSX in memory without
// touching thes filesystem.
func ReadZipReader(r *zip.Reader) (*File, error) {
	var err error
	var file *File
	var names []string
	var reftable []string
	var sharedStrings *zip.File
	var sheetMap map[string]*Sheet
	var sheetXMLMap map[string]string
	var sheets []*Sheet
	var style *xlsxStyles
	var styles *zip.File
	var v *zip.File
	var workbook *zip.File
	var workbookRels *zip.File
	var worksheets map[string]*zip.File

	file = new(File)
	worksheets = make(map[string]*zip.File, len(r.File))
	for _, v = range r.File {
		switch v.Name {
		case "xl/sharedStrings.xml":
			sharedStrings = v
		case "xl/workbook.xml":
			workbook = v
		case "xl/_rels/workbook.xml.rels":
			workbookRels = v
		case "xl/styles.xml":
			styles = v
		default:
			if len(v.Name) > 14 {
				if v.Name[0:13] == "xl/worksheets" {
					worksheets[v.Name[14:len(v.Name)-4]] = v
				}
			}
		}
	}
	sheetXMLMap, err = readWorkbookRelationsFromZipFile(workbookRels)
	if err != nil {
		return nil, err
	}
	file.worksheets = worksheets
	reftable, err = readSharedStringsFromZipFile(sharedStrings)
	if err != nil {
		return nil, err
	}
	if reftable == nil {
		readerErr := new(XLSXReaderError)
		readerErr.Err = "No valid sharedStrings.xml found in XLSX file"
		return nil, readerErr
	}
	file.referenceTable = reftable
	style, err = readStylesFromZipFile(styles)
	if err != nil {
		return nil, err
	}
	file.styles = style
	sheets, names, err = readSheetsFromZipFile(workbook, file, sheetXMLMap)
	if err != nil {
		return nil, err
	}
	if sheets == nil {
		readerErr := new(XLSXReaderError)
		readerErr.Err = "No sheets found in XLSX File"
		return nil, readerErr
	}
	file.Sheets = sheets
	sheetMap = make(map[string]*Sheet, len(names))
	for i := 0; i < len(names); i++ {
		sheetMap[names[i]] = sheets[i]
	}
	file.Sheet = sheetMap
	return file, nil
}


