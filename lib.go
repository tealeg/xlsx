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

type CellFilter func(cell Cell) bool

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
	formula    string
	styleIndex int
	styles     *xlsxStyles
}

// CellInterface defines the public API of the Cell.
type CellInterface interface {
	String() string
}

// String returns the value of a Cell as a string.
func (c *Cell) String() string {
	return c.Value
}

// String returns the formula of a Cell as a string.
func (c *Cell) Formula() string {
	return c.formula
}

// GetStyle returns the Style associated with a Cell
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
	Height float64
}

// zero-based cell index
type CellCoord struct {
	X int
	Y int
}

// Sheet is a high level structure intended to provide user access to
// the contents of a particular sheet within an XLSX file.
type Sheet struct {
	Name             string
	Cells            map[CellCoord]Cell
	Rows             map[int]Row
	MaxRow           int
	MaxCol           int
	DefaultRowHeight float64
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

// getCoordsFromCellIDString returns the zero based cartesian
// coordinates from a cell name in Excel format, e.g. the cellIDString
// "A1" returns 0, 0 and the "B3" return 1, 2.
func getCoordsFromCellIDRunes(cellIDString []rune) (x, y int, error error) {
	for i, v := range cellIDString {
		if !(v >= 'A' && v <= 'Z') {
			if i == 0 {
				return 0, 0, errors.New("no alphanum in rune array")
			}
			x := lettersToNumeric(string(cellIDString[:i]))
			y, error := strconv.Atoi(string(cellIDString[i:]))
			if error != nil {
				return x, y, error
			}
			y -= 1 // Zero based
			return x, y, nil
		}
	}
	return 0, 0, errors.New("no number in rune array")
}

func reverseRunesInPlace(runes []rune) {
	for i, j := 0, len(runes)-1; i < j; i, j = i+1, j-1 {
		near := runes[j]
		far := runes[i]
		runes[i] = near
		runes[j] = far
	}
}

func coordsToCellIDRunes(x, y int) []rune {
	var itoa []rune
	y += 1
	for {
		itoa = append(itoa, rune('0'+y%10))
		y /= 10
		if y == 0 {
			break
		}
	}
	reverseRunesInPlace(itoa)
	var retval []rune
	x += 1
	for x > 0 {
		rem := (x - 1) % 26
		retval = append(retval, rune('A'+rem))
		x -= rem
		x /= 26
	}
	reverseRunesInPlace(retval)
	return append(retval, itoa...)
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

// calculateMaxMinFromWorkSheet works out the dimensions of a spreadsheet
// that doesn't have a DimensionRef set.  The only case currently
// known where this is true is with XLSX exported from Google Docs.
func calculateMaxMinFromWorksheet(worksheet *xlsxWorksheet) (minx, miny, maxx, maxy int, err error) {
	// Note, this method could be very slow for large spreadsheets.
	var x, y int
	minx = 0
	miny = 0
	maxy = 0
	maxx = 0
	for _, row := range worksheet.SheetData.Row {
		for _, cell := range row.C {
			x, y, err = getCoordsFromCellIDString(cell.R)
			if err != nil {
				return -1, -1, -1, -1, err
			}
			if x < minx {
				minx = x
			}
			if x > maxx {
				maxx = x
			}
			if y < miny {
				miny = y
			}
			if y > maxy {
				maxy = y
			}
		}
	}
	return
}

// getValueFromCellData attempts to extract a valid value, usable in CSV form from the raw cell value.
// Note - this is not actually general enough - we should support retaining tabs and newlines.
func getValueFromCellData(rawcell xlsxC, reftable []string) string {
	var value string = ""
	var vval string = rawcell.V
	if len(vval) > 0 {
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
func isDelimiter(v rune) bool {
	if v >= 'A' && v <= 'Z' {
		return false
	}
	if v >= '0' && v <= '9' {
		return false
	}
	if v >= 'a' && v <= 'z' {
		return false
	}
	return true
}

func fixupFormulaToken(candidate []rune, dx int, dy int) []rune {
	candidateX, candidateY, error := getCoordsFromCellIDRunes(candidate)
	if error != nil {
		return candidate
	}
	candidateX += dx
	candidateY += dy
	return coordsToCellIDRunes(candidateX, candidateY)
}

func updateFormula(sharedFormula xlsxSharedFormula, cellX int, cellY int) string {
	dx := cellX - sharedFormula.cellX
	dy := cellY - sharedFormula.cellY
	var formula []rune
	var candidate []rune
	quoted := false
	for _, v := range sharedFormula.F {
		if isDelimiter(v) {
			if len(candidate) > 0 {
				formula = append(formula, fixupFormulaToken(candidate, dx, dy)...)
				candidate = candidate[0:0]
			}
			formula = append(formula, v)
		} else if quoted {
			formula = append(formula, v)
		} else {
			candidate = append(candidate, v)
		}
		if v == '"' {
			quoted = !quoted
		}
	}

	return string(append(formula, fixupFormulaToken(candidate, dx, dy)...))
}

func getFormulaFromCellData(rawcell xlsxC, cellX int, cellY int, si map[string]xlsxSharedFormula) string {
	var value string = ""
	var fval string = rawcell.F.F
	if len(fval) > 0 {
		value = fval
	}
	if len(rawcell.F.Si) > 0 && rawcell.F.T == "shared" {
		fvalSi := rawcell.F.Si
		if len(fval) > 0 {
			si[fvalSi] = xlsxSharedFormula{fval, rawcell.F.Ref, cellX, cellY}
		} else {
			sharedFormula, ok := si[fvalSi]
			if ok {
				value = updateFormula(sharedFormula, cellX, cellY)
			}
		}
	}
	return value
}

// readRowsFromSheet is an internal helper function that extracts the
// rows from a XSLXWorksheet, poulates them with Cells and resolves
// the value references from the reference table and stores them in
func readRowsFromSheet(Worksheet *xlsxWorksheet, file *File, si map[string]xlsxSharedFormula, cellFilter CellFilter) Sheet {
	var maxCol, maxRow, colCount, rowCount int
	var reftable []string
	var err error
	var insertRowIndex, insertColIndex int

	reftable = file.referenceTable
	if len(Worksheet.Dimension.Ref) > 0 {
		_, _, maxCol, maxRow, err = getMaxMinFromDimensionRef(Worksheet.Dimension.Ref)
	} else {
		_, _, maxCol, maxRow, err = calculateMaxMinFromWorksheet(Worksheet)
	}
	if err != nil {
		panic(err.Error())
	}
	rowCount = maxRow + 1
	colCount = maxCol + 1
	cells := make(map[CellCoord]Cell)
	rows := make(map[int]Row)
	insertRowIndex = 0
	for rowIndex := 0; rowIndex < len(Worksheet.SheetData.Row); rowIndex++ {
		rawrow := Worksheet.SheetData.Row[rowIndex]
		// Some spreadsheets will omit blank rows from the
		// stored data
		if insertRowIndex < rawrow.R {
			insertRowIndex = rawrow.R - 1
		}
		if rawrow.CustomHeight != 0 {
			rows[insertRowIndex] = Row{rawrow.Ht}
		}
		// range is not empty
		insertColIndex = 0
		for _, rawcell := range rawrow.C {
			x, _, error := getCoordsFromCellIDString(rawcell.R)
			if error == nil {
				insertColIndex = x
			}
			var cell Cell
			cell.Value = getValueFromCellData(rawcell, reftable)
			cell.formula = getFormulaFromCellData(rawcell, insertColIndex, insertRowIndex, si)
			cell.styleIndex = rawcell.S
			cell.styles = file.styles
			if cellFilter(cell) {
				cells[CellCoord{insertColIndex, insertRowIndex}] = cell
			}
			insertColIndex++
		}
		insertRowIndex++
	}
	var sheet Sheet
	sheet.Cells = cells
	sheet.Rows = rows
	sheet.MaxRow = rowCount
	sheet.MaxCol = colCount
	sheet.DefaultRowHeight = Worksheet.SheetFormatPr.DefaultRowHeight
	return sheet
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
func readSheetFromFile(sc chan *indexedSheet, index int, rsheet xlsxSheet, fi *File, sheetXMLMap map[string]string, cellFilter CellFilter) {
	result := &indexedSheet{Index: index, Sheet: nil, Error: nil}
	worksheet, error := getWorksheetFromSheet(rsheet, fi.worksheets, sheetXMLMap)
	if error != nil {
		result.Error = error
		sc <- result
		return
	}
	siIndex := make(map[string]xlsxSharedFormula)
	sheet := readRowsFromSheet(worksheet, fi, siIndex, cellFilter)
	result.Sheet = &sheet
	sc <- result
}

// readSheetsFromZipFile is an internal helper function that loops
// over the Worksheets defined in the XSLXWorkbook and loads them into
// Sheet objects stored in the Sheets slice of a xlsx.File struct.
func readSheetsFromZipFile(f *zip.File, file *File, sheetXMLMap map[string]string, cellFilter CellFilter) ([]*Sheet, error) {
	var workbook *xlsxWorkbook
	var error error
	var rc io.ReadCloser
	var decoder *xml.Decoder
	var sheetCount int
	workbook = new(xlsxWorkbook)
	rc, error = f.Open()
	if error != nil {
		return nil, error
	}
	decoder = xml.NewDecoder(rc)
	error = decoder.Decode(workbook)
	if error != nil {
		return nil, error
	}
	sheetCount = len(workbook.Sheets.Sheet)
	sheets := make([]*Sheet, sheetCount)
	sheetChan := make(chan *indexedSheet, sheetCount)
	for i, rawsheet := range workbook.Sheets.Sheet {
		go readSheetFromFile(sheetChan, i, rawsheet, file, sheetXMLMap, cellFilter)
	}
	for j := 0; j < sheetCount; j++ {
		sheet := <-sheetChan
		if sheet.Error != nil {
			return nil, sheet.Error
		}
		sheet.Sheet.Name = workbook.Sheets.Sheet[sheet.Index].Name
		sheets[sheet.Index] = sheet.Sheet
	}
	return sheets, nil
}

// readSharedStringsFromZipFile() is an internal helper function to
// extract a reference table from the sharedStrings.xml file within
// the XLSX zip file.
func readSharedStringsFromZipFile(f *zip.File) ([]string, error) {
	if f == nil {
		return []string{}, nil
	}
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

func HashT(cell Cell) bool {
	return true
}

// OpenFile() take the name of an XLSX file and returns a populated
// xlsx.File struct for it.
func OpenFileFilter(filename string, cellFilter CellFilter) (*File, error) {
	var f *zip.ReadCloser
	f, err := zip.OpenReader(filename)
	if err != nil {
		return nil, err
	}
	return ReadZip(f, cellFilter)
}

func OpenFile(filename string) (*File, error) {
	return OpenFileFilter(filename, HashT)
}

// ReadZip() takes a pointer to a zip.ReadCloser and returns a
// xlsx.File struct populated with its contents.  In most cases
// ReadZip is not used directly, but is called internally by OpenFile.
func ReadZip(f *zip.ReadCloser, cellFilter CellFilter) (*File, error) {
	defer f.Close()
	return ReadZipReader(&f.Reader, cellFilter)
}

// ReadZipReader() can be used to read xlsx in memory without touch filesystem.
func ReadZipReader(r *zip.Reader, cellFilter CellFilter) (*File, error) {
	var err error
	var file *File
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
	sheets, err = readSheetsFromZipFile(workbook, file, sheetXMLMap, cellFilter)
	if err != nil {
		return nil, err
	}
	if sheets == nil {
		readerErr := new(XLSXReaderError)
		readerErr.Err = "No sheets found in XLSX File"
		return nil, readerErr
	}
	file.Sheets = sheets
	sheetMap = make(map[string]*Sheet, len(sheets))
	for i := 0; i < len(sheets); i++ {
		sheetMap[sheets[i].Name] = sheets[i]
	}
	file.Sheet = sheetMap
	return file, nil
}

func NewFile() *File {
	return &File{}
}
