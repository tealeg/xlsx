package xlsx

import (
	"archive/zip"
	"fmt"
	"io"
	"math"
	"os"
	"strconv"
	"strings"
	"xml"
)

// XLSXReaderError is the standard error type for otherwise undefined
// errors in the XSLX reading process.
type XLSXReaderError struct {
	Error string
}

// String() returns a string value from an XLSXReaderError struct in
// order that it might comply with the os.Error interface.
func (e *XLSXReaderError) String() string {
	return e.Error
}


// Cell is a high level structure intended to provide user access to
// the contents of Cell within an xlsx.Row.
type Cell struct {
	data string
}

// CellInterface defines the public API of the Cell.
type CellInterface interface {
	String() string
}

func (c *Cell) String() string {
	return c.data
}

// Row is a high level structure indended to provide user access to a
// row within a xlsx.Sheet.  An xlsx.Row contains a slice of xlsx.Cell.
type Row struct {
	Cells []*Cell
}

// Sheet is a high level structure intended to provide user access to
// the contents of a particular sheet within an XLSX file.
type Sheet struct {
	Rows []*Row
}

// File is a high level structure providing a slice of Sheet structs
// to the user.
type File struct {
	worksheets     map[string]*zip.File
	referenceTable []string
	Sheets         []*Sheet
}


// getRangeFromString is an internal helper function that converts
// XLSX internal range syntax to a pair of integers.  For example,
// the range string "1:3" yield the upper and lower intergers 1 and 3.
func getRangeFromString(rangeString string) (lower int, upper int, error os.Error) {
	var parts []string
	parts = strings.SplitN(rangeString, ":", 2)
	if parts[0] == "" {
		error = os.NewError(fmt.Sprintf("Invalid range '%s'\n", rangeString))
	}
	if parts[1] == "" {
		error = os.NewError(fmt.Sprintf("Invalid range '%s'\n", rangeString))
	}
	lower, error = strconv.Atoi(parts[0])
	if error != nil {
		error = os.NewError(fmt.Sprintf("Invalid range (not integer in lower bound) %s\n", rangeString))
	}
	upper, error = strconv.Atoi(parts[1])
	if error != nil {
		error = os.NewError(fmt.Sprintf("Invalid range (not integer in upper bound) %s\n", rangeString))
	}
	return lower, upper, error
}

// positionalLetterMultiplier gives an integer multiplier to use for a
// position in a letter based column identifer. For example, the
// column ID "AA" is equivalent to 26*1 + 1, "BA" is equivalent to
// 26*2 + 1 and "ABA" is equivalent to (676 * 1)+(26 * 2)+1 or
// ((26**2)*1)+((26**1)*2)+((26**0))*1
func positionalLetterMultiplier(extent, pos int) int {
	var result float64
	var power float64
	var offset int
	offset = pos + 1
	power = float64(extent - offset)
	result = math.Pow(26, power)
	return int(result)
}


// lettersToNumeric is used to convert a character based column
// reference to a zero based numeric column identifier.
func lettersToNumeric(letters string) int {
	var sum int = 0
	var shift int
	extent := len(letters)
	for i, c := range letters {
		// Just to make life akward.  If we think of this base
		// 26 notation as being like HEX or binary we hit a
		// nasty little problem.  The issue is that we have no
		// 0s and therefore A can be both a 1 and a 0.  The
		// value range of a letter is different in the most
		// significant position if (and only if) there is more
		// than one positions.  For example: 
		// "A" = 0 
		//               676 | 26 | 0
		//               ----+----+----
		//                 0 |  0 | 0
		// 
		//  "Z" = 25
		//                676 | 26 | 0
		//                ----+----+----
		//                  0 |  0 |  25
		//   "AA" = 26
		//                676 | 26 | 0
		//                ----+----+----
		//                  0 |  1 | 0     <--- note here - the value of "A" maps to both 1 and 0.  
		if i == 0 && extent > 1 {
			shift = 1
		} else {
			shift = 0
		}
		multiplier := positionalLetterMultiplier(extent, i)
		switch {
		case 'A' <= c && c <= 'Z':
			sum += multiplier * ((c - 'A') + shift)
		case 'a' <= c && c <= 'z':
			sum += multiplier * ((c - 'a') + shift)
		}
	}
	return sum
}


// letterOnlyMapF is used in conjunction with strings.Map to return
// only the characters A-Z and a-z in a string
func letterOnlyMapF(rune int) int {
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
func intOnlyMapF(rune int) int {
	if rune >= 48 && rune < 58 {
		return rune
	}
	return -1
}


// getCoordsFromCellIDString returns the zero based cartesian
// coordinates from a cell name in Excel format, e.g. the cellIDString
// "A1" returns 0, 0 and the "B3" return 1, 2.
func getCoordsFromCellIDString(cellIDString string) (x, y int, error os.Error) {
	var letterPart string = strings.Map(letterOnlyMapF, cellIDString)
	y, error = strconv.Atoi(strings.Map(intOnlyMapF, cellIDString))
	if error != nil {
		return x, y, error
	}
	y -= 1 // Zero based
	x = lettersToNumeric(letterPart)
	return x, y, error
}


// readRowsFromSheet is an internal helper function that extracts the
// rows from a XSLXWorksheet, poulates them with Cells and resolves
// the value references from the reference table and stores them in
func readRowsFromSheet(worksheet *XLSXWorksheet, reftable []string) []*Row {

	// Note, this function needs tidying up!
	var rows []*Row
	var error os.Error
	var upper int
	var row *Row
	var cell *Cell

	rows = make([]*Row, len(worksheet.SheetData.Row))
	for i, rawrow := range worksheet.SheetData.Row {
		row = new(Row)
		_, upper, error = getRangeFromString(rawrow.Spans)
		if error != nil {
			panic(error)
		}
		error = nil
		row.Cells = make([]*Cell, upper)
		for i := 0; i < upper; i++ {
			cell = new(Cell)
			cell.data = ""
			row.Cells[i] = cell
		}
		for _, rawcell := range rawrow.C {
			x, _, error := getCoordsFromCellIDString(rawcell.R)
			if error != nil {
				panic(fmt.Sprintf("Invalid Cell Coord, %s\n", rawcell.R))
			}
			error = nil
			cell = new(Cell)
			cell.data = ""
			if len(rawcell.V.Data) > 0 {
				vval := strings.Trim(rawcell.V.Data, " \t\n\r")
				if rawcell.T == "s" {
					ref, error := strconv.Atoi(vval)
					if error != nil {
						panic(error)
						panic(fmt.Sprintf("Invalid reference in Excel Cell (not found in sharedStrings.xml) - the reference was %v\n", rawcell.V.Data))
					}
					cell.data = reftable[ref]
				} else {
					cell.data = vval
				}

			}
			row.Cells[x] = cell
		}
		rows[i] = row
	}
	return rows
}


// readSheetsFromZipFile is an internal helper function that loops
// over the Worksheets defined in the XSLXWorkbook and loads them into
// Sheet objects stored in the Sheets slice of a xlsx.File struct.
func readSheetsFromZipFile(f *zip.File, file *File) ([]*Sheet, os.Error) {
	var workbook *XLSXWorkbook
	var error os.Error
	var rc io.ReadCloser
	workbook = new(XLSXWorkbook)
	rc, error = f.Open()
	if error != nil {
		return nil, error
	}
	error = xml.Unmarshal(rc, workbook)
	if error != nil {
		return nil, error
	}
	sheets := make([]*Sheet, len(workbook.Sheets.Sheet))
	for i, rawsheet := range workbook.Sheets.Sheet {
		worksheet, error := getWorksheetFromSheet(rawsheet, file.worksheets) // 
		if error != nil {
			return nil, error
		}
		sheet := new(Sheet)
		sheet.Rows = readRowsFromSheet(worksheet, file.referenceTable)
		sheets[i] = sheet
	}
	return sheets, nil
}


// readSharedStringsFromZipFile() is an internal helper function to
// extract a reference table from the sharedStrings.xml file within
// the XLSX zip file.
func readSharedStringsFromZipFile(f *zip.File) ([]string, os.Error) {
	var sst *XLSXSST
	var error os.Error
	var rc io.ReadCloser
	var reftable []string
	rc, error = f.Open()
	if error != nil {
		return nil, error
	}
	sst = new(XLSXSST)
	error = xml.Unmarshal(rc, sst)
	if error != nil {
		return nil, error
	}
	reftable = MakeSharedStringRefTable(sst)
	return reftable, nil
}

// OpenFile() take the name of an XLSX file and returns a populated
// xlsx.File struct for it.
func OpenFile(filename string) (x *File, e os.Error) {
	var f *zip.ReadCloser
	var error os.Error
	var file *File
	var v *zip.File
	var workbook *zip.File
	var sharedStrings *zip.File
	var reftable []string
	var worksheets map[string]*zip.File
	f, error = zip.OpenReader(filename)
	if error != nil {
		return nil, error
	}
	file = new(File)
	worksheets = make(map[string]*zip.File, len(f.File))
	for _, v = range f.File {
		switch v.Name {
		case "xl/sharedStrings.xml":
			sharedStrings = v
		case "xl/workbook.xml":
			workbook = v
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
		error.Error = "No valid sharedStrings.xml found in XLSX file"
		return nil, error
	}
	file.referenceTable = reftable
	sheets, error := readSheetsFromZipFile(workbook, file)
	if error != nil {
		return nil, error
	}
	if sheets == nil {
		error := new(XLSXReaderError)
		error.Error = "No sheets found in XLSX File"
		return nil, error
	}
	file.Sheets = sheets
	f.Close()
	return file, nil
}
