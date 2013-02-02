package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"math"
	"os"
	"strconv"
	"strings"
)

const (
	// excel xml header
	Header = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
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

// getColorFromIndexed is used to convert a color Indexed 
// to the ARGB color
func getColorFromIndexed(colorIndexed string) string {
	switch colorIndexed {
	case "0":
		return "00000000"
	case "1":
		return "00FFFFFF"
	case "2":
		return "00FF0000"
	case "3":
		return "0000FF00"
	case "4":
		return "000000FF"
	case "5":
		return "00FFFF00"
	case "6":
		return "00FF00FF"
	case "7":
		return "0000FFFF"
	case "8":
		return "00000000"
	case "9":
		return "00FFFFFF"
	case "10":
		return "00FF0000"
	case "11":
		return "0000FF00"
	case "12":
		return "000000FF"
	case "13":
		return "00FFFF00"
	case "14":
		return "00FF00FF"
	case "15":
		return "0000FFFF"
	case "16":
		return "00800000"
	case "17":
		return "00008000"
	case "18":
		return "00000080"
	case "19":
		return "00808000"
	case "20":
		return "00800080"
	case "21":
		return "00008080"
	case "22":
		return "00C0C0C0"
	case "23":
		return "00808080"
	case "24":
		return "009999FF"
	case "25":
		return "00993366"
	case "26":
		return "00FFFFCC"
	case "27":
		return "00CCFFFF"
	case "28":
		return "00660066"
	case "29":
		return "00FF8080"
	case "30":
		return "000066CC"
	case "31":
		return "00CCCCFF"
	case "32":
		return "00000080"
	case "33":
		return "00FF00FF"
	case "34":
		return "00FFFF00"
	case "35":
		return "0000FFFF"
	case "36":
		return "00800080"
	case "37":
		return "00800000"
	case "38":
		return "00008080"
	case "39":
		return "000000FF"
	case "40":
		return "0000CCFF"
	case "41":
		return "00CCFFFF"
	case "42":
		return "00CCFFCC"
	case "43":
		return "00FFFF99"
	case "44":
		return "0099CCFF"
	case "45":
		return "00FF99CC"
	case "46":
		return "00CC99FF"
	case "47":
		return "00FFCC99"
	case "48":
		return "003366FF"
	case "49":
		return "0033CCCC"
	case "50":
		return "0099CC00"
	case "51":
		return "00FFCC00"
	case "52":
		return "00FF9900"
	case "53":
		return "00FF6600"
	case "54":
		return "00666699"
	case "55":
		return "00969696"
	case "56":
		return "00003366"
	case "57":
		return "00339966"
	case "58":
		return "00003300"
	case "59":
		return "00333300"
	case "60":
		return "00993300"
	case "61":
		return "00993366"
	case "62":
		return "00333399"
	case "63":
		return "00333333"
	case "64":
		return "" // System Foreground
	case "65":
		return "" // System Background
	}
	return ""
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

// File is a high level structure providing a slice of Sheet structs
// to the user.
type File struct {
	workbookinfo  *xlsxWorkbook             // book date in xml struct
	xlsxsheetinfo map[string]*xlsxWorksheet // sheet date in xml struct
	styleinfo     *xlsxStyles               // styles data in xml struct
	sstinfo       *xlsxSST                  // shared strings data in xml struct
	worksheets    map[string]*zip.File      // xml files
	xlsxName      string
	rc            *zip.ReadCloser

	referenceTable []string          // share string data
	sheet          map[string]*Sheet // sheet access by name
}

// getRangeFromString is an internal helper function that converts
// XLSX internal range syntax to a pair of integers.  For example,
// the range string "1:3" yield the upper and lower intergers 1 and 3.
func getRangeFromString(rangeString string) (lower int, upper int, error error) {
	parts := strings.SplitN(rangeString, ":", 2)
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

// positionalLetterMultiplier gives an integer multiplier to use for a
// position in a letter based column identifer. For example, the
// column ID "AA" is equivalent to 26*1 + 1, "BA" is equivalent to
// 26*2 + 1 and "ABA" is equivalent to (676 * 1)+(26 * 2)+1 or
// ((26**2)*1)+((26**1)*2)+((26**0))*1
func positionalLetterMultiplier(extent, pos int) int {
	offset := pos + 1
	power := float64(extent - offset)
	result := math.Pow(26, power)
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
			sum += multiplier * (int((c - 'A')) + shift)
		case 'a' <= c && c <= 'z':
			sum += multiplier * (int((c - 'a')) + shift)
		}
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
	row := new(Row)
	_, upper, error := getRangeFromString(spans)
	if error != nil {
		panic(error)
	}
	error = nil
	row.Cells = make([]*Cell, upper)
	for i := 0; i < upper; i++ {
		cell := new(Cell)
		cell.Value = ""
		row.Cells[i] = cell
	}
	return row
}

// get the max column 
// return the cells of columns
func makeRowFromRaw(rawrow xlsxRow) *Row {
	row := new(Row)
	upper := 0

	for _, rawcell := range rawrow.C {
		x, _, error := getCoordsFromCellIDString(rawcell.R)
		if error != nil {
			panic(fmt.Sprintf("Invalid Cell Coord, %s\n", rawcell.R))
		}
		if x > upper {
			upper = x
		}
	}

	row.Cells = make([]*Cell, upper)
	for i := 0; i < upper; i++ {
		cell := new(Cell)
		cell.Value = ""
		row.Cells[i] = cell
	}
	return row
}

// getValueFromCellData attempts to extract a valid value, usable in CSV form from the raw cell value.
// Note - this is not actually general enough - we should support retaining tabs and newlines.
func getValueFromCellData(rawcell xlsxC, reftable []string) string {
	var value string = ""

	data := rawcell.V
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
	var row *Row

	reftable := file.referenceTable
	maxCol := 0
	maxRow := 0
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
	rows := make([]*Row, maxRow)
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
			row.Cells[x].Value = getValueFromCellData(rawcell, reftable)
			row.Cells[x].styleIndex = rawcell.S
			row.Cells[x].styles = file.styleinfo
		}
		rows[rowno] = row
	}
	return rows, maxCol, maxRow
}

func (book *File) Sheet(name string) (sh *Sheet) {
	if book.sheet[name] == nil {
		error := book.readSheetFromZipFile(name)
		if error != nil {
			panic(error)
		}
	}

	// debug start
	// output, err := xml.MarshalIndent(book.xlsxsheetinfo[name], "  ", "    ")
	// if err != nil {
	// 	fmt.Printf("error: %v\n", err)
	// }
	// fout,error := os.Create("out.xml")
	// if error != nil {
	// 	fmt.Printf("error: %v\n", err)
	// 	return
	// }
	// defer fout.Close()
	// fout.WriteString(string(output)+"\n")
	// debug end

	return book.sheet[name]
}

// read date stored in the sheet by name
// return error
func (f *File) readSheetFromZipFile(name string) error {
	var shxmlname string
	for _, sheet := range f.workbookinfo.Sheets.Sheet {
		if sheet.Name == name {
			shxmlname = fmt.Sprintf("sheet%s", sheet.Id[3:])
		}
	}
	worksheet := new(xlsxWorksheet)
	shxml := f.worksheets[shxmlname]
	rc, error := shxml.Open()
	if error != nil {
		return error
	}
	decoder := xml.NewDecoder(rc)
	error = decoder.Decode(worksheet)
	if f.xlsxsheetinfo == nil {
		f.xlsxsheetinfo = make(map[string]*xlsxWorksheet)
	}
	worksheet.sst = f.sstinfo
	f.xlsxsheetinfo[name] = worksheet

	sheet := new(Sheet)
	sheet.Rows, sheet.MaxCol, sheet.MaxRow = readRowsFromSheet(worksheet, f)
	if f.sheet == nil {
		f.sheet = make(map[string]*Sheet)
	}
	f.sheet[name] = sheet
	return nil
}

// readSharedStringsFromZipFile() is an internal helper function to
// extract a reference table from the sharedStrings.xml file within
// the XLSX zip file.
func (f *File) readSharedStringsFromZipFile(sstxml *zip.File) error {
	rc, error := sstxml.Open()
	if error != nil {
		return error
	}
	sst := new(xlsxSST)
	decoder := xml.NewDecoder(rc)
	error = decoder.Decode(sst)
	if error != nil {
		return error
	}
	f.sstinfo = sst

	// debug start
	// output, err := xml.MarshalIndent(sst, "  ", "    ")
	// if err != nil {
	// 	fmt.Printf("error: %v\n", err)
	// }
	// fout,error := os.Create("sstout.xml")
	// if error != nil {
	// 	fmt.Printf("error: %v\n", err)
	// 	return error
	// }
	// defer fout.Close()
	// fout.WriteString(string(output)+"\n")
	// debug end

	reftable := MakeSharedStringRefTable(sst)
	f.referenceTable = reftable
	return nil
}

// readStylesFromZipFile() is an internal helper function to
// extract a style table from the style.xml file within
// the XLSX zip file.
func (f *File) readStylesFromZipFile(stylexml *zip.File) error {
	rc, error := stylexml.Open()
	if error != nil {
		return error
	}
	style := new(xlsxStyles)
	decoder := xml.NewDecoder(rc)
	error = decoder.Decode(style)
	if error != nil {
		return error
	}
	f.styleinfo = style
	return nil
}

// readWorkbookFromZipFile() is an internal helper function to
// extract a workbook from the workbook.xml file within
// the XLSX zip file.
func (f *File) readWorkbookFromZipFile(bookxml *zip.File) error {
	rc, error := bookxml.Open()
	if error != nil {
		return error
	}
	book := new(xlsxWorkbook)
	decoder := xml.NewDecoder(rc)
	error = decoder.Decode(book)
	if error != nil {
		return error
	}
	f.workbookinfo = book
	return nil
}

// close workbook
func (f *File) Close() {
	if f.rc != nil {
		f.rc.Close()
		f.rc = nil
	}
}

// OpenFile() take the name of an XLSX file and returns a populated
// xlsx.File struct for it.
func OpenFile(filename string) (x *File, e error) {
	var workbook *zip.File
	var styles *zip.File
	var sharedStrings *zip.File

	f, error := zip.OpenReader(filename)
	if error != nil {
		return nil, error
	}

	file := new(File)
	file.rc = f
	file.xlsxName = filename

	worksheets := make(map[string]*zip.File, len(f.File))
	for _, v := range f.File {
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

	error = file.readSharedStringsFromZipFile(sharedStrings)
	if error != nil {
		return nil, error
	}
	if file.referenceTable == nil {
		error := new(XLSXReaderError)
		error.Err = "No valid sharedStrings.xml found in XLSX file"
		return nil, error
	}

	error = file.readStylesFromZipFile(styles)
	if error != nil {
		return nil, error
	}

	error = file.readWorkbookFromZipFile(workbook)
	if error != nil {
		return nil, error
	}

	return file, nil
}

// save the workbook that has modified
func (file *File) Save(xlsxPath string) error {
	xlsxFile, err := os.Create(xlsxPath)
	if err != nil {
		return err
	}
	defer xlsxFile.Close()
	newXlsxZip := zip.NewWriter(xlsxFile)
	defer newXlsxZip.Close()
	oldZip, err := zip.OpenReader(file.xlsxName)
	if err != nil {
		return err
	}

	for _, oldf := range oldZip.File {
		if oldf.Name != "xl/sharedStrings.xml" {
			newf, err := newXlsxZip.Create(oldf.Name)
			if err != nil {
				return err
			}
			oldfrd, err := oldf.Open()
			if err != nil {
				return err
			}
			_, err = io.Copy(newf, oldfrd)
			if err != nil {
				return err
			}
		}
	}

	newSstFile, err := newXlsxZip.Create("xl/sharedStrings.xml")
	if err != nil {
		return err
	}
	err = file.sstinfo.WriteTo(newSstFile)
	if err != nil {
		return err
	}
	// for i, sheet := range this.Sheets{
	// 	fileName := fmt.Sprintf("xl/worksheets/sheet%d.xml", i+1)
	// 	sheetXml, err := newXlsxZip.Create(fileName)
	// 	if err != nil{
	// 		return nil
	// 	}
	// 	err = sheet.WriteTo(sheetXml)
	// 	if err != nil{
	// 		return nil
	// 	}
	// }
	return nil
}
