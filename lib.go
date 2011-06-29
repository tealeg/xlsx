package xlsx

import (
	"archive/zip"
	"io"
	"os"
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
	worksheets map[string] *zip.File
	referenceTable []string
	Sheets []*Sheet
}


// readRowsFromSheet is an internal helper function that extracts the
// rows from a XSLXWorksheet, poulates them with Cells and resolves
// the value references from the reference table and stores them in
func readRowsFromSheet(worksheet *XLSXWorksheet, reftable []string) []*Row {
	var rows []*Row
	rows = make([]*Row, len(worksheet.SheetData.Row))
	for i, rawrow := range worksheet.SheetData.Row {
		row := new(Row)
		row.Cells = make([]*Cell, len(rawrow.C))
		for j, rawcell := range rawrow.C {
			cell := new(Cell)
			cell.data = rawcell.V.Data
			row.Cells[j] = cell
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
		worksheet, error := getWorksheetFromSheet(rawsheet, file.worksheets)
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
					worksheets[v.Name[14:len(v.Name)-4]]= v
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
