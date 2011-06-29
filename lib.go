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

// Sheet is a high level structure intended to provide user access to
// the contents of a particular sheet within an XLSX file.
type Sheet struct {

}

// File is a high level structure providing a slice of Sheet structs
// to the user.
type File struct {
	Sheets []*Sheet
}


// readSheetsFromZipFile is an internal helper function that loops
// over the Worksheets defined in the XSLXWorkbook and loads them into
// Sheet objects stored in the Sheets slice of a xlsx.File struct.
func readSheetsFromZipFile(f *zip.File) ([]*Sheet, os.Error) {
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
	for i, _ := range workbook.Sheets.Sheet {
		sheet := new(Sheet)
		sheets[i] = sheet
	}
	return sheets, nil
}

// OpenFile() take the name of an XLSX file and returns a populated
// xlsx.File struct for it.
func OpenFile(filename string) (x *File, e os.Error) {
	var f *zip.ReadCloser
	var error os.Error
	var xlsxFile *File
	var v *zip.File
	f, error = zip.OpenReader(filename)
	if error != nil {
		return nil, error
	}
	xlsxFile = new(File)
	for _, v = range f.File {
		if v.Name == "xl/workbook.xml" {
			sheets, error := readSheetsFromZipFile(v)
			if error != nil {
				return nil, error
			}
			if sheets == nil {
				error := new(XLSXReaderError)
				error.Error = "No sheets found in XLSX File"
				return nil, error
			}
			xlsxFile.Sheets = sheets
		}
	}
	f.Close()
	return xlsxFile, nil
}
