package xlsx

import (
	"archive/zip"
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

// OpenFile() take the name of an XLSX file and returns a populated
// xlsx.File struct for it.
func OpenFile(filename string) (x *xlsxWorkbook, e error) {
	var workbook *zip.File
	var styles *zip.File
	var sharedStrings *zip.File

	f, error := zip.OpenReader(filename)
	if error != nil {
		return nil, error
	}

	book := new(xlsxWorkbook)

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

	error = book.readWorkbookFromZipFile(workbook)
	if error != nil {
		return nil, error
	}

	book.rc = f
	book.xlsxName = filename
	book.worksheets = worksheets

	error = book.readSharedStringsFromZipFile(sharedStrings)
	if error != nil {
		return nil, error
	}

	error = book.readStylesFromZipFile(styles)
	if error != nil {
		return nil, error
	}

	error = book.readWorkbookFromZipFile(workbook)
	if error != nil {
		return nil, error
	}

	return book, nil
}
