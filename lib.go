package xlsx

import (
	"archive/zip"
	"io"
	"os"
	"xml"
)

type XLSXReaderError struct {
	Error string
}

func (e *XLSXReaderError) String() string {
	return e.Error
}

type Sheet struct {

}

type File struct {
	Sheets []*Sheet
}

type FileInterface interface {
	GetSheet(sheetname string) Sheet
}


func MakeSharedStringRefTable(source *XLSXSST) []string {
	reftable := make([]string, len(source.SI))
	for i, si := range source.SI {
		reftable[i] = si.T.Data
	}
	return reftable
}

func ResolveSharedString(reftable []string, index int) string {
	return reftable[index]
}


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
