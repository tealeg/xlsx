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

type XLSXV struct {
	Data string "chardata"
}

type XLSXC struct {
	R string "attr"
	T string "attr"
	V XLSXV
}

type XLSXRow struct {
	R     string "attr"
	Spans string "attr"
	C     []XLSXC
}

type XLSXSheetData struct {
	Row []XLSXRow
}

type XLSXSheetFormatPr struct {
	BaseColWidth     string "attr"
	DefaultRowHeight string "attr"
}

type XLSXSelection struct {
	ActiveCell string "attr"
	SQRef      string "attr"
}

type XLSXSheetView struct {
	TabSelected    string "attr"
	WorkbookViewID string "attr"
	Selection      XLSXSelection
}

type XLSXSheetViews struct {
	SheetView []XLSXSheetView
}

type XLSXDimension struct {
	Ref string "attr"
}

type XLSXWorksheet struct {
	Dimension     XLSXDimension
	SheetViews    XLSXSheetViews
	SheetFormatPr XLSXSheetFormatPr
	SheetData     XLSXSheetData
}

type XLSXT struct {
	Data string "chardata"
}

type XLSXSI struct {
	T XLSXT
}

type XLSXSST struct {
	Count       string "attr"
	UniqueCount string "attr"
	SI          []XLSXSI
}











type XLSXSheetStruct struct {

}

type XLSXFile struct {
	Sheets []*XLSXSheetStruct
}

type XLSXFileInterface interface {
	GetSheet(sheetname string) XLSXSheetStruct
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


func readSheetsFromZipFile(f *zip.File) ([]*XLSXSheetStruct, os.Error) {
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
	sheets := make([]*XLSXSheetStruct, len(workbook.Sheets.Sheet))
	for i, _ := range workbook.Sheets.Sheet {
		sheet := new(XLSXSheetStruct)
		sheets[i] = sheet
	}
	return sheets, nil
}

func OpenXLSXFile(filename string) (x *XLSXFile, e os.Error) {
	var f *zip.ReadCloser
	var error os.Error
	var xlsxFile *XLSXFile
	var v *zip.File
	f, error = zip.OpenReader(filename)
	if error != nil {
		return nil, error
	}
	xlsxFile = new(XLSXFile)
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
