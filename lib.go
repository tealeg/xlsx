package xlsx

import (
	"archive/zip"
	"io"
	"os"
	"xml"

)

type XLSXV struct {
	Data string "chardata"
}

type XLSXC struct {
	R string "attr"
	T string "attr"
	V XLSXV
}

type XLSXRow struct {
	R string "attr"
	Spans string "attr"
	C []XLSXC
}

type XLSXSheetData struct {
	Row []XLSXRow
}

type XLSXSheetFormatPr struct {
	BaseColWidth string "attr"
	DefaultRowHeight string "attr"
}

type XLSXSelection struct {
	ActiveCell string "attr"
	SQRef string "attr"
}

type XLSXSheetView struct {
	TabSelected string "attr"
	WorkbookViewID string "attr"
	Selection XLSXSelection
}

type XLSXSheetViews struct {
	SheetView []XLSXSheetView
}

type XLSXDimension struct {
	Ref string "attr"
}

type XLSXWorksheet struct {
	Dimension XLSXDimension
	SheetViews XLSXSheetViews
	SheetFormatPr XLSXSheetFormatPr
	SheetData XLSXSheetData
}

type XLSXT struct {
	Data string "chardata"
}

type XLSXSI struct {
	T XLSXT
}

type XLSXSST struct {
	Count string "attr"
	UniqueCount string "attr"
	SI []XLSXSI
}

type XLSXFileVersion struct {
	AppName string "attr"
	LastEdited string "attr"
	LowestEdited string "attr"
	RupBuild string "attr"
}

type XLSXWorkbookPr struct {
	DefaultThemeVersion string "attr"
}

type XLSXWorkBookView struct {
	XWindow string "attr"
	YWindow string "attr"
	WindowWidth string "attr"
	WindowHeight string "attr"
}

type XLSXSheet struct {
	Name string "attr"
	SheetId string "attr"
	Id string "attr"
}

type XLSXDefinedName struct {
	Data string "chardata"
	Name string "attr"
	LocalSheetID string "attr"
}

type XLSXCalcPr struct {
	CalcId string "attr"
}

type XLSXBookViews struct {
	WorkBookView []XLSXWorkBookView
}

type XLSXSheets struct {
	Sheet []XLSXSheet
}

type XLSXDefinedNames struct {
	DefinedName []XLSXDefinedName
}

type XLSXWorkbook struct {
	FileVersion XLSXFileVersion
	WorkbookPr XLSXWorkbookPr
	BookViews XLSXBookViews
	Sheets XLSXSheets
	DefinedNames XLSXDefinedNames
	CalcPr XLSXCalcPr
}

type XLSXSheetStruct struct {

}

type XLSXFile struct {
	Sheets map [string]*XLSXSheetStruct
}

type XLSXFileInterface interface  {
	GetSheet(sheetname string) XLSXSheetStruct
}


func readSheetsFromZipFile(f *zip.File) os.Error {
	var workbook *XLSXWorkbook
	var error os.Error
	var rc io.ReadCloser
	workbook = new(XLSXWorkbook)
	rc, error = f.Open()
	if error != nil {
		return error
	}	
	error = xml.Unmarshal(rc, workbook)
	if error != nil {
		return error
	}
	return nil
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
	for _, v = range f.File {
		if v.Name == "xl/workbook.xml" {
			readSheetsFromZipFile(v)
		}

	}
	xlsxFile = new(XLSXFile)
	f.Close()
	return xlsxFile, nil
}


