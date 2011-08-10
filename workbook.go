package xlsx

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"xml"
)

// XLSXWorkbook directly maps the workbook element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXWorkbook struct {
	FileVersion  XLSXFileVersion
	WorkbookPr   XLSXWorkbookPr
	BookViews    XLSXBookViews
	Sheets       XLSXSheets
	DefinedNames XLSXDefinedNames
	CalcPr       XLSXCalcPr
}

// XLSXFileVersion directly maps the fileVersion element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type XLSXFileVersion struct {
	AppName      string `xml:"attr"`
	LastEdited   string `xml:"attr"`
	LowestEdited string `xml:"attr"`
	RupBuild     string `xml:"attr"`
}

// XLSXWorkbookPr directly maps the workbookPr element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type XLSXWorkbookPr struct {
	DefaultThemeVersion string `xml:"attr"`
}

// XLSXBookViews directly maps the bookViews element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type XLSXBookViews struct {
	WorkBookView []XLSXWorkBookView
}

// XLSXWorkBookView directly maps the workbookView element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type XLSXWorkBookView struct {
	XWindow      string `xml:"attr"`
	YWindow      string `xml:"attr"`
	WindowWidth  string `xml:"attr"`
	WindowHeight string `xml:"attr"`
}

// XLSXSheets directly maps the sheets element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXSheets struct {
	Sheet []XLSXSheet
}

// XLSXSheet directly maps the sheet element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXSheet struct {
	Name    string `xml:"attr"`
	SheetId string `xml:"attr"`
	Id      string `xml:"attr"`
}

// XLSXDefinedNames directly maps the definedNames element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type XLSXDefinedNames struct {
	DefinedName []XLSXDefinedName
}


// XLSXDefinedName directly maps the definedName element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type XLSXDefinedName struct {
	Data         string `xml:"chardata"`
	Name         string `xml:"attr"`
	LocalSheetID string `xml:"attr"`
}


// XLSXCalcPr directly maps the calcPr element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXCalcPr struct {
	CalcId string `xml:"attr"`
}



// getWorksheetFromSheet() is an internal helper function to open a sheetN.xml file, refered to by an xlsx.XLSXSheet struct, from the XLSX file and unmarshal it an xlsx.XLSXWorksheet struct 
func getWorksheetFromSheet(sheet XLSXSheet, worksheets map[string]*zip.File) (*XLSXWorksheet, os.Error) {
	var rc io.ReadCloser
	var worksheet *XLSXWorksheet
	var error os.Error
	worksheet = new(XLSXWorksheet)
	sheetName := fmt.Sprintf("sheet%s", sheet.SheetId)
	f := worksheets[sheetName]
	rc, error = f.Open()
	if error != nil {
		return nil, error
	}
	error = xml.Unmarshal(rc, worksheet)
	if error != nil {
		return nil, error
	}
	return worksheet, nil 
}
