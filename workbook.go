package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"io/ioutil"
	"fmt"
	"io"
)

// XLSXWorkbook directly maps the workbook element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXWorkbook struct {
	XMLName      xml.Name        `xml:"workbook"`
	FileVersion  XLSXFileVersion `xml:"fileVersion"`
	WorkbookPr   XLSXWorkbookPr  `xml:"workbookPr"`
	BookViews    XLSXBookViews   `xml:"bookViews"`
	Sheets       XLSXSheets      `xml:"sheets"`
	DefinedNames XLSXDefinedNames
	CalcPr       XLSXCalcPr      `xml:"calcPr"`
}

// XLSXFileVersion directly maps the fileVersion element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type XLSXFileVersion struct {
	AppName      string `xml:"appName,attr"`
	LastEdited   string `xml:"lastEdited,attr"`
	LowestEdited string `xml:"lowestEdited,attr"`
	RupBuild     string `xml:"rupBuild,attr"`
}

// XLSXWorkbookPr directly maps the workbookPr element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type XLSXWorkbookPr struct {
	DefaultThemeVersion string `xml:"defaultThemeVersion,attr"`
}

// XLSXBookViews directly maps the bookViews element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type XLSXBookViews struct {
	WorkBookView []XLSXWorkBookView `xml:"workbookView"`
}

// XLSXWorkBookView directly maps the workbookView element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type XLSXWorkBookView struct {
	XWindow      string `xml:"xWindow,attr"`
	YWindow      string `xml:"yWindow,attr"`
	WindowWidth  string `xml:"windowWidth,attr"`
	WindowHeight string `xml:"windowHeight,attr"`
}

// XLSXSheets directly maps the sheets element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXSheets struct {
	Sheet []XLSXSheet `xml:"sheet"`
}

// XLSXSheet directly maps the sheet element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXSheet struct {
	Name    string `xml:"name,attr"`
	SheetId string `xml:"sheetId,attr"`
	Id      string `xml:"id,attr"`
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
	Data         string ",chardata"
	Name         string ",attr"
	LocalSheetID string ",attr"
}

// XLSXCalcPr directly maps the calcPr element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXCalcPr struct {
	CalcId string `xml:"calcId,attr"`
}

// getWorksheetFromSheet() is an internal helper function to open a sheetN.xml file, refered to by an xlsx.XLSXSheet struct, from the XLSX file and unmarshal it an xlsx.XLSXWorksheet struct 
func getWorksheetFromSheet(sheet XLSXSheet, worksheets map[string]*zip.File) (*XLSXWorksheet, error) {
	var rc io.ReadCloser
	var worksheet *XLSXWorksheet
	var error error
	worksheet = new(XLSXWorksheet)
	sheetName := fmt.Sprintf("sheet%s", sheet.SheetId)
	f := worksheets[sheetName]
	rc, error = f.Open()
	if error != nil {
		return nil, error
	}

	content, err := ioutil.ReadAll(rc)
	if err != nil{
		return nil, err
	}
	error = xml.Unmarshal(content, worksheet)
	fmt.Println("Worksheet:>>>>>>")
	fmt.Println(worksheet)
	if error != nil {
		return nil, error
	}
	return worksheet, nil
}
