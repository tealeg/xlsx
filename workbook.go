package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
)

// xlsxWorkbook directly maps the workbook element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxWorkbook struct {
	FileVersion  xlsxFileVersion  `xml:"fileVersion"`
	WorkbookPr   xlsxWorkbookPr   `xml:"workbookPr"`
	BookViews    xlsxBookViews    `xml:"bookViews"`
	Sheets       xlsxSheets       `xml:"sheets"`
	DefinedNames xlsxDefinedNames `xml:"definedNames"`
	CalcPr       xlsxCalcPr       `xml:"calcPr"`
}

// xlsxFileVersion directly maps the fileVersion element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type xlsxFileVersion struct {
	AppName      string `xml:"appName,attr"`
	LastEdited   string `xml:"lastEdited,attr"`
	LowestEdited string `xml:"lowestEdited,attr"`
	RupBuild     string `xml:"rupBuild,attr"`
}

// xlsxWorkbookPr directly maps the workbookPr element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type xlsxWorkbookPr struct {
	DefaultThemeVersion string `xml:"defaultThemeVersion,attr"`
}

// xlsxBookViews directly maps the bookViews element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type xlsxBookViews struct {
	WorkBookView []xlsxWorkBookView `xml:"workbookView"`
}

// xlsxWorkBookView directly maps the workbookView element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type xlsxWorkBookView struct {
	XWindow      string `xml:"xWindow,attr"`
	YWindow      string `xml:"yWindow,attr"`
	WindowWidth  string `xml:"windowWidth,attr"`
	WindowHeight string `xml:"windowHeight,attr"`
}

// xlsxSheets directly maps the sheets element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheets struct {
	Sheet []xlsxSheet `xml:"sheet"`
}

// xlsxSheet directly maps the sheet element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheet struct {
	Name    string `xml:"name,attr"`
	SheetId string `xml:"sheetId,attr"`
	Id      string `xml:"id,attr"`
}

// xlsxDefinedNames directly maps the definedNames element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type xlsxDefinedNames struct {
	DefinedName []xlsxDefinedName `xml:"definedName"`
}

// xlsxDefinedName directly maps the definedName element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type xlsxDefinedName struct {
	Data         string `xml:",chardata"`
	Name         string `xml:"name,attr"`
	LocalSheetID string `xml:"localSheetId,attr"`
}

// xlsxCalcPr directly maps the calcPr element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxCalcPr struct {
	CalcId string `xml:"calcId,attr"`
}

// getWorksheetFromSheet() is an internal helper function to open a sheetN.xml file, refered to by an xlsx.xlsxSheet struct, from the XLSX file and unmarshal it an xlsx.xlsxWorksheet struct
func getWorksheetFromSheet(sheet xlsxSheet, worksheets map[string]*zip.File) (*xlsxWorksheet, error) {
	var rc io.ReadCloser
	var decoder *xml.Decoder
	var worksheet *xlsxWorksheet
	var error error
	worksheet = new(xlsxWorksheet)
	sheetName := fmt.Sprintf("sheet%s", sheet.SheetId)
	f := worksheets[sheetName]
	rc, error = f.Open()
	if error != nil {
		return nil, error
	}
	decoder = xml.NewDecoder(rc)
	error = decoder.Decode(worksheet)
	if error != nil {
		return nil, error
	}
	return worksheet, nil
}
