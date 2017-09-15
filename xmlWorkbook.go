package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
)

const (
	// sheet state values as defined by
	// http://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.sheetstatevalues.aspx
	sheetStateVisible    = "visible"
	sheetStateHidden     = "hidden"
	sheetStateVeryHidden = "veryHidden"
)

// xmlxWorkbookRels contains xmlxWorkbookRelations
// which maps sheet id and sheet XML
type xlsxWorkbookRels struct {
	XMLName       xml.Name               `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Relationships []xlsxWorkbookRelation `xml:"Relationship"`
}

// xmlxWorkbookRelation maps sheet id and xl/worksheets/sheet%d.xml
type xlsxWorkbookRelation struct {
	Id     string `xml:",attr"`
	Target string `xml:",attr"`
	Type   string `xml:",attr"`
}

// xlsxWorkbook directly maps the workbook element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxWorkbook struct {
	XMLName            xml.Name               `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main workbook"`
	FileVersion        xlsxFileVersion        `xml:"fileVersion"`
	WorkbookPr         xlsxWorkbookPr         `xml:"workbookPr"`
	WorkbookProtection xlsxWorkbookProtection `xml:"workbookProtection"`
	BookViews          xlsxBookViews          `xml:"bookViews"`
	Sheets             xlsxSheets             `xml:"sheets"`
	DefinedNames       xlsxDefinedNames       `xml:"definedNames"`
	CalcPr             xlsxCalcPr             `xml:"calcPr"`
}

// xlsxWorkbookProtection directly maps the workbookProtection element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type xlsxWorkbookProtection struct {
	// We don't need this, yet.
}

// xlsxFileVersion directly maps the fileVersion element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type xlsxFileVersion struct {
	AppName      string `xml:"appName,attr,omitempty"`
	LastEdited   string `xml:"lastEdited,attr,omitempty"`
	LowestEdited string `xml:"lowestEdited,attr,omitempty"`
	RupBuild     string `xml:"rupBuild,attr,omitempty"`
}

// xlsxWorkbookPr directly maps the workbookPr element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type xlsxWorkbookPr struct {
	DefaultThemeVersion string `xml:"defaultThemeVersion,attr,omitempty"`
	BackupFile          bool   `xml:"backupFile,attr,omitempty"`
	ShowObjects         string `xml:"showObjects,attr,omitempty"`
	Date1904            bool   `xml:"date1904,attr"`
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
	ActiveTab            int    `xml:"activeTab,attr,omitempty"`
	FirstSheet           int    `xml:"firstSheet,attr,omitempty"`
	ShowHorizontalScroll bool   `xml:"showHorizontalScroll,attr,omitempty"`
	ShowVerticalScroll   bool   `xml:"showVerticalScroll,attr,omitempty"`
	ShowSheetTabs        bool   `xml:"showSheetTabs,attr,omitempty"`
	TabRatio             int    `xml:"tabRatio,attr,omitempty"`
	WindowHeight         int    `xml:"windowHeight,attr,omitempty"`
	WindowWidth          int    `xml:"windowWidth,attr,omitempty"`
	XWindow              string `xml:"xWindow,attr,omitempty"`
	YWindow              string `xml:"yWindow,attr,omitempty"`
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
	Name    string `xml:"name,attr,omitempty"`
	SheetId string `xml:"sheetId,attr,omitempty"`
	Id      string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
	State   string `xml:"state,attr,omitempty"`
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
// for a descriptions of the attributes see
// https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.definedname.aspx
type xlsxDefinedName struct {
	Data              string `xml:",chardata"`
	Name              string `xml:"name,attr"`
	Comment           string `xml:"comment,attr,omitempty"`
	CustomMenu        string `xml:"customMenu,attr,omitempty"`
	Description       string `xml:"description,attr,omitempty"`
	Help              string `xml:"help,attr,omitempty"`
	ShortcutKey       string `xml:"shortcutKey,attr,omitempty"`
	StatusBar         string `xml:"statusBar,attr,omitempty"`
	LocalSheetID      int    `xml:"localSheetId,attr,omitempty"`
	FunctionGroupID   int    `xml:"functionGroupId,attr,omitempty"`
	Function          bool   `xml:"function,attr,omitempty"`
	Hidden            bool   `xml:"hidden,attr,omitempty"`
	VbProcedure       bool   `xml:"vbProcedure,attr,omitempty"`
	PublishToServer   bool   `xml:"publishToServer,attr,omitempty"`
	WorkbookParameter bool   `xml:"workbookParameter,attr,omitempty"`
	Xlm               bool   `xml:"xml,attr,omitempty"`
}

// xlsxCalcPr directly maps the calcPr element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxCalcPr struct {
	CalcId       string  `xml:"calcId,attr,omitempty"`
	IterateCount int     `xml:"iterateCount,attr,omitempty"`
	RefMode      string  `xml:"refMode,attr,omitempty"`
	Iterate      bool    `xml:"iterate,attr,omitempty"`
	IterateDelta float64 `xml:"iterateDelta,attr,omitempty"`
}

// Helper function to lookup the file corresponding to a xlsxSheet object in the worksheets map
func worksheetFileForSheet(sheet xlsxSheet, worksheets map[string]*zip.File, sheetXMLMap map[string]string) *zip.File {
	sheetName, ok := sheetXMLMap[sheet.Id]
	if !ok {
		if sheet.SheetId != "" {
			sheetName = fmt.Sprintf("sheet%s", sheet.SheetId)
		} else {
			sheetName = fmt.Sprintf("sheet%s", sheet.Id)
		}
	}
	return worksheets[sheetName]
}

// getWorksheetFromSheet() is an internal helper function to open a
// sheetN.xml file, referred to by an xlsx.xlsxSheet struct, from the XLSX
// file and unmarshal it an xlsx.xlsxWorksheet struct
func getWorksheetFromSheet(sheet xlsxSheet, worksheets map[string]*zip.File, sheetXMLMap map[string]string, rowLimit int) (*xlsxWorksheet, error) {
	var r io.Reader
	var decoder *xml.Decoder
	var worksheet *xlsxWorksheet
	var err error
	worksheet = new(xlsxWorksheet)

	f := worksheetFileForSheet(sheet, worksheets, sheetXMLMap)
	if f == nil {
		return nil, fmt.Errorf("Unable to find sheet '%s'", sheet)
	}
	if rc, err := f.Open(); err != nil {
		return nil, err
	} else {
		defer rc.Close()
		r = rc
	}

	if rowLimit != -1 {
		r, err = truncateSheetXML(r, rowLimit)
		if err != nil {
			return nil, err
		}
	}

	decoder = xml.NewDecoder(r)
	err = decoder.Decode(worksheet)
	if err != nil {
		return nil, err
	}
	return worksheet, nil
}
