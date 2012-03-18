package xlsx

// XLSXWorksheet directly maps the worksheet element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXWorksheet struct {
	Dimension     XLSXDimension   `xml:"dimension"`
	SheetViews    XLSXSheetViews  `xml:"sheetViews"`
	SheetFormatPr XLSXSheetFormatPr `xml:"sheetFormatPr"`
	SheetData     XLSXSheetData   `xml:"sheetData"`
}

// XLSXDimension directly maps the dimension element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXDimension struct {
	Ref string `xml:"ref,attr"`
}

// XLSXSheetViews directly maps the sheetViews element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXSheetViews struct {
	SheetView []XLSXSheetView `xml:"sheetView"`
}

// XLSXSheetView directly maps the sheetView element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXSheetView struct {
	TabSelected    string `xml:"tabSelected,attr"`
	WorkbookViewID string `xml:"workbookViewId,attr"`
	Selection      XLSXSelection `xml:"selection"`
}


// XLSXSelection directly maps the selection element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXSelection struct {
	ActiveCell string `xml:"activeCell,attr"`
	SQRef      string `xml:"sqref,attr"`
}

// XLSXSheetFormatPr directly maps the sheetFormatPr element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXSheetFormatPr struct {
	BaseColWidth     string `xml:"baseColWidth,attr"`
	DefaultRowHeight string `xml:"defaultRowHeight,attr"`
}

// XLSXSheetData directly maps the sheetData element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXSheetData struct {
	Row []XLSXRow `xml:"row"`
}

// XLSXRow directly maps the row element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXRow struct {
	R     string `xml:"r,attr"`
	Spans string `xml:"spans,attr"`
	C     []XLSXC `xml:"c"`
}

// XLSXC directly maps the c element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXC struct {
	R string `xml:"r,attr"`
	T string `xml:"t,attr"`
	V XLSXV  `xml:"v"`
}


// XLSXV directly maps the v element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type XLSXV struct {
	Data string `xml:",chardata"`
}

