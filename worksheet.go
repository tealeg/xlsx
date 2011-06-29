package xlsx

type XLSXWorksheet struct {
	Dimension     XLSXDimension
	SheetViews    XLSXSheetViews
	SheetFormatPr XLSXSheetFormatPr
	SheetData     XLSXSheetData
}


type XLSXDimension struct {
	Ref string "attr"
}

type XLSXSheetViews struct {
	SheetView []XLSXSheetView
}


type XLSXSheetView struct {
	TabSelected    string "attr"
	WorkbookViewID string "attr"
	Selection      XLSXSelection
}


type XLSXSelection struct {
	ActiveCell string "attr"
	SQRef      string "attr"
}

type XLSXSheetFormatPr struct {
	BaseColWidth     string "attr"
	DefaultRowHeight string "attr"
}

type XLSXSheetData struct {
	Row []XLSXRow
}


type XLSXRow struct {
	R     string "attr"
	Spans string "attr"
	C     []XLSXC
}


type XLSXC struct {
	R string "attr"
	T string "attr"
	V XLSXV
}


type XLSXV struct {
	Data string "chardata"
}

