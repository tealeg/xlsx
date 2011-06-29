package xlsx


type XLSXWorkbook struct {
	FileVersion  XLSXFileVersion
	WorkbookPr   XLSXWorkbookPr
	BookViews    XLSXBookViews
	Sheets       XLSXSheets
	DefinedNames XLSXDefinedNames
	CalcPr       XLSXCalcPr
}

type XLSXFileVersion struct {
	AppName      string "attr"
	LastEdited   string "attr"
	LowestEdited string "attr"
	RupBuild     string "attr"
}

type XLSXWorkbookPr struct {
	DefaultThemeVersion string "attr"
}

type XLSXBookViews struct {
	WorkBookView []XLSXWorkBookView
}


type XLSXWorkBookView struct {
	XWindow      string "attr"
	YWindow      string "attr"
	WindowWidth  string "attr"
	WindowHeight string "attr"
}

type XLSXSheets struct {
	Sheet []XLSXSheet
}


type XLSXSheet struct {
	Name    string "attr"
	SheetId string "attr"
	Id      string "attr"
}


type XLSXDefinedNames struct {
	DefinedName []XLSXDefinedName
}

type XLSXDefinedName struct {
	Data         string "chardata"
	Name         string "attr"
	LocalSheetID string "attr"
}

type XLSXCalcPr struct {
	CalcId string "attr"
}
