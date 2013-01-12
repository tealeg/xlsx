package xlsx

// xlsxWorksheet directly maps the worksheet element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxWorksheet struct {
	Dimension     xlsxDimension     `xml:"dimension"`
	SheetViews    xlsxSheetViews    `xml:"sheetViews"`
	SheetFormatPr xlsxSheetFormatPr `xml:"sheetFormatPr"`
	SheetData     xlsxSheetData     `xml:"sheetData"`
}

// xlsxDimension directly maps the dimension element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxDimension struct {
	Ref string `xml:"ref,attr"`
}

// xlsxSheetViews directly maps the sheetViews element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheetViews struct {
	SheetView []xlsxSheetView `xml:"sheetView"`
}

// xlsxSheetView directly maps the sheetView element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheetView struct {
	TabSelected    string        `xml:"tabSelected,attr"`
	WorkbookViewID string        `xml:"workbookViewId,attr"`
	Selection      xlsxSelection `xml:"selection"`
}

// xlsxSelection directly maps the selection element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.

type xlsxSelection struct {
	ActiveCell string `xml:"activeCell,attr"`
	SQRef      string `xml:"sqref,attr"`
}

// xlsxSheetFormatPr directly maps the sheetFormatPr element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheetFormatPr struct {
	BaseColWidth     string `xml:"baseColWidth,attr"`
	DefaultRowHeight string `xml:"defaultRowHeight,attr"`
}

// xlsxSheetData directly maps the sheetData element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheetData struct {
	Row []xlsxRow `xml:"row"`
}

// xlsxRow directly maps the row element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxRow struct {
	R     string  `xml:"r,attr"`
	Spans string  `xml:"spans,attr"`
	C     []xlsxC `xml:"c"`
}

// xlsxC directly maps the c element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxC struct {
	R string `xml:"r,attr"`
	S int    `xml:"s,attr"`
	T string `xml:"t,attr"`
	V string `xml:"v"`
}

// xlsxV directly maps the v element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
// type xlsxV struct {
// 	Data string `xml:"chardata"`
// }

// get cell
func (sh *Sheet) Cell(row, col int) *Cell {

	if len(sh.Rows) > row && len(sh.Rows[row].Cells) > col {
		return sh.Rows[row].Cells[col]
	}
	return new(Cell)
}
