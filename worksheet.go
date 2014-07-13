package xlsx

// xlsxWorksheet directly maps the worksheet element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxWorksheet struct {
	SheetFormatPr xlsxSheetFormatPr `xml:"sheetFormatPr"`
	Dimension     xlsxDimension     `xml:"dimension"`
	SheetData     xlsxSheetData     `xml:"sheetData"`
}

type xlsxSheetFormatPr struct {
	DefaultRowHeight float64 `xml:"defaultRowHeight,attr"`
}

// xlsxDimension directly maps the dimension element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxDimension struct {
	Ref string `xml:"ref,attr"`
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
	R            int     `xml:"r,attr"`
	Spans        string  `xml:"spans,attr"`
	C            []xlsxC `xml:"c"`
	Ht           float64 `xml:"ht,attr"`
	CustomHeight int     `xml:"customHeight,attr"`
}

type xlsxSharedFormula struct {
	F     string
	Ref   string
	cellX int
	cellY int
}

type xlsxF struct {
	F   string `xml:",innerxml"`
	Si  string `xml:"si,attr"`
	Ref string `xml:"ref,attr"`
	T   string `xml:"t,attr"`
}

// xlsxC directly maps the c element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxC struct {
	F xlsxF  `xml:"f"`
	R string `xml:"r,attr"`
	S int    `xml:"s,attr"`
	T string `xml:"t,attr"`
	V string `xml:"v"`
}

// get cell
func (sh *Sheet) Cell(row, col int) *Cell {

	cell, ok := sh.Cells[CellCoord{col, row}]
	if ok {
		return &cell
	}
	return nil
}
