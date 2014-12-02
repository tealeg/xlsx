package xlsx

import (
	"encoding/xml"
)

// xlsxWorksheet directly maps the worksheet element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxWorksheet struct {
	XMLName    xml.Name       `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main worksheet"`
	Dimension  xlsxDimension  `xml:"dimension"`
	Cols       xlsxCols       `xml:"cols,omitempty"`
	SheetData  xlsxSheetData  `xml:"sheetData"`
	Hyperlinks xlsxHyperlinks `xml:"hyperlinks"`
}

// xlsxCols directly maps the cols element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxCols struct {
	Col []xlsxCol `xml:"col"`
}

type xlsxHyperlinks struct {
	Links []xlsxHyperlink `xml:"hyperlink"`
}

// xlsxCol directly maps the col element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxCol struct {
	Min    int     `xml:"min,attr"`
	Max    int     `xml:"max,attr"`
	Hidden bool    `xml:"hidden,attr,omitempty"`
	Width  float64 `xml:"width,attr,omitempty"`
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
	XMLName xml.Name  `xml:"sheetData"`
	Row     []xlsxRow `xml:"row"`
}

type xlsxHyperlink struct {
	Ref     string `xml:"ref,attr"`
	Tooltip string `xml:"tooltip,attr"`
}

// xlsxRow directly maps the row element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxRow struct {
	R      int     `xml:"r,attr"`
	Spans  string  `xml:"spans,attr,omitempty"`
	Hidden bool    `xml:"hidden,attr,omitempty"`
	C      []xlsxC `xml:"c"`
}

// xlsxC directly maps the c element in the namespace
// http://schemas.openxmlformats.org/sprceadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxC struct {
	R string `xml:"r,attr"` // Cell ID, e.g. A1
	S int    `xml:"s,attr"` // Style reference.
	T string `xml:"t,attr"` // Type.
	V string `xml:"v"`      // Value
}
