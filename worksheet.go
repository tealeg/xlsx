package xlsx

import (
	"encoding/xml"
	"io"
)

// xlsxWorksheet directly maps the worksheet element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxWorksheet struct {
	XMLName       xml.Name          `xml:"worksheet"`
	Dimension     xlsxDimension     `xml:"dimension"`
	SheetViews    xlsxSheetViews    `xml:"sheetViews"`
	SheetFormatPr xlsxSheetFormatPr `xml:"sheetFormatPr"`
	SheetData     xlsxSheetData     `xml:"sheetData"`
	PageMargins   xlsxPageMargins   `xml:"pageMargins"`
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
	TabSelected    string         `xml:"tabSelected,attr"`
	WorkbookViewID string         `xml:"workbookViewId,attr"`
	Selection      *xlsxSelection `xml:"selection,omitempty"`
}

// xlsxSelection directly maps the selection element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.

type xlsxSelection struct {
	ActiveCell string `xml:"activeCell,attr,omitempty"`
	SQRef      string `xml:"sqref,attr,omitempty"`
}

// xlsxSheetFormatPr directly maps the sheetFormatPr element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheetFormatPr struct {
	BaseColWidth     string `xml:"baseColWidth,attr,omitempty"`
	DefaultRowHeight string `xml:"defaultRowHeight,attr,omitempty"`
}

// xlsxSheetData directly maps the sheetData element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheetData struct {
	Row []xlsxRow `xml:"row"`
}

// xlsxPageMargins directly maps the pageMargins element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxPageMargins struct {
	Bottom string `xml:"bottom,attr,omitempty"`
	Footer string `xml:"footer,attr,omitempty"`
	Header string `xml:"header,attr,omitempty"`
	Left   string `xml:"left,attr,omitempty"`
	Right  string `xml:"right,attr,omitempty"`
	Top    string `xml:"top,attr,omitempty"`
}

// xlsxRow directly maps the row element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxRow struct {
	R     string  `xml:"r,attr,omitempty"`
	Spans string  `xml:"spans,attr,omitempty"`
	C     []xlsxC `xml:"c"`
}

// xlsxC directly maps the c element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxC struct {
	R string `xml:"r,attr,omitempty"`
	S int    `xml:"s,attr,omitempty"`
	T string `xml:"t,attr,omitempty"`
	V string `xml:"v,omitempty"`
}

// xlsxV directly maps the v element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
// type xlsxV struct {
// 	Data string `xml:"chardata"`
// }

// write sheet to xml file
func (sh *xlsxWorksheet) WriteTo(w io.Writer) error {
	data, err := xml.MarshalIndent(sh, "", "    ")
	if err != nil {
		return err
	}
	content := string(data)
	_, err = w.Write([]byte(xml.Header))
	_, err = w.Write([]byte(content))
	return err
}

// get cell
func (sh *Sheet) Cell(row, col int) *Cell {

	if len(sh.Rows) > row && sh.Rows[row] != nil && len(sh.Rows[row].Cells) > col {
		return sh.Rows[row].Cells[col]
	}
	return new(Cell)
}
