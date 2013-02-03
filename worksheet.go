package xlsx

import (
	"encoding/xml"
	"errors"
	"fmt"
	"io"
)

// xlsxWorksheet directly maps the worksheet element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxWorksheet struct {
	XMLName       xml.Name
	R             string            `xml:"xmlns r,attr"`
	Dimension     xlsxDimension     `xml:"dimension"`
	SheetViews    xlsxSheetViews    `xml:"sheetViews"`
	SheetFormatPr xlsxSheetFormatPr `xml:"sheetFormatPr"`
	SheetData     xlsxSheetData     `xml:"sheetData"`
	PageMargins   xlsxPageMargins   `xml:"pageMargins"`

	sst     *xlsxSST // shared string
	changed bool     // changed flg
	rows    []*Row   // row data
	maxRow  int      // max row count
	maxCol  int      // max col count
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

// get cell info
func (sh *xlsxWorksheet) getCell(rowIndex int, colIndex int) (*xlsxC, error) {
	cellName := getCellName(rowIndex, colIndex)
	for i, row := range sh.SheetData.Row {
		if row.R == fmt.Sprintf("%d", rowIndex+1) {
			for j, col := range row.C {
				if col.R == cellName {
					return &sh.SheetData.Row[i].C[j], nil
				}
			}
			newCell := xlsxC{R: cellName}
			lenEnd := len(sh.SheetData.Row[i].C)
			sh.SheetData.Row[i].C = append(sh.SheetData.Row[i].C, newCell)
			return &sh.SheetData.Row[i].C[lenEnd+1], nil
		}
	}
	//didn't find the row
	newRow := xlsxRow{R: fmt.Sprintf("%d", rowIndex+1)}
	newRow.C = make([]xlsxC, 1)
	newRow.C[0] = xlsxC{R: cellName}
	sh.SheetData.Row = append(sh.SheetData.Row, newRow)
	return &newRow.C[0], nil
}

// set cell value from row and col
func (sh *xlsxWorksheet) SetCell(row int, col int, value interface{}) error {

	colomnData, _ := sh.getCell(row, col)

	if sh.sst == nil {
		return errors.New("The shared string table is nil")
	}
	if strValue, ok := value.(string); ok {
		index, err := sh.sst.getIndex(strValue)
		if err != nil {
			return err
		}
		colomnData.V = fmt.Sprintf("%d", index)
		colomnData.T = "s"
	} else if intValue, ok := value.(int); ok {
		colomnData.V = fmt.Sprintf("%d", intValue)
		colomnData.T = ""
	} else if floatValue, ok := value.(float32); ok {
		colomnData.V = fmt.Sprintf("%f", floatValue)
		colomnData.T = ""
	} else {
		return errors.New("Unknow type")
	}
	sh.changed = true
	return nil
}

// from int to letter
// e.g. 0->A 1->B
func getColumnName(columnIndex int) string {
	intOfA := int([]byte("A")[0])
	colName := string(intOfA + columnIndex%26)
	columnIndex = columnIndex / 26
	if columnIndex != 0 {
		return string(intOfA-1+columnIndex) + colName
	}
	return colName
}

// from row and colum get cell name
// e.g. 0,0->A0 1,1->B1
func getCellName(row, column int) string {
	return getColumnName(column) + fmt.Sprintf("%d", row+1)
}

// write sheet to xml file
func (sh *xlsxWorksheet) WriteTo(w io.Writer) error {
	data, err := xml.Marshal(sh)
	if err != nil {
		return err
	}
	content := string(data)
	_, err = w.Write([]byte(Header))
	_, err = w.Write([]byte(content))
	return err
}

// get the max col
func (sh *xlsxWorksheet) MaxCol() int {
	return sh.maxCol
}

// get the max row
func (sh *xlsxWorksheet) MaxRow() int {
	return sh.maxRow
}

// get cell
func (sh *xlsxWorksheet) Cell(row, col int) *Cell {

	if len(sh.rows) > row && sh.rows[row] != nil && len(sh.rows[row].Cells) > col {
		return sh.rows[row].Cells[col]
	}
	return new(Cell)
}
