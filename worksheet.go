package xlsx

import (
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"strings"
	"strconv"
	"math"
)

// xlsxWorksheet directly maps the worksheet element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxWorksheet struct {
	XMLName xml.Name
	//R             string            `xml:"xmlns:r,attr"`
	Dimension     xlsxDimension     `xml:"dimension"`
	SheetViews    xlsxSheetViews    `xml:"sheetViews"`
	SheetFormatPr xlsxSheetFormatPr `xml:"sheetFormatPr"`
	SheetData     xlsxSheetData     `xml:"sheetData"`
	PageMargins   xlsxPageMargins   `xml:"pageMargins"`

	sst     *xlsxSST // shared string
	changed bool     // changed flg
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
			return &sh.SheetData.Row[i].C[lenEnd], nil
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

// getCoordsFromCellIDString returns the zero based cartesian
// coordinates from a cell name in Excel format, e.g. the cellIDString
// "A1" returns 0, 0 and the "B3" return 1, 2.
func getCoordsFromCellIDString(cellIDString string) (x, y int, error error) {
	var letterPart string = strings.Map(letterOnlyMapF, cellIDString)
	y, error = strconv.Atoi(strings.Map(intOnlyMapF, cellIDString))
	if error != nil {
		return x, y, error
	}
	y -= 1 // Zero based
	x = lettersToNumeric(letterPart)
	return x, y, error
}

// letterOnlyMapF is used in conjunction with strings.Map to return
// only the characters A-Z and a-z in a string
func letterOnlyMapF(rune rune) rune {
	switch {
	case 'A' <= rune && rune <= 'Z':
		return rune
	case 'a' <= rune && rune <= 'z':
		return rune - 32
	}
	return -1
}

// intOnlyMapF is used in conjunction with strings.Map to return only
// the numeric portions of a string.
func intOnlyMapF(rune rune) rune {
	if rune >= 48 && rune < 58 {
		return rune
	}
	return -1
}

// lettersToNumeric is used to convert a character based column
// reference to a zero based numeric column identifier.
func lettersToNumeric(letters string) int {
	var sum int = 0
	var shift int
	extent := len(letters)
	for i, c := range letters {
		// Just to make life akward.  If we think of this base
		// 26 notation as being like HEX or binary we hit a
		// nasty little problem.  The issue is that we have no
		// 0s and therefore A can be both a 1 and a 0.  The
		// value range of a letter is different in the most
		// significant position if (and only if) there is more
		// than one positions.  For example:
		// "A" = 0
		//               676 | 26 | 0
		//               ----+----+----
		//                 0 |  0 | 0
		//
		//  "Z" = 25
		//                676 | 26 | 0
		//                ----+----+----
		//                  0 |  0 |  25
		//   "AA" = 26
		//                676 | 26 | 0
		//                ----+----+----
		//                  0 |  1 | 0     <--- note here - the value of "A" maps to both 1 and 0.
		if i == 0 && extent > 1 {
			shift = 1
		} else {
			shift = 0
		}
		multiplier := positionalLetterMultiplier(extent, i)
		switch {
		case 'A' <= c && c <= 'Z':
			sum += multiplier * (int((c - 'A')) + shift)
		case 'a' <= c && c <= 'z':
			sum += multiplier * (int((c - 'a')) + shift)
		}
	}
	return sum
}

// positionalLetterMultiplier gives an integer multiplier to use for a
// position in a letter based column identifer. For example, the
// column ID "AA" is equivalent to 26*1 + 1, "BA" is equivalent to
// 26*2 + 1 and "ABA" is equivalent to (676 * 1)+(26 * 2)+1 or
// ((26**2)*1)+((26**1)*2)+((26**0))*1
func positionalLetterMultiplier(extent, pos int) int {
	offset := pos + 1
	power := float64(extent - offset)
	result := math.Pow(26, power)
	return int(result)
}

// get the max col
func (sh *xlsxWorksheet) MaxCol() int {
	maxCol := 0
	for _, rawrow := range sh.SheetData.Row {
		for _, rawcell := range rawrow.C {
			x, _, error := getCoordsFromCellIDString(rawcell.R)
			if error != nil {
				panic(fmt.Sprintf("Invalid Cell Coord, %s\n", rawcell.R))
			}
			if x > maxCol {
				maxCol = x
			}
		}
	}
	return maxCol + 1
}

// get the max row
func (sh *xlsxWorksheet) MaxRow() int {
	maxRow := 0
	for _, rawrow := range sh.SheetData.Row {
		for _, rawcell := range rawrow.C {
			_, y, error := getCoordsFromCellIDString(rawcell.R)
			if error != nil {
				panic(fmt.Sprintf("Invalid Cell Coord, %s\n", rawcell.R))
			}
			if y > maxRow {
				maxRow = y
			}
		}
	}
	return maxRow + 1
}

// get cell
func (sh *xlsxWorksheet) Cell(rowIndex, colIndex int) *Cell {

	cell := new(Cell)
	cellName := getCellName(rowIndex, colIndex)
	for i, row := range sh.SheetData.Row {
		if row.R == fmt.Sprintf("%d", rowIndex+1) {
			for j, col := range row.C {
				if col.R == cellName {
					c := sh.SheetData.Row[i].C[j]
					data := c.V
					var value string
					if len(data) > 0 {
						vval := strings.Trim(data, " \t\n\r")
						if c.T == "s" {
							ref, error := strconv.Atoi(vval)
							if error != nil {
								panic(error)
							}
							si := sh.sst.SI[ref]
							if len(si.R) > 0 {
								for j := 0; j < len(si.R); j++ {
									value = value + si.R[j].T
								}
							} else {
								value = si.T
							}
						} else {
							value = vval
						}
					}
					cell.Value = value
					return cell
				}
			}
		}
	}
	return cell
}
