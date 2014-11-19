// xslx is a package designed to help with reading data from
// spreadsheets stored in the XLSX format used in recent versions of
// Microsoft's Excel spreadsheet.
//
// For a concise example of how to use this library why not check out
// the source for xlsx2csv here: https://github.com/tealeg/xlsx2csv

package xlsx

import (
	"strconv"
	"strings"
)

// xlsxStyle directly maps the style element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxStyles struct {
	Fonts        []xlsxFont   `xml:"fonts>font,"`
	Fills        []xlsxFill   `xml:"fills>fill"`
	Borders      []xlsxBorder `xml:"borders>border"`
	CellStyleXfs []xlsxXf     `xml:"cellStyleXfs>xf"`
	CellXfs      []xlsxXf     `xml:"cellXfs>xf"`
	NumFmts      []xlsxNumFmt `xml:numFmts>numFmt"`
}

// xlsxNumFmt directly maps the numFmt element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxNumFmt struct {
	NumFmtId   int    `xml:"numFmtId"`
	FormatCode string `xml:"formatCode"`
}

// xlsxFont directly maps the font element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxFont struct {
	Sz      xlsxVal `xml:"sz"`
	Name    xlsxVal `xml:"name"`
	Family  xlsxVal `xml:"family"`
	Charset xlsxVal `xml:"charset"`
}

// xlsxVal directly maps the val element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxVal struct {
	Val string `xml:"val,attr"`
}

// xlsxFill directly maps the fill element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxFill struct {
	PatternFill xlsxPatternFill `xml:"patternFill"`
}

// xlsxPatternFill directly maps the patternFill element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxPatternFill struct {
	PatternType string    `xml:"patternType,attr"`
	FgColor     xlsxColor `xml:"fgColor"`
	BgColor     xlsxColor `xml:"bgColor"`
}

// xlsxColor is a common mapping used for both the fgColor and bgColor
// elements in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxColor struct {
	RGB string `xml:"rgb,attr"`
}

// xlsxBorder directly maps the border element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxBorder struct {
	Left   xlsxLine `xml:"left"`
	Right  xlsxLine `xml:"right"`
	Top    xlsxLine `xml:"top"`
	Bottom xlsxLine `xml:"bottom"`
}

// xlsxLine directly maps the line style element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxLine struct {
	Style string `xml:"style,attr"`
}

// xlsxXf directly maps the xf element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxXf struct {
	ApplyAlignment  bool          `xml:"applyAlignment,attr"`
	ApplyBorder     bool          `xml:"applyBorder,attr"`
	ApplyFont       bool          `xml:"applyFont,attr"`
	ApplyFill       bool          `xml:"applyFill,attr"`
	ApplyProtection bool          `xml:"applyProtection,attr"`
	BorderId        int           `xml:"borderId,attr"`
	FillId          int           `xml:"fillId,attr"`
	FontId          int           `xml:"fontId,attr"`
	NumFmtId        int           `xml:"numFmtId,attr"`
	alignment       xlsxAlignment `xml:"alignement"`
}

type xlsxAlignment struct {
	Horizontal   string `xml:"horizontal,attr"`
	Indent       int    `xml:"indent,attr"`
	ShrinkToFit  bool   `xml:"shrinkToFit,attr"`
	TextRotation int    `xml:"textRotation,attr"`
	Vertical     string `xml:"vertical,attr"`
	WrapText     bool   `xml:"wrapText,attr"`
}

func (styles *xlsxStyles) getStyle(styleIndex int) (style Style) {
	style = Style{}
	style.Border = Border{}
	style.Fill = Fill{}
	style.Font = Font{}

	xfCount := len(styles.CellXfs)
	if styleIndex > -1 && xfCount > 0 && styleIndex <= xfCount {
		xf := styles.CellXfs[styleIndex]

		if xf.ApplyBorder {
			style.Border.Left = styles.Borders[xf.BorderId].Left.Style
			style.Border.Right = styles.Borders[xf.BorderId].Right.Style
			style.Border.Top = styles.Borders[xf.BorderId].Top.Style
			style.Border.Bottom = styles.Borders[xf.BorderId].Bottom.Style
		}
		// TODO - consider how to handle the fact that
		// ApplyFill can be missing.  At the moment the XML
		// unmarshaller simply sets it to false, which creates
		// a bug.

		// if xf.ApplyFill {
		if xf.FillId > -1 && xf.FillId < len(styles.Fills) {
			xFill := styles.Fills[xf.FillId]
			style.Fill.PatternType = xFill.PatternFill.PatternType
			style.Fill.FgColor = xFill.PatternFill.FgColor.RGB
			style.Fill.BgColor = xFill.PatternFill.BgColor.RGB
		}
		// }
		if xf.ApplyFont {
			xfont := styles.Fonts[xf.FontId]
			style.Font.Size, _ = strconv.Atoi(xfont.Sz.Val)
			style.Font.Name = xfont.Name.Val
			style.Font.Family, _ = strconv.Atoi(xfont.Family.Val)
			style.Font.Charset, _ = strconv.Atoi(xfont.Charset.Val)
		}
	}
	return style

}

func (styles *xlsxStyles) getNumberFormat(styleIndex int, numFmtRefTable map[int]xlsxNumFmt) string {
	if styles.CellXfs == nil {
		return ""
	}
	var numberFormat string = ""
	if styleIndex > -1 && styleIndex <= len(styles.CellXfs) {
		xf := styles.CellXfs[styleIndex]
		numFmt := numFmtRefTable[xf.NumFmtId]
		numberFormat = numFmt.FormatCode
	}
	return strings.ToLower(numberFormat)
}

func (styles *xlsxStyles) addFont(xFont xlsxFont) int {
	styles.Fonts = append(styles.Fonts, xFont)
	return len(styles.Fonts) - 1
}

func (styles *xlsxStyles) addFill(xFill xlsxFill) int {
	styles.Fills = append(styles.Fills, xFill)
	return len(styles.Fills) - 1
}

func (styles *xlsxStyles) addBorder(xBorder xlsxBorder) int {
	styles.Borders = append(styles.Borders, xBorder)
	return len(styles.Borders) - 1
}

func (styles *xlsxStyles) addCellStyleXf(xCellStyleXf xlsxXf) {
	styles.CellStyleXfs = append(styles.CellStyleXfs, xCellStyleXf)
}

func (styles *xlsxStyles) addCellXf(xCellXf xlsxXf) {
	styles.CellXfs = append(styles.CellXfs, xCellXf)
}
