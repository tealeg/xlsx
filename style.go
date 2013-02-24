// xslx is a package designed to help with reading data from
// spreadsheets stored in the XLSX format used in recent versions of
// Microsoft's Excel spreadsheet.
//
// For a concise example of how to use this library why not check out
// the source for xlsx2csv here: https://github.com/tealeg/xlsx2csv

package xlsx

// xlsxStyle directly maps the style element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxStyles struct {
	Fonts        []xlsxFont   `xml:"fonts>font"`
	Fills        []xlsxFill   `xml:"fills>fill"`
	Borders      []xlsxBorder `xml:"borders>border"`
	CellStyleXfs []xlsxXf     `xml:"cellStyleXfs>xf"`
	CellXfs      []xlsxXf     `xml:"cellXfs>xf"`
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
	FgColor xlsxColor `xml:"patternFill>fgColor"`
	BgColor xlsxColor `xml:"patternFill>bgColor"`
}

// xlsxColor directly maps the Color index or Rgb Color element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
// Color Indexed ARGBValue
// 0             00000000
// 1             00FFFFFF
// 2             00FF0000
// 3             0000FF00
// ...............
// ...............
type xlsxColor struct {
	Indexed string `xml:"indexed,attr"`
	Rgb     string `xml:"rgb,attr"`
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
	ApplyBorder string `xml:"applyBorder,attr"`
	BorderId    int    `xml:"borderId,attr"`
	ApplyFill   string `xml:"applyFill,attr"`
	FillId      int    `xml:"fillId,attr"`
}
