// xslx is a package designed to help with reading data from
// spreadsheets stored in the XLSX format used in recent versions of
// Microsoft's Excel spreadsheet.

package xlsx

// Cell is a high level structure intended to provide user access to
// the contents of Cell within an xlsx.Row.
type Cell struct {
	Value      string
	styleIndex int
	styles     *xlsxStyles
}

// get cell borders(Left,Right,Top,Bottom)
func (c *Cell) Border() border {
	var bd border
	if c.styleIndex > 0 && c.styleIndex < len(c.styles.CellXfs) {
		xf := c.styles.CellXfs[c.styleIndex]
		borders := c.styles.Borders
		if xf.ApplyBorder != "0" {
			bd.Left = borders[xf.BorderId].Left.Style
			bd.Right = borders[xf.BorderId].Right.Style
			bd.Top = borders[xf.BorderId].Top.Style
			bd.Bottom = borders[xf.BorderId].Bottom.Style
		}
	}
	return bd
}

// get cell fills(background color and foreground color)
func (c *Cell) Fill() fill {
	var fi fill
	if c.styleIndex > 0 && c.styleIndex < len(c.styles.CellXfs) {
		xf := c.styles.CellXfs[c.styleIndex]
		fills := c.styles.Fills
		if xf.ApplyFill != "0" {
			if len(fills[xf.FillId].BgColor.Indexed) > 0 {
				index := fills[xf.FillId].BgColor.Indexed
				fi.BgColor = getColorFromIndexed(index)
			} else {
				fi.BgColor = fills[xf.FillId].BgColor.Rgb
			}
			if len(fills[xf.FillId].FgColor.Indexed) > 0 {
				index := fills[xf.FillId].FgColor.Indexed
				fi.FgColor = getColorFromIndexed(index)
			} else {
				fi.FgColor = fills[xf.FillId].FgColor.Rgb
			}
		}
	}
	return fi
}

// get cell styles (borders, fills..)
func (c *Cell) Style() *style {
	sty := new(style)
	// get borders
	sty.Borders = c.Border()
	// get colors
	sty.Fills = c.Fill()
	return sty
}

// Style is a high level structure intended to provide user access to
// the contents of Style within an XLSX file.
type style struct {
	Borders border
	Fills   fill
}

// Border is a high level structure intended to provide user access to
// the contents of Border Style within an Sheet.
type border struct {
	Left   string
	Right  string
	Top    string
	Bottom string
}

// Fill is a high level structure intended to provide user access to
// the contents of background and foreground color index within an Sheet.
type fill struct {
	BgColor string
	FgColor string
}

// getColorFromIndexed is used to convert a color Indexed 
// to the ARGB color
func getColorFromIndexed(colorIndexed string) string {
	switch colorIndexed {
	case "0":
		return "00000000"
	case "1":
		return "00FFFFFF"
	case "2":
		return "00FF0000"
	case "3":
		return "0000FF00"
	case "4":
		return "000000FF"
	case "5":
		return "00FFFF00"
	case "6":
		return "00FF00FF"
	case "7":
		return "0000FFFF"
	case "8":
		return "00000000"
	case "9":
		return "00FFFFFF"
	case "10":
		return "00FF0000"
	case "11":
		return "0000FF00"
	case "12":
		return "000000FF"
	case "13":
		return "00FFFF00"
	case "14":
		return "00FF00FF"
	case "15":
		return "0000FFFF"
	case "16":
		return "00800000"
	case "17":
		return "00008000"
	case "18":
		return "00000080"
	case "19":
		return "00808000"
	case "20":
		return "00800080"
	case "21":
		return "00008080"
	case "22":
		return "00C0C0C0"
	case "23":
		return "00808080"
	case "24":
		return "009999FF"
	case "25":
		return "00993366"
	case "26":
		return "00FFFFCC"
	case "27":
		return "00CCFFFF"
	case "28":
		return "00660066"
	case "29":
		return "00FF8080"
	case "30":
		return "000066CC"
	case "31":
		return "00CCCCFF"
	case "32":
		return "00000080"
	case "33":
		return "00FF00FF"
	case "34":
		return "00FFFF00"
	case "35":
		return "0000FFFF"
	case "36":
		return "00800080"
	case "37":
		return "00800000"
	case "38":
		return "00008080"
	case "39":
		return "000000FF"
	case "40":
		return "0000CCFF"
	case "41":
		return "00CCFFFF"
	case "42":
		return "00CCFFCC"
	case "43":
		return "00FFFF99"
	case "44":
		return "0099CCFF"
	case "45":
		return "00FF99CC"
	case "46":
		return "00CC99FF"
	case "47":
		return "00FFCC99"
	case "48":
		return "003366FF"
	case "49":
		return "0033CCCC"
	case "50":
		return "0099CC00"
	case "51":
		return "00FFCC00"
	case "52":
		return "00FF9900"
	case "53":
		return "00FF6600"
	case "54":
		return "00666699"
	case "55":
		return "00969696"
	case "56":
		return "00003366"
	case "57":
		return "00339966"
	case "58":
		return "00003300"
	case "59":
		return "00333300"
	case "60":
		return "00993300"
	case "61":
		return "00993366"
	case "62":
		return "00333399"
	case "63":
		return "00333333"
	case "64":
		return "" // System Foreground
	case "65":
		return "" // System Background
	}
	return ""
}
