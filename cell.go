// xslx is a package designed to help with reading data from
// spreadsheets stored in the XLSX format used in recent versions of
// Microsoft's Excel spreadsheet.
//
// For a concise example of how to use this library why not check out
// the source for xlsx2csv here: https://github.com/tealeg/xlsx2csv

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
		if xf.ApplyBorder != "0" {
			bd.Left = c.styles.Borders[xf.BorderId].Left.Style
			bd.Right = c.styles.Borders[xf.BorderId].Right.Style
			bd.Top = c.styles.Borders[xf.BorderId].Top.Style
			bd.Bottom = c.styles.Borders[xf.BorderId].Bottom.Style
		}
	}
	return bd
}

// get cell fills(background color and foreground color)
func (c *Cell) Fill() fill {
	var fi fill
	if c.styleIndex > 0 && c.styleIndex < len(c.styles.CellXfs) {
		xf := c.styles.CellXfs[c.styleIndex]
		if xf.ApplyFill != "0" {
			if len(c.styles.Fills[xf.FillId].BgColor.Indexed) > 0 {
				fi.BgColor = getColorFromIndexed(c.styles.Fills[xf.FillId].BgColor.Indexed)
			} else {
				fi.BgColor = c.styles.Fills[xf.FillId].BgColor.Rgb
			}
			if len(c.styles.Fills[xf.FillId].FgColor.Indexed) > 0 {
				fi.FgColor = getColorFromIndexed(c.styles.Fills[xf.FillId].FgColor.Indexed)
			} else {
				fi.FgColor = c.styles.Fills[xf.FillId].FgColor.Rgb
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
