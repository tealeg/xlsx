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

// CellInterface defines the public API of the Cell.
type CellInterface interface {
	String() string
}

// return cell's string value
func (c *Cell) String() string {
	return c.Value
}

// get cell borders(Left,Right,Top,Bottom)
func (c *Cell) GetBorder() Border {
	var border Border
	if c.styleIndex > 0 && c.styleIndex < len(c.styles.CellXfs) {
		xf := c.styles.CellXfs[c.styleIndex]
		if xf.ApplyBorder != "0" {
			border.Left = c.styles.Borders[xf.BorderId].Left.Style
			border.Right = c.styles.Borders[xf.BorderId].Right.Style
			border.Top = c.styles.Borders[xf.BorderId].Top.Style
			border.Bottom = c.styles.Borders[xf.BorderId].Bottom.Style
		}
	}
	return border
}

// get cell fills(background color and foreground color)
func (c *Cell) GetFill() Fill {
	var fill Fill
	if c.styleIndex > 0 && c.styleIndex < len(c.styles.CellXfs) {
		xf := c.styles.CellXfs[c.styleIndex]
		if xf.ApplyFill != "0" {
			if len(c.styles.Fills[xf.FillId].BgColor.Indexed) > 0 {
				fill.BgColor = getColorFromIndexed(c.styles.Fills[xf.FillId].BgColor.Indexed)
			} else {
				fill.BgColor = c.styles.Fills[xf.FillId].BgColor.Rgb
			}
			if len(c.styles.Fills[xf.FillId].FgColor.Indexed) > 0 {
				fill.FgColor = getColorFromIndexed(c.styles.Fills[xf.FillId].FgColor.Indexed)
			} else {
				fill.FgColor = c.styles.Fills[xf.FillId].FgColor.Rgb
			}
		}
	}
	return fill
}

// get cell styles (borders, fills..)
func (c *Cell) GetStyle() *Style {
	style := new(Style)
	// get borders
	style.Borders = c.GetBorder()
	// get colors
	style.Fills = c.GetFill()

	// if c.styleIndex > 0 && c.styleIndex < len(c.styles.CellXfs) {
	// 	xf := c.styles.CellXfs[c.styleIndex]
	// 	if xf.ApplyBorder != "0" {
	// 		var border Border
	// 		border.Left = c.styles.Borders[xf.BorderId].Left.Style
	// 		border.Right = c.styles.Borders[xf.BorderId].Right.Style
	// 		border.Top = c.styles.Borders[xf.BorderId].Top.Style
	// 		border.Bottom = c.styles.Borders[xf.BorderId].Bottom.Style
	// 		style.Boders = border
	// 	}
	// 	if xf.ApplyFill != "0" {
	// 		var fill Fill
	// 		if len(c.styles.Fills[xf.FillId].BgColor.Indexed) > 0 {
	// 			fill.BgColor = getColorFromIndexed(c.styles.Fills[xf.FillId].BgColor.Indexed)
	// 		} else {
	// 			fill.BgColor = c.styles.Fills[xf.FillId].BgColor.Rgb
	// 		}
	// 		if len(c.styles.Fills[xf.FillId].FgColor.Indexed) > 0 {
	// 			fill.FgColor = getColorFromIndexed(c.styles.Fills[xf.FillId].FgColor.Indexed)
	// 		} else {
	// 			fill.FgColor = c.styles.Fills[xf.FillId].FgColor.Rgb
	// 		}
	// 		style.Fills = fill
	// 	}
	// }
	return style
}

// Style is a high level structure intended to provide user access to
// the contents of Style within an XLSX file.
type Style struct {
	Borders Border
	Fills  Fill
}

// Border is a high level structure intended to provide user access to
// the contents of Border Style within an Sheet.
type Border struct {
	Left   string
	Right  string
	Top    string
	Bottom string
}

// Fill is a high level structure intended to provide user access to
// the contents of background and foreground color index within an Sheet.
type Fill struct {
	BgColor string
	FgColor string
}
