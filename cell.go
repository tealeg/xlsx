package xlsx

import (
	"strconv"
)

// Style is a high level structure intended to provide user access to
// the contents of Style within an XLSX file.
type Style struct {
	Border Border
	Fill   Fill
	Font   Font
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
	PatternType string
	BgColor     string
	FgColor     string
}

type Font struct {
	Size    int
	Name    string
	Family  int
	Charset int
}

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

// String returns the value of a Cell as a string.
func (c *Cell) String() string {
	return c.Value
}

// GetStyle returns the Style associated with a Cell
func (c *Cell) GetStyle() *Style {
	style := &Style{}

	if c.styleIndex > 0 && c.styleIndex <= len(c.styles.CellXfs) {
		xf := c.styles.CellXfs[c.styleIndex-1]
		if xf.ApplyBorder {
			var border Border
			border.Left = c.styles.Borders[xf.BorderId].Left.Style
			border.Right = c.styles.Borders[xf.BorderId].Right.Style
			border.Top = c.styles.Borders[xf.BorderId].Top.Style
			border.Bottom = c.styles.Borders[xf.BorderId].Bottom.Style
			style.Border = border
		}
		if xf.ApplyFill {
			var fill Fill
			fill.PatternType = c.styles.Fills[xf.FillId].PatternFill.PatternType
			fill.BgColor = c.styles.Fills[xf.FillId].PatternFill.BgColor.RGB
			fill.FgColor = c.styles.Fills[xf.FillId].PatternFill.FgColor.RGB
			style.Fill = fill
		}
		if xf.ApplyFont {
			font := c.styles.Fonts[xf.FontId]
			style.Font = Font{}
			style.Font.Size, _ = strconv.Atoi(font.Sz.Val)
			style.Font.Name = font.Name.Val
			style.Font.Family, _ = strconv.Atoi(font.Family.Val)
			style.Font.Charset, _ = strconv.Atoi(font.Charset.Val)
		}
	}
	return style
}
