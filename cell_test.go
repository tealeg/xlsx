package xlsx

import (
	. "gopkg.in/check.v1"
)

type StyleSuite struct {}

var _ = Suite(&StyleSuite{})


func (s *StyleSuite) TestNewStyle(c *C){
	style := NewStyle()
	c.Assert(style, NotNil)
}


// Test that SetFont correctly updates the Font associated with a Style.
func (s *StyleSuite) TestSetFont(c *C) {
	font := NewFont(12, "Calibra")
	style := Style{}
	style.SetFont(*font)
	c.Assert(style.Font.Size, Equals, 12)
	c.Assert(style.Font.Name, Equals, "Calibra")
}


type FontSuite struct {}

var _ = Suite(&FontSuite{})

func (s *FontSuite) TestNewFont(c *C) {
	font := NewFont(12, "Verdana")
	c.Assert(font, NotNil)
	c.Assert(font.Name, Equals, "Verdana")
	c.Assert(font.Size, Equals, 12)
}

type CellSuite struct {}

var _ = Suite(&CellSuite{})

// Test that we can set and get a Value from a Cell
func (s *CellSuite) TestValueSet(c *C) {
	// Note, this test is fairly pointless, it serves mostly to
	// reinforce that this functionality is important, and should
	// the mechanics of this all change at some point, to remind
	// us not to lose this.
	cell := Cell{}
	cell.Value = "A string"
	c.Assert(cell.Value, Equals, "A string")
}

// Test that GetStyle correctly converts the xlsxStyle.Fonts.
func (s *CellSuite) TestGetStyleWithFonts(c *C) {
	var cell *Cell
	var style *Style
	var xStyles *xlsxStyles
	var fonts []xlsxFont
	var cellXfs []xlsxXf

	fonts = make([]xlsxFont, 1)
	fonts[0] = xlsxFont{
		Sz:   xlsxVal{Val: "10"},
		Name: xlsxVal{Val: "Calibra"}}

	cellXfs = make([]xlsxXf, 1)
	cellXfs[0] = xlsxXf{ApplyFont: true, FontId: 0}

	xStyles = &xlsxStyles{Fonts: fonts, CellXfs: cellXfs}

	cell = &Cell{Value: "123", styleIndex: 1, styles: xStyles}
	style = cell.GetStyle()
	c.Assert(style, NotNil)
	c.Assert(style.Font.Size, Equals, 10)
	c.Assert(style.Font.Name, Equals, "Calibra")
}


// Test that SetStyle correctly updates the xlsxStyle.Fonts.
func (s *CellSuite) TestSetStyleWithFonts(c *C) {
	file := NewFile()
	sheet := file.AddSheet("Test")
	row := sheet.AddRow()
	cell := row.AddCell()
	font := NewFont(12, "Calibra")
	style := NewStyle()
	style.SetFont(*font)
	cell.SetStyle(style)
	c.Assert(cell.styleIndex, Equals, 0)
	c.Assert(cell.styles.Fonts, HasLen, 1)
	xFont := cell.styles.Fonts[0]
	c.Assert(xFont.Sz.Val, Equals, "12")
	c.Assert(xFont.Name.Val, Equals, "Calibra")
}


// Test that GetStyle correctly converts the xlsxStyle.Fills.
func (s *CellSuite) TestGetStyleWithFills(c *C) {
	var cell *Cell
	var style *Style
	var xStyles *xlsxStyles
	var fills []xlsxFill
	var cellXfs []xlsxXf

	fills = make([]xlsxFill, 1)
	fills[0] = xlsxFill{
		PatternFill: xlsxPatternFill{
			PatternType: "solid",
			FgColor:     xlsxColor{RGB: "FF000000"},
			BgColor:     xlsxColor{RGB: "00FF0000"}}}
	cellXfs = make([]xlsxXf, 1)
	cellXfs[0] = xlsxXf{ApplyFill: true, FillId: 0}

	xStyles = &xlsxStyles{Fills: fills, CellXfs: cellXfs}

	cell = &Cell{Value: "123", styleIndex: 1, styles: xStyles}
	style = cell.GetStyle()
	fill := style.Fill
	c.Assert(fill.PatternType, Equals, "solid")
	c.Assert(fill.BgColor, Equals, "00FF0000")
	c.Assert(fill.FgColor, Equals, "FF000000")
}

// Test that SetStyle correctly updates xlsxStyle.Fills.
func (s *CellSuite) TestSetStyleWithFills(c *C) {
	file := NewFile()
	sheet := file.AddSheet("Test")
	row := sheet.AddRow()
	cell := row.AddCell()
	fill := NewFill(
}


// Test that GetStyle correctly converts the xlsxStyle.Borders.
func (s *CellSuite) TestGetStyleWithBorders(c *C) {
	var cell *Cell
	var style *Style
	var xStyles *xlsxStyles
	var borders []xlsxBorder
	var cellXfs []xlsxXf

	borders = make([]xlsxBorder, 1)
	borders[0] = xlsxBorder{
		Left:   xlsxLine{Style: "thin"},
		Right:  xlsxLine{Style: "thin"},
		Top:    xlsxLine{Style: "thin"},
		Bottom: xlsxLine{Style: "thin"}}

	cellXfs = make([]xlsxXf, 1)
	cellXfs[0] = xlsxXf{ApplyBorder: true, BorderId: 0}

	xStyles = &xlsxStyles{Borders: borders, CellXfs: cellXfs}

	cell = &Cell{Value: "123", styleIndex: 1, styles: xStyles}
	style = cell.GetStyle()
	border := style.Border
	c.Assert(border.Left, Equals, "thin")
	c.Assert(border.Right, Equals, "thin")
	c.Assert(border.Top, Equals, "thin")
	c.Assert(border.Bottom, Equals, "thin")
}
