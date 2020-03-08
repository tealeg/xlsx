package xlsx

import (
	"testing"

	qt "github.com/frankban/quicktest"
	. "gopkg.in/check.v1"
)

type StyleSuite struct{}

var _ = Suite(&StyleSuite{})

func (s *StyleSuite) TestNewStyle(c *C) {
	style := NewStyle()
	c.Assert(style, NotNil)
}

func (s *StyleSuite) TestNewStyleDefaultts(c *C) {
	style := NewStyle()
	c.Assert(style.Font, Equals, *DefaultFont())
	c.Assert(style.Fill, Equals, *DefaultFill())
	c.Assert(style.Border, Equals, *DefaultBorder())
}

func (s *StyleSuite) TestMakeXLSXStyleElements(c *C) {
	style := NewStyle()
	font := *NewFont(12, "Verdana")
	font.Bold = true
	font.Italic = true
	font.Underline = true
	style.Font = font
	fill := *NewFill("solid", "00FF0000", "FF000000")
	style.Fill = fill
	border := *NewBorder("thin", "thin", "thin", "thin")
	style.Border = border
	style.ApplyBorder = true
	style.ApplyFill = true

	style.ApplyFont = true
	xFont, xFill, xBorder, xCellXf := style.makeXLSXStyleElements()
	c.Assert(xFont.Sz.Val, Equals, "12")
	c.Assert(xFont.Name.Val, Equals, "Verdana")
	c.Assert(xFont.B, NotNil)
	c.Assert(xFont.I, NotNil)
	c.Assert(xFont.U, NotNil)
	c.Assert(xFill.PatternFill.PatternType, Equals, "solid")
	c.Assert(xFill.PatternFill.FgColor.RGB, Equals, "00FF0000")
	c.Assert(xFill.PatternFill.BgColor.RGB, Equals, "FF000000")
	c.Assert(xBorder.Left.Style, Equals, "thin")
	c.Assert(xBorder.Right.Style, Equals, "thin")
	c.Assert(xBorder.Top.Style, Equals, "thin")
	c.Assert(xBorder.Bottom.Style, Equals, "thin")
	c.Assert(xCellXf.ApplyBorder, Equals, true)
	c.Assert(xCellXf.ApplyFill, Equals, true)
	c.Assert(xCellXf.ApplyFont, Equals, true)

}

func TestReadCellColorBackground(t *testing.T) {
	c := qt.New(t)
	csRunO(c, "ReadCellColorBackground", func(c *qt.C, option FileOption) {
		xFile, err := OpenFile("./testdocs/color_stylesheet.xlsx", option)
		c.Assert(err, qt.Equals, nil)
		c.Assert(xFile.styles.Fills.Fill, qt.HasLen, 4)
		c.Assert(xFile.styles.Colors.IndexedColors, qt.HasLen, 64)
		sheet := xFile.Sheets[0]
		cell, err := sheet.Cell(0, 1)
		c.Assert(err, qt.Equals, nil)
		style := cell.GetStyle()
		c.Assert(style.Fill, qt.Equals, *NewFill("none", "", ""))
		cell, err = sheet.Cell(1, 1)
		c.Assert(err, qt.Equals, nil)
		style = cell.GetStyle()
		c.Assert(style.Fill, qt.Equals, *NewFill("solid", "00CC99FF", "00333333"))
		cell, err = sheet.Cell(2, 1)
		c.Assert(err, qt.Equals, nil)
		style = cell.GetStyle()
		c.Assert(style.Fill, qt.Equals, *NewFill("solid", "FF990099", "00333333"))
	})
}

type FontSuite struct{}

var _ = Suite(&FontSuite{})

func (s *FontSuite) TestNewFont(c *C) {
	font := NewFont(12, "Verdana")
	c.Assert(font, NotNil)
	c.Assert(font.Name, Equals, "Verdana")
	c.Assert(font.Size, Equals, 12)
}
