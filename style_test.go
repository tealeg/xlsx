package xlsx

import (
	"testing"

	qt "github.com/frankban/quicktest"
)

func TestStyle(t *testing.T) {

	c := qt.New(t)

	c.Run("TestNewStyle", func(c *qt.C) {
		style := NewStyle()
		c.Assert(style, qt.Not(qt.IsNil))
	})

	c.Run("TestNewStyleDefaultts", func(c *qt.C) {
		style := NewStyle()
		c.Assert(style.Font, qt.Equals, *DefaultFont())
		c.Assert(style.Fill, qt.Equals, *DefaultFill())
		c.Assert(style.Border, qt.Equals, *DefaultBorder())
	})

	c.Run("TestMakeXLSXStyleElements", func(c *qt.C) {
		style := NewStyle()
		font := *NewFont(12, "Verdana")
		font.Bold = true
		font.Italic = true
		font.Underline = true
		font.Strike = true
		style.Font = font
		fill := *NewFill("solid", "00FF0000", "FF000000")
		style.Fill = fill
		border := *NewBorder("thin", "thin", "thin", "thin")
		style.Border = border
		style.ApplyBorder = true
		style.ApplyFill = true

		style.ApplyFont = true
		xFont, xFill, xBorder, xCellXf := style.makeXLSXStyleElements()
		c.Assert(xFont.Sz.Val, qt.Equals, "12")
		c.Assert(xFont.Name.Val, qt.Equals, "Verdana")
		c.Assert(xFont.B, qt.Not(qt.IsNil))
		c.Assert(xFont.I, qt.Not(qt.IsNil))
		c.Assert(xFont.U, qt.Not(qt.IsNil))
		c.Assert(xFont.Strike, qt.Not(qt.IsNil))
		c.Assert(xFill.PatternFill.PatternType, qt.Equals, "solid")
		c.Assert(xFill.PatternFill.FgColor.RGB, qt.Equals, "00FF0000")
		c.Assert(xFill.PatternFill.BgColor.RGB, qt.Equals, "FF000000")
		c.Assert(xBorder.Left.Style, qt.Equals, "thin")
		c.Assert(xBorder.Right.Style, qt.Equals, "thin")
		c.Assert(xBorder.Top.Style, qt.Equals, "thin")
		c.Assert(xBorder.Bottom.Style, qt.Equals, "thin")
		c.Assert(xCellXf.ApplyBorder, qt.Equals, true)
		c.Assert(xCellXf.ApplyFill, qt.Equals, true)
		c.Assert(xCellXf.ApplyFont, qt.Equals, true)

	})
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

func TestNewFont(t *testing.T) {
	c := qt.New(t)
	font := NewFont(12.2, "Verdana")
	c.Assert(font, qt.Not(qt.IsNil))
	c.Assert(font.Name, qt.Equals, "Verdana")
	c.Assert(font.Size, qt.Equals, 12.2)
}
