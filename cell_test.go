package xlsx

import (
	. "gopkg.in/check.v1"
)

type StyleSuite struct{}

var _ = Suite(&StyleSuite{})

func (s *StyleSuite) TestNewStyle(c *C) {
	style := NewStyle()
	c.Assert(style, NotNil)
}

type FontSuite struct{}

var _ = Suite(&FontSuite{})

func (s *FontSuite) TestNewFont(c *C) {
	font := NewFont(12, "Verdana")
	c.Assert(font, NotNil)
	c.Assert(font.Name, Equals, "Verdana")
	c.Assert(font.Size, Equals, 12)
}

type CellSuite struct{}

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
	var style Style
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

	cell = &Cell{Value: "123", styleIndex: 0, styles: xStyles}
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
	style.Font = *font
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
	var style Style
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

	cell = &Cell{Value: "123", styleIndex: 0, styles: xStyles}
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
	fill := NewFill("solid", "00FF0000", "FF000000")
	style := NewStyle()
	style.Fill = *fill
	cell.SetStyle(style)
	c.Assert(cell.styleIndex, Equals, 0)
	c.Assert(cell.styles.Fills, HasLen, 1)
	xFill := cell.styles.Fills[0]
	xPatternFill := xFill.PatternFill
	c.Assert(xPatternFill.PatternType, Equals, "solid")
	c.Assert(xPatternFill.FgColor.RGB, Equals, "00FF0000")
	c.Assert(xPatternFill.BgColor.RGB, Equals, "FF000000")
}

// Test that GetStyle correctly converts the xlsxStyle.Borders.
func (s *CellSuite) TestGetStyleWithBorders(c *C) {
	var cell *Cell
	var style Style
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

	cell = &Cell{Value: "123", styleIndex: 0, styles: xStyles}
	style = cell.GetStyle()
	border := style.Border
	c.Assert(border.Left, Equals, "thin")
	c.Assert(border.Right, Equals, "thin")
	c.Assert(border.Top, Equals, "thin")
	c.Assert(border.Bottom, Equals, "thin")
}

func (s *CellSuite) TestGetNumberFormat(c *C) {
	var cell *Cell
	var cellXfs []xlsxXf
	var numFmt xlsxNumFmt
	var numFmts []xlsxNumFmt
	var xStyles *xlsxStyles
	var numFmtRefTable map[int]xlsxNumFmt

	cellXfs = make([]xlsxXf, 1)
	cellXfs[0] = xlsxXf{NumFmtId: 0}

	numFmts = make([]xlsxNumFmt, 1)
	numFmtRefTable = make(map[int]xlsxNumFmt)

	xStyles = &xlsxStyles{NumFmts: numFmts, CellXfs: cellXfs}

	cell = &Cell{Value: "123.123", numFmtRefTable: numFmtRefTable, styleIndex: 0, styles: xStyles}

	numFmt = xlsxNumFmt{NumFmtId: 0, FormatCode: "dd/mm/yy"}
	numFmts[0] = numFmt
	numFmtRefTable[0] = numFmt
	c.Assert(cell.GetNumberFormat(), Equals, "dd/mm/yy")
}

// We can return a string representation of the formatted data
func (l *CellSuite) TestFormattedValue(c *C) {
	var cell, earlyCell, negativeCell, smallCell *Cell
	var cellXfs []xlsxXf
	var numFmt xlsxNumFmt
	var numFmts []xlsxNumFmt
	var xStyles *xlsxStyles
	var numFmtRefTable map[int]xlsxNumFmt

	cellXfs = make([]xlsxXf, 1)
	cellXfs[0] = xlsxXf{NumFmtId: 1}

	numFmts = make([]xlsxNumFmt, 1)
	numFmtRefTable = make(map[int]xlsxNumFmt)

	xStyles = &xlsxStyles{NumFmts: numFmts, CellXfs: cellXfs}
	cell = &Cell{Value: "37947.7500001", numFmtRefTable: numFmtRefTable, styleIndex: 0, styles: xStyles}
	negativeCell = &Cell{Value: "-37947.7500001", numFmtRefTable: numFmtRefTable, styleIndex: 0, styles: xStyles}
	smallCell = &Cell{Value: "0.007", numFmtRefTable: numFmtRefTable, styleIndex: 0, styles: xStyles}
	earlyCell = &Cell{Value: "2.1", numFmtRefTable: numFmtRefTable, styleIndex: 0, styles: xStyles}
	setCode := func(code string) {
		numFmt = xlsxNumFmt{NumFmtId: 1, FormatCode: code}
		numFmts[0] = numFmt
		numFmtRefTable[1] = numFmt
	}

	setCode("general")
	c.Assert(cell.FormattedValue(), Equals, "37947.7500001")
	c.Assert(negativeCell.FormattedValue(), Equals, "-37947.7500001")

	setCode("0")
	c.Assert(cell.FormattedValue(), Equals, "37947")

	setCode("#,##0") // For the time being we're not doing this
	// comma formatting, so it'll fall back to
	// the related non-comma form.
	c.Assert(cell.FormattedValue(), Equals, "37947")

	setCode("0.00")
	c.Assert(cell.FormattedValue(), Equals, "37947.75")

	setCode("#,##0.00") // For the time being we're not doing this
	// comma formatting, so it'll fall back to
	// the related non-comma form.
	c.Assert(cell.FormattedValue(), Equals, "37947.75")

	setCode("#,##0 ;(#,##0)")
	c.Assert(cell.FormattedValue(), Equals, "37947")
	c.Assert(negativeCell.FormattedValue(), Equals, "(37947)")

	setCode("#,##0 ;[red](#,##0)")
	c.Assert(cell.FormattedValue(), Equals, "37947")
	c.Assert(negativeCell.FormattedValue(), Equals, "(37947)")

	setCode("0%")
	c.Assert(cell.FormattedValue(), Equals, "3794775%")

	setCode("0.00%")
	c.Assert(cell.FormattedValue(), Equals, "3794775.00%")

	setCode("0.00e+00")
	c.Assert(cell.FormattedValue(), Equals, "3.794775e+04")

	setCode("##0.0e+0") // This is wrong, but we'll use it for now.
	c.Assert(cell.FormattedValue(), Equals, "3.794775e+04")

	setCode("mm-dd-yy")
	c.Assert(cell.FormattedValue(), Equals, "11-22-03")

	setCode("d-mmm-yy")
	c.Assert(cell.FormattedValue(), Equals, "22-Nov-03")
	c.Assert(earlyCell.FormattedValue(), Equals, "1-Jan-00")

	setCode("d-mmm")
	c.Assert(cell.FormattedValue(), Equals, "22-Nov")
	c.Assert(earlyCell.FormattedValue(), Equals, "1-Jan")

	setCode("mmm-yy")
	c.Assert(cell.FormattedValue(), Equals, "Nov-03")

	setCode("h:mm am/pm")
	c.Assert(cell.FormattedValue(), Equals, "6:00 pm")
	c.Assert(smallCell.FormattedValue(), Equals, "12:14 am")

	setCode("h:mm:ss am/pm")
	c.Assert(cell.FormattedValue(), Equals, "6:00:00 pm")
	c.Assert(smallCell.FormattedValue(), Equals, "12:14:47 am")

	setCode("h:mm")
	c.Assert(cell.FormattedValue(), Equals, "18:00")
	c.Assert(smallCell.FormattedValue(), Equals, "00:14")

	setCode("h:mm:ss")
	c.Assert(cell.FormattedValue(), Equals, "18:00:00")
	// This is wrong, but there's no eary way aroud it in Go right now, AFAICT.
	c.Assert(smallCell.FormattedValue(), Equals, "00:14:47")

	setCode("m/d/yy h:mm")
	c.Assert(cell.FormattedValue(), Equals, "11/22/03 18:00")
	c.Assert(smallCell.FormattedValue(), Equals, "12/30/99 00:14") // Note, that's 1899
	c.Assert(earlyCell.FormattedValue(), Equals, "1/1/00 02:24")   // and 1900

	setCode("mm:ss")
	c.Assert(cell.FormattedValue(), Equals, "00:00")
	c.Assert(smallCell.FormattedValue(), Equals, "14:47")

	setCode("[h]:mm:ss")
	c.Assert(cell.FormattedValue(), Equals, "18:00:00")
	c.Assert(smallCell.FormattedValue(), Equals, "14:47")

	setCode("mmss.0") // I'm not sure about these.
	c.Assert(cell.FormattedValue(), Equals, "00.8640")
	c.Assert(smallCell.FormattedValue(), Equals, "1447.999997")

	setCode("yyyy\\-mm\\-dd")
	c.Assert(cell.FormattedValue(), Equals, "2003\\-11\\-22")

	setCode("dd/mm/yy")
	c.Assert(cell.FormattedValue(), Equals, "22/11/03")
	c.Assert(earlyCell.FormattedValue(), Equals, "01/01/00")

	setCode("hh:mm:ss")
	c.Assert(cell.FormattedValue(), Equals, "18:00:00")
	c.Assert(smallCell.FormattedValue(), Equals, "00:14:47")

	setCode("dd/mm/yy\\ hh:mm")
	c.Assert(cell.FormattedValue(), Equals, "22/11/03\\ 18:00")

	setCode("yy-mm-dd")
	c.Assert(cell.FormattedValue(), Equals, "03-11-22")

	setCode("d-mmm-yyyy")
	c.Assert(cell.FormattedValue(), Equals, "22-Nov-2003")
	c.Assert(earlyCell.FormattedValue(), Equals, "1-Jan-1900")

	setCode("m/d/yy")
	c.Assert(cell.FormattedValue(), Equals, "11/22/03")
	c.Assert(earlyCell.FormattedValue(), Equals, "1/1/00")

	setCode("m/d/yyyy")
	c.Assert(cell.FormattedValue(), Equals, "11/22/2003")
	c.Assert(earlyCell.FormattedValue(), Equals, "1/1/1900")

	setCode("dd-mmm-yyyy")
	c.Assert(cell.FormattedValue(), Equals, "22-Nov-2003")

	setCode("dd/mm/yyyy")
	c.Assert(cell.FormattedValue(), Equals, "22/11/2003")

	setCode("mm/dd/yy hh:mm am/pm")
	c.Assert(cell.FormattedValue(), Equals, "11/22/03 06:00 pm")

	setCode("mm/dd/yyyy hh:mm:ss")
	c.Assert(cell.FormattedValue(), Equals, "11/22/2003 18:00:00")
	c.Assert(smallCell.FormattedValue(), Equals, "12/30/1899 00:14:47")

	setCode("yyyy-mm-dd hh:mm:ss")
	c.Assert(cell.FormattedValue(), Equals, "2003-11-22 18:00:00")
	c.Assert(smallCell.FormattedValue(), Equals, "1899-12-30 00:14:47")
}
