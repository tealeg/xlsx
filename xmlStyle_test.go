package xlsx

import (
	. "gopkg.in/check.v1"
)

type XMLStyleSuite struct{}

var _ = Suite(&XMLStyleSuite{})

// Test we produce valid output for an empty style file.
func (x *XMLStyleSuite) TestMarshalEmptyXlsxStyleSheet(c *C) {
	styles := &xlsxStyleSheet{}
	result, err := styles.Marshal()
	c.Assert(err, IsNil)
	c.Assert(string(result), Equals, `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></styleSheet>`)
}

// Test we produce valid output for a style file with one font definition.
func (x *XMLStyleSuite) TestMarshalXlsxStyleSheetWithAFont(c *C) {
	styles := &xlsxStyleSheet{}
	styles.Fonts = xlsxFonts{}
	styles.Fonts.Count = 1
	styles.Fonts.Font = make([]xlsxFont, 1)
	font := xlsxFont{}
	font.Sz.Val = "10"
	font.Name.Val = "Andale Mono"
	styles.Fonts.Font[0] = font

	expected := `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="10"/><name val="Andale Mono"/></font></fonts></styleSheet>`
	result, err := styles.Marshal()
	c.Assert(err, IsNil)
	c.Assert(string(result), Equals, expected)
}

// Test we produce valid output for a style file with one fill definition.
func (x *XMLStyleSuite) TestMarshalXlsxStyleSheetWithAFill(c *C) {
	styles := &xlsxStyleSheet{}
	styles.Fills = xlsxFills{}
	styles.Fills.Count = 1
	styles.Fills.Fill = make([]xlsxFill, 1)
	fill := xlsxFill{}
	patternFill := xlsxPatternFill{
		PatternType: "solid",
		FgColor:     xlsxColor{RGB: "#FFFFFF"},
		BgColor:     xlsxColor{RGB: "#000000"}}
	fill.PatternFill = patternFill
	styles.Fills.Fill[0] = fill

	expected := `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fills count="1"><fill><patternFill patternType="solid"><fgColor rgb="#FFFFFF"/><bgColor rgb="#000000"/></patternFill></fill></fills></styleSheet>`
	result, err := styles.Marshal()
	c.Assert(err, IsNil)
	c.Assert(string(result), Equals, expected)
}

// Test we produce valid output for a style file with one border definition.
func (x *XMLStyleSuite) TestMarshalXlsxStyleSheetWithABorder(c *C) {
	styles := &xlsxStyleSheet{}
	styles.Borders = xlsxBorders{}
	styles.Borders.Count = 1
	styles.Borders.Border = make([]xlsxBorder, 1)
	border := xlsxBorder{}
	border.Left.Style = "solid"
	border.Top.Style = "none"
	styles.Borders.Border[0] = border

	expected := `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><borders count="1"><border><left style="solid"/><top style="none"/></border></borders></styleSheet>`
	result, err := styles.Marshal()
	c.Assert(err, IsNil)
	c.Assert(string(result), Equals, expected)
}

// Test we produce valid output for a style file with one cellStyleXf definition.
func (x *XMLStyleSuite) TestMarshalXlsxStyleSheetWithACellStyleXf(c *C) {
	styles := &xlsxStyleSheet{}
	styles.CellStyleXfs = xlsxCellStyleXfs{}
	styles.CellStyleXfs.Count = 1
	styles.CellStyleXfs.Xf = make([]xlsxXf, 1)
	xf := xlsxXf{}
	xf.ApplyAlignment = true
	xf.ApplyBorder = true
	xf.ApplyFont = true
	xf.ApplyFill = true
	xf.ApplyProtection = true
	xf.BorderId = 0
	xf.FillId = 0
	xf.FontId = 0
	xf.NumFmtId = 0
	xf.alignment = xlsxAlignment{
		Horizontal:   "left",
		Indent:       1,
		ShrinkToFit:  true,
		TextRotation: 0,
		Vertical:     "middle",
		WrapText:     false}
	styles.CellStyleXfs.Xf[0] = xf

	expected := `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellStyleXfs count="1"><xf applyAlignment="1" applyBorder="1" applyFont="1" applyFill="1" applyProtection="1" borderId="0" fillId="0" fontId="0" numFmtId="0"><alignment horizontal="left" indent="1" shrinkToFit="1" textRotation="0" vertical="middle" wrapText="0"/></xf></cellStyleXfs></styleSheet>`
	result, err := styles.Marshal()
	c.Assert(err, IsNil)
	c.Assert(string(result), Equals, expected)
}

// Test we produce valid output for a style file with one cellXf
// definition.
func (x *XMLStyleSuite) TestMarshalXlsxStyleSheetWithACellXf(c *C) {
	styles := &xlsxStyleSheet{}
	styles.CellXfs = xlsxCellXfs{}
	styles.CellXfs.Count = 1
	styles.CellXfs.Xf = make([]xlsxXf, 1)
	xf := xlsxXf{}
	xf.ApplyAlignment = true
	xf.ApplyBorder = true
	xf.ApplyFont = true
	xf.ApplyFill = true
	xf.ApplyProtection = true
	xf.BorderId = 0
	xf.FillId = 0
	xf.FontId = 0
	xf.NumFmtId = 0
	xf.alignment = xlsxAlignment{
		Horizontal:   "left",
		Indent:       1,
		ShrinkToFit:  true,
		TextRotation: 0,
		Vertical:     "middle",
		WrapText:     false}
	styles.CellXfs.Xf[0] = xf

	expected := `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="1"><xf applyAlignment="1" applyBorder="1" applyFont="1" applyFill="1" applyProtection="1" borderId="0" fillId="0" fontId="0" numFmtId="0"><alignment horizontal="left" indent="1" shrinkToFit="1" textRotation="0" vertical="middle" wrapText="0"/></xf></cellXfs></styleSheet>`
	result, err := styles.Marshal()
	c.Assert(err, IsNil)
	c.Assert(string(result), Equals, expected)
}

// Test we produce valid output for a style file with one NumFmt
// definition.
func (x *XMLStyleSuite) TestMarshalXlsxStyleSheetWithANumFmt(c *C) {
	styles := &xlsxStyleSheet{}
	styles.NumFmts = xlsxNumFmts{}
	styles.NumFmts.NumFmt = make([]xlsxNumFmt, 0)
	numFmt := xlsxNumFmt{NumFmtId: 164, FormatCode: "GENERAL"}
	styles.addNumFmt(numFmt)

	expected := `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="GENERAL"/></numFmts></styleSheet>`
	result, err := styles.Marshal()
	c.Assert(err, IsNil)
	c.Assert(string(result), Equals, expected)
}
