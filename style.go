package xlsx

import (
	"fmt"
	"log"
	"strconv"
	"strings"
)

// Several popular font names that can be used to create fonts
const (
	Helvetica     = "Helvetica"
	Baskerville   = "Baskerville Old Face"
	TimesNewRoman = "Times New Roman"
	Bodoni        = "Bodoni MT"
	GillSans      = "Gill Sans MT"
	Courier       = "Courier"
)

const (
	RGB_Light_Green = "FFC6EFCE"
	RGB_Dark_Green  = "FF006100"
	RGB_Light_Red   = "FFFFC7CE"
	RGB_Dark_Red    = "FF9C0006"
	RGB_White       = "FFFFFFFF"
	RGB_Black       = "00000000"
)

const (
	Solid_Cell_Fill = "solid"
)

// Style is a high level structure intended to provide user access to
// the contents of Style within an XLSX file.
type Style struct {
	Border          Border
	Fill            Fill
	Font            Font
	ApplyBorder     bool
	ApplyFill       bool
	ApplyFont       bool
	ApplyAlignment  bool
	Alignment       Alignment
	NamedStyleIndex *int
}

// Return a new Style structure initialised with the default values.
func NewStyle() *Style {
	return &Style{
		Alignment: *DefaultAlignment(),
		Border:    *DefaultBorder(),
		Fill:      *DefaultFill(),
		Font:      *DefaultFont(),
	}
}

// Generate the underlying XLSX style elements that correspond to the Style.
func (style *Style) makeXLSXStyleElements() (xFont xlsxFont, xFill xlsxFill, xBorder xlsxBorder, xCellXf xlsxXf) {
	if style == nil {
		panic("Called makeXLSXStyleElements on a nil *Style!")
	}

	xFont = xlsxFont{}
	xFill = xlsxFill{}
	xBorder = xlsxBorder{}
	xCellXf = xlsxXf{}
	xFont.Sz.Val = strconv.FormatFloat(style.Font.Size, 'f', -1, 64)
	xFont.Name.Val = style.Font.Name
	xFont.Family.Val = strconv.Itoa(style.Font.Family)
	xFont.Charset.Val = strconv.Itoa(style.Font.Charset)
	xFont.Color = style.Font.Color.asXlsxColor()

	if style.Font.Bold {
		xFont.B = &xlsxVal{}
	} else {
		xFont.B = nil
	}
	if style.Font.Italic {
		xFont.I = &xlsxVal{}
	} else {
		xFont.I = nil
	}
	if style.Font.Underline {
		xFont.U = &xlsxVal{}
	} else {
		xFont.U = nil
	}
	if style.Font.Strike {
		xFont.Strike = &xlsxVal{}
	} else {
		xFont.Strike = nil
	}
	xPatternFill := xlsxPatternFill{}
	xPatternFill.PatternType = style.Fill.PatternType
	xPatternFill.FgColor = style.Fill.FgColor.asXlsxColor()
	xPatternFill.BgColor = style.Fill.BgColor.asXlsxColor()
	xFill.PatternFill = xPatternFill
	xBorder.Left = xlsxLine{
		Style: style.Border.Left,
		Color: style.Border.LeftColor.asXlsxColor(),
	}
	xBorder.Right = xlsxLine{
		Style: style.Border.Right,
		Color: style.Border.RightColor.asXlsxColor(),
	}
	xBorder.Top = xlsxLine{
		Style: style.Border.Top,
		Color: style.Border.TopColor.asXlsxColor(),
	}
	xBorder.Bottom = xlsxLine{
		Style: style.Border.Bottom,
		Color: style.Border.BottomColor.asXlsxColor(),
	}
	xCellXf = makeXLSXCellElement()
	xCellXf.ApplyBorder = style.ApplyBorder
	xCellXf.ApplyFill = style.ApplyFill
	xCellXf.ApplyFont = style.ApplyFont
	xCellXf.ApplyAlignment = style.ApplyAlignment
	if style.NamedStyleIndex != nil {
		xCellXf.XfId = style.NamedStyleIndex
	}
	return
}

func makeXLSXCellElement() (xCellXf xlsxXf) {
	xCellXf.NumFmtId = 0
	return
}

// Color is a high level structure intended to provide user access to color defiinitions.
type Color struct {
	RGB     *string
	Theme   *int
	Tint    *float64
	Indexed *int
	Auto    *int
}

func NewColorFromRGB(rgb string) *Color {
	return &Color{RGB: sPtr(rgb)}
}

func NewColorFromXlsxColor(xC *xlsxColor) *Color {
	if xC == nil {
		return nil
	}
	return &Color{RGB: xC.RGB,
		Theme:   xC.Theme,
		Tint:    xC.Tint,
		Indexed: xC.Indexed,
		Auto:    xC.Auto,
	}
}

func (c *Color) String() string {

	if c == nil {
		return "<nil>"
	}

	var sb strings.Builder
	sb.WriteString("&Color{")
	if c.RGB != nil {
		sb.WriteString("RGB: ")
		sb.WriteString(*c.RGB)
	}
	if c.Theme != nil {
		sb.WriteString("Theme: ")
		sb.WriteString(strconv.Itoa(*c.Theme))
	}
	if c.Tint != nil {
		sb.WriteString("Tint: ")
		sb.WriteString(fmt.Sprintf("%f", *c.Tint))
	}
	if c.Indexed != nil {
		sb.WriteString("Indexed: ")
		sb.WriteString(strconv.Itoa(*c.Indexed))
	}
	if c.Auto != nil {
		sb.WriteString("Auto: ")
		sb.WriteString(strconv.Itoa(*c.Auto))
	}
	return sb.String()
}

func (c *Color) asXlsxColor() *xlsxColor {
	if c == nil {
		return nil
	}
	return &xlsxColor{RGB: c.RGB, Indexed: c.Indexed, Auto: c.Auto, Theme: c.Theme, Tint: c.Tint}
}

func (c *Color) Equals(o *Color) bool {
	if c == nil {
		return o == nil
	}
	if o == nil {
		return false
	}
	rgbMatch := c.RGB == o.RGB || (c.RGB != nil && *c.RGB == *o.RGB)
	log.Printf("RGB %v == %v => %t\n", c.RGB, o.RGB, rgbMatch)
	themeMatch := c.Theme == o.Theme || (c.Theme != nil && *c.Theme == *o.Theme)
	tintMatch := c.Tint == o.Tint || (c.Tint != nil && *c.Tint == *o.Tint)
	indexedMatch := c.Indexed == o.Indexed || (c.Indexed != nil && *c.Indexed == *o.Indexed)
	autoMatch := c.Auto == o.Auto || (c.Auto != nil && *c.Auto == *o.Auto)
	return rgbMatch && themeMatch && tintMatch && indexedMatch && autoMatch
}

// Border is a high level structure intended to provide user access to
// the contents of Border Style within an Sheet.
type Border struct {
	Left        string
	LeftColor   *Color
	Right       string
	RightColor  *Color
	Top         string
	TopColor    *Color
	Bottom      string
	BottomColor *Color
}

func NewBorder(left, right, top, bottom string) *Border {
	return &Border{
		Left:   left,
		Right:  right,
		Top:    top,
		Bottom: bottom,
	}
}

func (b *Border) Equals(o *Border) bool {
	if b == nil {
		return o == nil
	}
	log.Printf("%+v == %+v\n", b, o)
	valuesMatch := b.Left == o.Left && b.Right == o.Right && b.Top == o.Top && b.Bottom == o.Bottom
	leftMatch := b.LeftColor.Equals(o.LeftColor)
	rightMatch := b.RightColor.Equals(o.RightColor)
	topMatch := b.TopColor.Equals(o.TopColor)
	bottomMatch := b.BottomColor.Equals(o.BottomColor)
	return valuesMatch && leftMatch && rightMatch && topMatch && bottomMatch
}

// Fill is a high level structure intended to provide user access to
// the contents of background and foreground color index within an Sheet.
type Fill struct {
	PatternType string
	BgColor     *Color
	FgColor     *Color
}

func NewFill(patternType string, fgColor, bgColor *Color) *Fill {
	return &Fill{
		PatternType: patternType,
		FgColor:     fgColor,
		BgColor:     bgColor,
	}
}

func (f *Fill) Equals(o *Fill) bool {
	if f == nil {
		return o == nil
	}
	if o == nil {
		return false
	}
	return f.PatternType == o.PatternType &&
		f.FgColor.Equals(o.FgColor) && f.BgColor.Equals(o.BgColor)
}

type Font struct {
	Size      float64
	Name      string
	Family    int
	Charset   int
	Color     *Color
	Bold      bool
	Italic    bool
	Underline bool
	Strike    bool
}

func NewFont(size float64, name string) *Font {
	return &Font{Size: size, Name: name}
}

type Alignment struct {
	Horizontal   string
	Indent       int
	ShrinkToFit  bool
	TextRotation int
	Vertical     string
	WrapText     bool
}

var defaultFontSize = 12.0
var defaultFontName = "Verdana"

func SetDefaultFont(size float64, name string) {
	defaultFontSize = size
	defaultFontName = name
}

func DefaultFont() *Font {
	return NewFont(defaultFontSize, defaultFontName)
}

func DefaultFill() *Fill {
	return NewFill("none", nil, nil)

}

func DefaultBorder() *Border {
	return NewBorder("none", "none", "none", "none")
}

func DefaultAlignment() *Alignment {
	return &Alignment{
		Horizontal: "general",
		Vertical:   "bottom",
	}
}
