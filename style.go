package xlsx

import "strconv"

// Style is a high level structure intended to provide user access to
// the contents of Style within an XLSX file.
type Style struct {
	Border      Border
	Fill        Fill
	Font        Font
	ApplyBorder bool
	ApplyFill   bool
	ApplyFont   bool
}

// Return a new Style structure initialised with the default values.
func NewStyle() *Style {
	return &Style{Font: *DefaulFont()}
}

// Generate the underlying XLSX style elements that correspond to the Style.
func (style *Style) makeXLSXStyleElements() (xNumFmt xlsxNumFmt, xFont xlsxFont, xFill xlsxFill, xBorder xlsxBorder, xCellStyleXf xlsxXf, xCellXf xlsxXf) {
	xFont = xlsxFont{}
	xFill = xlsxFill{}
	xBorder = xlsxBorder{}
	xCellStyleXf = xlsxXf{}
	xCellXf = xlsxXf{}
	xFont.Sz.Val = strconv.Itoa(style.Font.Size)
	xFont.Name.Val = style.Font.Name
	xFont.Family.Val = strconv.Itoa(style.Font.Family)
	xFont.Charset.Val = strconv.Itoa(style.Font.Charset)
	xFont.Color.RGB = style.Font.Color
	xPatternFill := xlsxPatternFill{}
	xPatternFill.PatternType = style.Fill.PatternType
	xPatternFill.FgColor.RGB = style.Fill.FgColor
	xPatternFill.BgColor.RGB = style.Fill.BgColor
	xFill.PatternFill = xPatternFill
	xBorder.Left = xlsxLine{Style: style.Border.Left}
	xBorder.Right = xlsxLine{Style: style.Border.Right}
	xBorder.Top = xlsxLine{Style: style.Border.Top}
	xBorder.Bottom = xlsxLine{Style: style.Border.Bottom}
	xCellXf.ApplyBorder = style.ApplyBorder
	xCellXf.ApplyFill = style.ApplyFill
	xCellXf.ApplyFont = style.ApplyFont
	xCellXf.NumFmtId = 0
	xCellStyleXf.ApplyBorder = style.ApplyBorder
	xCellStyleXf.ApplyFill = style.ApplyFill
	xCellStyleXf.ApplyFont = style.ApplyFont
	xCellStyleXf.NumFmtId = 0
	return
}

// Border is a high level structure intended to provide user access to
// the contents of Border Style within an Sheet.
type Border struct {
	Left   string
	Right  string
	Top    string
	Bottom string
}

func NewBorder(left, right, top, bottom string) *Border {
	return &Border{Left: left, Right: right, Top: top, Bottom: bottom}
}

// Fill is a high level structure intended to provide user access to
// the contents of background and foreground color index within an Sheet.
type Fill struct {
	PatternType string
	BgColor     string
	FgColor     string
}

func NewFill(patternType, fgColor, bgColor string) *Fill {
	return &Fill{PatternType: patternType, FgColor: fgColor, BgColor: bgColor}
}

type Font struct {
	Size    int
	Name    string
	Family  int
	Charset int
	Color   string
}

func NewFont(size int, name string) *Font {
	return &Font{Size: size, Name: name}
}

func DefaulFont() *Font {
	return NewFont(12, "Verdana")
}
