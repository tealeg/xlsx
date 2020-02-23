package xlsx

import (
	"fmt"
	"reflect"
)

type RichTextFontFamily int
type RichTextCharset int
type RichTextVertAlign string
type RichTextUnderline string

const (
	// RichTextFontFamilyUnspecified indicates that the font family was not specified
	RichTextFontFamilyUnspecified   RichTextFontFamily = -1
	RichTextFontFamilyNotApplicable RichTextFontFamily = 0
	RichTextFontFamilyRoman         RichTextFontFamily = 1
	RichTextFontFamilySwiss         RichTextFontFamily = 2
	RichTextFontFamilyModern        RichTextFontFamily = 3
	RichTextFontFamilyScript        RichTextFontFamily = 4
	RichTextFontFamilyDecorative    RichTextFontFamily = 5

	// RichTextCharsetUnspecified indicates that the font charset was not specified
	RichTextCharsetUnspecified RichTextCharset = -1
	RichTextCharsetANSI        RichTextCharset = 0
	RichTextCharsetDefault     RichTextCharset = 1
	RichTextCharsetSymbol      RichTextCharset = 2
	RichTextCharsetMac         RichTextCharset = 77
	RichTextCharsetShiftJIS    RichTextCharset = 128
	RichTextCharsetHangul      RichTextCharset = 129
	RichTextCharsetJohab       RichTextCharset = 130
	RichTextCharsetGB2312      RichTextCharset = 134
	RichTextCharsetBIG5        RichTextCharset = 136
	RichTextCharsetGreek       RichTextCharset = 161
	RichTextCharsetTurkish     RichTextCharset = 162
	RichTextCharsetVietnamese  RichTextCharset = 163
	RichTextCharsetHebrew      RichTextCharset = 177
	RichTextCharsetArabic      RichTextCharset = 178
	RichTextCharsetBaltic      RichTextCharset = 186
	RichTextCharsetRussian     RichTextCharset = 204
	RichTextCharsetThai        RichTextCharset = 222
	RichTextCharsetEastEurope  RichTextCharset = 238
	RichTextCharsetOEM         RichTextCharset = 255

	RichTextVertAlignSuperscript RichTextVertAlign = "superscript"
	RichTextVertAlignSubscript   RichTextVertAlign = "subscript"

	RichTextUnderlineSingle RichTextUnderline = "single"
	RichTextUnderlineDouble RichTextUnderline = "double"

	// These underline styles doesn't work on the RichTextRun,
	// and should be set as a part of cell style.
	// "singleAccounting"
	// "doubleAccounting"
)

// RichTextColor is the color of the RichTextRun.
type RichTextColor struct {
	coreColor xlsxColor
}

// NewRichTextColorFromARGB creates a new RichTextColor from ARGB component values.
// Each component must have a value in range of 0 to 255.
func NewRichTextColorFromARGB(alpha, red, green, blue int) *RichTextColor {
	argb := fmt.Sprintf("%02X%02X%02X%02X", alpha, red, green, blue)
	return &RichTextColor{coreColor: xlsxColor{RGB: argb}}
}

// NewRichTextColorFromThemeColor creates a new RichTextColor from the theme color.
// The argument `themeColor` is a zero-based index of the theme color.
func NewRichTextColorFromThemeColor(themeColor int) *RichTextColor {
	return &RichTextColor{coreColor: xlsxColor{Theme: &themeColor}}
}

// RichTextFont is the font spec of the RichTextRun.
type RichTextFont struct {
	// Name is the font name. If Name is empty, Size, Family and Charset will be ignored.
	Name string
	// Size is the font size.
	Size float64
	// Family is a value of the font family. Use one of the RichTextFontFamily constants.
	Family RichTextFontFamily
	// Charset is a value of the charset of the font. Use one of the RichTextCharset constants.
	Charset RichTextCharset
	// Color is the text color.
	Color *RichTextColor
	// Bold specifies the bold face font style.
	Bold bool
	// Italic specifies the italic font style.
	Italic bool
	// Strike specifies a strikethrough line.
	Strike bool
	// VertAlign specifies the vertical position of the text. Use one of the RichTextVertAlign constants, or empty.
	VertAlign RichTextVertAlign
	// Underline specifies the underline style. Use one of the RichTextUnderline constants, or empty.
	Underline RichTextUnderline
}

// RichTextRun is a run of the decorated text.
type RichTextRun struct {
	Font *RichTextFont
	Text string
}

func (rt *RichTextRun) Equals(other *RichTextRun) bool {
	return reflect.DeepEqual(rt, other)
}

func richTextToXml(r []RichTextRun) []xlsxR {
	var xrs []xlsxR
	for _, rt := range r {
		xr := xlsxR{}
		xr.T = xlsxT{Text: rt.Text}
		if rt.Font != nil {
			rpr := xlsxRunProperties{}
			if len(rt.Font.Name) > 0 {
				rpr.RFont = &xlsxVal{Val: rt.Font.Name}
			}
			if rt.Font.Size > 0.0 {
				rpr.Sz = &xlsxFloatVal{Val: rt.Font.Size}
			}
			if rt.Font.Family != RichTextFontFamilyUnspecified {
				rpr.Family = &xlsxIntVal{Val: int(rt.Font.Family)}
			}
			if rt.Font.Charset != RichTextCharsetUnspecified {
				rpr.Charset = &xlsxIntVal{Val: int(rt.Font.Charset)}
			}
			if rt.Font.Color != nil {
				xcolor := rt.Font.Color.coreColor
				rpr.Color = &xcolor
			}
			if rt.Font.Bold {
				rpr.B.Val = true
			}
			if rt.Font.Italic {
				rpr.I.Val = true
			}
			if rt.Font.Strike {
				rpr.Strike.Val = true
			}
			if len(rt.Font.VertAlign) > 0 {
				rpr.VertAlign = &xlsxVal{Val: string(rt.Font.VertAlign)}
			}
			if len(rt.Font.Underline) > 0 {
				rpr.U = &xlsxVal{Val: string(rt.Font.Underline)}
			}
			xr.RPr = &rpr
		}
		xrs = append(xrs, xr)
	}
	return xrs
}

func xmlToRichText(r []xlsxR) []RichTextRun {
	richiText := []RichTextRun(nil)
	for _, rr := range r {
		rtr := RichTextRun{Text: rr.T.Text}
		rpr := rr.RPr
		if rpr != nil {
			rtr.Font = &RichTextFont{}
			if rpr.RFont != nil {
				rtr.Font.Name = rpr.RFont.Val
			}
			if rpr.Sz != nil {
				rtr.Font.Size = rpr.Sz.Val
			}
			if rpr.Family != nil {
				rtr.Font.Family = RichTextFontFamily(rpr.Family.Val)
			} else {
				rtr.Font.Family = RichTextFontFamilyUnspecified
			}
			if rpr.Charset != nil {
				rtr.Font.Charset = RichTextCharset(rpr.Charset.Val)
			} else {
				rtr.Font.Charset = RichTextCharsetUnspecified
			}
			if rpr.Color != nil {
				rtr.Font.Color = &RichTextColor{coreColor: *rpr.Color}
			}
			if rpr.B.Val {
				rtr.Font.Bold = true
			}
			if rpr.I.Val {
				rtr.Font.Italic = true
			}
			if rpr.Strike.Val {
				rtr.Font.Strike = true
			}
			if rpr.VertAlign != nil {
				rtr.Font.VertAlign = RichTextVertAlign(rpr.VertAlign.Val)
			}
			if rpr.U != nil {
				rtr.Font.Underline = RichTextUnderline(rpr.U.Val)
			}
		}
		richiText = append(richiText, rtr)
	}
	return richiText
}

func richTextToPlainText(richText []RichTextRun) string {
	var s string
	for _, r := range richText {
		s += r.Text
	}
	return s
}
