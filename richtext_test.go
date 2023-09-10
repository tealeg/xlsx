package xlsx

import (
	"testing"

	qt "github.com/frankban/quicktest"
)

func TestNewRichTextColorFromARGB(t *testing.T) {
	c := qt.New(t)
	rtColor := NewRichTextColorFromARGB(127, 128, 129, 130)
	c.Assert(rtColor.coreColor.RGB, qt.Equals, "7F808182")
}

func TestNewRichTextColorFromThemeColor(t *testing.T) {
	c := qt.New(t)
	rtColor := NewRichTextColorFromThemeColor(123)
	c.Assert(*rtColor.coreColor.Theme, qt.Equals, 123)
}

func TestRichTextRunEquals(t *testing.T) {
	c := qt.New(t)
	r1color := 1
	r1 := &RichTextRun{
		Font: &RichTextFont{
			Family:  RichTextFontFamilyUnspecified,
			Charset: RichTextCharsetUnspecified,
			Color:   &RichTextColor{coreColor: xlsxColor{Theme: &r1color}},
			Bold:    true,
			Italic:  true,
		},
		Text: "X",
	}

	r2color := r1color
	r2 := &RichTextRun{ // same with r1
		Font: &RichTextFont{
			Family:  RichTextFontFamilyUnspecified,
			Charset: RichTextCharsetUnspecified,
			Color:   &RichTextColor{coreColor: xlsxColor{Theme: &r2color}},
			Bold:    true,
			Italic:  true,
		},
		Text: "X",
	}

	r3color := r1color
	r3 := &RichTextRun{ // different font setting from r1
		Font: &RichTextFont{
			Family:  RichTextFontFamilyUnspecified,
			Charset: RichTextCharsetUnspecified,
			Color:   &RichTextColor{coreColor: xlsxColor{Theme: &r3color}},
			Bold:    true,
			Italic:  false,
		},
		Text: "X",
	}

	r4color := 2
	r4 := &RichTextRun{ // different color setting from r1
		Font: &RichTextFont{
			Family:  RichTextFontFamilyUnspecified,
			Charset: RichTextCharsetUnspecified,
			Color:   &RichTextColor{coreColor: xlsxColor{Theme: &r4color}},
			Bold:    true,
			Italic:  true,
		},
		Text: "X",
	}

	r5 := &RichTextRun{ // no font setting
		Text: "X",
	}

	r6color := r1color
	r6 := &RichTextRun{ // different text from r1
		Font: &RichTextFont{
			Family:  RichTextFontFamilyUnspecified,
			Charset: RichTextCharsetUnspecified,
			Color:   &RichTextColor{coreColor: xlsxColor{Theme: &r6color}},
			Bold:    true,
			Italic:  true,
		},
		Text: "Y",
	}

	var r7 *RichTextRun = nil

	c.Assert(r1.Equals(r2), qt.Equals, true)
	c.Assert(r1.Equals(r3), qt.Equals, false)
	c.Assert(r1.Equals(r4), qt.Equals, false)
	c.Assert(r1.Equals(r5), qt.Equals, false)
	c.Assert(r1.Equals(r6), qt.Equals, false)
	c.Assert(r1.Equals(r7), qt.Equals, false)

	c.Assert(r2.Equals(r1), qt.Equals, true)
	c.Assert(r3.Equals(r1), qt.Equals, false)
	c.Assert(r4.Equals(r1), qt.Equals, false)
	c.Assert(r5.Equals(r1), qt.Equals, false)
	c.Assert(r6.Equals(r1), qt.Equals, false)
	c.Assert(r7.Equals(r1), qt.Equals, false)

	c.Assert(r7.Equals(nil), qt.Equals, true)
}

func TestRichTextToXml(t *testing.T) {
	c := qt.New(t)
	rtr := []RichTextRun{
		{
			Font: &RichTextFont{
				Name:      "Font",
				Size:      12.345,
				Family:    RichTextFontFamilyScript,
				Charset:   RichTextCharsetHebrew,
				Color:     &RichTextColor{coreColor: xlsxColor{RGB: "DEADBEEF"}},
				Bold:      true,
				Italic:    false,
				Strike:    false,
				VertAlign: RichTextVertAlignSuperscript,
				Underline: RichTextUnderlineSingle,
			},
			Text: "Bold",
		},
		{
			Font: &RichTextFont{
				Family:  RichTextFontFamilyUnspecified,
				Charset: RichTextCharsetUnspecified,
				Italic:  true,
			},
			Text: "Italic",
		},
		{
			Font: &RichTextFont{
				Family:  RichTextFontFamilyUnspecified,
				Charset: RichTextCharsetUnspecified,
				Strike:  true,
			},
			Text: "Strike",
		},
		{
			Font: &RichTextFont{},
			Text: "Empty",
		},
		{
			Text: "No Font",
		},
	}

	xmlr := richTextToXml(rtr)
	c.Assert(xmlr, qt.HasLen, 5)

	r := xmlr[0]
	c.Assert(r.RPr.RFont.Val, qt.Equals, "Font")
	c.Assert(r.RPr.Charset.Val, qt.Equals, int(RichTextCharsetHebrew))
	c.Assert(r.RPr.Family.Val, qt.Equals, int(RichTextFontFamilyScript))
	c.Assert(r.RPr.B.Val, qt.Equals, true)
	c.Assert(r.RPr.I.Val, qt.Equals, false)
	c.Assert(r.RPr.Strike.Val, qt.Equals, false)
	c.Assert(r.RPr.Outline.Val, qt.Equals, false)
	c.Assert(r.RPr.Shadow.Val, qt.Equals, false)
	c.Assert(r.RPr.Condense.Val, qt.Equals, false)
	c.Assert(r.RPr.Extend.Val, qt.Equals, false)
	c.Assert(r.RPr.Color.RGB, qt.Equals, "DEADBEEF")
	c.Assert(r.RPr.Sz.Val, qt.Equals, 12.345)
	c.Assert(r.RPr.U.Val, qt.Equals, string(RichTextUnderlineSingle))
	c.Assert(r.RPr.VertAlign.Val, qt.Equals, string(RichTextVertAlignSuperscript))
	c.Assert(r.RPr.Scheme, qt.IsNil)
	c.Assert(r.T.Text, qt.Equals, "Bold")

	r = xmlr[1]
	c.Assert(r.RPr.RFont, qt.IsNil)
	c.Assert(r.RPr.Charset, qt.IsNil)
	c.Assert(r.RPr.Family, qt.IsNil)
	c.Assert(r.RPr.B.Val, qt.Equals, false)
	c.Assert(r.RPr.I.Val, qt.Equals, true)
	c.Assert(r.RPr.Strike.Val, qt.Equals, false)
	c.Assert(r.RPr.Outline.Val, qt.Equals, false)
	c.Assert(r.RPr.Shadow.Val, qt.Equals, false)
	c.Assert(r.RPr.Condense.Val, qt.Equals, false)
	c.Assert(r.RPr.Extend.Val, qt.Equals, false)
	c.Assert(r.RPr.Color, qt.IsNil)
	c.Assert(r.RPr.Sz, qt.IsNil)
	c.Assert(r.RPr.U, qt.IsNil)
	c.Assert(r.RPr.VertAlign, qt.IsNil)
	c.Assert(r.RPr.Scheme, qt.IsNil)
	c.Assert(r.T.Text, qt.Equals, "Italic")

	r = xmlr[2]
	c.Assert(r.RPr.RFont, qt.IsNil)
	c.Assert(r.RPr.Charset, qt.IsNil)
	c.Assert(r.RPr.Family, qt.IsNil)
	c.Assert(r.RPr.B.Val, qt.Equals, false)
	c.Assert(r.RPr.I.Val, qt.Equals, false)
	c.Assert(r.RPr.Strike.Val, qt.Equals, true)
	c.Assert(r.RPr.Outline.Val, qt.Equals, false)
	c.Assert(r.RPr.Shadow.Val, qt.Equals, false)
	c.Assert(r.RPr.Condense.Val, qt.Equals, false)
	c.Assert(r.RPr.Extend.Val, qt.Equals, false)
	c.Assert(r.RPr.Color, qt.IsNil)
	c.Assert(r.RPr.Sz, qt.IsNil)
	c.Assert(r.RPr.U, qt.IsNil)
	c.Assert(r.RPr.VertAlign, qt.IsNil)
	c.Assert(r.RPr.Scheme, qt.IsNil)
	c.Assert(r.T.Text, qt.Equals, "Strike")

	r = xmlr[3]
	c.Assert(r.RPr.RFont, qt.IsNil)
	c.Assert(r.RPr.Charset.Val, qt.Equals, int(RichTextCharsetANSI))
	c.Assert(r.RPr.Family.Val, qt.Equals, int(RichTextFontFamilyNotApplicable))
	c.Assert(r.RPr.B.Val, qt.Equals, false)
	c.Assert(r.RPr.I.Val, qt.Equals, false)
	c.Assert(r.RPr.Strike.Val, qt.Equals, false)
	c.Assert(r.RPr.Outline.Val, qt.Equals, false)
	c.Assert(r.RPr.Shadow.Val, qt.Equals, false)
	c.Assert(r.RPr.Condense.Val, qt.Equals, false)
	c.Assert(r.RPr.Extend.Val, qt.Equals, false)
	c.Assert(r.RPr.Color, qt.IsNil)
	c.Assert(r.RPr.Sz, qt.IsNil)
	c.Assert(r.RPr.U, qt.IsNil)
	c.Assert(r.RPr.VertAlign, qt.IsNil)
	c.Assert(r.RPr.Scheme, qt.IsNil)
	c.Assert(r.T.Text, qt.Equals, "Empty")

	r = xmlr[4]
	c.Assert(r.RPr, qt.IsNil)
	c.Assert(r.T.Text, qt.Equals, "No Font")
}

func TestXmlToRichText(t *testing.T) {
	c := qt.New(t)
	xmlr := []xlsxR{
		{
			RPr: &xlsxRunProperties{
				RFont:     &xlsxVal{Val: "Font"},
				Charset:   &xlsxIntVal{Val: int(RichTextCharsetGreek)},
				Family:    &xlsxIntVal{Val: int(RichTextFontFamilySwiss)},
				B:         xlsxBoolProp{Val: true},
				I:         xlsxBoolProp{Val: false},
				Strike:    xlsxBoolProp{Val: false},
				Outline:   xlsxBoolProp{Val: false},
				Shadow:    xlsxBoolProp{Val: false},
				Condense:  xlsxBoolProp{Val: false},
				Extend:    xlsxBoolProp{Val: false},
				Color:     &xlsxColor{RGB: "DEADBEEF"},
				Sz:        &xlsxFloatVal{Val: 12.345},
				U:         &xlsxVal{Val: string(RichTextUnderlineDouble)},
				VertAlign: &xlsxVal{Val: string(RichTextVertAlignSuperscript)},
				Scheme:    nil,
			},
			T: xlsxT{Text: "Bold"},
		},
		{
			RPr: &xlsxRunProperties{
				RFont:     nil,
				Charset:   nil,
				Family:    nil,
				B:         xlsxBoolProp{Val: false},
				I:         xlsxBoolProp{Val: true},
				Strike:    xlsxBoolProp{Val: false},
				Outline:   xlsxBoolProp{Val: false},
				Shadow:    xlsxBoolProp{Val: false},
				Condense:  xlsxBoolProp{Val: false},
				Extend:    xlsxBoolProp{Val: false},
				Color:     nil,
				Sz:        nil,
				U:         nil,
				VertAlign: nil,
				Scheme:    nil,
			},
			T: xlsxT{Text: "Italic"},
		},
		{
			RPr: &xlsxRunProperties{
				RFont:     nil,
				Charset:   nil,
				Family:    nil,
				B:         xlsxBoolProp{Val: false},
				I:         xlsxBoolProp{Val: false},
				Strike:    xlsxBoolProp{Val: true},
				Outline:   xlsxBoolProp{Val: false},
				Shadow:    xlsxBoolProp{Val: false},
				Condense:  xlsxBoolProp{Val: false},
				Extend:    xlsxBoolProp{Val: false},
				Color:     nil,
				Sz:        nil,
				U:         nil,
				VertAlign: nil,
				Scheme:    nil,
			},
			T: xlsxT{Text: "Strike"},
		},
		{
			RPr: &xlsxRunProperties{},
			T:   xlsxT{Text: "Empty"},
		},
		{
			RPr: nil,
			T:   xlsxT{Text: "No Font"},
		},
	}

	rtr := xmlToRichText(xmlr)
	c.Assert(rtr, qt.HasLen, 5)

	r := rtr[0]
	c.Assert(r.Font.Name, qt.Equals, "Font")
	c.Assert(r.Font.Size, qt.Equals, 12.345)
	c.Assert(r.Font.Family, qt.Equals, RichTextFontFamilySwiss)
	c.Assert(r.Font.Charset, qt.Equals, RichTextCharsetGreek)
	c.Assert(r.Font.Color.coreColor.RGB, qt.Equals, "DEADBEEF")
	c.Assert(r.Font.Bold, qt.Equals, true)
	c.Assert(r.Font.Italic, qt.Equals, false)
	c.Assert(r.Font.Strike, qt.Equals, false)
	c.Assert(r.Font.VertAlign, qt.Equals, RichTextVertAlignSuperscript)
	c.Assert(r.Font.Underline, qt.Equals, RichTextUnderlineDouble)
	c.Assert(r.Text, qt.Equals, "Bold")

	r = rtr[1]
	c.Assert(r.Font.Name, qt.Equals, "")
	c.Assert(r.Font.Size, qt.Equals, 0.0)
	c.Assert(r.Font.Family, qt.Equals, RichTextFontFamilyUnspecified)
	c.Assert(r.Font.Charset, qt.Equals, RichTextCharsetUnspecified)
	c.Assert(r.Font.Color, qt.IsNil)
	c.Assert(r.Font.Bold, qt.Equals, false)
	c.Assert(r.Font.Italic, qt.Equals, true)
	c.Assert(r.Font.Strike, qt.Equals, false)
	c.Assert(r.Font.VertAlign, qt.Equals, RichTextVertAlign(""))
	c.Assert(r.Font.Underline, qt.Equals, RichTextUnderline(""))
	c.Assert(r.Text, qt.Equals, "Italic")

	r = rtr[2]
	c.Assert(r.Font.Name, qt.Equals, "")
	c.Assert(r.Font.Size, qt.Equals, 0.0)
	c.Assert(r.Font.Family, qt.Equals, RichTextFontFamilyUnspecified)
	c.Assert(r.Font.Charset, qt.Equals, RichTextCharsetUnspecified)
	c.Assert(r.Font.Color, qt.IsNil)
	c.Assert(r.Font.Bold, qt.Equals, false)
	c.Assert(r.Font.Italic, qt.Equals, false)
	c.Assert(r.Font.Strike, qt.Equals, true)
	c.Assert(r.Font.VertAlign, qt.Equals, RichTextVertAlign(""))
	c.Assert(r.Font.Underline, qt.Equals, RichTextUnderline(""))
	c.Assert(r.Text, qt.Equals, "Strike")

	r = rtr[3]
	c.Assert(r.Font.Name, qt.Equals, "")
	c.Assert(r.Font.Size, qt.Equals, 0.0)
	c.Assert(r.Font.Family, qt.Equals, RichTextFontFamilyUnspecified)
	c.Assert(r.Font.Charset, qt.Equals, RichTextCharsetUnspecified)
	c.Assert(r.Font.Color, qt.IsNil)
	c.Assert(r.Font.Bold, qt.Equals, false)
	c.Assert(r.Font.Italic, qt.Equals, false)
	c.Assert(r.Font.Strike, qt.Equals, false)
	c.Assert(r.Font.VertAlign, qt.Equals, RichTextVertAlign(""))
	c.Assert(r.Font.Underline, qt.Equals, RichTextUnderline(""))
	c.Assert(r.Text, qt.Equals, "Empty")

	r = rtr[4]
	c.Assert(r.Font, qt.IsNil)
	c.Assert(r.Text, qt.Equals, "No Font")
}

func TestRichTextToPlainText(t *testing.T) {
	c := qt.New(t)
	rt := []RichTextRun{
		{
			Font: &RichTextFont{
				Bold: true,
			},
			Text: "Bold",
		},
		{
			Font: &RichTextFont{
				Italic: true,
			},
			Text: "Italic",
		},
		{
			Font: &RichTextFont{
				Strike: true,
			},
			Text: "Strike",
		},
	}
	plainText := richTextToPlainText(rt)
	c.Assert(plainText, qt.Equals, "BoldItalicStrike")
}

func TestRichTextToPlainTextEmpty(t *testing.T) {
	c := qt.New(t)
	rt := []RichTextRun{}
	plainText := richTextToPlainText(rt)
	c.Assert(plainText, qt.Equals, "")
}
