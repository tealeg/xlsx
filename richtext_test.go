package xlsx

import (
	. "gopkg.in/check.v1"
)

type RichTextSuite struct{}

var _ = Suite(&RichTextSuite{})

func (s *RichTextSuite) TestNewRichTextColorFromARGB(c *C) {
	rtColor := NewRichTextColorFromARGB(127, 128, 129, 130)
	c.Assert(rtColor.coreColor.RGB, Equals, "7F808182")
}

func (s *RichTextSuite) TestNewRichTextColorFromThemeColor(c *C) {
	rtColor := NewRichTextColorFromThemeColor(123)
	c.Assert(*rtColor.coreColor.Theme, Equals, 123)
}

func (s *RichTextSuite) TestRichTextRunEquals(c *C) {
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

	c.Assert(r1.Equals(r2), Equals, true)
	c.Assert(r1.Equals(r3), Equals, false)
	c.Assert(r1.Equals(r4), Equals, false)
	c.Assert(r1.Equals(r5), Equals, false)
	c.Assert(r1.Equals(r6), Equals, false)
	c.Assert(r1.Equals(r7), Equals, false)

	c.Assert(r2.Equals(r1), Equals, true)
	c.Assert(r3.Equals(r1), Equals, false)
	c.Assert(r4.Equals(r1), Equals, false)
	c.Assert(r5.Equals(r1), Equals, false)
	c.Assert(r6.Equals(r1), Equals, false)
	c.Assert(r7.Equals(r1), Equals, false)

	c.Assert(r7.Equals(nil), Equals, true)
}

func (s *RichTextSuite) TestRichTextToXml(c *C) {
	rtr := []RichTextRun{
		RichTextRun{
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
		RichTextRun{
			Font: &RichTextFont{
				Family:  RichTextFontFamilyUnspecified,
				Charset: RichTextCharsetUnspecified,
				Italic:  true,
			},
			Text: "Italic",
		},
		RichTextRun{
			Font: &RichTextFont{
				Family:  RichTextFontFamilyUnspecified,
				Charset: RichTextCharsetUnspecified,
				Strike:  true,
			},
			Text: "Strike",
		},
		RichTextRun{
			Font: &RichTextFont{},
			Text: "Empty",
		},
		RichTextRun{
			Text: "No Font",
		},
	}

	xmlr := richTextToXml(rtr)
	c.Assert(xmlr, HasLen, 5)

	r := xmlr[0]
	c.Assert(r.RPr.RFont.Val, Equals, "Font")
	c.Assert(r.RPr.Charset.Val, Equals, int(RichTextCharsetHebrew))
	c.Assert(r.RPr.Family.Val, Equals, int(RichTextFontFamilyScript))
	c.Assert(r.RPr.B.Val, Equals, true)
	c.Assert(r.RPr.I.Val, Equals, false)
	c.Assert(r.RPr.Strike.Val, Equals, false)
	c.Assert(r.RPr.Outline.Val, Equals, false)
	c.Assert(r.RPr.Shadow.Val, Equals, false)
	c.Assert(r.RPr.Condense.Val, Equals, false)
	c.Assert(r.RPr.Extend.Val, Equals, false)
	c.Assert(r.RPr.Color.RGB, Equals, "DEADBEEF")
	c.Assert(r.RPr.Sz.Val, Equals, 12.345)
	c.Assert(r.RPr.U.Val, Equals, string(RichTextUnderlineSingle))
	c.Assert(r.RPr.VertAlign.Val, Equals, string(RichTextVertAlignSuperscript))
	c.Assert(r.RPr.Scheme, IsNil)
	c.Assert(r.T.Text, Equals, "Bold")

	r = xmlr[1]
	c.Assert(r.RPr.RFont, IsNil)
	c.Assert(r.RPr.Charset, IsNil)
	c.Assert(r.RPr.Family, IsNil)
	c.Assert(r.RPr.B.Val, Equals, false)
	c.Assert(r.RPr.I.Val, Equals, true)
	c.Assert(r.RPr.Strike.Val, Equals, false)
	c.Assert(r.RPr.Outline.Val, Equals, false)
	c.Assert(r.RPr.Shadow.Val, Equals, false)
	c.Assert(r.RPr.Condense.Val, Equals, false)
	c.Assert(r.RPr.Extend.Val, Equals, false)
	c.Assert(r.RPr.Color, IsNil)
	c.Assert(r.RPr.Sz, IsNil)
	c.Assert(r.RPr.U, IsNil)
	c.Assert(r.RPr.VertAlign, IsNil)
	c.Assert(r.RPr.Scheme, IsNil)
	c.Assert(r.T.Text, Equals, "Italic")

	r = xmlr[2]
	c.Assert(r.RPr.RFont, IsNil)
	c.Assert(r.RPr.Charset, IsNil)
	c.Assert(r.RPr.Family, IsNil)
	c.Assert(r.RPr.B.Val, Equals, false)
	c.Assert(r.RPr.I.Val, Equals, false)
	c.Assert(r.RPr.Strike.Val, Equals, true)
	c.Assert(r.RPr.Outline.Val, Equals, false)
	c.Assert(r.RPr.Shadow.Val, Equals, false)
	c.Assert(r.RPr.Condense.Val, Equals, false)
	c.Assert(r.RPr.Extend.Val, Equals, false)
	c.Assert(r.RPr.Color, IsNil)
	c.Assert(r.RPr.Sz, IsNil)
	c.Assert(r.RPr.U, IsNil)
	c.Assert(r.RPr.VertAlign, IsNil)
	c.Assert(r.RPr.Scheme, IsNil)
	c.Assert(r.T.Text, Equals, "Strike")

	r = xmlr[3]
	c.Assert(r.RPr.RFont, IsNil)
	c.Assert(r.RPr.Charset.Val, Equals, int(RichTextCharsetANSI))
	c.Assert(r.RPr.Family.Val, Equals, int(RichTextFontFamilyNotApplicable))
	c.Assert(r.RPr.B.Val, Equals, false)
	c.Assert(r.RPr.I.Val, Equals, false)
	c.Assert(r.RPr.Strike.Val, Equals, false)
	c.Assert(r.RPr.Outline.Val, Equals, false)
	c.Assert(r.RPr.Shadow.Val, Equals, false)
	c.Assert(r.RPr.Condense.Val, Equals, false)
	c.Assert(r.RPr.Extend.Val, Equals, false)
	c.Assert(r.RPr.Color, IsNil)
	c.Assert(r.RPr.Sz, IsNil)
	c.Assert(r.RPr.U, IsNil)
	c.Assert(r.RPr.VertAlign, IsNil)
	c.Assert(r.RPr.Scheme, IsNil)
	c.Assert(r.T.Text, Equals, "Empty")

	r = xmlr[4]
	c.Assert(r.RPr, IsNil)
	c.Assert(r.T.Text, Equals, "No Font")
}

func (s *RichTextSuite) TestXmlToRichText(c *C) {
	xmlr := []xlsxR{
		xlsxR{
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
		xlsxR{
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
		xlsxR{
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
		xlsxR{
			RPr: &xlsxRunProperties{},
			T:   xlsxT{Text: "Empty"},
		},
		xlsxR{
			RPr: nil,
			T:   xlsxT{Text: "No Font"},
		},
	}

	rtr := xmlToRichText(xmlr)
	c.Assert(rtr, HasLen, 5)

	r := rtr[0]
	c.Assert(r.Font.Name, Equals, "Font")
	c.Assert(r.Font.Size, Equals, 12.345)
	c.Assert(r.Font.Family, Equals, RichTextFontFamilySwiss)
	c.Assert(r.Font.Charset, Equals, RichTextCharsetGreek)
	c.Assert(r.Font.Color.coreColor.RGB, Equals, "DEADBEEF")
	c.Assert(r.Font.Bold, Equals, true)
	c.Assert(r.Font.Italic, Equals, false)
	c.Assert(r.Font.Strike, Equals, false)
	c.Assert(r.Font.VertAlign, Equals, RichTextVertAlignSuperscript)
	c.Assert(r.Font.Underline, Equals, RichTextUnderlineDouble)
	c.Assert(r.Text, Equals, "Bold")

	r = rtr[1]
	c.Assert(r.Font.Name, Equals, "")
	c.Assert(r.Font.Size, Equals, 0.0)
	c.Assert(r.Font.Family, Equals, RichTextFontFamilyUnspecified)
	c.Assert(r.Font.Charset, Equals, RichTextCharsetUnspecified)
	c.Assert(r.Font.Color, IsNil)
	c.Assert(r.Font.Bold, Equals, false)
	c.Assert(r.Font.Italic, Equals, true)
	c.Assert(r.Font.Strike, Equals, false)
	c.Assert(r.Font.VertAlign, Equals, RichTextVertAlign(""))
	c.Assert(r.Font.Underline, Equals, RichTextUnderline(""))
	c.Assert(r.Text, Equals, "Italic")

	r = rtr[2]
	c.Assert(r.Font.Name, Equals, "")
	c.Assert(r.Font.Size, Equals, 0.0)
	c.Assert(r.Font.Family, Equals, RichTextFontFamilyUnspecified)
	c.Assert(r.Font.Charset, Equals, RichTextCharsetUnspecified)
	c.Assert(r.Font.Color, IsNil)
	c.Assert(r.Font.Bold, Equals, false)
	c.Assert(r.Font.Italic, Equals, false)
	c.Assert(r.Font.Strike, Equals, true)
	c.Assert(r.Font.VertAlign, Equals, RichTextVertAlign(""))
	c.Assert(r.Font.Underline, Equals, RichTextUnderline(""))
	c.Assert(r.Text, Equals, "Strike")

	r = rtr[3]
	c.Assert(r.Font.Name, Equals, "")
	c.Assert(r.Font.Size, Equals, 0.0)
	c.Assert(r.Font.Family, Equals, RichTextFontFamilyUnspecified)
	c.Assert(r.Font.Charset, Equals, RichTextCharsetUnspecified)
	c.Assert(r.Font.Color, IsNil)
	c.Assert(r.Font.Bold, Equals, false)
	c.Assert(r.Font.Italic, Equals, false)
	c.Assert(r.Font.Strike, Equals, false)
	c.Assert(r.Font.VertAlign, Equals, RichTextVertAlign(""))
	c.Assert(r.Font.Underline, Equals, RichTextUnderline(""))
	c.Assert(r.Text, Equals, "Empty")

	r = rtr[4]
	c.Assert(r.Font, IsNil)
	c.Assert(r.Text, Equals, "No Font")
}

func (s *RichTextSuite) TestRichTextToPlainText(c *C) {
	rt := []RichTextRun{
		RichTextRun{
			Font: &RichTextFont{
				Bold: true,
			},
			Text: "Bold",
		},
		RichTextRun{
			Font: &RichTextFont{
				Italic: true,
			},
			Text: "Italic",
		},
		RichTextRun{
			Font: &RichTextFont{
				Strike: true,
			},
			Text: "Strike",
		},
	}
	plainText := richTextToPlainText(rt)
	c.Assert(plainText, Equals, "BoldItalicStrike")
}

func (s *RichTextSuite) TestRichTextToPlainTextEmpty(c *C) {
	rt := []RichTextRun{}
	plainText := richTextToPlainText(rt)
	c.Assert(plainText, Equals, "")
}
