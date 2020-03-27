package xlsx

import (
	"bytes"
	"encoding/xml"

	. "gopkg.in/check.v1"
)

type SharedStringsSuite struct {
	SharedStringsXML *bytes.Buffer
}

var _ = Suite(&SharedStringsSuite{})

func (s *SharedStringsSuite) SetUpTest(c *C) {
	s.SharedStringsXML = bytes.NewBufferString(
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
             count="5"
             uniqueCount="5">
          <si>
            <t>Foo</t>
          </si>
          <si>
            <t>Bar</t>
          </si>
          <si>
            <t xml:space="preserve">Baz </t>
          </si>
          <si>
            <t>Quuk</t>
          </si>
          <si>
			<r>
				<t>Normal</t>
			</r>
			<r>
				<rPr>
				</rPr>
				<t>Normal2</t>
			</r>
			<r>
				<rPr>
					<b val="true"/>
					<i val="false"/>
					<strike/>
					<condense val="1"/>
					<extend val="0"/>
				</rPr>
				<t>Bools</t>
			</r>
			<r>
				<rPr>
					<sz val="13.5"/><color theme="1"/><rFont val="FontZ"/><family val="2"/><charset val="128"/><scheme val="minor"/>
				</rPr>
				<t>Font Spec</t>
			</r>
			<r>
				<rPr>
					<u val="single"/>
					<vertAlign val="superscript"/>
				</rPr>
				<t>Misc</t>
			</r>
          </si>
        </sst>`)
}

// Test we can correctly unmarshal an the sharedstrings.xml file into
// an xlsx.xlsxSST struct and it's associated children.
func (s *SharedStringsSuite) TestUnmarshallSharedStrings(c *C) {
	sst := new(xlsxSST)
	err := xml.NewDecoder(s.SharedStringsXML).Decode(sst)
	c.Assert(err, IsNil)
	c.Assert(sst.Count, Equals, 5)
	c.Assert(sst.UniqueCount, Equals, 5)
	c.Assert(sst.SI, HasLen, 5)

	si := sst.SI[0]
	c.Assert(si.T.Text, Equals, "Foo")
	c.Assert(si.R, IsNil)
	si = sst.SI[1]
	c.Assert(si.T.Text, Equals, "Bar")
	c.Assert(si.R, IsNil)
	si = sst.SI[2]
	c.Assert(si.T.Text, Equals, "Baz ")
	c.Assert(si.R, IsNil)
	si = sst.SI[3]
	c.Assert(si.T.Text, Equals, "Quuk")
	c.Assert(si.R, IsNil)
	si = sst.SI[4]
	c.Assert(si.T, IsNil)
	c.Assert(len(si.R), Equals, 5)
	r := si.R[0]
	c.Assert(r.T.Text, Equals, "Normal")
	c.Assert(r.RPr, IsNil)
	r = si.R[1]
	c.Assert(r.T.Text, Equals, "Normal2")
	c.Assert(r.RPr.RFont, IsNil)
	c.Assert(r.RPr.Sz, IsNil)
	c.Assert(r.RPr.Color, IsNil)
	c.Assert(r.RPr.Family, IsNil)
	c.Assert(r.RPr.Charset, IsNil)
	c.Assert(r.RPr.Scheme, IsNil)
	c.Assert(r.RPr.B.Val, Equals, false)
	c.Assert(r.RPr.I.Val, Equals, false)
	c.Assert(r.RPr.Strike.Val, Equals, false)
	c.Assert(r.RPr.Outline.Val, Equals, false)
	c.Assert(r.RPr.Shadow.Val, Equals, false)
	c.Assert(r.RPr.Condense.Val, Equals, false)
	c.Assert(r.RPr.Extend.Val, Equals, false)
	c.Assert(r.RPr.U, IsNil)
	c.Assert(r.RPr.VertAlign, IsNil)
	r = si.R[2]
	c.Assert(r.T.Text, Equals, "Bools")
	c.Assert(r.RPr.RFont, IsNil)
	c.Assert(r.RPr.B.Val, Equals, true)
	c.Assert(r.RPr.I.Val, Equals, false)
	c.Assert(r.RPr.Strike.Val, Equals, true)
	c.Assert(r.RPr.Condense.Val, Equals, true)
	c.Assert(r.RPr.Extend.Val, Equals, false)
	r = si.R[3]
	c.Assert(r.T.Text, Equals, "Font Spec")
	c.Assert(r.RPr.RFont.Val, Equals, "FontZ")
	c.Assert(r.RPr.Sz.Val, Equals, 13.5)
	c.Assert(*r.RPr.Color.Theme, Equals, 1)
	c.Assert(r.RPr.Family.Val, Equals, 2)
	c.Assert(r.RPr.Charset.Val, Equals, 128)
	c.Assert(r.RPr.Scheme.Val, Equals, "minor")
	r = si.R[4]
	c.Assert(r.T.Text, Equals, "Misc")
	c.Assert(r.RPr.U.Val, Equals, "single")
	c.Assert(r.RPr.VertAlign.Val, Equals, "superscript")
}

// TestMarshalSI_T tests that xlsxT is marshaled as it is expected.
func (s *SharedStringsSuite) TestMarshalSI_T(c *C) {
	testMarshalSIT(c, "", "<xlsxSI><t></t></xlsxSI>")
	testMarshalSIT(c, "a b c", "<xlsxSI><t>a b c</t></xlsxSI>")
	testMarshalSIT(c, " abc", "<xlsxSI><t xml:space=\"preserve\"> abc</t></xlsxSI>")
	testMarshalSIT(c, "abc ", "<xlsxSI><t xml:space=\"preserve\">abc </t></xlsxSI>")
	testMarshalSIT(c, "\nabc", "<xlsxSI><t xml:space=\"preserve\">\nabc</t></xlsxSI>")
	testMarshalSIT(c, "abc\n", "<xlsxSI><t xml:space=\"preserve\">abc\n</t></xlsxSI>")
	testMarshalSIT(c, "ab\nc", "<xlsxSI><t xml:space=\"preserve\">ab\nc</t></xlsxSI>")
}

func testMarshalSIT(c *C, t string, expected string) {
	si := xlsxSI{T: &xlsxT{Text: t}}
	bytes, err := xml.Marshal(&si)
	c.Assert(err, IsNil)
	c.Assert(string(bytes), Equals, expected)
}

// TestMarshalSI_R tests that xlsxR is marshaled as it is expected.
func (s *SharedStringsSuite) TestMarshalSI_R(c *C) {
	testMarshalSIR(c, xlsxR{}, "<xlsxSI><r><t></t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "a b c"}}, "<xlsxSI><r><t>a b c</t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: " abc"}}, "<xlsxSI><r><t xml:space=\"preserve\"> abc</t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "abc "}}, "<xlsxSI><r><t xml:space=\"preserve\">abc </t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "\nabc"}}, "<xlsxSI><r><t xml:space=\"preserve\">\nabc</t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "abc\n"}}, "<xlsxSI><r><t xml:space=\"preserve\">abc\n</t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "ab\nc"}}, "<xlsxSI><r><t xml:space=\"preserve\">ab\nc</t></r></xlsxSI>")

	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "a"}, RPr: &xlsxRunProperties{RFont: &xlsxVal{Val: "Times New Roman"}}},
		"<xlsxSI><r><rPr><rFont val=\"Times New Roman\"></rFont></rPr><t>a</t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "a"}, RPr: &xlsxRunProperties{Charset: &xlsxIntVal{Val: 1}}},
		"<xlsxSI><r><rPr><charset val=\"1\"></charset></rPr><t>a</t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "a"}, RPr: &xlsxRunProperties{Family: &xlsxIntVal{Val: 1}}},
		"<xlsxSI><r><rPr><family val=\"1\"></family></rPr><t>a</t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "a"}, RPr: &xlsxRunProperties{B: xlsxBoolProp{Val: true}}},
		"<xlsxSI><r><rPr><b></b></rPr><t>a</t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "a"}, RPr: &xlsxRunProperties{I: xlsxBoolProp{Val: true}}},
		"<xlsxSI><r><rPr><i></i></rPr><t>a</t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "a"}, RPr: &xlsxRunProperties{Strike: xlsxBoolProp{Val: true}}},
		"<xlsxSI><r><rPr><strike></strike></rPr><t>a</t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "a"}, RPr: &xlsxRunProperties{Outline: xlsxBoolProp{Val: true}}},
		"<xlsxSI><r><rPr><outline></outline></rPr><t>a</t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "a"}, RPr: &xlsxRunProperties{Shadow: xlsxBoolProp{Val: true}}},
		"<xlsxSI><r><rPr><shadow></shadow></rPr><t>a</t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "a"}, RPr: &xlsxRunProperties{Condense: xlsxBoolProp{Val: true}}},
		"<xlsxSI><r><rPr><condense></condense></rPr><t>a</t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "a"}, RPr: &xlsxRunProperties{Extend: xlsxBoolProp{Val: true}}},
		"<xlsxSI><r><rPr><extend></extend></rPr><t>a</t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "a"}, RPr: &xlsxRunProperties{Color: &xlsxColor{RGB: "FF123456"}}},
		"<xlsxSI><r><rPr><color rgb=\"FF123456\"></color></rPr><t>a</t></r></xlsxSI>")
	colorIndex := 11
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "a"}, RPr: &xlsxRunProperties{Color: &xlsxColor{Indexed: &colorIndex}}},
		"<xlsxSI><r><rPr><color indexed=\"11\"></color></rPr><t>a</t></r></xlsxSI>")
	colorTheme := 5
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "a"}, RPr: &xlsxRunProperties{Color: &xlsxColor{Theme: &colorTheme}}},
		"<xlsxSI><r><rPr><color theme=\"5\"></color></rPr><t>a</t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "a"}, RPr: &xlsxRunProperties{Color: &xlsxColor{Theme: &colorTheme, Tint: 0.1}}},
		"<xlsxSI><r><rPr><color theme=\"5\" tint=\"0.1\"></color></rPr><t>a</t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "a"}, RPr: &xlsxRunProperties{Sz: &xlsxFloatVal{Val: 12.5}}},
		"<xlsxSI><r><rPr><sz val=\"12.5\"></sz></rPr><t>a</t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "a"}, RPr: &xlsxRunProperties{U: &xlsxVal{Val: "single"}}},
		"<xlsxSI><r><rPr><u val=\"single\"></u></rPr><t>a</t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "a"}, RPr: &xlsxRunProperties{VertAlign: &xlsxVal{Val: "superscript"}}},
		"<xlsxSI><r><rPr><vertAlign val=\"superscript\"></vertAlign></rPr><t>a</t></r></xlsxSI>")
	testMarshalSIR(c, xlsxR{T: xlsxT{Text: "a"}, RPr: &xlsxRunProperties{Scheme: &xlsxVal{Val: "major"}}},
		"<xlsxSI><r><rPr><scheme val=\"major\"></scheme></rPr><t>a</t></r></xlsxSI>")
}

func testMarshalSIR(c *C, r xlsxR, expected string) {
	si := xlsxSI{R: []xlsxR{r}}
	bytes, err := xml.Marshal(&si)
	c.Assert(err, IsNil)
	c.Assert(string(bytes), Equals, expected)
}
