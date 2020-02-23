package xlsx

import (
	"bytes"
	"encoding/xml"

	. "gopkg.in/check.v1"
)

type RefTableSuite struct {
	SharedStringsXML *bytes.Buffer
}

var _ = Suite(&RefTableSuite{})

func (s *RefTableSuite) SetUpTest(c *C) {
	s.SharedStringsXML = bytes.NewBufferString(
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
             count="4"
             uniqueCount="4">
          <si>
            <t>Foo</t>
          </si>
          <si>
            <t>Bar</t>
          </si>
          <si>
            <t xml:space="preserve">Baz 
</t>
          </si>
          <si>
            <t>Quuk</t>
          </si>
          <si>
            <r>
              <rPr>
                <sz val="11.5"/>
                <rFont val="Font1"/>
              </rPr>
              <t>Text1</t>
            </r>
            <r>
              <rPr>
                <sz val="12.5"/>
                <rFont val="Font2"/>
              </rPr>
              <t>Text2</t>
            </r>
          </si>
		  </sst>`)
}

// We can add a new string to the RefTable
func (s *RefTableSuite) TestRefTableAddString(c *C) {
	refTable := NewSharedStringRefTable()
	index := refTable.AddString("Foo")
	c.Assert(index, Equals, 0)
	p, r := refTable.ResolveSharedString(0)
	c.Assert(p, Equals, "Foo")
	c.Assert(r, IsNil)
}

func (s *RefTableSuite) TestCreateNewSharedStringRefTable(c *C) {
	refTable := NewSharedStringRefTable()
	refTable.AddString("Foo")
	refTable.AddString("Bar")
	p, r := refTable.ResolveSharedString(0)
	c.Assert(p, Equals, "Foo")
	c.Assert(r, IsNil)
	p, r = refTable.ResolveSharedString(1)
	c.Assert(p, Equals, "Bar")
	c.Assert(r, IsNil)
}

// Test we can correctly convert a xlsxSST into a reference table
// using xlsx.MakeSharedStringRefTable().
func (s *RefTableSuite) TestMakeSharedStringRefTable(c *C) {
	sst := new(xlsxSST)
	err := xml.NewDecoder(s.SharedStringsXML).Decode(sst)
	c.Assert(err, IsNil)
	reftable := MakeSharedStringRefTable(sst)
	c.Assert(reftable.Length(), Equals, 5)
	p, r := reftable.ResolveSharedString(0)
	c.Assert(p, Equals, "Foo")
	c.Assert(r, IsNil)
	p, r = reftable.ResolveSharedString(1)
	c.Assert(p, Equals, "Bar")
	c.Assert(r, IsNil)
	p, r = reftable.ResolveSharedString(2)
	c.Assert(p, Equals, "Baz \n")
	c.Assert(r, IsNil)
	p, r = reftable.ResolveSharedString(3)
	c.Assert(p, Equals, "Quuk")
	c.Assert(r, IsNil)
	p, r = reftable.ResolveSharedString(4)
	c.Assert(p, Equals, "")
	c.Assert(r, HasLen, 2)
	c.Assert(r[0].Font.Size, Equals, 11.5)
	c.Assert(r[0].Font.Name, Equals, "Font1")
	c.Assert(r[1].Font.Size, Equals, 12.5)
	c.Assert(r[1].Font.Name, Equals, "Font2")
}

// Test we can correctly resolve a numeric reference in the reference
// table to a string value using RefTable.ResolveSharedString().
func (s *RefTableSuite) TestResolveSharedString(c *C) {
	sst := new(xlsxSST)
	err := xml.NewDecoder(s.SharedStringsXML).Decode(sst)
	c.Assert(err, IsNil)
	reftable := MakeSharedStringRefTable(sst)
	p, r := reftable.ResolveSharedString(0)
	c.Assert(p, Equals, "Foo")
	c.Assert(r, IsNil)
}

// Test we can correctly create the xlsx.xlsxSST struct from a RefTable
func (s *RefTableSuite) TestMakeXLSXSST(c *C) {
	refTable := NewSharedStringRefTable()
	refTable.AddString("Foo")
	refTable.AddString("Bar")
	refTable.AddRichText([]RichTextRun{
		RichTextRun{
			Font: &RichTextFont{
				Family:  RichTextFontFamilyUnspecified,
				Charset: RichTextCharsetUnspecified,
				Bold:    true,
			},
			Text: "Text1",
		},
		RichTextRun{
			Text: "Text2",
		},
	})
	sst := refTable.makeXLSXSST()
	c.Assert(sst, NotNil)
	c.Assert(sst.Count, Equals, 3)
	c.Assert(sst.UniqueCount, Equals, 3)
	c.Assert(sst.SI, HasLen, 3)
	si := sst.SI[0]
	c.Assert(si.T.Text, Equals, "Foo")
	c.Assert(si.R, IsNil)
	si = sst.SI[2]
	c.Assert(si.T, IsNil)
	c.Assert(si.R, HasLen, 2)
	c.Assert(si.R[0].RPr.B, NotNil)
	c.Assert(si.R[0].T.Text, Equals, "Text1")
	c.Assert(si.R[1].RPr, IsNil)
	c.Assert(si.R[1].T.Text, Equals, "Text2")
}

func (s *RefTableSuite) TestMarshalSST(c *C) {
	refTable := NewSharedStringRefTable()
	refTable.AddString("Foo")
	refTable.AddRichText([]RichTextRun{
		RichTextRun{
			Font: &RichTextFont{
				Family:  RichTextFontFamilyUnspecified,
				Charset: RichTextCharsetUnspecified,
				Bold:    true,
			},
			Text: "Text1",
		},
		RichTextRun{
			Text: "Text2",
		},
	})
	sst := refTable.makeXLSXSST()

	output := bytes.NewBufferString(xml.Header)
	body, err := xml.Marshal(sst)
	c.Assert(err, IsNil)
	c.Assert(body, NotNil)
	_, err = output.Write(body)
	c.Assert(err, IsNil)

	expectedXLSXSST := `<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t>Foo</t></si><si><r><rPr><b></b></rPr><t>Text1</t></r><r><t>Text2</t></r></si></sst>`
	c.Assert(output.String(), Equals, expectedXLSXSST)
}

func (s *RefTableSuite) TestRefTableReadAddString(c *C) {
	refTable := NewSharedStringRefTable()
	refTable.isWrite = false
	index1 := refTable.AddString("Foo")
	index2 := refTable.AddString("Foo")
	c.Assert(index1, Equals, 0)
	c.Assert(index2, Equals, 1)
	p, r := refTable.ResolveSharedString(0)
	c.Assert(p, Equals, "Foo")
	c.Assert(r, IsNil)
	p, r = refTable.ResolveSharedString(1)
	c.Assert(p, Equals, "Foo")
	c.Assert(r, IsNil)
}

func (s *RefTableSuite) TestRefTableWriteAddString(c *C) {
	refTable := NewSharedStringRefTable()
	refTable.isWrite = true
	index1 := refTable.AddString("Foo")
	index2 := refTable.AddString("Foo")
	c.Assert(index1, Equals, 0)
	c.Assert(index2, Equals, 0)
	p, r := refTable.ResolveSharedString(0)
	c.Assert(p, Equals, "Foo")
	c.Assert(r, IsNil)
}

func (s *RefTableSuite) TestRefTableReadAddRichText(c *C) {
	refTable := NewSharedStringRefTable()
	refTable.isWrite = false
	index1 := refTable.AddRichText([]RichTextRun{
		RichTextRun{
			Font: &RichTextFont{
				Family:  RichTextFontFamilyUnspecified,
				Charset: RichTextCharsetUnspecified,
				Bold:    true,
			},
			Text: "Text1",
		},
	})
	index2 := refTable.AddRichText([]RichTextRun{
		RichTextRun{
			Font: &RichTextFont{
				Family:  RichTextFontFamilyUnspecified,
				Charset: RichTextCharsetUnspecified,
				Bold:    true,
			},
			Text: "Text1",
		},
	})

	c.Assert(index1, Equals, 0)
	c.Assert(index2, Equals, 1)
	p, r := refTable.ResolveSharedString(0)
	c.Assert(p, Equals, "")
	c.Assert(r, HasLen, 1)
	c.Assert(r[0].Font.Bold, NotNil)
	c.Assert(r[0].Text, Equals, "Text1")
	p, r = refTable.ResolveSharedString(1)
	c.Assert(p, Equals, "")
	c.Assert(r, HasLen, 1)
	c.Assert(r[0].Font.Bold, NotNil)
	c.Assert(r[0].Text, Equals, "Text1")
}

func (s *RefTableSuite) TestRefTableWriteAddRichText(c *C) {
	refTable := NewSharedStringRefTable()
	refTable.isWrite = true
	index1 := refTable.AddRichText([]RichTextRun{
		RichTextRun{
			Font: &RichTextFont{
				Family:  RichTextFontFamilyUnspecified,
				Charset: RichTextCharsetUnspecified,
				Bold:    true,
			},
			Text: "Text1",
		},
	})
	index2 := refTable.AddRichText([]RichTextRun{
		RichTextRun{
			Font: &RichTextFont{
				Family:  RichTextFontFamilyUnspecified,
				Charset: RichTextCharsetUnspecified,
				Bold:    true,
			},
			Text: "Text1",
		},
	})

	c.Assert(index1, Equals, 0)
	c.Assert(index2, Equals, 0)
	p, r := refTable.ResolveSharedString(0)
	c.Assert(p, Equals, "")
	c.Assert(r, HasLen, 1)
	c.Assert(r[0].Font.Bold, NotNil)
	c.Assert(r[0].Text, Equals, "Text1")
}
