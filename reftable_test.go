package xlsx

import (
	"bytes"
	"encoding/xml"
	"testing"

	qt "github.com/frankban/quicktest"
)

type RefTableSuite struct {
	SharedStringsXML *bytes.Buffer
}

var reftabletest_sharedStringsXMLStr = (`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

// We can add a new string to the RefTable
func TestRefTableAddString(t *testing.T) {
	c := qt.New(t)
	refTable := NewSharedStringRefTable()
	index := refTable.AddString("Foo")
	c.Assert(index, qt.Equals, 0)
	p, r := refTable.ResolveSharedString(0)
	c.Assert(p, qt.Equals, "Foo")
	c.Assert(r, qt.IsNil)
}

func TestCreateNewSharedStringRefTable(t *testing.T) {
	c := qt.New(t)
	refTable := NewSharedStringRefTable()
	refTable.AddString("Foo")
	refTable.AddString("Bar")
	p, r := refTable.ResolveSharedString(0)
	c.Assert(p, qt.Equals, "Foo")
	c.Assert(r, qt.IsNil)
	p, r = refTable.ResolveSharedString(1)
	c.Assert(p, qt.Equals, "Bar")
	c.Assert(r, qt.IsNil)
}

// Test we can correctly convert a xlsxSST into a reference table
// using xlsx.MakeSharedStringRefTable().
func TestMakeSharedStringRefTable(t *testing.T) {
	c := qt.New(t)
	sst := new(xlsxSST)
	sharedStringsXML := bytes.NewBufferString(reftabletest_sharedStringsXMLStr)
	err := xml.NewDecoder(sharedStringsXML).Decode(sst)
	c.Assert(err, qt.IsNil)
	reftable := MakeSharedStringRefTable(sst)
	c.Assert(reftable.Length(), qt.Equals, 5)
	p, r := reftable.ResolveSharedString(0)
	c.Assert(p, qt.Equals, "Foo")
	c.Assert(r, qt.IsNil)
	p, r = reftable.ResolveSharedString(1)
	c.Assert(p, qt.Equals, "Bar")
	c.Assert(r, qt.IsNil)
	p, r = reftable.ResolveSharedString(2)
	c.Assert(p, qt.Equals, "Baz \n")
	c.Assert(r, qt.IsNil)
	p, r = reftable.ResolveSharedString(3)
	c.Assert(p, qt.Equals, "Quuk")
	c.Assert(r, qt.IsNil)
	p, r = reftable.ResolveSharedString(4)
	c.Assert(p, qt.Equals, "")
	c.Assert(r, qt.HasLen, 2)
	c.Assert(r[0].Font.Size, qt.Equals, 11.5)
	c.Assert(r[0].Font.Name, qt.Equals, "Font1")
	c.Assert(r[1].Font.Size, qt.Equals, 12.5)
	c.Assert(r[1].Font.Name, qt.Equals, "Font2")
}

// Test we can correctly resolve a numeric reference in the reference
// table to a string value using RefTable.ResolveSharedString().
func TestResolveSharedString(t *testing.T) {
	c := qt.New(t)
	sst := new(xlsxSST)
	sharedStringsXML := bytes.NewBufferString(reftabletest_sharedStringsXMLStr)
	err := xml.NewDecoder(sharedStringsXML).Decode(sst)
	c.Assert(err, qt.IsNil)
	reftable := MakeSharedStringRefTable(sst)
	p, r := reftable.ResolveSharedString(0)
	c.Assert(p, qt.Equals, "Foo")
	c.Assert(r, qt.IsNil)
}

// Test we can correctly create the xlsx.xlsxSST struct from a RefTable
func TestMakeXLSXSST(t *testing.T) {
	c := qt.New(t)
	refTable := NewSharedStringRefTable()
	refTable.AddString("Foo")
	refTable.AddString("Bar")
	refTable.AddRichText([]RichTextRun{
		{
			Font: &RichTextFont{
				Family:  RichTextFontFamilyUnspecified,
				Charset: RichTextCharsetUnspecified,
				Bold:    true,
			},
			Text: "Text1",
		},
		{
			Text: "Text2",
		},
	})
	sst := refTable.makeXLSXSST()
	c.Assert(sst, qt.IsNotNil)
	c.Assert(sst.Count, qt.Equals, 3)
	c.Assert(sst.UniqueCount, qt.Equals, 3)
	c.Assert(sst.SI, qt.HasLen, 3)
	si := sst.SI[0]
	c.Assert(si.T.Text, qt.Equals, "Foo")
	c.Assert(si.R, qt.IsNil)
	si = sst.SI[2]
	c.Assert(si.T, qt.IsNil)
	c.Assert(si.R, qt.HasLen, 2)
	c.Assert(si.R[0].RPr.B, qt.IsNotNil)
	c.Assert(si.R[0].T.Text, qt.Equals, "Text1")
	c.Assert(si.R[1].RPr, qt.IsNil)
	c.Assert(si.R[1].T.Text, qt.Equals, "Text2")
}

func TestMarshalSST(t *testing.T) {
	c := qt.New(t)
	refTable := NewSharedStringRefTable()
	refTable.AddString("Foo")
	refTable.AddRichText([]RichTextRun{
		{
			Font: &RichTextFont{
				Family:  RichTextFontFamilyUnspecified,
				Charset: RichTextCharsetUnspecified,
				Bold:    true,
			},
			Text: "Text1",
		},
		{
			Text: "Text2",
		},
	})
	sst := refTable.makeXLSXSST()

	output := bytes.NewBufferString(xml.Header)
	body, err := xml.Marshal(sst)
	c.Assert(err, qt.IsNil)
	c.Assert(body, qt.IsNotNil)
	_, err = output.Write(body)
	c.Assert(err, qt.IsNil)

	expectedXLSXSST := `<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t>Foo</t></si><si><r><rPr><b></b></rPr><t>Text1</t></r><r><t>Text2</t></r></si></sst>`
	c.Assert(output.String(), qt.Equals, expectedXLSXSST)
}

func TestRefTableReadAddString(t *testing.T) {
	c := qt.New(t)
	refTable := NewSharedStringRefTable()
	refTable.isWrite = false
	index1 := refTable.AddString("Foo")
	index2 := refTable.AddString("Foo")
	c.Assert(index1, qt.Equals, 0)
	c.Assert(index2, qt.Equals, 1)
	p, r := refTable.ResolveSharedString(0)
	c.Assert(p, qt.Equals, "Foo")
	c.Assert(r, qt.IsNil)
	p, r = refTable.ResolveSharedString(1)
	c.Assert(p, qt.Equals, "Foo")
	c.Assert(r, qt.IsNil)
}

func TestRefTableWriteAddString(t *testing.T) {
	c := qt.New(t)
	refTable := NewSharedStringRefTable()
	refTable.isWrite = true
	index1 := refTable.AddString("Foo")
	index2 := refTable.AddString("Foo")
	c.Assert(index1, qt.Equals, 0)
	c.Assert(index2, qt.Equals, 0)
	p, r := refTable.ResolveSharedString(0)
	c.Assert(p, qt.Equals, "Foo")
	c.Assert(r, qt.IsNil)
}

func TestRefTableReadAddRichText(t *testing.T) {
	c := qt.New(t)
	refTable := NewSharedStringRefTable()
	refTable.isWrite = false
	index1 := refTable.AddRichText([]RichTextRun{
		{
			Font: &RichTextFont{
				Family:  RichTextFontFamilyUnspecified,
				Charset: RichTextCharsetUnspecified,
				Bold:    true,
			},
			Text: "Text1",
		},
	})
	index2 := refTable.AddRichText([]RichTextRun{
		{
			Font: &RichTextFont{
				Family:  RichTextFontFamilyUnspecified,
				Charset: RichTextCharsetUnspecified,
				Bold:    true,
			},
			Text: "Text1",
		},
	})

	c.Assert(index1, qt.Equals, 0)
	c.Assert(index2, qt.Equals, 1)
	p, r := refTable.ResolveSharedString(0)
	c.Assert(p, qt.Equals, "")
	c.Assert(r, qt.HasLen, 1)
	c.Assert(r[0].Font.Bold, qt.IsNotNil)
	c.Assert(r[0].Text, qt.Equals, "Text1")
	p, r = refTable.ResolveSharedString(1)
	c.Assert(p, qt.Equals, "")
	c.Assert(r, qt.HasLen, 1)
	c.Assert(r[0].Font.Bold, qt.IsNotNil)
	c.Assert(r[0].Text, qt.Equals, "Text1")
}

func TestRefTableWriteAddRichText(t *testing.T) {
	c := qt.New(t)
	refTable := NewSharedStringRefTable()
	refTable.isWrite = true
	index1 := refTable.AddRichText([]RichTextRun{
		{
			Font: &RichTextFont{
				Family:  RichTextFontFamilyUnspecified,
				Charset: RichTextCharsetUnspecified,
				Bold:    true,
			},
			Text: "Text1",
		},
	})
	index2 := refTable.AddRichText([]RichTextRun{
		{
			Font: &RichTextFont{
				Family:  RichTextFontFamilyUnspecified,
				Charset: RichTextCharsetUnspecified,
				Bold:    true,
			},
			Text: "Text1",
		},
	})

	c.Assert(index1, qt.Equals, 0)
	c.Assert(index2, qt.Equals, 0)
	p, r := refTable.ResolveSharedString(0)
	c.Assert(p, qt.Equals, "")
	c.Assert(r, qt.HasLen, 1)
	c.Assert(r[0].Font.Bold, qt.IsNotNil)
	c.Assert(r[0].Text, qt.Equals, "Text1")
}
