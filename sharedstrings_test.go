package xlsx

import (
	"bytes"
	"encoding/xml"
	"testing"
)

// Test we can correctly convert a XLSXSST into a reference table using xlsx.MakeSharedStringRefTable().
func TestMakeSharedStringRefTable(t *testing.T) {
	var sharedstringsXML = bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4"><si><t>Foo</t></si><si><t>Bar</t></si><si><t xml:space="preserve">Baz </t></si><si><t>Quuk</t></si></sst>`)
	sst := new(XLSXSST)
	error := xml.NewDecoder(sharedstringsXML).Decode(sst)
	if error != nil {
		t.Error(error)
		return
	}
	reftable := MakeSharedStringRefTable(sst)
	if len(reftable) == 0 {
		t.Error("Reftable is zero length.")
		return
	}
	if reftable[0] != "Foo" {
		t.Error("RefTable lookup failed, expected reftable[0] == 'Foo'")
	}
	if reftable[1] != "Bar" {
		t.Error("RefTable lookup failed, expected reftable[1] == 'Bar'")
	}

}

// Test we can correctly resolve a numeric reference in the reference table to a string value using xlsx.ResolveSharedString().
func TestResolveSharedString(t *testing.T) {
	var sharedstringsXML = bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4"><si><t>Foo</t></si><si><t>Bar</t></si><si><t xml:space="preserve">Baz </t></si><si><t>Quuk</t></si></sst>`)
	sst := new(XLSXSST)
	error := xml.NewDecoder(sharedstringsXML).Decode(sst)
	if error != nil {
		t.Error(error)
		return
	}
	reftable := MakeSharedStringRefTable(sst)
	if ResolveSharedString(reftable, 0) != "Foo" {
		t.Error("Expected ResolveSharedString(reftable, 0) == 'Foo'")
	}
}

// Test we can correctly unmarshal an the sharedstrings.xml file into
// an xlsx.XLSXSST struct and it's associated children.
func TestUnmarshallSharedStrings(t *testing.T) {
	var sharedstringsXML = bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4"><si><t>Foo</t></si><si><t>Bar</t></si><si><t xml:space="preserve">Baz </t></si><si><t>Quuk</t></si></sst>`)
	sst := new(XLSXSST)
	error := xml.NewDecoder(sharedstringsXML).Decode(sst)
	if error != nil {
		t.Error(error)
		return
	}
	if sst.Count != "4" {
		t.Error(`sst.Count != "4"`)
	}
	if sst.UniqueCount != "4" {
		t.Error(`sst.UniqueCount != 4`)
	}
	if len(sst.SI) == 0 {
		t.Error("Expected 4 sst.SI but found none")
	}
	si := sst.SI[0]
	if si.T.Data != "Foo" {
		t.Error("Expected s.T.Data == 'Foo'")
	}

}
