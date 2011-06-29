package xlsx


import (
	"bytes"
	"testing"
	"xml"
)


func TestMakeSharedStringRefTable(t *testing.T) {
	var sharedstringsXML = bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4"><si><t>Foo</t></si><si><t>Bar</t></si><si><t xml:space="preserve">Baz </t></si><si><t>Quuk</t></si></sst>`)
	sst := new(XLSXSST)
	error := xml.Unmarshal(sharedstringsXML, sst)
	if error != nil {
		t.Error(error.String())
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


func TestResolveSharedString(t *testing.T) {
	var sharedstringsXML = bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4"><si><t>Foo</t></si><si><t>Bar</t></si><si><t xml:space="preserve">Baz </t></si><si><t>Quuk</t></si></sst>`)
	sst := new(XLSXSST)
	error := xml.Unmarshal(sharedstringsXML, sst)
	if error != nil {
		t.Error(error.String())
		return
	}
	reftable := MakeSharedStringRefTable(sst)
	if ResolveSharedString(reftable, 0) != "Foo" {
		t.Error("Expected ResolveSharedString(reftable, 0) == 'Foo'")
	}
}


func TestUnmarshallSharedStrings(t *testing.T) {
	var sharedstringsXML = bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4"><si><t>Foo</t></si><si><t>Bar</t></si><si><t xml:space="preserve">Baz </t></si><si><t>Quuk</t></si></sst>`)
	sst := new(XLSXSST)
	error := xml.Unmarshal(sharedstringsXML, sst)
	if error != nil {
		t.Error(error.String())
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
