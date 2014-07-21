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
             count="4"
             uniqueCount="4">
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
        </sst>`)
}

// Test we can correctly convert a xlsxSST into a reference table
// using xlsx.MakeSharedStringRefTable().
func (s *SharedStringsSuite) TestMakeSharedStringRefTable(c *C) {
	sst := new(xlsxSST)
	err := xml.NewDecoder(s.SharedStringsXML).Decode(sst)
	c.Assert(err, IsNil)
	reftable := MakeSharedStringRefTable(sst)
	c.Assert(len(reftable), Equals, 4)
	c.Assert(reftable[0], Equals, "Foo")
	c.Assert(reftable[1], Equals, "Bar")
}

// Test we can correctly resolve a numeric reference in the reference table to a string value using xlsx.ResolveSharedString().
func (s *SharedStringsSuite) TestResolveSharedString(c *C) {
	sst := new(xlsxSST)
	err := xml.NewDecoder(s.SharedStringsXML).Decode(sst)
	c.Assert(err, IsNil)
	reftable := MakeSharedStringRefTable(sst)
	c.Assert(ResolveSharedString(reftable, 0), Equals, "Foo")
}

// Test we can correctly unmarshal an the sharedstrings.xml file into
// an xlsx.xlsxSST struct and it's associated children.
func (s *SharedStringsSuite) TestUnmarshallSharedStrings(c *C) {
	sst := new(xlsxSST)
	err := xml.NewDecoder(s.SharedStringsXML).Decode(sst)
	c.Assert(err, IsNil)
	c.Assert(sst.Count, Equals, "4")
	c.Assert(sst.UniqueCount, Equals, "4")
	c.Assert(sst.SI, HasLen, 4)
	si := sst.SI[0]
	c.Assert(si.T, Equals, "Foo")
}
