package xlsx

import (
	. "gopkg.in/check.v1"
)

type SliceReaderSuite struct{}

var _ = Suite(&SliceReaderSuite{})

func (s *SliceReaderSuite) TestFileToSlice(c *C) {
	output, err := FileToSlice("testfile.xlsx")
	c.Assert(err, IsNil)
	fileToSliceCheckOutput(c, output)
}

func (s *SliceReaderSuite) TestFileObjToSlice(c *C) {
	f, err := OpenFile("testfile.xlsx")
	output, err := f.ToSlice()
	c.Assert(err, IsNil)
	fileToSliceCheckOutput(c, output)
}

func fileToSliceCheckOutput(c *C, output [][][]string) {
	c.Assert(len(output), Equals, 3)
	c.Assert(len(output[0]), Equals, 2)
	c.Assert(len(output[0][0]), Equals, 2)
	c.Assert(output[0][0][0], Equals, "Foo")
	c.Assert(output[0][0][1], Equals, "Bar")
	c.Assert(len(output[0][1]), Equals, 2)
	c.Assert(output[0][1][0], Equals, "Baz")
	c.Assert(output[0][1][1], Equals, "Quuk")
	c.Assert(len(output[1]), Equals, 0)
	c.Assert(len(output[2]), Equals, 0)
}
