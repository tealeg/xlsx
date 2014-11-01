package xlsx

import (
	. "gopkg.in/check.v1"
)

type StyleSuite struct{}

var _ = Suite(&StyleSuite{})

func (s *StyleSuite) TestNewStyle(c *C) {
	style := NewStyle()
	c.Assert(style, NotNil)
}

type FontSuite struct{}

var _ = Suite(&FontSuite{})

func (s *FontSuite) TestNewFont(c *C) {
	font := NewFont(12, "Verdana")
	c.Assert(font, NotNil)
	c.Assert(font.Name, Equals, "Verdana")
	c.Assert(font.Size, Equals, 12)
}
