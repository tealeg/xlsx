package xlsx

import (
	. "gopkg.in/check.v1"
)

type ColStoreSuite struct{}

var _ = Suite(&ColStoreSuite{})

func (css *ColStoreSuite) TestAddOneNode(c *C) {
	col := &Col{Min: 0, Max: 1}
	cs := ColStore{}
	err := cs.Add(col)
	c.Assert(err, IsNil)
	c.Assert(cs.Root.Col, Equals, col)
}

func (css *ColStoreSuite) TestAddTwoNonOverlappingSequentialNodes(c *C) {
	col1 := &Col{Min: 0, Max: 1}
	col2 := &Col{Min: 2, Max: 4}
	cs := ColStore{}
	err := cs.Add(col1)
	c.Assert(err, IsNil)
	err = cs.Add(col2)
	c.Assert(err, IsNil)
	c.Assert(cs.Root.Col, Equals, col1)
	c.Assert(cs.Root.Next.Col, Equals, col2)
}
