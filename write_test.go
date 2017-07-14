package xlsx

import (
	"math"
	"time"

	. "gopkg.in/check.v1"
)

type WriteSuite struct{}

var _ = Suite(&WriteSuite{})

type testStringerImpl struct {
	Value string
}

func (this testStringerImpl) String() string {
	return this.Value
}

// Test if we can write a struct to a row
func (r *RowSuite) TestWriteStruct(c *C) {
	var f *File
	f = NewFile()
	sheet, _ := f.AddSheet("Test1")
	row := sheet.AddRow()
	type e struct {
		FirstName   string
		Age         int
		GPA         float64
		LikesPHP    bool
		Stringer    testStringerImpl
		StringerPtr *testStringerImpl
		Time        time.Time
	}
	testStruct := e{
		"Eric",
		20,
		3.94,
		false,
		testStringerImpl{"Stringer"},
		&testStringerImpl{"Pointer to Stringer"},
		time.Unix(0, 0),
	}
	cnt := row.WriteStruct(&testStruct, -1)
	c.Assert(cnt, Equals, 7)
	c.Assert(row, NotNil)

	var (
		c0, c4, c5 string
		err        error
		c6         float64
	)
	if c0, err = row.Cells[0].FormattedValue(); err != nil {
		c.Error(err)
	}
	c1, e1 := row.Cells[1].Int()
	c2, e2 := row.Cells[2].Float()
	c3 := row.Cells[3].Bool()
	if c4, err = row.Cells[4].FormattedValue(); err != nil {
		c.Error(err)
	}
	if c5, err = row.Cells[5].FormattedValue(); err != nil {
		c.Error(err)
	}
	if c6, err = row.Cells[6].Float(); err != nil {
		c.Error(err)
	}

	c.Assert(c0, Equals, "Eric")
	c.Assert(c1, Equals, 20)
	c.Assert(c2, Equals, 3.94)
	c.Assert(c3, Equals, false)
	c.Assert(c4, Equals, "Stringer")
	c.Assert(c5, Equals, "Pointer to Stringer")
	c.Assert(math.Floor(c6), Equals, 25569.0)

	c.Assert(e1, Equals, nil)
	c.Assert(e2, Equals, nil)
}

// Test if we can write a slice to a row
func (r *RowSuite) TestWriteSlice(c *C) {
	var f *File
	f = NewFile()
	sheet, _ := f.AddSheet("Test1")

	type strA []string
	type intA []int
	type floatA []float64
	type boolA []bool
	type interfaceA []interface{}
	type stringerA []testStringerImpl
	type stringerPtrA []*testStringerImpl

	s0 := strA{"Eric"}
	row0 := sheet.AddRow()
	row0.WriteSlice(&s0, -1)
	c.Assert(row0, NotNil)

	if val, err := row0.Cells[0].FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "Eric")
	}

	s1 := intA{10}
	row1 := sheet.AddRow()
	row1.WriteSlice(&s1, -1)
	c.Assert(row1, NotNil)
	c1, e1 := row1.Cells[0].Int()
	c.Assert(e1, Equals, nil)
	c.Assert(c1, Equals, 10)

	s2 := floatA{3.94}
	row2 := sheet.AddRow()
	row2.WriteSlice(&s2, -1)
	c.Assert(row2, NotNil)
	c2, e2 := row2.Cells[0].Float()
	c.Assert(e2, Equals, nil)
	c.Assert(c2, Equals, 3.94)

	s3 := boolA{true}
	row3 := sheet.AddRow()
	row3.WriteSlice(&s3, -1)
	c.Assert(row3, NotNil)
	c3 := row3.Cells[0].Bool()
	c.Assert(c3, Equals, true)

	s4 := interfaceA{"Eric", 10, 3.94, true, time.Unix(0, 0)}
	row4 := sheet.AddRow()
	row4.WriteSlice(&s4, -1)
	c.Assert(row4, NotNil)
	if val, err := row4.Cells[0].FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "Eric")
	}
	c41, e41 := row4.Cells[1].Int()
	c.Assert(e41, Equals, nil)
	c.Assert(c41, Equals, 10)
	c42, e42 := row4.Cells[2].Float()
	c.Assert(e42, Equals, nil)
	c.Assert(c42, Equals, 3.94)
	c43 := row4.Cells[3].Bool()
	c.Assert(c43, Equals, true)

	c44, e44 := row4.Cells[4].Float()
	c.Assert(e44, Equals, nil)
	c.Assert(math.Floor(c44), Equals, 25569.0)

	s5 := stringerA{testStringerImpl{"Stringer"}}
	row5 := sheet.AddRow()
	row5.WriteSlice(&s5, -1)
	c.Assert(row5, NotNil)

	if val, err := row5.Cells[0].FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "Stringer")
	}

	s6 := stringerPtrA{&testStringerImpl{"Pointer to Stringer"}}
	row6 := sheet.AddRow()
	row6.WriteSlice(&s6, -1)
	c.Assert(row6, NotNil)

	if val, err := row6.Cells[0].FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "Pointer to Stringer")
	}

	s7 := "expects -1 on non pointer to slice"
	row7 := sheet.AddRow()
	c.Assert(row7, NotNil)
	s7_ret := row7.WriteSlice(s7, -1)
	c.Assert(s7_ret, Equals, -1)
	s7_ret = row7.WriteSlice(&s7, -1)
	c.Assert(s7_ret, Equals, -1)
	s7_ret = row7.WriteSlice([]string{s7}, -1)
	c.Assert(s7_ret, Equals, -1)
}
