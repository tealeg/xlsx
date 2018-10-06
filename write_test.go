package xlsx

import (
	"database/sql"
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
		FirstName       string
		Age             int
		GPA             float64
		LikesPHP        bool
		Stringer        testStringerImpl
		StringerPtr     *testStringerImpl
		Time            time.Time
		LastName        sql.NullString
		HasPhd          sql.NullBool
		GithubStars     sql.NullInt64
		Raiting         sql.NullFloat64
		NullLastName    sql.NullString
		NullHasPhd      sql.NullBool
		NullGithubStars sql.NullInt64
		NullRaiting     sql.NullFloat64
	}
	testStruct := e{
		"Eric",
		20,
		3.94,
		false,
		testStringerImpl{"Stringer"},
		&testStringerImpl{"Pointer to Stringer"},
		time.Unix(0, 0),
		sql.NullString{`Smith`, true},
		sql.NullBool{false, true},
		sql.NullInt64{100, true},
		sql.NullFloat64{0.123, true},
		sql.NullString{`What ever`, false},
		sql.NullBool{true, false},
		sql.NullInt64{100, false},
		sql.NullFloat64{0.123, false},
	}
	cnt := row.WriteStruct(&testStruct, -1)
	c.Assert(cnt, Equals, 15)
	c.Assert(row, NotNil)

	var (
		c0, c4, c5, c7, c11, c12, c13, c14 string
		err                                error
		c6                                 float64
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
	if c7, err = row.Cells[7].FormattedValue(); err != nil {
		c.Error(err)
	}

	c8 := row.Cells[8].Bool()
	c9, e9 := row.Cells[9].Int()
	c10, e10 := row.Cells[10].Float()

	if c11, err = row.Cells[11].FormattedValue(); err != nil {
		c.Error(err)
	}
	if c12, err = row.Cells[12].FormattedValue(); err != nil {
		c.Error(err)
	}
	if c13, err = row.Cells[13].FormattedValue(); err != nil {
		c.Error(err)
	}
	if c14, err = row.Cells[14].FormattedValue(); err != nil {
		c.Error(err)
	}

	c.Assert(c0, Equals, "Eric")
	c.Assert(c1, Equals, 20)
	c.Assert(c2, Equals, 3.94)
	c.Assert(c3, Equals, false)
	c.Assert(c4, Equals, "Stringer")
	c.Assert(c5, Equals, "Pointer to Stringer")
	c.Assert(math.Floor(c6), Equals, 25569.0)
	c.Assert(c7, Equals, `Smith`)
	c.Assert(c8, Equals, false)
	c.Assert(c9, Equals, 100)
	c.Assert(c10, Equals, 0.123)
	c.Assert(c11, Equals, ``)
	c.Assert(c12, Equals, ``)
	c.Assert(c13, Equals, ``)
	c.Assert(c14, Equals, ``)

	c.Assert(e1, Equals, nil)
	c.Assert(e2, Equals, nil)
	c.Assert(e9, Equals, nil)
	c.Assert(e10, Equals, nil)

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
	type nullStringA []sql.NullString
	type nullBoolA []sql.NullBool
	type nullFloatA []sql.NullFloat64
	type nullIntA []sql.NullInt64

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

	s8 := nullStringA{sql.NullString{"Smith", true}, sql.NullString{`What ever`, false}}
	row8 := sheet.AddRow()
	row8.WriteSlice(&s8, -1)
	c.Assert(row8, NotNil)

	if val, err := row8.Cells[0].FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "Smith")
	}
	// check second cell on empty string ""

	if val2, err := row8.Cells[1].FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val2, Equals, "")
	}

	s9 := nullBoolA{sql.NullBool{false, true}, sql.NullBool{true, false}}
	row9 := sheet.AddRow()
	row9.WriteSlice(&s9, -1)
	c.Assert(row9, NotNil)
	c9 := row9.Cells[0].Bool()
	c9Null := row9.Cells[1].String()
	c.Assert(c9, Equals, false)
	c.Assert(c9Null, Equals, "")

	s10 := nullIntA{sql.NullInt64{100, true}, sql.NullInt64{100, false}}
	row10 := sheet.AddRow()
	row10.WriteSlice(&s10, -1)
	c.Assert(row10, NotNil)
	c10, e10 := row10.Cells[0].Int()
	c10Null, e10Null := row10.Cells[1].FormattedValue()
	c.Assert(e10, Equals, nil)
	c.Assert(c10, Equals, 100)
	c.Assert(e10Null, Equals, nil)
	c.Assert(c10Null, Equals, "")

	s11 := nullFloatA{sql.NullFloat64{0.123, true}, sql.NullFloat64{0.123, false}}
	row11 := sheet.AddRow()
	row11.WriteSlice(&s11, -1)
	c.Assert(row11, NotNil)
	c11, e11 := row11.Cells[0].Float()
	c11Null, e11Null := row11.Cells[1].FormattedValue()
	c.Assert(e11, Equals, nil)
	c.Assert(c11, Equals, 0.123)
	c.Assert(e11Null, Equals, nil)
	c.Assert(c11Null, Equals, "")
}
