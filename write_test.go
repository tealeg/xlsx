package xlsx

import (
	"database/sql"
	"math"
	"testing"
	"time"

	qt "github.com/frankban/quicktest"
)

type testStringerImpl struct {
	Value string
}

func (this testStringerImpl) String() string {
	return this.Value
}

func TestWrite(t *testing.T) {
	c := qt.New(t)

	// Test if we can write a struct to a row
	csRunO(c, "TestWriteStruct", func(c *qt.C, option FileOption) {
		var f *File
		f = NewFile(option)
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
			sql.NullString{String: `Smith`, Valid: true},
			sql.NullBool{Bool: false, Valid: true},
			sql.NullInt64{Int64: 100, Valid: true},
			sql.NullFloat64{Float64: 0.123, Valid: true},
			sql.NullString{String: `What ever`, Valid: false},
			sql.NullBool{Bool: true, Valid: false},
			sql.NullInt64{Int64: 100, Valid: false},
			sql.NullFloat64{Float64: 0.123, Valid: false},
		}
		cnt := row.WriteStruct(&testStruct, -1)
		c.Assert(cnt, qt.Equals, 15)
		c.Assert(row, qt.Not(qt.IsNil))

		var (
			c0, c4, c5, c7, c11, c12, c13, c14 string
			err                                error
			c6                                 float64
		)
		if c0, err = row.GetCell(0).FormattedValue(); err != nil {
			c.Error(err)
		}
		c1, e1 := row.GetCell(1).Int()
		c2, e2 := row.GetCell(2).Float()
		c3 := row.GetCell(3).Bool()
		if c4, err = row.GetCell(4).FormattedValue(); err != nil {
			c.Error(err)
		}
		if c5, err = row.GetCell(5).FormattedValue(); err != nil {
			c.Error(err)
		}
		if c6, err = row.GetCell(6).Float(); err != nil {
			c.Error(err)
		}
		if c7, err = row.GetCell(7).FormattedValue(); err != nil {
			c.Error(err)
		}

		c8 := row.GetCell(8).Bool()
		c9, e9 := row.GetCell(9).Int()
		c10, e10 := row.GetCell(10).Float()

		if c11, err = row.GetCell(11).FormattedValue(); err != nil {
			c.Error(err)
		}
		if c12, err = row.GetCell(12).FormattedValue(); err != nil {
			c.Error(err)
		}
		if c13, err = row.GetCell(13).FormattedValue(); err != nil {
			c.Error(err)
		}
		if c14, err = row.GetCell(14).FormattedValue(); err != nil {
			c.Error(err)
		}

		c.Assert(c0, qt.Equals, "Eric")
		c.Assert(c1, qt.Equals, 20)
		c.Assert(c2, qt.Equals, 3.94)
		c.Assert(c3, qt.Equals, false)
		c.Assert(c4, qt.Equals, "Stringer")
		c.Assert(c5, qt.Equals, "Pointer to Stringer")
		c.Assert(math.Floor(c6), qt.Equals, 25569.0)
		c.Assert(c7, qt.Equals, `Smith`)
		c.Assert(c8, qt.Equals, false)
		c.Assert(c9, qt.Equals, 100)
		c.Assert(c10, qt.Equals, 0.123)
		c.Assert(c11, qt.Equals, ``)
		c.Assert(c12, qt.Equals, ``)
		c.Assert(c13, qt.Equals, ``)
		c.Assert(c14, qt.Equals, ``)

		c.Assert(e1, qt.Equals, nil)
		c.Assert(e2, qt.Equals, nil)
		c.Assert(e9, qt.Equals, nil)
		c.Assert(e10, qt.Equals, nil)

	})

	// Test if we can write a slice to a row
	csRunO(c, "TestWriteSlice", func(c *qt.C, option FileOption) {
		var f *File
		f = NewFile(option)
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
		c.Assert(row0, qt.Not(qt.IsNil))

		if val, err := row0.GetCell(0).FormattedValue(); err != nil {
			c.Error(err)
		} else {
			c.Assert(val, qt.Equals, "Eric")
		}

		s1 := intA{10}
		row1 := sheet.AddRow()
		row1.WriteSlice(&s1, -1)
		c.Assert(row1, qt.Not(qt.IsNil))
		c1, e1 := row1.GetCell(0).Int()
		c.Assert(e1, qt.Equals, nil)
		c.Assert(c1, qt.Equals, 10)

		s2 := floatA{3.94}
		row2 := sheet.AddRow()
		row2.WriteSlice(&s2, -1)
		c.Assert(row2, qt.Not(qt.IsNil))
		c2, e2 := row2.GetCell(0).Float()
		c.Assert(e2, qt.Equals, nil)
		c.Assert(c2, qt.Equals, 3.94)

		s3 := boolA{true}
		row3 := sheet.AddRow()
		row3.WriteSlice(&s3, -1)
		c.Assert(row3, qt.Not(qt.IsNil))
		c3 := row3.GetCell(0).Bool()
		c.Assert(c3, qt.Equals, true)

		s4 := interfaceA{"Eric", 10, 3.94, true, time.Unix(0, 0)}
		row4 := sheet.AddRow()
		row4.WriteSlice(&s4, -1)
		c.Assert(row4, qt.Not(qt.IsNil))
		if val, err := row4.GetCell(0).FormattedValue(); err != nil {
			c.Error(err)
		} else {
			c.Assert(val, qt.Equals, "Eric")
		}
		c41, e41 := row4.GetCell(1).Int()
		c.Assert(e41, qt.Equals, nil)
		c.Assert(c41, qt.Equals, 10)
		c42, e42 := row4.GetCell(2).Float()
		c.Assert(e42, qt.Equals, nil)
		c.Assert(c42, qt.Equals, 3.94)
		c43 := row4.GetCell(3).Bool()
		c.Assert(c43, qt.Equals, true)

		c44, e44 := row4.GetCell(4).Float()
		c.Assert(e44, qt.Equals, nil)
		c.Assert(math.Floor(c44), qt.Equals, 25569.0)

		s5 := stringerA{testStringerImpl{"Stringer"}}
		row5 := sheet.AddRow()
		row5.WriteSlice(&s5, -1)
		c.Assert(row5, qt.Not(qt.IsNil))

		if val, err := row5.GetCell(0).FormattedValue(); err != nil {
			c.Error(err)
		} else {
			c.Assert(val, qt.Equals, "Stringer")
		}

		s6 := stringerPtrA{&testStringerImpl{"Pointer to Stringer"}}
		row6 := sheet.AddRow()
		row6.WriteSlice(&s6, -1)
		c.Assert(row6, qt.Not(qt.IsNil))

		if val, err := row6.GetCell(0).FormattedValue(); err != nil {
			c.Error(err)
		} else {
			c.Assert(val, qt.Equals, "Pointer to Stringer")
		}

		s7 := "expects -1 on non pointer to slice"
		row7 := sheet.AddRow()
		c.Assert(row7, qt.Not(qt.IsNil))
		s7_ret := row7.WriteSlice(s7, -1)
		c.Assert(s7_ret, qt.Equals, -1)
		s7_ret = row7.WriteSlice(&s7, -1)
		c.Assert(s7_ret, qt.Equals, -1)

		s8 := nullStringA{sql.NullString{String: "Smith", Valid: true}, sql.NullString{String: `What ever`, Valid: false}}
		row8 := sheet.AddRow()
		row8.WriteSlice(&s8, -1)
		c.Assert(row8, qt.Not(qt.IsNil))

		if val, err := row8.GetCell(0).FormattedValue(); err != nil {
			c.Error(err)
		} else {
			c.Assert(val, qt.Equals, "Smith")
		}
		// check second cell on empty string ""

		if val2, err := row8.GetCell(1).FormattedValue(); err != nil {
			c.Error(err)
		} else {
			c.Assert(val2, qt.Equals, "")
		}

		s9 := nullBoolA{sql.NullBool{Bool: false, Valid: true}, sql.NullBool{Bool: true, Valid: false}}
		row9 := sheet.AddRow()
		row9.WriteSlice(&s9, -1)
		c.Assert(row9, qt.Not(qt.IsNil))
		c9 := row9.GetCell(0).Bool()
		c9Null := row9.GetCell(1).String()
		c.Assert(c9, qt.Equals, false)
		c.Assert(c9Null, qt.Equals, "")

		s10 := nullIntA{sql.NullInt64{Int64: 100, Valid: true}, sql.NullInt64{Int64: 100, Valid: false}}
		row10 := sheet.AddRow()
		row10.WriteSlice(&s10, -1)
		c.Assert(row10, qt.Not(qt.IsNil))
		c10, e10 := row10.GetCell(0).Int()
		c10Null, e10Null := row10.GetCell(1).FormattedValue()
		c.Assert(e10, qt.Equals, nil)
		c.Assert(c10, qt.Equals, 100)
		c.Assert(e10Null, qt.Equals, nil)
		c.Assert(c10Null, qt.Equals, "")

		s11 := nullFloatA{sql.NullFloat64{Float64: 0.123, Valid: true}, sql.NullFloat64{Float64: 0.123, Valid: false}}
		row11 := sheet.AddRow()
		row11.WriteSlice(&s11, -1)
		c.Assert(row11, qt.Not(qt.IsNil))
		c11, e11 := row11.GetCell(0).Float()
		c11Null, e11Null := row11.GetCell(1).FormattedValue()
		c.Assert(e11, qt.Equals, nil)
		c.Assert(c11, qt.Equals, 0.123)
		c.Assert(e11Null, qt.Equals, nil)
		c.Assert(c11Null, qt.Equals, "")
	})
}
