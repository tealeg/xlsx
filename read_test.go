package xlsx

import (
	"errors"
	"fmt"
	"testing"
	"time"

	qt "github.com/frankban/quicktest"
)

var (
	errorNoPair         = errors.New("Integer to be unmarshaled is not a pair")
	errorNotEnoughCells = errors.New("Row has not enough cells")
)

type pairUnmarshaler int

func (i *pairUnmarshaler) Unmarshal(row *Row) error {
	if row.cellCount == 0 {
		return errorNotEnoughCells
	}
	cellInt, err := row.GetCell(0).Int()
	if err != nil {
		return err
	}
	if cellInt%2 != 0 {
		return errorNoPair
	}
	*i = pairUnmarshaler(cellInt)
	return nil
}

type structUnmarshaler struct {
	private bool
	custom  string
	normal  int
}

func (s *structUnmarshaler) Unmarshal(r *Row) error {
	if r.cellCount < 3 {
		return errorNotEnoughCells
	}
	s.private = r.GetCell(0).Bool()
	var err error
	s.normal, err = r.GetCell(2).Int()
	if err != nil {
		return err
	}
	currency, err := r.GetCell(1).FormattedValue()
	if err != nil {
		return err
	}
	s.custom = fmt.Sprintf("$ %s", currency)
	return nil
}

func TestRead(t *testing.T) {
	c := qt.New(t)

	csRunO(c, "TestInterface", func(c *qt.C, option FileOption) {
		var p pairUnmarshaler
		var s structUnmarshaler
		f := NewFile(option)
		sheet, _ := f.AddSheet("TestReadTime")
		row := sheet.AddRow()
		values := []interface{}{1, "500", true}
		row.WriteSlice(&values, -1)
		errPair := row.ReadStruct(&p)
		err := row.ReadStruct(&s)
		c.Assert(errPair, qt.Equals, errorNoPair)
		c.Assert(err, qt.Equals, nil)
		var empty pairUnmarshaler
		c.Assert(p, qt.Equals, empty)
		c.Assert(s.normal, qt.Equals, 1)
		c.Assert(s.private, qt.Equals, true)
		c.Assert(s.custom, qt.Equals, "$ 500")
	})

	csRunO(c, "TestTime", func(c *qt.C, option FileOption) {
		type Timer struct {
			Initial time.Time `xlsx:"0"`
			Final   time.Time `xlsx:"1"`
		}
		initial := time.Date(1990, 12, 30, 10, 30, 30, 0, time.UTC)
		t := Timer{
			Initial: initial,
			Final:   initial.Add(time.Hour * 24),
		}
		f := NewFile(option)
		sheet, _ := f.AddSheet("TestReadTime")
		row := sheet.AddRow()
		cell := row.AddCell()
		cell.SetDateTime(t.Initial)
		ctime2 := row.AddCell()
		ctime2.SetDate(t.Final)
		t2 := Timer{}
		err := row.ReadStruct(&t2)
		if err != nil {
			c.Error(err)
			c.FailNow()
		}
		//removing ns precition
		t2.Initial = t2.Initial.Add(time.Duration(-1 * t2.Initial.Nanosecond()))
		t2.Final = t2.Final.Add(time.Duration(-1 * t2.Final.Nanosecond()))
		c.Assert(t2.Initial, qt.Equals, t.Initial)
		c.Assert(t2.Final, qt.Equals, t.Final)
	})

	csRunO(c, "TestEmbedStruct", func(c *qt.C, option FileOption) {
		type Embed struct {
			privateVal bool   `xlsx:"0"`
			IgnoredVal int    `xlsx:"-"`
			VisibleVal string `xlsx:"2"`
		}
		type structTest struct {
			Embed
			FinalVal string `xlsx:"3"`
		}
		f := NewFile(option)
		sheet, _ := f.AddSheet("TestRead")
		row := sheet.AddRow()
		v := structTest{
			Embed: Embed{
				privateVal: true,
				IgnoredVal: 10,
				VisibleVal: "--This is a test value--",
			},
			FinalVal: "--end of struct",
		}
		values := []string{
			fmt.Sprint(v.privateVal),
			fmt.Sprint(v.IgnoredVal),
			fmt.Sprint(v.VisibleVal),
			fmt.Sprint(v.FinalVal),
		}
		row.WriteSlice(&values, -1)
		read := new(structTest)
		err := row.ReadStruct(read)
		if err != nil {
			c.Error(err)
			c.FailNow()
		}
		c.Assert(read.privateVal, qt.Equals, false)
		c.Assert(read.VisibleVal, qt.Equals, v.VisibleVal)
		c.Assert(read.IgnoredVal, qt.Equals, 0)
		c.Assert(read.FinalVal, qt.Equals, v.FinalVal)
	})

	csRunO(c, "TestReadStructPrivateFields", func(c *qt.C, option FileOption) {
		type nested struct {
			IgnoredVal int    `xlsx:"-"`
			VisibleVal string `xlsx:"6"`
			privateVal bool   `xlsx:"7"`
		}
		type structTest struct {
			IntVal     int16   `xlsx:"0"`
			StringVal  string  `xlsx:"1"`
			FloatVal   float64 `xlsx:"2"`
			IgnoredVal int     `xlsx:"-"`
			BoolVal    bool    `xlsx:"4"`
			Nested     nested
		}
		val := structTest{
			IntVal:     16,
			StringVal:  "heyheyhey :)!",
			FloatVal:   3.14159216,
			IgnoredVal: 7,
			BoolVal:    true,
			Nested: nested{
				privateVal: true,
				IgnoredVal: 90,
				VisibleVal: "Hello",
			},
		}
		writtenValues := []string{
			fmt.Sprint(val.IntVal), val.StringVal, fmt.Sprint(val.FloatVal),
			fmt.Sprint(val.IgnoredVal), fmt.Sprint(val.BoolVal),
			fmt.Sprint(val.Nested.IgnoredVal), val.Nested.VisibleVal,
			fmt.Sprint(val.Nested.privateVal),
		}
		f := NewFile(option)
		sheet, _ := f.AddSheet("TestRead")
		row := sheet.AddRow()
		row.WriteSlice(&writtenValues, -1)
		readStruct := structTest{}
		err := row.ReadStruct(&readStruct)
		if err != nil {
			c.Error(err)
			c.FailNow()
		}
		c.Assert(err, qt.Equals, nil)
		c.Assert(readStruct.IntVal, qt.Equals, val.IntVal)
		c.Assert(readStruct.StringVal, qt.Equals, val.StringVal)
		c.Assert(readStruct.IgnoredVal, qt.Equals, 0)
		c.Assert(readStruct.FloatVal, qt.Equals, val.FloatVal)
		c.Assert(readStruct.BoolVal, qt.Equals, val.BoolVal)
		c.Assert(readStruct.Nested.IgnoredVal, qt.Equals, 0)
		c.Assert(readStruct.Nested.VisibleVal, qt.Equals, "Hello")
		c.Assert(readStruct.Nested.privateVal, qt.Equals, false)
	})

	csRunO(c, "TestReadStruct", func(c *qt.C, option FileOption) {
		type structTest struct {
			IntVal     int8    `xlsx:"0"`
			StringVal  string  `xlsx:"1"`
			FloatVal   float64 `xlsx:"2"`
			IgnoredVal int     `xlsx:"-"`
			BoolVal    bool    `xlsx:"4"`
		}
		structVal := structTest{
			IntVal:     10,
			StringVal:  "heyheyhey :)!",
			FloatVal:   3.14159216,
			IgnoredVal: 7,
			BoolVal:    true,
		}
		f := NewFile(option)
		sheet, _ := f.AddSheet("TestRead")
		row := sheet.AddRow()
		row.WriteStruct(&structVal, -1)

		readStruct := &structTest{}
		err := row.ReadStruct(readStruct)
		c.Assert(err, qt.Equals, nil)
		c.Assert(readStruct.IntVal, qt.Equals, structVal.IntVal)
		c.Assert(readStruct.StringVal, qt.Equals, structVal.StringVal)
		c.Assert(readStruct.IgnoredVal, qt.Equals, 0)
		c.Assert(readStruct.FloatVal, qt.Equals, structVal.FloatVal)
		c.Assert(readStruct.BoolVal, qt.Equals, structVal.BoolVal)
	})

}
