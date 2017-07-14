package xlsx

import (
	"errors"
	"fmt"
	"time"

	. "gopkg.in/check.v1"
)

type ReadSuite struct{}

var _ = Suite(&ReadSuite{})

var (
	errorNoPair         = errors.New("Integer to be unmarshaled is not a pair")
	errorNotEnoughCells = errors.New("Row has not enough cells")
)

type pairUnmarshaler int

func (i *pairUnmarshaler) Unmarshal(row *Row) error {
	if len(row.Cells) == 0 {
		return errorNotEnoughCells
	}
	cellInt, err := row.Cells[0].Int()
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
	if len(r.Cells) < 3 {
		return errorNotEnoughCells
	}
	s.private = r.Cells[0].Bool()
	var err error
	s.normal, err = r.Cells[2].Int()
	if err != nil {
		return err
	}
	currency, err := r.Cells[1].FormattedValue()
	if err != nil {
		return err
	}
	s.custom = fmt.Sprintf("$ %s", currency)
	return nil
}

func (r *RowSuite) TestInterface(c *C) {
	var p pairUnmarshaler
	var s structUnmarshaler
	f := NewFile()
	sheet, _ := f.AddSheet("TestReadTime")
	row := sheet.AddRow()
	values := []interface{}{1, "500", true}
	row.WriteSlice(&values, -1)
	errPair := row.ReadStruct(&p)
	err := row.ReadStruct(&s)
	c.Assert(errPair, Equals, errorNoPair)
	c.Assert(err, Equals, nil)
	var empty pairUnmarshaler
	c.Assert(p, Equals, empty)
	c.Assert(s.normal, Equals, 1)
	c.Assert(s.private, Equals, true)
	c.Assert(s.custom, Equals, "$ 500")
}

func (r *RowSuite) TestTime(c *C) {
	type Timer struct {
		Initial time.Time `xlsx:"0"`
		Final   time.Time `xlsx:"1"`
	}
	initial := time.Date(1990, 12, 30, 10, 30, 30, 0, time.UTC)
	t := Timer{
		Initial: initial,
		Final:   initial.Add(time.Hour * 24),
	}
	f := NewFile()
	sheet, _ := f.AddSheet("TestReadTime")
	row := sheet.AddRow()
	row.AddCell().SetDateTime(t.Initial)
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
	c.Assert(t2.Initial, Equals, t.Initial)
	c.Assert(t2.Final, Equals, t.Final)
}

func (r *RowSuite) TestEmbedStruct(c *C) {
	type Embed struct {
		privateVal bool   `xlsx:"0"`
		IgnoredVal int    `xlsx:"-"`
		VisibleVal string `xlsx:"2"`
	}
	type structTest struct {
		Embed
		FinalVal string `xlsx:"3"`
	}
	f := NewFile()
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
	for _, cell := range row.Cells {
		v := cell.String()
		c.Log(v)
	}
	read := new(structTest)
	err := row.ReadStruct(read)
	if err != nil {
		c.Error(err)
		c.FailNow()
	}
	c.Assert(read.privateVal, Equals, false)
	c.Assert(read.VisibleVal, Equals, v.VisibleVal)
	c.Assert(read.IgnoredVal, Equals, 0)
	c.Assert(read.FinalVal, Equals, v.FinalVal)
}

func (r *RowSuite) TestReadStructPrivateFields(c *C) {
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
	f := NewFile()
	sheet, _ := f.AddSheet("TestRead")
	row := sheet.AddRow()
	row.WriteSlice(&writtenValues, -1)
	for i, cell := range row.Cells {
		str := cell.String()
		c.Log(i, " ", str)
	}
	readStruct := structTest{}
	err := row.ReadStruct(&readStruct)
	if err != nil {
		c.Error(err)
		c.FailNow()
	}
	c.Assert(err, Equals, nil)
	c.Assert(readStruct.IntVal, Equals, val.IntVal)
	c.Assert(readStruct.StringVal, Equals, val.StringVal)
	c.Assert(readStruct.IgnoredVal, Equals, 0)
	c.Assert(readStruct.FloatVal, Equals, val.FloatVal)
	c.Assert(readStruct.BoolVal, Equals, val.BoolVal)
	c.Assert(readStruct.Nested.IgnoredVal, Equals, 0)
	c.Assert(readStruct.Nested.VisibleVal, Equals, "Hello")
	c.Assert(readStruct.Nested.privateVal, Equals, false)
}

func (r *RowSuite) TestReadStruct(c *C) {
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
	f := NewFile()
	sheet, _ := f.AddSheet("TestRead")
	row := sheet.AddRow()
	row.WriteStruct(&structVal, -1)
	for i, cell := range row.Cells {
		str := cell.String()
		c.Log(i, " ", str)
	}
	readStruct := &structTest{}
	err := row.ReadStruct(readStruct)
	c.Log(readStruct)
	c.Log(structVal)
	c.Assert(err, Equals, nil)
	c.Assert(readStruct.IntVal, Equals, structVal.IntVal)
	c.Assert(readStruct.StringVal, Equals, structVal.StringVal)
	c.Assert(readStruct.IgnoredVal, Equals, 0)
	c.Assert(readStruct.FloatVal, Equals, structVal.FloatVal)
	c.Assert(readStruct.BoolVal, Equals, structVal.BoolVal)
}
