package xlsx

import (
	"math"
	"time"

	. "gopkg.in/check.v1"
)

type CellSuite struct{}

var _ = Suite(&CellSuite{})

// Test that we can set and get a Value from a Cell
func (s *CellSuite) TestValueSet(c *C) {
	// Note, this test is fairly pointless, it serves mostly to
	// reinforce that this functionality is important, and should
	// the mechanics of this all change at some point, to remind
	// us not to lose this.
	cell := Cell{}
	cell.Value = "A string"
	c.Assert(cell.Value, Equals, "A string")
}

// Test that GetStyle correctly converts the xlsxStyle.Fonts.
func (s *CellSuite) TestGetStyleWithFonts(c *C) {
	font := NewFont(10, "Calibra")
	style := NewStyle()
	style.Font = *font

	cell := &Cell{Value: "123", style: style}
	style = cell.GetStyle()
	c.Assert(style, NotNil)
	c.Assert(style.Font.Size, Equals, 10)
	c.Assert(style.Font.Name, Equals, "Calibra")
}

// Test that SetStyle correctly translates into a xlsxFont element
func (s *CellSuite) TestSetStyleWithFonts(c *C) {
	file := NewFile()
	sheet, _ := file.AddSheet("Test")
	row := sheet.AddRow()
	cell := row.AddCell()
	font := NewFont(12, "Calibra")
	style := NewStyle()
	style.Font = *font
	cell.SetStyle(style)
	style = cell.GetStyle()
	xFont, _, _, _ := style.makeXLSXStyleElements()
	c.Assert(xFont.Sz.Val, Equals, "12")
	c.Assert(xFont.Name.Val, Equals, "Calibra")
}

// Test that GetStyle correctly converts the xlsxStyle.Fills.
func (s *CellSuite) TestGetStyleWithFills(c *C) {
	fill := *NewFill("solid", "FF000000", "00FF0000")
	style := NewStyle()
	style.Fill = fill
	cell := &Cell{Value: "123", style: style}
	style = cell.GetStyle()
	_, xFill, _, _ := style.makeXLSXStyleElements()
	c.Assert(xFill.PatternFill.PatternType, Equals, "solid")
	c.Assert(xFill.PatternFill.BgColor.RGB, Equals, "00FF0000")
	c.Assert(xFill.PatternFill.FgColor.RGB, Equals, "FF000000")
}

// Test that SetStyle correctly updates xlsxStyle.Fills.
func (s *CellSuite) TestSetStyleWithFills(c *C) {
	file := NewFile()
	sheet, _ := file.AddSheet("Test")
	row := sheet.AddRow()
	cell := row.AddCell()
	fill := NewFill("solid", "00FF0000", "FF000000")
	style := NewStyle()
	style.Fill = *fill
	cell.SetStyle(style)
	style = cell.GetStyle()
	_, xFill, _, _ := style.makeXLSXStyleElements()
	xPatternFill := xFill.PatternFill
	c.Assert(xPatternFill.PatternType, Equals, "solid")
	c.Assert(xPatternFill.FgColor.RGB, Equals, "00FF0000")
	c.Assert(xPatternFill.BgColor.RGB, Equals, "FF000000")
}

// Test that GetStyle correctly converts the xlsxStyle.Borders.
func (s *CellSuite) TestGetStyleWithBorders(c *C) {
	border := *NewBorder("thin", "thin", "thin", "thin")
	style := NewStyle()
	style.Border = border
	cell := Cell{Value: "123", style: style}
	style = cell.GetStyle()
	_, _, xBorder, _ := style.makeXLSXStyleElements()
	c.Assert(xBorder.Left.Style, Equals, "thin")
	c.Assert(xBorder.Right.Style, Equals, "thin")
	c.Assert(xBorder.Top.Style, Equals, "thin")
	c.Assert(xBorder.Bottom.Style, Equals, "thin")
}

// We can return a string representation of the formatted data
func (l *CellSuite) TestSetFloatWithFormat(c *C) {
	cell := Cell{}
	cell.SetFloatWithFormat(37947.75334343, "yyyy/mm/dd")
	c.Assert(cell.Value, Equals, "37947.75334343")
	c.Assert(cell.NumFmt, Equals, "yyyy/mm/dd")
	c.Assert(cell.Type(), Equals, CellTypeNumeric)
}

func (l *CellSuite) TestSetFloat(c *C) {
	cell := Cell{}
	cell.SetFloat(0)
	c.Assert(cell.Value, Equals, "0")
	cell.SetFloat(0.000005)
	c.Assert(cell.Value, Equals, "5e-06")
	cell.SetFloat(100.0)
	c.Assert(cell.Value, Equals, "100")
	cell.SetFloat(37947.75334343)
	c.Assert(cell.Value, Equals, "37947.75334343")
}

func (l *CellSuite) TestGeneralNumberHandling(c *C) {
	// If you go to Excel, make a new file, type 18.99 in a cell, and save, what you will get is a
	// cell where the format is General and the storage type is Number, that contains the value 18.989999999999998.
	// The correct way to format this should be 18.99.
	// 1.1 will get you the same, with a stored value of 1.1000000000000001.
	// Also, numbers greater than 1e11 and less than 1e-9 wil be shown as scientific notation.
	testCases := []struct {
		value  string
		output string
	}{
		{
			value:  "18.989999999999998",
			output: "18.99",
		},
		{
			value:  "1.1000000000000001",
			output: "1.1",
		},
		{
			value:  "0.0000000000000001",
			output: "1E-16",
		},
		{
			value:  "0.000000000000008",
			output: "8E-15",
		},
		{
			value:  "1000000000000000000",
			output: "1E+18",
		},
		{
			value:  "1230000000000000000",
			output: "1.23E+18",
		},
		{
			value:  "12345678",
			output: "12345678",
		},
	}
	for _, testCase := range testCases {
		cell := Cell{
			cellType: CellTypeNumeric,
			NumFmt:   builtInNumFmt[builtInNumFmtIndex_GENERAL],
			Value:    testCase.value,
		}
		val, err := cell.FormattedValue()
		if err != nil {
			c.Fatal(err)
		}
		c.Assert(val, Equals, testCase.output)
	}
}

func (s *CellSuite) TestGetTime(c *C) {
	cell := Cell{}
	cell.SetFloat(0)
	date, err := cell.GetTime(false)
	c.Assert(err, Equals, nil)
	c.Assert(date, Equals, time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC))
	cell.SetFloat(39813.0)
	date, err = cell.GetTime(true)
	c.Assert(err, Equals, nil)
	c.Assert(date, Equals, time.Date(2013, 1, 1, 0, 0, 0, 0, time.UTC))
	cell.Value = "d"
	_, err = cell.GetTime(false)
	c.Assert(err, NotNil)
}

// FormattedValue returns an error for formatting errors
func (l *CellSuite) TestFormattedValueErrorsOnBadFormat(c *C) {
	cell := Cell{Value: "Fudge Cake"}
	cell.NumFmt = "#,##0 ;(#,##0)"
	value, err := cell.FormattedValue()
	c.Assert(value, Equals, "Fudge Cake")
	c.Assert(err, NotNil)
	c.Assert(err.Error(), Equals, "strconv.ParseFloat: parsing \"Fudge Cake\": invalid syntax")
}

// FormattedValue returns a string containing error text for formatting errors
func (l *CellSuite) TestFormattedValueReturnsErrorAsValueForBadFormat(c *C) {
	cell := Cell{Value: "Fudge Cake"}
	cell.NumFmt = "#,##0 ;(#,##0)"
	_, err := cell.FormattedValue()
	c.Assert(err.Error(), Equals, "strconv.ParseFloat: parsing \"Fudge Cake\": invalid syntax")
}

// formattedValueChecker removes all the boilerplate for testing Cell.FormattedValue
// after its change from returning one value (a string) to two values (string, error)
// This allows all the old one-line asserts in the test to continue to be one
// line, instead of multi-line with error checking.
type formattedValueChecker struct {
	c *C
}

func (fvc *formattedValueChecker) Equals(cell Cell, expected string) {
	val, err := cell.FormattedValue()
	if err != nil {
		fvc.c.Error(err)
	}
	fvc.c.Assert(val, Equals, expected)
}

// We can return a string representation of the formatted data
func (l *CellSuite) TestFormattedValue(c *C) {
	// XXX TODO, this test should probably be split down, and made
	// in terms of SafeFormattedValue, as FormattedValue wraps
	// that function now.
	cell := Cell{Value: "37947.7500001"}
	negativeCell := Cell{Value: "-37947.7500001"}
	smallCell := Cell{Value: "0.007"}
	earlyCell := Cell{Value: "2.1"}

	fvc := formattedValueChecker{c: c}

	cell.NumFmt = "general"
	fvc.Equals(cell, "37947.7500001")
	negativeCell.NumFmt = "general"
	fvc.Equals(negativeCell, "-37947.7500001")

	// TODO: This test is currently broken.  For a string type cell, I
	// don't think FormattedValue() should be doing a numeric conversion on the value
	// before returning the string.
	cell.NumFmt = "0"
	fvc.Equals(cell, "37947")

	cell.NumFmt = "#,##0" // For the time being we're not doing
	// this comma formatting, so it'll fall back to the related
	// non-comma form.
	fvc.Equals(cell, "37947")

	cell.NumFmt = "#,##0.00;(#,##0.00)"
	fvc.Equals(cell, "37947.75")

	cell.NumFmt = "0.00"
	fvc.Equals(cell, "37947.75")

	cell.NumFmt = "#,##0.00" // For the time being we're not doing
	// this comma formatting, so it'll fall back to the related
	// non-comma form.
	fvc.Equals(cell, "37947.75")

	cell.NumFmt = "#,##0 ;(#,##0)"
	fvc.Equals(cell, "37947")
	negativeCell.NumFmt = "#,##0 ;(#,##0)"
	fvc.Equals(negativeCell, "(37947)")

	cell.NumFmt = "#,##0 ;[red](#,##0)"
	fvc.Equals(cell, "37947")
	negativeCell.NumFmt = "#,##0 ;[red](#,##0)"
	fvc.Equals(negativeCell, "(37947)")

	negativeCell.NumFmt = "#,##0.00;(#,##0.00)"
	fvc.Equals(negativeCell, "(-37947.75)")

	cell.NumFmt = "0%"
	fvc.Equals(cell, "3794775%")

	cell.NumFmt = "0.00%"
	fvc.Equals(cell, "3794775.00%")

	cell.NumFmt = "0.00e+00"
	fvc.Equals(cell, "3.794775e+04")

	cell.NumFmt = "##0.0e+0" // This is wrong, but we'll use it for now.
	fvc.Equals(cell, "3.794775e+04")

	cell.NumFmt = "mm-dd-yy"
	fvc.Equals(cell, "11-22-03")

	cell.NumFmt = "d-mmm-yy"
	fvc.Equals(cell, "22-Nov-03")
	earlyCell.NumFmt = "d-mmm-yy"
	fvc.Equals(earlyCell, "1-Jan-00")

	cell.NumFmt = "d-mmm"
	fvc.Equals(cell, "22-Nov")
	earlyCell.NumFmt = "d-mmm"
	fvc.Equals(earlyCell, "1-Jan")

	cell.NumFmt = "mmm-yy"
	fvc.Equals(cell, "Nov-03")

	cell.NumFmt = "h:mm am/pm"
	fvc.Equals(cell, "6:00 pm")
	smallCell.NumFmt = "h:mm am/pm"
	fvc.Equals(smallCell, "12:10 am")

	cell.NumFmt = "h:mm:ss am/pm"
	fvc.Equals(cell, "6:00:00 pm")
	cell.NumFmt = "hh:mm:ss"
	fvc.Equals(cell, "18:00:00")
	smallCell.NumFmt = "h:mm:ss am/pm"
	fvc.Equals(smallCell, "12:10:04 am")

	cell.NumFmt = "h:mm"
	fvc.Equals(cell, "18:00")
	smallCell.NumFmt = "h:mm"
	fvc.Equals(smallCell, "00:10")
	smallCell.NumFmt = "hh:mm"
	fvc.Equals(smallCell, "00:10")

	cell.NumFmt = "h:mm:ss"
	fvc.Equals(cell, "18:00:00")
	cell.NumFmt = "hh:mm:ss"
	fvc.Equals(cell, "18:00:00")

	smallCell.NumFmt = "hh:mm:ss"
	fvc.Equals(smallCell, "00:10:04")
	smallCell.NumFmt = "h:mm:ss"
	fvc.Equals(smallCell, "00:10:04")

	cell.NumFmt = "m/d/yy h:mm"
	fvc.Equals(cell, "11/22/03 18:00")
	cell.NumFmt = "m/d/yy hh:mm"
	fvc.Equals(cell, "11/22/03 18:00")
	smallCell.NumFmt = "m/d/yy h:mm"
	fvc.Equals(smallCell, "12/30/99 00:10")
	smallCell.NumFmt = "m/d/yy hh:mm"
	fvc.Equals(smallCell, "12/30/99 00:10")
	earlyCell.NumFmt = "m/d/yy hh:mm"
	fvc.Equals(earlyCell, "1/1/00 02:24")
	earlyCell.NumFmt = "m/d/yy h:mm"
	fvc.Equals(earlyCell, "1/1/00 02:24")

	cell.NumFmt = "mm:ss"
	fvc.Equals(cell, "00:00")
	smallCell.NumFmt = "mm:ss"
	fvc.Equals(smallCell, "10:04")

	cell.NumFmt = "[hh]:mm:ss"
	fvc.Equals(cell, "18:00:00")
	cell.NumFmt = "[h]:mm:ss"
	fvc.Equals(cell, "18:00:00")
	smallCell.NumFmt = "[h]:mm:ss"
	fvc.Equals(smallCell, "10:04")

	const (
		expect1 = "0000.0086"
		expect2 = "1004.8000"
		format  = "mmss.0000"
		tlen    = len(format)
	)

	for i := 0; i < 3; i++ {
		tfmt := format[0 : tlen-i]
		cell.NumFmt = tfmt
		fvc.Equals(cell, expect1[0:tlen-i])
		smallCell.NumFmt = tfmt
		fvc.Equals(smallCell, expect2[0:tlen-i])
	}

	cell.NumFmt = "yyyy\\-mm\\-dd"
	fvc.Equals(cell, "2003\\-11\\-22")

	cell.NumFmt = "dd/mm/yyyy hh:mm:ss"
	fvc.Equals(cell, "22/11/2003 18:00:00")

	cell.NumFmt = "dd/mm/yy"
	fvc.Equals(cell, "22/11/03")
	earlyCell.NumFmt = "dd/mm/yy"
	fvc.Equals(earlyCell, "01/01/00")

	cell.NumFmt = "hh:mm:ss"
	fvc.Equals(cell, "18:00:00")
	smallCell.NumFmt = "hh:mm:ss"
	fvc.Equals(smallCell, "00:10:04")

	cell.NumFmt = "dd/mm/yy\\ hh:mm"
	fvc.Equals(cell, "22/11/03\\ 18:00")

	cell.NumFmt = "yyyy/mm/dd"
	fvc.Equals(cell, "2003/11/22")

	cell.NumFmt = "yy-mm-dd"
	fvc.Equals(cell, "03-11-22")

	cell.NumFmt = "d-mmm-yyyy"
	fvc.Equals(cell, "22-Nov-2003")
	earlyCell.NumFmt = "d-mmm-yyyy"
	fvc.Equals(earlyCell, "1-Jan-1900")

	cell.NumFmt = "m/d/yy"
	fvc.Equals(cell, "11/22/03")
	earlyCell.NumFmt = "m/d/yy"
	fvc.Equals(earlyCell, "1/1/00")

	cell.NumFmt = "m/d/yyyy"
	fvc.Equals(cell, "11/22/2003")
	earlyCell.NumFmt = "m/d/yyyy"
	fvc.Equals(earlyCell, "1/1/1900")

	cell.NumFmt = "dd-mmm-yyyy"
	fvc.Equals(cell, "22-Nov-2003")

	cell.NumFmt = "dd/mm/yyyy"
	fvc.Equals(cell, "22/11/2003")

	cell.NumFmt = "mm/dd/yy hh:mm am/pm"
	fvc.Equals(cell, "11/22/03 06:00 pm")
	cell.NumFmt = "mm/dd/yy h:mm am/pm"
	fvc.Equals(cell, "11/22/03 6:00 pm")

	cell.NumFmt = "mm/dd/yyyy hh:mm:ss"
	fvc.Equals(cell, "11/22/2003 18:00:00")
	smallCell.NumFmt = "mm/dd/yyyy hh:mm:ss"
	fvc.Equals(smallCell, "12/30/1899 00:10:04")

	cell.NumFmt = "yyyy-mm-dd hh:mm:ss"
	fvc.Equals(cell, "2003-11-22 18:00:00")
	smallCell.NumFmt = "yyyy-mm-dd hh:mm:ss"
	fvc.Equals(smallCell, "1899-12-30 00:10:04")

	cell.NumFmt = "mmmm d, yyyy"
	fvc.Equals(cell, "November 22, 2003")
	smallCell.NumFmt = "mmmm d, yyyy"
	fvc.Equals(smallCell, "December 30, 1899")

	cell.NumFmt = "dddd, mmmm dd, yyyy"
	fvc.Equals(cell, "Saturday, November 22, 2003")
	smallCell.NumFmt = "dddd, mmmm dd, yyyy"
	fvc.Equals(smallCell, "Saturday, December 30, 1899")
}

// test setters and getters
func (s *CellSuite) TestSetterGetters(c *C) {
	cell := Cell{}

	cell.SetString("hello world")
	if val, err := cell.FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "hello world")
	}
	c.Assert(cell.Type(), Equals, CellTypeString)

	cell.SetInt(1024)
	intValue, _ := cell.Int()
	c.Assert(intValue, Equals, 1024)
	c.Assert(cell.NumFmt, Equals, builtInNumFmt[builtInNumFmtIndex_GENERAL])
	c.Assert(cell.Type(), Equals, CellTypeGeneral)

	cell.SetInt64(1024)
	int64Value, _ := cell.Int64()
	c.Assert(int64Value, Equals, int64(1024))
	c.Assert(cell.NumFmt, Equals, builtInNumFmt[builtInNumFmtIndex_GENERAL])
	c.Assert(cell.Type(), Equals, CellTypeGeneral)

	cell.SetFloat(1.024)
	float, _ := cell.Float()
	intValue, _ = cell.Int() // convert
	c.Assert(float, Equals, 1.024)
	c.Assert(intValue, Equals, 1)
	c.Assert(cell.NumFmt, Equals, builtInNumFmt[builtInNumFmtIndex_GENERAL])
	c.Assert(cell.Type(), Equals, CellTypeGeneral)

	cell.SetFormula("10+20")
	c.Assert(cell.Formula(), Equals, "10+20")
	c.Assert(cell.Type(), Equals, CellTypeFormula)
}

// TestOddInput is a regression test for #101. When the number format
// was "@" (string), the input below caused a crash in strconv.ParseFloat.
// The solution was to check if cell.Value was both a CellTypeString and
// had a NumFmt of "general" or "@" and short-circuit FormattedValue() if so.
func (s *CellSuite) TestOddInput(c *C) {
	cell := Cell{}
	odd := `[1],[12,"DATE NOT NULL DEFAULT '0000-00-00'"]`
	cell.Value = odd
	cell.NumFmt = "@"
	if val, err := cell.FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, odd)
	}
}

// TestBool tests basic Bool getting and setting booleans.
func (s *CellSuite) TestBool(c *C) {
	cell := Cell{}
	cell.SetBool(true)
	c.Assert(cell.Value, Equals, "1")
	c.Assert(cell.Bool(), Equals, true)
	cell.SetBool(false)
	c.Assert(cell.Value, Equals, "0")
	c.Assert(cell.Bool(), Equals, false)
}

// TestStringBool tests calling Bool on a non CellTypeBool value.
func (s *CellSuite) TestStringBool(c *C) {
	cell := Cell{}
	cell.SetInt(0)
	c.Assert(cell.Bool(), Equals, false)
	cell.SetInt(1)
	c.Assert(cell.Bool(), Equals, true)
	cell.SetString("")
	c.Assert(cell.Bool(), Equals, false)
	cell.SetString("0")
	c.Assert(cell.Bool(), Equals, true)
}

// TestSetValue tests whether SetValue handle properly for different type values.
func (s *CellSuite) TestSetValue(c *C) {
	cell := Cell{}

	// int
	for _, i := range []interface{}{1, int8(1), int16(1), int32(1), int64(1)} {
		cell.SetValue(i)
		val, err := cell.Int64()
		c.Assert(err, IsNil)
		c.Assert(val, Equals, int64(1))
	}

	// float
	for _, i := range []interface{}{1.11, float32(1.11), float64(1.11)} {
		cell.SetValue(i)
		val, err := cell.Float()
		c.Assert(err, IsNil)
		c.Assert(val, Equals, 1.11)
	}

	// time
	cell.SetValue(time.Unix(0, 0))
	val, err := cell.Float()
	c.Assert(err, IsNil)
	c.Assert(math.Floor(val), Equals, 25569.0)

	// string and nil
	for _, i := range []interface{}{nil, "", []byte("")} {
		cell.SetValue(i)
		c.Assert(cell.Value, Equals, "")
	}

	// others
	cell.SetValue([]string{"test"})
	c.Assert(cell.Value, Equals, "[test]")
}

func (s *CellSuite) TestSetDateWithOptions(c *C) {
	cell := Cell{}

	// time
	cell.SetDate(time.Unix(0, 0))
	val, err := cell.Float()
	c.Assert(err, IsNil)
	c.Assert(math.Floor(val), Equals, 25569.0)

	// our test subject
	date2016UTC := time.Date(2016, 1, 1, 12, 0, 0, 0, time.UTC)

	// test ny timezone
	nyTZ, err := time.LoadLocation("America/New_York")
	c.Assert(err, IsNil)
	cell.SetDateWithOptions(date2016UTC, DateTimeOptions{
		ExcelTimeFormat: "test_format1",
		Location:        nyTZ,
	})
	val, err = cell.Float()
	c.Assert(err, IsNil)
	c.Assert(val, Equals, TimeToExcelTime(time.Date(2016, 1, 1, 7, 0, 0, 0, time.UTC)))

	// test jp timezone
	jpTZ, err := time.LoadLocation("Asia/Tokyo")
	c.Assert(err, IsNil)
	cell.SetDateWithOptions(date2016UTC, DateTimeOptions{
		ExcelTimeFormat: "test_format2",
		Location:        jpTZ,
	})
	val, err = cell.Float()
	c.Assert(err, IsNil)
	c.Assert(val, Equals, TimeToExcelTime(time.Date(2016, 1, 1, 21, 0, 0, 0, time.UTC)))
}

func (s *CellSuite) TestIsTimeFormat(c *C) {
	c.Assert(isTimeFormat("yy"), Equals, true)
	c.Assert(isTimeFormat("hh"), Equals, true)
	c.Assert(isTimeFormat("h"), Equals, true)
	c.Assert(isTimeFormat("am/pm"), Equals, true)
	c.Assert(isTimeFormat("AM/PM"), Equals, true)
	c.Assert(isTimeFormat("A/P"), Equals, true)
	c.Assert(isTimeFormat("a/p"), Equals, true)
	c.Assert(isTimeFormat("ss"), Equals, true)
	c.Assert(isTimeFormat("mm"), Equals, true)
	c.Assert(isTimeFormat(":"), Equals, true)
	c.Assert(isTimeFormat("z"), Equals, false)
}

func (s *CellSuite) TestIs12HourtTime(c *C) {
	c.Assert(is12HourTime("am/pm"), Equals, true)
	c.Assert(is12HourTime("AM/PM"), Equals, true)
	c.Assert(is12HourTime("a/p"), Equals, true)
	c.Assert(is12HourTime("A/P"), Equals, true)
	c.Assert(is12HourTime("x"), Equals, false)
}
