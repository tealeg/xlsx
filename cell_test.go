package xlsx

import (
	"math"
	"path/filepath"
	"testing"
	"time"

	qt "github.com/frankban/quicktest"
)

func TestCell(t *testing.T) {
	c := qt.New(t)
	// Test that we can set and get a Value from a Cell
	c.Run("TestValueSet", func(c *qt.C) {
		// Note, this test is fairly pointless, it serves mostly to
		// reinforce that this functionality is important, and should
		// the mechanics of this all change at some point, to remind
		// us not to lose this.
		cell := Cell{}
		cell.Value = "A string"
		c.Assert(cell.Value, qt.Equals, "A string")
	})

	// Test that GetStyle correctly converts the xlsxStyle.Fonts.
	c.Run("TestGetStyleWithFonts", func(c *qt.C) {
		font := NewFont(10, "Calibra")
		style := NewStyle()
		style.Font = *font

		cell := &Cell{Value: "123", style: style}
		style = cell.GetStyle()
		c.Assert(style, qt.Not(qt.IsNil))
		c.Assert(style.Font.Size, qt.Equals, 10.0)
		c.Assert(style.Font.Name, qt.Equals, "Calibra")
	})

	// Test that SetStyle correctly translates into a xlsxFont element
	csRunO(c, "TestSetStyleWithFonts", func(c *qt.C, option FileOption) {
		file := NewFile(option)
		sheet, _ := file.AddSheet("Test")
		row := sheet.AddRow()
		cell := row.AddCell()
		font := NewFont(12, "Calibra")
		style := NewStyle()
		style.Font = *font
		cell.SetStyle(style)
		style = cell.GetStyle()
		xFont, _, _, _ := style.makeXLSXStyleElements()
		c.Assert(xFont.Sz.Val, qt.Equals, "12")
		c.Assert(xFont.Name.Val, qt.Equals, "Calibra")
	})

	// Test that GetStyle correctly converts the xlsxStyle.Fills.
	c.Run("TestGetStyleWithFills", func(c *qt.C) {
		fill := *NewFill("solid", "FF000000", "00FF0000")
		style := NewStyle()
		style.Fill = fill
		cell := &Cell{Value: "123", style: style}
		style = cell.GetStyle()
		_, xFill, _, _ := style.makeXLSXStyleElements()
		c.Assert(xFill.PatternFill.PatternType, qt.Equals, "solid")
		c.Assert(xFill.PatternFill.BgColor.RGB, qt.Equals, "00FF0000")
		c.Assert(xFill.PatternFill.FgColor.RGB, qt.Equals, "FF000000")
	})

	// Test that SetStyle correctly updates xlsxStyle.Fills.
	csRunO(c, "TestSetStyleWithFills", func(c *qt.C, option FileOption) {
		file := NewFile(option)
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
		c.Assert(xPatternFill.PatternType, qt.Equals, "solid")
		c.Assert(xPatternFill.FgColor.RGB, qt.Equals, "00FF0000")
		c.Assert(xPatternFill.BgColor.RGB, qt.Equals, "FF000000")
	})

	// Test that GetStyle correctly converts the xlsxStyle.Borders.
	c.Run("TestGetStyleWithBorders", func(c *qt.C) {
		border := *NewBorder("thin", "thin", "thin", "thin")
		style := NewStyle()
		style.Border = border
		cell := Cell{Value: "123", style: style}
		style = cell.GetStyle()
		_, _, xBorder, _ := style.makeXLSXStyleElements()
		c.Assert(xBorder.Left.Style, qt.Equals, "thin")
		c.Assert(xBorder.Right.Style, qt.Equals, "thin")
		c.Assert(xBorder.Top.Style, qt.Equals, "thin")
		c.Assert(xBorder.Bottom.Style, qt.Equals, "thin")
	})

	// We can return a string representation of the formatted data
	c.Run("TestSetFloatWithFormat", func(c *qt.C) {
		cell := Cell{}
		cell.SetFloatWithFormat(37947.75334343, "yyyy/mm/dd")
		c.Assert(cell.Value, qt.Equals, "37947.75334343")
		c.Assert(cell.NumFmt, qt.Equals, "yyyy/mm/dd")
		c.Assert(cell.Type(), qt.Equals, CellTypeNumeric)
	})

	c.Run("TestSetFloat", func(c *qt.C) {
		cell := Cell{}
		cell.SetFloat(0)
		c.Assert(cell.Value, qt.Equals, "0")
		cell.SetFloat(0.000005)
		c.Assert(cell.Value, qt.Equals, "0.000005")
		cell.SetFloat(100.0)
		c.Assert(cell.Value, qt.Equals, "100")
		cell.SetFloat(37947.75334343)
		c.Assert(cell.Value, qt.Equals, "37947.75334343")
	})

	c.Run("TestGeneralNumberHandling", func(c *qt.C) {
		// If you go to Excel, make a new file, type 18.99 in a cell, and save, what you will get is a
		// cell where the format is General and the storage type is Number, that contains the value 18.989999999999998.
		// The correct way to format this should be 18.99.
		// 1.1 will get you the same, with a stored value of 1.1000000000000001.
		// Also, numbers greater than 1e11 and less than 1e-9 wil be shown as scientific notation.
		testCases := []struct {
			value                   string
			formattedValueOutput    string
			noScientificValueOutput string
		}{
			{
				value:                   "18.989999999999998",
				formattedValueOutput:    "18.99",
				noScientificValueOutput: "18.99",
			},
			{
				value:                   "1.1000000000000001",
				formattedValueOutput:    "1.1",
				noScientificValueOutput: "1.1",
			},
			{
				value:                   "0.0000000000000001",
				formattedValueOutput:    "1E-16",
				noScientificValueOutput: "0.0000000000000001",
			},
			{
				value:                   "0.000000000000008",
				formattedValueOutput:    "8E-15",
				noScientificValueOutput: "0.000000000000008",
			},
			{
				value:                   "1000000000000000000",
				formattedValueOutput:    "1E+18",
				noScientificValueOutput: "1000000000000000000",
			},
			{
				value:                   "1230000000000000000",
				formattedValueOutput:    "1.23E+18",
				noScientificValueOutput: "1230000000000000000",
			},
			{
				value:                   "12345678",
				formattedValueOutput:    "12345678",
				noScientificValueOutput: "12345678",
			},
			{
				value:                   "0",
				formattedValueOutput:    "0",
				noScientificValueOutput: "0",
			},
			{
				value:                   "-18.989999999999998",
				formattedValueOutput:    "-18.99",
				noScientificValueOutput: "-18.99",
			},
			{
				value:                   "-1.1000000000000001",
				formattedValueOutput:    "-1.1",
				noScientificValueOutput: "-1.1",
			},
			{
				value:                   "-0.0000000000000001",
				formattedValueOutput:    "-1E-16",
				noScientificValueOutput: "-0.0000000000000001",
			},
			{
				value:                   "-0.000000000000008",
				formattedValueOutput:    "-8E-15",
				noScientificValueOutput: "-0.000000000000008",
			},
			{
				value:                   "-1000000000000000000",
				formattedValueOutput:    "-1E+18",
				noScientificValueOutput: "-1000000000000000000",
			},
			{
				value:                   "-1230000000000000000",
				formattedValueOutput:    "-1.23E+18",
				noScientificValueOutput: "-1230000000000000000",
			},
			{
				value:                   "-12345678",
				formattedValueOutput:    "-12345678",
				noScientificValueOutput: "-12345678",
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
			c.Assert(val, qt.Equals, testCase.formattedValueOutput)
			val, err = cell.GeneralNumericWithoutScientific()
			if err != nil {
				c.Fatal(err)
			}
			c.Assert(val, qt.Equals, testCase.noScientificValueOutput)
		}
	})

	// TestCellTypeFormatHandling tests all cell types other than numeric. Numeric cells are tested above since those
	// cells have so many edge cases.
	c.Run("TestCellTypeFormatHandling", func(c *qt.C) {
		testCases := []struct {
			cellType             CellType
			numFmt               string
			value                string
			formattedValueOutput string
			expectError          bool
		}{
			// All of the string cell types, will return only the string format if there is no @ symbol in the format.
			{
				cellType:             CellTypeInline,
				numFmt:               `0;0;0;"Error"`,
				value:                "asdf",
				formattedValueOutput: "Error",
			},
			{
				cellType:             CellTypeString,
				numFmt:               `0;0;0;"Error"`,
				value:                "asdf",
				formattedValueOutput: "Error",
			},
			{
				cellType:             CellTypeStringFormula,
				numFmt:               `0;0;0;"Error"`,
				value:                "asdf",
				formattedValueOutput: "Error",
			},
			// Errors are returned as is regardless of what the format shows
			{
				cellType:             CellTypeError,
				numFmt:               `0;0;0;"Error"`,
				value:                "#NAME?",
				formattedValueOutput: "#NAME?",
			},
			{
				cellType:             CellTypeError,
				numFmt:               `"$"@`,
				value:                "#######",
				formattedValueOutput: "#######",
			},
			// Dates are returned as is regardless of what the format shows
			{
				cellType:             CellTypeDate,
				numFmt:               `"$"@`,
				value:                "2017-10-24T15:29:30+00:00",
				formattedValueOutput: "2017-10-24T15:29:30+00:00",
			},
			// Make sure the format used above would have done something for a string
			{
				cellType:             CellTypeString,
				numFmt:               `"$"@`,
				value:                "#######",
				formattedValueOutput: "$#######",
			},
			// For bool cells, 0 is false, 1 is true, anything else will error
			{
				cellType:             CellTypeBool,
				numFmt:               `"$"@`,
				value:                "1",
				formattedValueOutput: "TRUE",
			},
			{
				cellType:             CellTypeBool,
				numFmt:               `"$"@`,
				value:                "0",
				formattedValueOutput: "FALSE",
			},
			{
				cellType:             CellTypeBool,
				numFmt:               `"$"@`,
				value:                "2",
				expectError:          true,
				formattedValueOutput: "2",
			},
			{
				cellType:             CellTypeBool,
				numFmt:               `"$"@`,
				value:                "2",
				expectError:          true,
				formattedValueOutput: "2",
			},
			// Invalid cell type should cause an error
			{
				cellType:             CellType(7),
				numFmt:               `0`,
				value:                "1.0",
				expectError:          true,
				formattedValueOutput: "1.0",
			},
		}
		for _, testCase := range testCases {
			cell := Cell{
				cellType: testCase.cellType,
				NumFmt:   testCase.numFmt,
				Value:    testCase.value,
			}
			val, err := cell.FormattedValue()
			if err != nil != testCase.expectError {
				c.Fatal(err)
			}
			c.Assert(val, qt.Equals, testCase.formattedValueOutput)
		}
	})

	c.Run("TestIsTime", func(c *qt.C) {
		cell := Cell{}
		isTime := cell.IsTime()
		c.Assert(isTime, qt.Equals, false)
		cell.Value = "43221"
		c.Assert(isTime, qt.Equals, false)
		cell.NumFmt = "d-mmm-yy"
		cell.Value = "43221"
		isTime = cell.IsTime()
		c.Assert(isTime, qt.Equals, true)
	})

	c.Run("TestGetTime", func(c *qt.C) {
		cell := Cell{}
		cell.SetFloat(0)
		date, err := cell.GetTime(false)
		c.Assert(err, qt.Equals, nil)
		c.Assert(date, qt.Equals, time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC))
		cell.SetFloat(39813.0)
		date, err = cell.GetTime(true)
		c.Assert(err, qt.Equals, nil)
		c.Assert(date, qt.Equals, time.Date(2013, 1, 1, 0, 0, 0, 0, time.UTC))
		cell.Value = "d"
		_, err = cell.GetTime(false)
		c.Assert(err, qt.Not(qt.IsNil))
	})

	// FormattedValue returns an error for formatting errors
	c.Run("TestFormattedValueErrorsOnBadFormat", func(c *qt.C) {
		cell := Cell{Value: "Fudge Cake", cellType: CellTypeNumeric}
		cell.NumFmt = "#,##0 ;(#,##0)"
		value, err := cell.FormattedValue()
		c.Assert(value, qt.Equals, "Fudge Cake")
		c.Assert(err, qt.Not(qt.IsNil))
		c.Assert(err.Error(), qt.Equals, "strconv.ParseFloat: parsing \"Fudge Cake\": invalid syntax")
	})

	// We can return a string representation of the formatted data
	c.Run("TestFormattedValue", func(c *qt.C) {
		cell := Cell{Value: "37947.7500001", cellType: CellTypeNumeric}
		negativeCell := Cell{Value: "-37947.7500001", cellType: CellTypeNumeric}
		smallCell := Cell{Value: "0.007", cellType: CellTypeNumeric}
		earlyCell := Cell{Value: "2.1", cellType: CellTypeNumeric}

		fvc := formattedValueChecker{c: c}

		cell.NumFmt = "general"
		fvc.Equals(cell, "37947.7500001")
		negativeCell.NumFmt = "general"
		fvc.Equals(negativeCell, "-37947.7500001")

		// TODO: This test is currently broken.  For a string type cell, I
		// don't think FormattedValue() should be doing a numeric conversion on the value
		// before returning the string.
		cell.NumFmt = "0"
		fvc.Equals(cell, "37948")

		cell.NumFmt = "#,##0" // For the time being we're not doing
		// this comma formatting, so it'll fall back to the related
		// non-comma form.
		fvc.Equals(cell, "37948")

		cell.NumFmt = "#,##0.00;(#,##0.00)"
		fvc.Equals(cell, "37947.75")

		cell.NumFmt = "0.00"
		fvc.Equals(cell, "37947.75")

		cell.NumFmt = "#,##0.00" // For the time being we're not doing
		// this comma formatting, so it'll fall back to the related
		// non-comma form.
		fvc.Equals(cell, "37947.75")

		cell.NumFmt = "#,##0 ;(#,##0)"
		fvc.Equals(cell, "37948")
		negativeCell.NumFmt = "#,##0 ;(#,##0)"
		fvc.Equals(negativeCell, "(37948)")

		cell.NumFmt = "#,##0 ;[red](#,##0)"
		fvc.Equals(cell, "37948")
		negativeCell.NumFmt = "#,##0 ;[red](#,##0)"
		fvc.Equals(negativeCell, "(37948)")

		negativeCell.NumFmt = "#,##0.00;(#,##0.00)"
		fvc.Equals(negativeCell, "(37947.75)")

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
	})

	c.Run("TestTimeToExcelTime", func(c *qt.C) {
		c.Assert(0.0, qt.Equals, TimeToExcelTime(time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC), false))
		c.Assert(-1462.0, qt.Equals, TimeToExcelTime(time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC), true))
		c.Assert(25569.0, qt.Equals, TimeToExcelTime(time.Unix(0, 0), false))
		c.Assert(43269.0, qt.Equals, TimeToExcelTime(time.Date(2018, 6, 18, 0, 0, 0, 0, time.UTC), false))
		c.Assert(401769.0, qt.Equals, TimeToExcelTime(time.Date(3000, 1, 1, 0, 0, 0, 0, time.UTC), false))
		smallDate := time.Date(1899, 12, 30, 0, 0, 0, 1000, time.UTC)
		smallExcelTime := TimeToExcelTime(smallDate, false)

		c.Assert(true, qt.Equals, 0.0 != smallExcelTime)
		roundTrippedDate := TimeFromExcelTime(smallExcelTime, false)
		c.Assert(roundTrippedDate, qt.Equals, smallDate)
	})

	// test setters and getters
	c.Run("TestSetterGetters", func(c *qt.C) {
		cell := Cell{}

		cell.SetString("hello world")
		if val, err := cell.FormattedValue(); err != nil {
			c.Error(err)
		} else {
			c.Assert(val, qt.Equals, "hello world")
		}
		c.Assert(cell.Type(), qt.Equals, CellTypeString)

		cell.SetInt(1024)
		intValue, _ := cell.Int()
		c.Assert(intValue, qt.Equals, 1024)
		c.Assert(cell.NumFmt, qt.Equals, builtInNumFmt[builtInNumFmtIndex_GENERAL])
		c.Assert(cell.Type(), qt.Equals, CellTypeNumeric)

		cell.SetInt64(1024)
		int64Value, _ := cell.Int64()
		c.Assert(int64Value, qt.Equals, int64(1024))
		c.Assert(cell.NumFmt, qt.Equals, builtInNumFmt[builtInNumFmtIndex_GENERAL])
		c.Assert(cell.Type(), qt.Equals, CellTypeNumeric)

		cell.SetFloat(1.024)
		float, _ := cell.Float()
		intValue, _ = cell.Int() // convert
		c.Assert(float, qt.Equals, 1.024)
		c.Assert(intValue, qt.Equals, 1)
		c.Assert(cell.NumFmt, qt.Equals, builtInNumFmt[builtInNumFmtIndex_GENERAL])
		c.Assert(cell.Type(), qt.Equals, CellTypeNumeric)

		cell.SetFormula("10+20")
		c.Assert(cell.Formula(), qt.Equals, "10+20")
		c.Assert(cell.Type(), qt.Equals, CellTypeNumeric)

		cell.SetStringFormula("A1")
		c.Assert(cell.Formula(), qt.Equals, "A1")
		c.Assert(cell.Type(), qt.Equals, CellTypeStringFormula)
	})

	// TestOddInput is a regression test for #101. When the number format
	// was "@" (string), the input below caused a crash in strconv.ParseFloat.
	// The solution was to check if cell.Value was both a CellTypeString and
	// had a NumFmt of "general" or "@" and short-circuit FormattedValue() if so.
	c.Run("TestOddInput", func(c *qt.C) {
		cell := Cell{}
		odd := `[1],[12,"DATE NOT NULL DEFAULT '0000-00-00'"]`
		cell.Value = odd
		cell.NumFmt = "@"
		if val, err := cell.FormattedValue(); err != nil {
			c.Error(err)
		} else {
			c.Assert(val, qt.Equals, odd)
		}
	})

	// TestBool tests basic Bool getting and setting booleans.
	c.Run("TestBool", func(c *qt.C) {
		cell := Cell{}
		cell.SetBool(true)
		c.Assert(cell.Value, qt.Equals, "1")
		c.Assert(cell.Bool(), qt.Equals, true)
		cell.SetBool(false)
		c.Assert(cell.Value, qt.Equals, "0")
		c.Assert(cell.Bool(), qt.Equals, false)
	})

	// TestStringBool tests calling Bool on a non CellTypeBool value.
	c.Run("TestStringBool", func(c *qt.C) {
		cell := Cell{}
		cell.SetInt(0)
		c.Assert(cell.Bool(), qt.Equals, false)
		cell.SetInt(1)
		c.Assert(cell.Bool(), qt.Equals, true)
		cell.SetString("")
		c.Assert(cell.Bool(), qt.Equals, false)
		cell.SetString("0")
		c.Assert(cell.Bool(), qt.Equals, true)
	})

	// TestSetValue tests whether SetValue handle properly for different type values.
	c.Run("TestSetValue", func(c *qt.C) {
		cell := Cell{}

		// int
		for _, i := range []interface{}{1, int8(1), int16(1), int32(1), int64(1)} {
			cell.SetValue(i)
			val, err := cell.Int64()
			c.Assert(err, qt.IsNil)
			c.Assert(val, qt.Equals, int64(1))
		}

		// float
		for _, i := range []interface{}{1.11, float32(1.11), float64(1.11)} {
			cell.SetValue(i)
			val, err := cell.Float()
			c.Assert(err, qt.IsNil)
			c.Assert(val, qt.Equals, 1.11)
		}
		// In the naive implementation using go fmt "%v", this test would fail and the cell.Value would be "1e-06"
		for _, i := range []interface{}{0.000001, float32(0.000001), float64(0.000001)} {
			cell.SetValue(i)
			c.Assert(cell.Value, qt.Equals, "0.000001")
			val, err := cell.Float()
			c.Assert(err, qt.IsNil)
			c.Assert(val, qt.Equals, 0.000001)
		}

		// time
		cell.SetValue(time.Unix(0, 0))
		val, err := cell.Float()
		c.Assert(err, qt.IsNil)
		c.Assert(math.Floor(val), qt.Equals, 25569.0)

		// string and nil
		for _, i := range []interface{}{nil, "", []byte("")} {
			cell.SetValue(i)
			c.Assert(cell.Value, qt.Equals, "")
		}

		// others
		cell.SetValue([]string{"test"})
		c.Assert(cell.Value, qt.Equals, "[test]")
	})

	c.Run("TestSetDateWithOptions", func(c *qt.C) {
		cell := Cell{}

		// time
		cell.SetDate(time.Unix(0, 0))
		val, err := cell.Float()
		c.Assert(err, qt.IsNil)
		c.Assert(math.Floor(val), qt.Equals, 25569.0)

		// our test subject
		date2016UTC := time.Date(2016, 1, 1, 12, 0, 0, 0, time.UTC)

		// test ny timezone
		nyTZ, err := time.LoadLocation("America/New_York")
		c.Assert(err, qt.IsNil)
		cell.SetDateWithOptions(date2016UTC, DateTimeOptions{
			ExcelTimeFormat: "test_format1",
			Location:        nyTZ,
		})
		val, err = cell.Float()
		c.Assert(err, qt.IsNil)
		c.Assert(val, qt.Equals, TimeToExcelTime(time.Date(2016, 1, 1, 7, 0, 0, 0, time.UTC), false))

		// test jp timezone
		jpTZ, err := time.LoadLocation("Asia/Tokyo")
		c.Assert(err, qt.IsNil)
		cell.SetDateWithOptions(date2016UTC, DateTimeOptions{
			ExcelTimeFormat: "test_format2",
			Location:        jpTZ,
		})
		val, err = cell.Float()
		c.Assert(err, qt.IsNil)
		c.Assert(val, qt.Equals, TimeToExcelTime(time.Date(2016, 1, 1, 21, 0, 0, 0, time.UTC), false))
	})

	c.Run("TestIsTimeFormat", func(c *qt.C) {
		c.Assert(isTimeFormat("yy"), qt.Equals, true)
		c.Assert(isTimeFormat("hh"), qt.Equals, true)
		c.Assert(isTimeFormat("h"), qt.Equals, true)
		c.Assert(isTimeFormat("am/pm"), qt.Equals, true)
		c.Assert(isTimeFormat("AM/PM"), qt.Equals, true)
		c.Assert(isTimeFormat("A/P"), qt.Equals, true)
		c.Assert(isTimeFormat("a/p"), qt.Equals, true)
		c.Assert(isTimeFormat("ss"), qt.Equals, true)
		c.Assert(isTimeFormat("mm"), qt.Equals, true)
		c.Assert(isTimeFormat(":"), qt.Equals, false)
		c.Assert(isTimeFormat("z"), qt.Equals, false)
	})

	c.Run("TestIs12HourtTime", func(c *qt.C) {
		c.Assert(is12HourTime("am/pm"), qt.Equals, true)
		c.Assert(is12HourTime("AM/PM"), qt.Equals, true)
		c.Assert(is12HourTime("a/p"), qt.Equals, true)
		c.Assert(is12HourTime("A/P"), qt.Equals, true)
		c.Assert(is12HourTime("x"), qt.Equals, false)
	})

	c.Run("TestFallbackTo", func(c *qt.C) {
		testCases := []struct {
			cellType       *CellType
			cellData       string
			fallback       CellType
			expectedReturn CellType
		}{
			{
				cellType:       CellTypeNumeric.Ptr(),
				cellData:       `string`,
				fallback:       CellTypeString,
				expectedReturn: CellTypeString,
			},
			{
				cellType:       nil,
				cellData:       `string`,
				fallback:       CellTypeNumeric,
				expectedReturn: CellTypeNumeric,
			},
			{
				cellType:       CellTypeNumeric.Ptr(),
				cellData:       `300.24`,
				fallback:       CellTypeString,
				expectedReturn: CellTypeNumeric,
			},
			{
				cellType:       CellTypeNumeric.Ptr(),
				cellData:       `300`,
				fallback:       CellTypeString,
				expectedReturn: CellTypeNumeric,
			},
		}
		for _, testCase := range testCases {
			c.Assert(testCase.cellType.fallbackTo(testCase.cellData, testCase.fallback), qt.Equals, testCase.expectedReturn)
		}
	})

	// Test that GetCoordinates returns accurate numbers..
	csRunO(c, "GetCoordinates", func(c *qt.C, option FileOption) {
		file := NewFile(option)
		sheet, _ := file.AddSheet("Test")
		row := sheet.AddRow()
		cell := row.AddCell()
		x, y := cell.GetCoordinates()
		c.Assert(x, qt.Equals, 0)
		c.Assert(y, qt.Equals, 0)
		cell = row.AddCell()
		x, y = cell.GetCoordinates()
		c.Assert(x, qt.Equals, 1)
		c.Assert(y, qt.Equals, 0)
		row = sheet.AddRow()
		cell = row.AddCell()
		x, y = cell.GetCoordinates()
		c.Assert(x, qt.Equals, 0)
		c.Assert(y, qt.Equals, 1)
	})

}

// formattedValueChecker removes all the boilerplate for testing Cell.FormattedValue
// after its change from returning one value (a string) to two values (string, error)
// This allows all the old one-line asserts in the test to continue to be one
// line, instead of multi-line with error checking.
type formattedValueChecker struct {
	c *qt.C
}

func (fvc *formattedValueChecker) Equals(cell Cell, expected string) {
	val, err := cell.FormattedValue()
	if err != nil {
		fvc.c.Error(err)
	}
	fvc.c.Assert(val, qt.Equals, expected)
}

func cellsFormattedValueEquals(t *testing.T, cell *Cell, expected string) {
	val, err := cell.FormattedValue()
	if err != nil {
		t.Error(err)
	}
	if val != expected {
		t.Errorf("Expected cell.FormattedValue() to be %v, got %v", expected, val)
	}
}

func TestCellMerge(t *testing.T) {
	c := qt.New(t)
	csRunO(c, "MergeAndSave", func(c *qt.C, option FileOption) {
		// This test exposed issue #559 with the custom XML writer for xlsxWorksheet
		f := NewFile(option)
		sht, err := f.AddSheet("sheet1")
		if err != nil {
			t.Fatal(err)
		}
		row := sht.AddRow()
		cell := row.AddCell()
		cell.Value = "test"
		cell.Merge(1, 0)
		path := filepath.Join(c.Mkdir(), "merged.xlsx")
		err = f.Save(path)
		c.Assert(err, qt.Equals, nil)
	})
}
