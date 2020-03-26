package xlsx

import (
	"testing"
	"time"

	qt "github.com/frankban/quicktest"
)

func TestCellFormatCode(t *testing.T) {
	c := qt.New(t)

	c.Run("TestMoreFormattingFeatures", func(c *qt.C) {

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

	c.Run("TestFormatStringSupport", func(c *qt.C) {
		testCases := []struct {
			formatString         string
			value                string
			formattedValueOutput string
			cellType             CellType
			expectError          bool
		}{
			{
				formatString:         `[red]0`,
				value:                "18.989999999999998",
				formattedValueOutput: "19",
				cellType:             CellTypeNumeric,
			},
			{
				formatString:         `[blue]0`,
				value:                "18.989999999999998",
				formattedValueOutput: "19",
				cellType:             CellTypeNumeric,
			},
			{
				formatString:         `[color50]0`,
				value:                "18.989999999999998",
				formattedValueOutput: "19",
				cellType:             CellTypeNumeric,
			},
			{
				formatString:         `[$$-409]0`,
				value:                "18.989999999999998",
				formattedValueOutput: "$19",
				cellType:             CellTypeNumeric,
			},
			{
				formatString:         `[$¥-409]0`,
				value:                "18.989999999999998",
				formattedValueOutput: "¥19",
				cellType:             CellTypeNumeric,
			},
			{
				formatString:         `[$€-409]0`,
				value:                "18.989999999999998",
				formattedValueOutput: "€19",
				cellType:             CellTypeNumeric,
			},
			{
				formatString:         `[$£-409]0`,
				value:                "18.989999999999998",
				formattedValueOutput: "£19",
				cellType:             CellTypeNumeric,
			},
			{
				formatString:         `[$USD-409] 0`,
				value:                "18.989999999999998",
				formattedValueOutput: "USD 19",
				cellType:             CellTypeNumeric,
			},
			{
				formatString:         `0[$USD-409]`,
				value:                "18.989999999999998",
				formattedValueOutput: "19USD",
				cellType:             CellTypeNumeric,
			},
			{
				formatString:         `-[$USD-409]0`,
				value:                "18.989999999999998",
				formattedValueOutput: "-USD19",
				cellType:             CellTypeNumeric,
			},
			{
				formatString:         `\[0`,
				value:                "18.989999999999998",
				formattedValueOutput: "[19",
				cellType:             CellTypeNumeric,
			},
			{
				formatString:         `"["0`,
				value:                "18.989999999999998",
				formattedValueOutput: "[19",
				cellType:             CellTypeNumeric,
			},
			{
				formatString:         "_[0",
				value:                "18.989999999999998",
				formattedValueOutput: "19",
				cellType:             CellTypeNumeric,
			},
			{
				formatString:         `"asdf"0`,
				value:                "18.989999999999998",
				formattedValueOutput: "asdf19",
				cellType:             CellTypeNumeric,
			},
			{
				formatString:         `"$"0`,
				value:                "18.989999999999998",
				formattedValueOutput: "$19",
				cellType:             CellTypeNumeric,
			},
			{
				formatString:         `$0`,
				value:                "18.989999999999998",
				formattedValueOutput: "$19",
				cellType:             CellTypeNumeric,
			},
			{
				formatString:         `%0`, // The percent sign can be anywhere in the format.
				value:                "18.989999999999998",
				formattedValueOutput: "%1899",
				cellType:             CellTypeNumeric,
			},
			{
				formatString:         `$-+/()!^&'~{}<>=: 0 :=><}{~'&^)(/+-$`,
				value:                "18.989999999999998",
				formattedValueOutput: "$-+/()!^&'~{}<>=: 19 :=><}{~'&^)(/+-$",
				cellType:             CellTypeNumeric,
			},
			{
				formatString:         `0;-0;"zero"`,
				value:                "18.989999999999998",
				formattedValueOutput: "19",
				cellType:             CellTypeNumeric,
			},
			{ // 2 formats
				formatString:         `0;(0)`,
				value:                "0",
				formattedValueOutput: "0",
				cellType:             CellTypeNumeric,
			},
			{ // 2 formats
				formatString:         `0;(0)`,
				value:                "4.1",
				formattedValueOutput: "4",
				cellType:             CellTypeNumeric,
			},
			{ // 2 formats
				formatString:         `0;(0)`,
				value:                "-1",
				formattedValueOutput: "(1)",
				cellType:             CellTypeNumeric,
			},
			{ // 2 formats
				formatString:         `0;(0)`,
				value:                "asdf",
				formattedValueOutput: "asdf",
				cellType:             CellTypeNumeric,
				expectError:          true,
			},
			{ // 2 formats
				formatString:         `0;(0)`,
				value:                "asdf",
				formattedValueOutput: "asdf",
				cellType:             CellTypeString,
			},
			{ // 3 formats
				formatString:         `0;(0);"zero"`,
				value:                "59.6",
				formattedValueOutput: "60",
				cellType:             CellTypeNumeric,
			},
			{ // 3 formats
				formatString:         `0;(0);"zero"`,
				value:                "-39",
				formattedValueOutput: "(39)",
				cellType:             CellTypeNumeric,
			},
			{ // 3 formats
				formatString:         `0;(0);"zero"`,
				value:                "0",
				formattedValueOutput: "zero",
				cellType:             CellTypeNumeric,
			},
			{ // 3 formats
				formatString:         `0;(0);"zero"`,
				value:                "asdf",
				formattedValueOutput: "asdf",
				cellType:             CellTypeNumeric,
				expectError:          true,
			},
			{ // 3 formats
				formatString:         `0;(0);"zero"`,
				value:                "asdf",
				formattedValueOutput: "asdf",
				cellType:             CellTypeString,
			},
			{ // 4 formats, also note that the case of the format is maintained. Format codes should not be lower cased.
				formatString:         `0;(0);"zero";"Behold: "@`,
				value:                "asdf",
				formattedValueOutput: "Behold: asdf",
				cellType:             CellTypeString,
			},
			{ // 4 formats
				formatString:         `0;(0);"zero";"Behold": @`,
				value:                "asdf",
				formattedValueOutput: "Behold: asdf",
				cellType:             CellTypeString,
			},
			{ // 4 formats. This format contains an extra
				formatString:         `0;(0);"zero";"Behold; "@`,
				value:                "asdf",
				formattedValueOutput: "Behold; asdf",
				cellType:             CellTypeString,
			},
		}
		for _, testCase := range testCases {
			cell := &Cell{
				cellType: testCase.cellType,
				NumFmt:   testCase.formatString,
				Value:    testCase.value,
			}
			val, err := cell.FormattedValue()
			if err != nil != testCase.expectError {
				c.Fatal(err, testCase)
			}
			if val != testCase.formattedValueOutput {
				c.Fatalf("Expected %v but got %v", testCase.formattedValueOutput, val)
			}
		}
	})

}

func TestIsNumberFormat(t *testing.T) {
	c := qt.New(t)

	c.Assert(isTimeFormat("General"), qt.Equals, false)
	c.Assert(isTimeFormat("0"), qt.Equals, false)
	c.Assert(isTimeFormat("0.00"), qt.Equals, false)
	c.Assert(isTimeFormat("#,##0"), qt.Equals, false)
	c.Assert(isTimeFormat("#,##0.00"), qt.Equals, false)
	c.Assert(isTimeFormat("0%"), qt.Equals, false)
	c.Assert(isTimeFormat("0.00%"), qt.Equals, false)
	c.Assert(isTimeFormat("0.00E+00"), qt.Equals, false)
	c.Assert(isTimeFormat(`mm-dd-yy`), qt.Equals, true)
	c.Assert(isTimeFormat(`d-mmm-yy`), qt.Equals, true)
	c.Assert(isTimeFormat(`d-mmm`), qt.Equals, true)
	c.Assert(isTimeFormat(`mmm-yy`), qt.Equals, true)
	c.Assert(isTimeFormat(`h:mm AM/PM`), qt.Equals, true)
	c.Assert(isTimeFormat(`h:mm:ss AM/PM`), qt.Equals, true)
	c.Assert(isTimeFormat(`h:mm`), qt.Equals, true)
	c.Assert(isTimeFormat(`h:mm:ss`), qt.Equals, true)
	c.Assert(isTimeFormat(`m/d/yy h:mm`), qt.Equals, true)
	c.Assert(isTimeFormat(`mm:ss`), qt.Equals, true)
	c.Assert(isTimeFormat(`[h]:mm:ss`), qt.Equals, true)
	c.Assert(isTimeFormat(`mmss.0`), qt.Equals, true)
	c.Assert(isTimeFormat(`[$-404]e/m/d`), qt.Equals, true)
	c.Assert(isTimeFormat(`yyyy"年"m"月"`), qt.Equals, true)
	c.Assert(isTimeFormat(`[$-404]e"年"`), qt.Equals, true)
	c.Assert(isTimeFormat(`m"月"d"日" m"月"d"日"`), qt.Equals, true)
	c.Assert(isTimeFormat(`[$-404]e"年"m"月"d"日"`), qt.Equals, true)
	c.Assert(isTimeFormat(`m"月"d"日"`), qt.Equals, true)
	c.Assert(isTimeFormat(` m/d/yy`), qt.Equals, true)
	c.Assert(isTimeFormat(`m-d-yy`), qt.Equals, true)
	c.Assert(isTimeFormat(`yyyy"年"m"月"d"日"`), qt.Equals, true)
	c.Assert(isTimeFormat(`yyyy"年"m"月"d"日"`), qt.Equals, true)
	c.Assert(isTimeFormat(`hh"時"mm"分"`), qt.Equals, true)
	c.Assert(isTimeFormat(`h"时"mm"分"`), qt.Equals, true)
	c.Assert(isTimeFormat(`[$-404]e/m/d`), qt.Equals, true)
	c.Assert(isTimeFormat(`yyyy"年"m"月"`), qt.Equals, true)
	c.Assert(isTimeFormat(`[$-404]e"年"m"月"d"日"`), qt.Equals, true)
	c.Assert(isTimeFormat(`m"月"d"日"`), qt.Equals, true)
	c.Assert(isTimeFormat(`[$-404]e"年"m"月"d"日"`), qt.Equals, true)
	c.Assert(isTimeFormat(`m"月"d"日"`), qt.Equals, true)
	c.Assert(isTimeFormat(`m/d/yy`), qt.Equals, true)
	c.Assert(isTimeFormat(`m-d-yy`), qt.Equals, true)
	c.Assert(isTimeFormat(`yyyy"年"m"月"d"日"`), qt.Equals, true)
	c.Assert(isTimeFormat(`yyyy"年"m"月"d"日"`), qt.Equals, true)
	c.Assert(isTimeFormat(`hh"時"mm"分"`), qt.Equals, true)
	c.Assert(isTimeFormat(`h"时"mm"分"`), qt.Equals, true)
	c.Assert(isTimeFormat(`hh"時"mm"分"ss"秒"`), qt.Equals, true)
	c.Assert(isTimeFormat(`h"时"mm"分"ss"秒"`), qt.Equals, true)
	c.Assert(isTimeFormat(`上午/下午 hh"時"mm"分"ss"秒"`), qt.Equals, true)
	c.Assert(isTimeFormat(`上午/下午 h"时"mm"分"ss"秒"`), qt.Equals, true)
	c.Assert(isTimeFormat(`[$-404]e/m/d`), qt.Equals, true)
	c.Assert(isTimeFormat(`yyyy"年"m"月"`), qt.Equals, true)
	c.Assert(isTimeFormat(`[$-404]e/m/d`), qt.Equals, true)
	c.Assert(isTimeFormat(`yyyy"年"m"月"`), qt.Equals, true)
	c.Assert(isTimeFormat(`[$-404]e"年"m"月"d"日"`), qt.Equals, true)
	c.Assert(isTimeFormat(`m"月"d"日"`), qt.Equals, true)
	c.Assert(isTimeFormat(`上午/下午`), qt.Equals, true)
	c.Assert(isTimeFormat(`hh"時"mm"分" yyyy"年"m"月"`), qt.Equals, true)
	c.Assert(isTimeFormat(`上午/下午`), qt.Equals, true)
	c.Assert(isTimeFormat(`hh"時"mm"分"ss"秒 " m"月"d"日"`), qt.Equals, true)
	c.Assert(isTimeFormat(`[$-409]M/D/YYYY`), qt.Equals, true)
	c.Assert(isTimeFormat(`[$-404]e"年"m"月"d"日"`), qt.Equals, true)
	c.Assert(isTimeFormat(`m"月"d"日"`), qt.Equals, true)
	c.Assert(isTimeFormat(`上午/下午`), qt.Equals, true)
	c.Assert(isTimeFormat(`hh"時"mm"分" 上午/下午 h"时"mm"分"`), qt.Equals, true)
	c.Assert(isTimeFormat(`上午/下午`), qt.Equals, true)
	c.Assert(isTimeFormat(`hh"時"mm"分"ss"秒"`), qt.Equals, true)
	c.Assert(isTimeFormat(`下午`), qt.Equals, true)
	c.Assert(isTimeFormat(`h"时"mm"分"ss"秒"`), qt.Equals, true)
	c.Assert(isTimeFormat(`[$-404]e/m/d`), qt.Equals, true)
	c.Assert(isTimeFormat(`yyyy"年"m"月"`), qt.Equals, true)
	c.Assert(isTimeFormat(`[$-404]e"年"m"月"d"日"`), qt.Equals, true)
	c.Assert(isTimeFormat(`m"月"d"日"`), qt.Equals, true)
	c.Assert(isTimeFormat(`[$-404]e/m/d`), qt.Equals, true)
	c.Assert(isTimeFormat(`yyyy"5E74"m"6708"`), qt.Equals, true)
	c.Assert(isTimeFormat(`[$-404]e"5E74"m"6708"d"65E5"`), qt.Equals, true)
	c.Assert(isTimeFormat(`m"6708"d"65E5"`), qt.Equals, true)
	c.Assert(isTimeFormat(`[$-404]e"5E74"m"6708"d"65E5"`), qt.Equals, true)
	c.Assert(isTimeFormat(`m"6708"d"65E5"`), qt.Equals, true)
	c.Assert(isTimeFormat(`m/d/yy`), qt.Equals, true)
	c.Assert(isTimeFormat(`m-d-yy`), qt.Equals, true)
}
