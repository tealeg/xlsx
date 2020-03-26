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
}





// 54 [$-404]e"年"m"月"d"日" m"月"d"日"
// 55 上午/下午 hh"時"mm"分" 上午/下午 h"时"mm"分"
// 56 上午/下午 hh"時"mm"分"ss"秒
// " 上午/下午 h"时"mm"分"ss"秒
// "
// 57 [$-404]e/m/d yyyy"年"m"月"
// 58 [$-404]e"年"m"月"d"日" m"月"d"日"
// zh-tw and zh-cn (with unicode values provided for language glyphs where they occur)
// ID
// 1770
// zh-tw formatCode
// 	zh-cn formatCode

// 	27 [$-404]e/m/d yyyy"5E74"m"6708"
// 28 [$-404]e"5E74"m"6708"d"65E5" m"6708"d"65E5"
// 29 [$-404]e"5E74"m"6708"d"65E5" m"6708"d"65E5"
// 30 m/d/yy m-d-yy
// 31 yyyy"5E74"m"6708"d"65E5" yyyy"5E74"m"6708"d"65E5"
// 32 hh"6642"mm"5206" h"65F6"mm"5206"
// 33 hh"6642"mm"5206"ss"79D2" h"65F6"mm"5206"ss"79D2"
// 34 4E0A5348/4E0B5348hh"6642"mm"5206" 4E0A5348/4E0B5348h"65F6"mm"5206"
// 35 4E0A5348/4E0B5348hh"6642"mm"5206"ss"79
// D2" 4E0A5348/4E0B5348h"65F6"mm"5206"ss"79
// D2"
// 36 [$-404]e/m/d yyyy"5E74"m"6708"
// 50 [$-404]e/m/d yyyy"5E74"m"6708"
// 	51 [$-404]e"5E74"m"6708"d"65E5" m"6708"d"65E5"
// 	ID
// zh-tw formatCode
// zh-cn formatCode
// 52 4E0A5348/4E0B5348hh"6642"mm"5206" yyyy"5E74"m"6708"
// 53 4E0A5348/4E0B5348hh"6642"mm"5206"ss"79
// D2" m"6708"d"65E5"
// 54 [$-404]e"5E74"m"6708"d"65E5" m"6708"d"65E5"
// 55 4E0A5348/4E0B5348hh"6642"mm"5206" 4E0A5348/4E0B5348h"65F6"mm"5206"
// 56 4E0A5348/4E0B5348hh"6642"mm"5206"ss"79
// D2" 4E0A5348/4E0B5348h"65F6"mm"5206"ss"79
// D2"
// 57 [$-404]e/m/d yyyy"5E74"m"6708"
// 	58 [$-404]e"5E74"m"6708"d"65E5" m"6708"d"65E5"
// 	ID
// ja-jp formatCode
// ko-kr formatCode
// 27 [$-411]ge.m.d yyyy"年" mm"月" dd"日"
// 28 [$-411]ggge"年"m"月"d"日" mm-dd
// 29 [$-411]ggge"年"m"月"d"日" mm-dd
// 30 m/d/yy mm-dd-yy
// 31 yyyy"年"m"月"d"日" yyyy"년" mm"월" dd"일"
// 32 h"時"mm"分" h"시" mm"분"
// 33 h"時"mm"分"ss"秒" h"시" mm"분" ss"초"
// 34 yyyy"年"m"月" yyyy-mm-dd
// 35 m"月"d"日" yyyy-mm-dd
// 	36 [$-411]ge.m.d yyyy"年" mm"月" dd"日"
// 	50 [$-411]ge.m.d yyyy"年" mm"月" dd"日"
// 51 [$-411]ggge"年"m"月"d"日" mm-dd
// 52 yyyy"年"m"月" yyyy-mm-dd
// 53 m"月"d"日" yyyy-mm-dd
// 54 [$-411]ggge"年"m"月"d"日" mm-dd
// 55 yyyy"年"m"月" yyyy-mm-dd
// 56 m"月"d"日" yyyy-mm-dd
// 57 [$-411]ge.m.d yyyy"年" mm"月" dd"日"
// 	58 [$-411]ggge"年"m"月"d"日" mm-dd
// 	ja-jp and ko-kr (with unicode values provided for language glyphs where they occur)
// ID
// ja-jp formatCode
// 27 [$-411]ge.m.d yyyy"5E74" mm"6708" dd"65E5"
// 28 [$-411]ggge"5E74"m"6708"d"65E5" mm-dd
// 29 [$-411]ggge"5E74"m"6708"d"65E5" mm-dd
// 30 m/d/yy mm-dd-yy
// 31 yyyy"5E74"m"6708"d"65E5" yyyy"B144" mm"C6D4" dd"C77C"
// 32 h"6642"mm"5206" h"C2DC" mm"BD84"
// 33 h"6642"mm"5206"ss"79D2" h"C2DC" mm"BD84" ss"CD08"
// 34 yyyy"5E74"m"6708" yyyy-mm-dd
// 35 m"6708"d"65E5" yyyy-mm-dd
// 36 [$-411]ge.m.d yyyy"5E74" mm"6708" dd"65E5"
// 50 [$-411]ge.m.d yyyy"5E74" mm"6708" dd"65E5"
// 51 [$-411]ggge"5E74"m"6708"d"65E5" mm-dd
// 52 yyyy"5E74"m"6708" yyyy-mm-dd
// 53 m"6708"d"65E5" yyyy-mm-dd
// 54 [$-411]ggge"5E74"m"6708"d"65E5" mm-dd
// 55 yyyy"5E74"m"6708" yyyy-mm-dd
// 56 m"6708"d"65E5" yyyy-mm-dd
// 57 [$-411]ge.m.d yyyy"5E74" mm"6708" dd"65E5"
// 	58 [$-411]ggge"5E74"m"6708"d"65E5" mm-dd
// 	th-th
// ID
// 1772
// ko-kr formatCode
// th-th formatCode
// 59 t0
// 60 t0.00
// 61 t#,##0
// 62 t#,##0.00
// 67 t0%
// 68 t0.00%
// 69 t# ?/?18. SpreadsheetML Reference Material
// ID
// th-th formatCode
// 70 t# ??/??
// 71 /ด/ปปปป
// 72 -ดดด-ปป
// 73 -ดดด
// 74 ดดด-ปป
// 75 ช:
// 76 ช: :
// /ด/ปปปป ช:
// 77
// 78 :
// 79 [ช]: :
// 80 : .0
// 81 d/m/bb
// th-th (with unicode values provided for language glyphs where they occur)
// ID
// th-th formatCode
// 59 t0
// 60 t0.00
// 61 t#,##0
// 62 t#,##0.00
// 67 t0%
// 68 t0.00%
// 69 t# ?/?
// 70 t# ??/??
// 71 0E27/0E14/0E1B0E1B0E1B0E1B
// 72 0E27-0E140E140E14-0E1B0E1B
// 73 0E27-0E140E140E14
// 74 0E140E140E14-0E1B0E1B
// 75 0E0A:0E190E19
// 76 0E0A:0E190E19:0E170E17
// 77 0E27/0E14/0E1B0E1B0E1B0E1B 0E0A:0E190E19
// 1773ECMA-376 Part 1
// ID
// th-th formatCode
// 78 0E190E19:0E170E17
// 79 [0E0A]:0E190E19:0E170E17
// 80 0E190E19:0E170E17.0
// 81 d/m/bb
