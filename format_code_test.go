package xlsx

import (
	"time"

	. "gopkg.in/check.v1"
)

func (s *CellSuite) TestMoreFormattingFeatures(c *C) {

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

func (l *CellSuite) TestFormatStringSupport(c *C) {
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
}
