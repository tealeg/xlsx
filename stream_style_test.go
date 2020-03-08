package xlsx

import (
	"bytes"
	"errors"
	"fmt"
	"io"
	"strconv"
	"testing"
	"time"

	qt "github.com/frankban/quicktest"
)

const (
	StyleStreamTestsShouldMakeRealFiles = false
)

func TestStreamTestsShouldMakeRealFilesShouldBeFalse(t *testing.T) {
	if StyleStreamTestsShouldMakeRealFiles {
		t.Fatal("TestsShouldMakeRealFiles should only be true for local debugging. Don't forget to switch back before commiting.")
	}
}

func TestXlsxStreamWriteWithStyle(t *testing.T) {
	// When shouldMakeRealFiles is set to true this test will make actual XLSX files in the file system.
	// This is useful to ensure files open in Excel, Numbers, Google Docs, etc.
	// In case of issues you can use "Open XML SDK 2.5" to diagnose issues in generated XLSX files:
	// https://www.microsoft.com/en-us/download/details.aspx?id=30425
	testCases := []struct {
		testName      string
		sheetNames    []string
		workbookData  [][][]StreamCell
		expectedError error
	}{
		{
			testName: "Style Test",
			sheetNames: []string{
				"Sheet1",
			},
			workbookData: [][][]StreamCell{
				{
					{
						NewStyledStringStreamCell("1", StreamStyleUnderlinedString),
						NewStyledStringStreamCell("25", StreamStyleItalicString),
						NewStyledStringStreamCell("A", StreamStyleBoldString),
						NewStringStreamCell("B"),
					},
					{
						NewIntegerStreamCell(1234),
						NewStyledIntegerStreamCell(98, StreamStyleBoldInteger),
						NewStyledIntegerStreamCell(34, StreamStyleItalicInteger),
						NewStyledIntegerStreamCell(26, StreamStyleUnderlinedInteger),
					},
				},
			},
		},
		{
			testName: "One Sheet",
			sheetNames: []string{
				"Sheet1",
			},
			workbookData: [][][]StreamCell{
				{
					{NewStringStreamCell("Token"), NewStringStreamCell("Name"),
						NewStringStreamCell("Price"), NewStringStreamCell("SKU")},
					{NewIntegerStreamCell(123), NewStringStreamCell("Taco"),
						NewIntegerStreamCell(300), NewIntegerStreamCell(123)},
				},
			},
		},
		{
			testName: "One Column",
			sheetNames: []string{
				"Sheet1",
			},
			workbookData: [][][]StreamCell{
				{
					{NewStringStreamCell("Token")},
					{NewIntegerStreamCell(123)},
				},
			},
		},
		{
			testName: "Several Sheets, with different numbers of columns and rows",
			sheetNames: []string{
				"Sheet 1", "Sheet 2", "Sheet3",
			},
			workbookData: [][][]StreamCell{
				{
					{NewStringStreamCell("Token"), NewStringStreamCell("Name"),
						NewStringStreamCell("Price"), NewStringStreamCell("SKU")},

					{NewIntegerStreamCell(123), NewStringStreamCell("Taco"),
						NewIntegerStreamCell(300), NewIntegerStreamCell(123)},
				},
				{
					{NewStringStreamCell("Token"), NewStringStreamCell("Name"),
						NewStringStreamCell("Price"), NewStringStreamCell("SKU"),
						NewStringStreamCell("Stock")},

					{NewIntegerStreamCell(456), NewStringStreamCell("Salsa"),
						NewIntegerStreamCell(200), NewIntegerStreamCell(346),
						NewIntegerStreamCell(1)},

					{NewIntegerStreamCell(789), NewStringStreamCell("Burritos"),
						NewIntegerStreamCell(400), NewIntegerStreamCell(754),
						NewIntegerStreamCell(3)},
				},
				{
					{NewStringStreamCell("Token"), NewStringStreamCell("Name"),
						NewStringStreamCell("Price")},

					{NewIntegerStreamCell(9853), NewStringStreamCell("Guacamole"),
						NewIntegerStreamCell(500)},

					{NewIntegerStreamCell(2357), NewStringStreamCell("Margarita"),
						NewIntegerStreamCell(700)},
				},
			},
		},
		{
			testName: "Two Sheets with same the name",
			sheetNames: []string{
				"Sheet 1", "Sheet 1",
			},
			workbookData: [][][]StreamCell{
				{
					{NewStringStreamCell("Token"), NewStringStreamCell("Name"),
						NewStringStreamCell("Price"), NewStringStreamCell("SKU")},

					{NewIntegerStreamCell(123), NewStringStreamCell("Taco"),
						NewIntegerStreamCell(300), NewIntegerStreamCell(123)},
				},
				{
					{NewStringStreamCell("Token"), NewStringStreamCell("Name"),
						NewStringStreamCell("Price"), NewStringStreamCell("SKU"),
						NewStringStreamCell("Stock")},

					{NewIntegerStreamCell(456), NewStringStreamCell("Salsa"),
						NewIntegerStreamCell(200), NewIntegerStreamCell(346),
						NewIntegerStreamCell(1)},

					{NewIntegerStreamCell(789), NewStringStreamCell("Burritos"),
						NewIntegerStreamCell(400), NewIntegerStreamCell(754),
						NewIntegerStreamCell(3)},
				},
			},
			expectedError: fmt.Errorf("duplicate sheet name '%s'.", "Sheet 1"),
		},
		{
			testName: "One Sheet Registered, tries to write to two",
			sheetNames: []string{
				"Sheet 1",
			},
			workbookData: [][][]StreamCell{
				{
					{NewStringStreamCell("Token"), NewStringStreamCell("Name"),
						NewStringStreamCell("Price"), NewStringStreamCell("SKU")},

					{NewIntegerStreamCell(123), NewStringStreamCell("Taco"),
						NewIntegerStreamCell(300), NewIntegerStreamCell(123)},
				},
				{
					{NewStringStreamCell("Token"), NewStringStreamCell("Name"),
						NewStringStreamCell("Price"), NewStringStreamCell("SKU")},

					{NewIntegerStreamCell(456), NewStringStreamCell("Salsa"),
						NewIntegerStreamCell(200), NewIntegerStreamCell(346)},
				},
			},
			expectedError: AlreadyOnLastSheetError,
		},
		{
			testName: "One Sheet, too many columns in row 1",
			sheetNames: []string{
				"Sheet 1",
			},
			workbookData: [][][]StreamCell{
				{
					{NewStringStreamCell("Token"), NewStringStreamCell("Name"),
						NewStringStreamCell("Price"), NewStringStreamCell("SKU")},

					{NewIntegerStreamCell(123), NewStringStreamCell("Taco"),
						NewIntegerStreamCell(300), NewIntegerStreamCell(123),
						NewStringStreamCell("asdf")},
				},
			},
			expectedError: WrongNumberOfRowsError,
		},
		{
			testName: "One Sheet, too few columns in row 1",
			sheetNames: []string{
				"Sheet 1",
			},
			workbookData: [][][]StreamCell{
				{
					{NewStringStreamCell("Token"), NewStringStreamCell("Name"),
						NewStringStreamCell("Price"), NewStringStreamCell("SKU")},

					{NewIntegerStreamCell(123), NewStringStreamCell("Taco"),
						NewIntegerStreamCell(300)},
				},
			},
			expectedError: WrongNumberOfRowsError,
		},
		{
			testName: "Lots of Sheets, only writes rows to one, only writes headers to one, should not error and should still create a valid file",
			sheetNames: []string{
				"Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4", "Sheet 5", "Sheet 6",
			},
			workbookData: [][][]StreamCell{
				{
					{NewStringStreamCell("Token"), NewStringStreamCell("Name"),
						NewStringStreamCell("Price"), NewStringStreamCell("SKU")},

					{NewIntegerStreamCell(123), NewStringStreamCell("Taco"),
						NewIntegerStreamCell(300), NewIntegerStreamCell(123)},
				},
				{{}},
				{{NewStringStreamCell("Id"), NewStringStreamCell("Unit Cost")}},
				{{}},
				{{}},
				{{}},
			},
		},
		{
			testName: "Two Sheets, only writes to one, should not error and should still create a valid file",
			sheetNames: []string{
				"Sheet 1", "Sheet 2",
			},
			workbookData: [][][]StreamCell{
				{
					{NewStringStreamCell("Token"), NewStringStreamCell("Name"),
						NewStringStreamCell("Price"), NewStringStreamCell("SKU")},

					{NewIntegerStreamCell(123), NewStringStreamCell("Taco"),
						NewIntegerStreamCell(300), NewIntegerStreamCell(123)},
				},
				{{}},
			},
		},
		{
			testName: "UTF-8 Characters. This XLSX File loads correctly with Excel, Numbers, and Google Docs. It also passes Microsoft's Office File Format Validator.",
			sheetNames: []string{
				"Sheet1",
			},
			workbookData: [][][]StreamCell{
				{
					// String courtesy of https://github.com/minimaxir/big-list-of-naughty-strings/
					// Header row contains the tags that I am filtering on
					{NewStringStreamCell("Token"), NewStringStreamCell(endSheetDataTag),
						NewStringStreamCell("Price"), NewStringStreamCell(fmt.Sprintf(dimensionTag, "A1:D1"))},
					// Japanese and emojis
					{NewIntegerStreamCell(123), NewStringStreamCell("パーティーへ行かないか"),
						NewIntegerStreamCell(300), NewStringStreamCell("🍕🐵 🙈 🙉 🙊")},
					// XML encoder/parser test strings
					{NewIntegerStreamCell(123), NewStringStreamCell(`<?xml version="1.0" encoding="ISO-8859-1"?>`),
						NewIntegerStreamCell(300), NewStringStreamCell(`<?xml version="1.0" encoding="ISO-8859-1"?><!DOCTYPE foo [ <!ELEMENT foo ANY ><!ENTITY xxe SYSTEM "file:///etc/passwd" >]><foo>&xxe;</foo>`)},
					// Upside down text and Right to Left Arabic text
					{NewIntegerStreamCell(123), NewStringStreamCell(`˙ɐnbᴉlɐ ɐuƃɐɯ ǝɹolop ʇǝ ǝɹoqɐl ʇn ʇunpᴉpᴉɔuᴉ ɹodɯǝʇ poɯsnᴉǝ op pǝs 'ʇᴉlǝ ƃuᴉɔsᴉdᴉpɐ ɹnʇǝʇɔǝsuoɔ 'ʇǝɯɐ ʇᴉs ɹolop ɯnsdᴉ ɯǝɹo˥
					00˙Ɩ$-`), NewIntegerStreamCell(300), NewStringStreamCell(`ﷺ`)},
					{NewIntegerStreamCell(123), NewStringStreamCell("Taco"),
						NewIntegerStreamCell(300), NewIntegerStreamCell(123)},
				},
			},
		},
	}

	c := qt.New(t)
	for i, testCase := range testCases {
		csRunO(c, testCase.testName, func(c *qt.C, option FileOption) {
			var filePath string
			var buffer bytes.Buffer
			if StyleStreamTestsShouldMakeRealFiles {
				filePath = fmt.Sprintf("WorkbookWithStyle%d.xlsx", i)
			}

			err := writeStreamFileWithStyle(filePath, &buffer, testCase.sheetNames, testCase.workbookData, StyleStreamTestsShouldMakeRealFiles, []StreamStyle{}, option)
			switch {
			case err == nil && testCase.expectedError != nil:
				t.Fatalf("Expected error but none was returned")
			case err != nil && testCase.expectedError == nil:
				t.Fatalf("Unexpected error: %q", err.Error())
			case err != testCase.expectedError && err.Error() != testCase.expectedError.Error():
				t.Fatalf("Error differs from expected error. Error: %v, Expected Error: %v ", err, testCase.expectedError)
			case err != nil:
				// We got an expected error
				return
			}

			// read the file back with the xlsx package
			var bufReader *bytes.Reader
			var size int64
			if !StyleStreamTestsShouldMakeRealFiles {
				bufReader = bytes.NewReader(buffer.Bytes())
				size = bufReader.Size()
			}
			actualSheetNames, actualWorkbookData, actualWorkbookCells := readXLSXFileS(t, filePath, bufReader, size, StyleStreamTestsShouldMakeRealFiles, option)
			// check if data was able to be read correctly
			c.Assert(actualSheetNames, qt.DeepEquals, testCase.sheetNames)
			expectedWorkbookDataStrings := [][][]string{}
			for j := range testCase.workbookData {
				expectedWorkbookDataStrings = append(expectedWorkbookDataStrings, [][]string{})
				for k := range testCase.workbookData[j] {
					if len(testCase.workbookData[j][k]) == 0 {
						expectedWorkbookDataStrings[j] = append(expectedWorkbookDataStrings[j], nil)
					} else {
						expectedWorkbookDataStrings[j] = append(expectedWorkbookDataStrings[j], []string{})
						for _, cell := range testCase.workbookData[j][k] {
							expectedWorkbookDataStrings[j][k] = append(expectedWorkbookDataStrings[j][k], cell.cellData)
						}
					}
				}

			}
			c.Assert(actualWorkbookData, qt.DeepEquals, expectedWorkbookDataStrings)

			c.Assert(checkForCorrectCellStyles(actualWorkbookCells, testCase.workbookData), qt.Equals, nil)

		})
	}
}

// writeStreamFile will write the file using this stream package
func writeStreamFileWithStyle(filePath string, fileBuffer io.Writer, sheetNames []string, workbookData [][][]StreamCell,
	shouldMakeRealFiles bool, customStyles []StreamStyle, options ...FileOption) error {
	var file *StreamFileBuilder
	var err error
	if shouldMakeRealFiles {
		file, err = NewStreamFileBuilderForPath(filePath, options...)
		if err != nil {
			return err
		}
	} else {
		file = NewStreamFileBuilder(fileBuffer, options...)
	}

	defaultStyles := []StreamStyle{StreamStyleDefaultString, StreamStyleBoldString, StreamStyleItalicString, StreamStyleUnderlinedString,
		StreamStyleDefaultInteger, StreamStyleBoldInteger, StreamStyleItalicInteger, StreamStyleUnderlinedInteger,
		StreamStyleDefaultDate}
	allStylesToBeAdded := append(defaultStyles, customStyles...)
	err = file.AddStreamStyleList(allStylesToBeAdded)
	if err != nil {
		return err
	}

	for i, sheetName := range sheetNames {
		var colStyles []StreamStyle
		for range workbookData[i][0] {
			colStyles = append(colStyles, StreamStyleDefaultString)
		}

		err := file.AddSheetS(sheetName, colStyles)
		if err != nil {
			return err
		}
	}
	streamFile, err := file.Build()
	if err != nil {
		return err
	}
	for i, sheetData := range workbookData {

		if i != 0 {
			err = streamFile.NextSheet()
			if err != nil {
				return err
			}
		}
		if i%2 == 0 {
			err = streamFile.WriteAllS(sheetData)
		} else {
			for _, row := range sheetData {
				err = streamFile.WriteS(row)
				if err != nil {
					return err
				}
			}
		}
	}
	err = streamFile.Close()
	if err != nil {
		return err
	}
	return nil
}

// readXLSXFileS will read the file using the xlsx package.
func readXLSXFileS(t *testing.T, filePath string, fileBuffer io.ReaderAt, size int64, shouldMakeRealFiles bool, options ...FileOption) ([]string, [][][]string, [][][]Cell) {
	var readFile *File
	var err error
	if shouldMakeRealFiles {
		readFile, err = OpenFile(filePath, options...)
		if err != nil {
			t.Fatal(err)
		}
	} else {
		readFile, err = OpenReaderAt(fileBuffer, size, options...)
		if err != nil {
			t.Fatal(err)
		}
	}
	var actualWorkbookData [][][]string
	var sheetNames []string
	var actualWorkBookCells [][][]Cell
	for i, sheet := range readFile.Sheets {
		actualWorkBookCells = append(actualWorkBookCells, [][]Cell{})
		var sheetData [][]string
		err = sheet.ForEachRow(func(row *Row) error {
			j := row.num
			actualWorkBookCells[i] = append(actualWorkBookCells[i], []Cell{})
			var data []string
			err := row.ForEachCell(func(cell *Cell) error {
				actualWorkBookCells[i][j] = append(actualWorkBookCells[i][j], *cell)
				str, err := cell.FormattedValue()
				if err != nil {
					return err
				}
				data = append(data, str)
				return nil
			})
			if err != nil {
				return err
			}
			sheetData = append(sheetData, data)
			return nil
		})
		if err != nil {
			t.Fatal(err)
		}
		sheetNames = append(sheetNames, sheet.Name)
		actualWorkbookData = append(actualWorkbookData, sheetData)
	}
	return sheetNames, actualWorkbookData, actualWorkBookCells
}

func TestStreamStyleDates(t *testing.T) {
	c := qt.New(t)
	csRunO(c, "StreamStyleDates", func(c *qt.C, option FileOption) {
		var filePath string
		var buffer bytes.Buffer
		if StyleStreamTestsShouldMakeRealFiles {
			filePath = fmt.Sprintf("Workbook_Date_test.xlsx")
		}

		// We use a fixed time to avoid weird errors around midnight
		fixedTime := time.Date(2009, time.November, 10, 23, 0, 0, 0, time.UTC)
		sheetNames := []string{"Sheet1"}
		workbookData := [][][]StreamCell{
			{
				{NewStringStreamCell("Date:")},
				{NewDateStreamCell(fixedTime)},
			},
		}

		err := writeStreamFileWithStyle(filePath, &buffer, sheetNames, workbookData, StyleStreamTestsShouldMakeRealFiles, []StreamStyle{}, option)
		c.Assert(err, qt.Equals, nil)

		// read the file back with the xlsx package
		var bufReader *bytes.Reader
		var size int64
		if !StyleStreamTestsShouldMakeRealFiles {
			bufReader = bytes.NewReader(buffer.Bytes())
			size = bufReader.Size()
		}
		actualSheetNames, actualWorkbookData, actualWorkbookCells := readXLSXFileS(t, filePath, bufReader, size, StyleStreamTestsShouldMakeRealFiles, option)
		// check if data was able to be read correctly
		c.Assert(actualSheetNames, qt.DeepEquals, sheetNames)

		expectedWorkbookDataStrings := [][][]string{}
		for j := range workbookData {
			expectedWorkbookDataStrings = append(expectedWorkbookDataStrings, [][]string{})
			for range workbookData[j] {
				expectedWorkbookDataStrings[j] = append(expectedWorkbookDataStrings[j], []string{})
			}
		}

		expectedWorkbookDataStrings[0][0] = append(expectedWorkbookDataStrings[0][0], workbookData[0][0][0].cellData)
		year, month, day := fixedTime.Date()
		monthString := strconv.Itoa(int(month))
		if int(month) < 10 {
			monthString = "0" + monthString
		}
		dayString := strconv.Itoa(day)
		if day < 10 {
			dayString = "0" + dayString
		}
		yearString := strconv.Itoa(year - 2000)
		if (year - 2000) < 10 {
			yearString = "0" + yearString
		}
		expectedWorkbookDataStrings[0][1] = append(
			expectedWorkbookDataStrings[0][1],
			monthString+"-"+dayString+"-"+yearString)
		c.Assert(actualWorkbookData, qt.DeepEquals, expectedWorkbookDataStrings)

		c.Assert(checkForCorrectCellStyles(actualWorkbookCells, workbookData), qt.Equals, nil)
	})
}

func TestMakeNewStylesAndUseIt(t *testing.T) {
	c := qt.New(t)
	csRunO(c, "MakeNewStyles", func(c *qt.C, option FileOption) {
		var filePath string
		var buffer bytes.Buffer
		if StyleStreamTestsShouldMakeRealFiles {
			filePath = fmt.Sprintf("Workbook_newStyle.xlsx")
		}

		timesNewRoman12 := NewFont(12, TimesNewRoman)
		timesNewRoman12.Color = RGB_Dark_Green
		courier12 := NewFont(12, Courier)
		courier12.Color = RGB_Dark_Red

		greenFill := NewFill(Solid_Cell_Fill, RGB_Light_Green, RGB_White)
		redFill := NewFill(Solid_Cell_Fill, RGB_Light_Red, RGB_White)

		greenStyle := MakeStyle(GeneralFormat, timesNewRoman12, greenFill, DefaultAlignment(), DefaultBorder())
		redStyle := MakeStyle(GeneralFormat, courier12, redFill, DefaultAlignment(), DefaultBorder())

		sheetNames := []string{"Sheet1"}
		workbookData := [][][]StreamCell{
			{
				{NewStringStreamCell("TRUE"), NewStringStreamCell("False")},
				{NewStyledStringStreamCell("Good", greenStyle), NewStyledStringStreamCell("Bad", redStyle)},
			},
		}

		err := writeStreamFileWithStyle(filePath, &buffer, sheetNames, workbookData, StyleStreamTestsShouldMakeRealFiles, []StreamStyle{greenStyle, redStyle}, option)
		c.Assert(err, qt.Equals, nil)

		// read the file back with the xlsx package
		var bufReader *bytes.Reader
		var size int64
		if !StyleStreamTestsShouldMakeRealFiles {
			bufReader = bytes.NewReader(buffer.Bytes())
			size = bufReader.Size()
		}
		actualSheetNames, actualWorkbookData, actualWorkbookCells := readXLSXFileS(t, filePath, bufReader, size, StyleStreamTestsShouldMakeRealFiles, option)
		// check if data was able to be read correctly
		c.Assert(actualSheetNames, qt.DeepEquals, sheetNames)

		expectedWorkbookDataStrings := [][][]string{}
		for j := range workbookData {
			expectedWorkbookDataStrings = append(expectedWorkbookDataStrings, [][]string{})
			for k := range workbookData[j] {
				expectedWorkbookDataStrings[j] = append(expectedWorkbookDataStrings[j], []string{})
				for _, cell := range workbookData[j][k] {
					expectedWorkbookDataStrings[j][k] = append(expectedWorkbookDataStrings[j][k], cell.cellData)
				}
			}

		}
		c.Assert(actualWorkbookData, qt.DeepEquals, expectedWorkbookDataStrings)

		c.Assert(checkForCorrectCellStyles(actualWorkbookCells, workbookData), qt.Equals, nil)
	})
}

func TestStreamNewTypes(t *testing.T) {

	c := qt.New(t)
	csRunO(c, "NewTypes", func(c *qt.C, option FileOption) {
		var filePath string
		var buffer bytes.Buffer
		if StyleStreamTestsShouldMakeRealFiles {
			filePath = fmt.Sprintf("Workbook_newStyle.xlsx")
		}

		sheetNames := []string{"Sheet1"}
		workbookData := [][][]StreamCell{
			{
				{NewStreamCell("1", StreamStyleDefaultString, CellTypeBool),
					NewStreamCell("InLine", StreamStyleBoldString, CellTypeInline),
					NewStreamCell("Error", StreamStyleDefaultString, CellTypeError)},
			},
		}

		err := writeStreamFileWithStyle(filePath, &buffer, sheetNames, workbookData, StyleStreamTestsShouldMakeRealFiles, []StreamStyle{}, option)
		c.Assert(err, qt.Equals, nil)

		// read the file back with the xlsx package
		var bufReader *bytes.Reader
		var size int64
		if !StyleStreamTestsShouldMakeRealFiles {
			bufReader = bytes.NewReader(buffer.Bytes())
			size = bufReader.Size()
		}
		actualSheetNames, actualWorkbookData, actualWorkbookCells := readXLSXFileS(t, filePath, bufReader, size, StyleStreamTestsShouldMakeRealFiles, option)
		// check if data was able to be read correctly
		c.Assert(actualSheetNames, qt.DeepEquals, sheetNames)

		expectedWorkbookDataStrings := [][][]string{}
		for j := range workbookData {
			expectedWorkbookDataStrings = append(expectedWorkbookDataStrings, [][]string{})
			for k := range workbookData[j] {
				expectedWorkbookDataStrings[j] = append(expectedWorkbookDataStrings[j], []string{})
				for _, cell := range workbookData[j][k] {
					if cell.cellData == "1" {
						expectedWorkbookDataStrings[j][k] = append(expectedWorkbookDataStrings[j][k], "TRUE")
					} else {
						expectedWorkbookDataStrings[j][k] = append(expectedWorkbookDataStrings[j][k], cell.cellData)
					}
				}
			}

		}
		c.Assert(actualWorkbookData, qt.DeepEquals, expectedWorkbookDataStrings)
		c.Assert(checkForCorrectCellStyles(actualWorkbookCells, workbookData), qt.Equals, nil)
	})
}

func TestStreamCloseWithNothingWrittenToSheetsWithStyle(t *testing.T) {
	c := qt.New(t)
	csRunO(c, "NothingWrittenToSheets", func(c *qt.C, option FileOption) {
		buffer := bytes.NewBuffer(nil)
		file := NewStreamFileBuilder(buffer, option)

		sheetNames := []string{"Sheet1", "Sheet2"}
		workbookData := [][][]StreamCell{
			{{NewStringStreamCell("Header1"), NewStringStreamCell("Header2")}},
			{{NewStringStreamCell("Header3"), NewStringStreamCell("Header4")}},
		}

		defaultStyles := []StreamStyle{StreamStyleDefaultString, StreamStyleBoldString, StreamStyleItalicInteger, StreamStyleUnderlinedString,
			StreamStyleDefaultInteger, StreamStyleBoldInteger, StreamStyleItalicInteger, StreamStyleUnderlinedInteger,
			StreamStyleDefaultDate}
		err := file.AddStreamStyleList(defaultStyles)
		c.Assert(err, qt.Equals, nil)

		colStyles0 := []StreamStyle{}
		for range workbookData[0][0] {
			colStyles0 = append(colStyles0, StreamStyleDefaultString)
		}

		colStyles1 := []StreamStyle{}
		for range workbookData[1][0] {
			colStyles1 = append(colStyles1, StreamStyleDefaultString)
		}

		err = file.AddSheetS(sheetNames[0], colStyles0)
		c.Assert(err, qt.Equals, nil)

		err = file.AddSheetS(sheetNames[1], colStyles1)
		c.Assert(err, qt.Equals, nil)

		stream, err := file.Build()
		c.Assert(err, qt.Equals, nil)
		err = stream.Close()
		c.Assert(err, qt.Equals, nil)

		bufReader := bytes.NewReader(buffer.Bytes())
		size := bufReader.Size()

		actualSheetNames, actualWorkbookData, _ := readXLSXFileS(t, "", bufReader, size, false, option)
		// check if data was able to be read correctly
		c.Assert(actualSheetNames, qt.DeepEquals, sheetNames)

		expectedWorkbookDataStrings := [][][]string{}
		for range workbookData {
			expectedWorkbookDataStrings = append(expectedWorkbookDataStrings, nil)
		}
		c.Assert(actualWorkbookData, qt.DeepEquals, expectedWorkbookDataStrings)
	})
}

func TestStreamBuildErrorsAfterBuildWithStyle(t *testing.T) {
	c := qt.New(t)
	csRunO(c, "BuildErrorsAfterBuildWithStyle", func(c *qt.C, option FileOption) {
		file := NewStreamFileBuilder(bytes.NewBuffer(nil), option)

		defaultStyles := []StreamStyle{StreamStyleDefaultString, StreamStyleBoldString, StreamStyleItalicInteger, StreamStyleUnderlinedString,
			StreamStyleDefaultInteger, StreamStyleBoldInteger, StreamStyleItalicInteger, StreamStyleUnderlinedInteger,
			StreamStyleDefaultDate}
		err := file.AddStreamStyleList(defaultStyles)
		if err != nil {
			t.Fatal(err)
		}

		err = file.AddSheetS("Sheet1", []StreamStyle{StreamStyleDefaultString})
		c.Assert(err, qt.Equals, nil)
		err = file.AddSheetS("Sheet2", []StreamStyle{StreamStyleDefaultString})
		c.Assert(err, qt.Equals, nil)

		_, err = file.Build()
		c.Assert(err, qt.Equals, nil)
		_, err = file.Build()
		c.Assert(err, qt.Equals, BuiltStreamFileBuilderError)
	})
}

func TestStreamAddSheetSWithErrorsAfterBuild(t *testing.T) {
	c := qt.New(t)
	csRunO(c, "AddSheetSWithErrorsAfterBuild", func(c *qt.C, option FileOption) {
		file := NewStreamFileBuilder(bytes.NewBuffer(nil), option)

		defaultStyles := []StreamStyle{StreamStyleDefaultString, StreamStyleBoldString, StreamStyleItalicInteger, StreamStyleUnderlinedString,
			StreamStyleDefaultInteger, StreamStyleBoldInteger, StreamStyleItalicInteger, StreamStyleUnderlinedInteger,
			StreamStyleDefaultDate}
		err := file.AddStreamStyleList(defaultStyles)
		c.Assert(err, qt.Equals, nil)

		err = file.AddSheetS("Sheet1", []StreamStyle{StreamStyleDefaultString})
		c.Assert(err, qt.Equals, nil)
		err = file.AddSheetS("Sheet2", []StreamStyle{StreamStyleDefaultString})
		c.Assert(err, qt.Equals, nil)

		_, err = file.Build()
		c.Assert(err, qt.Equals, nil)
		err = file.AddSheetS("Sheet3", []StreamStyle{StreamStyleDefaultString})
		c.Assert(err, qt.Equals, BuiltStreamFileBuilderError)
	})
}

func TestStreamNoStylesAddSheetSError(t *testing.T) {
	c := qt.New(t)
	csRunO(c, "NoStylesAddSheetSError", func(c *qt.C, option FileOption) {
		buffer := bytes.NewBuffer(nil)
		file := NewStreamFileBuilder(buffer, option)

		sheetNames := []string{"Sheet1", "Sheet2"}
		workbookData := [][][]StreamCell{
			{{NewStringStreamCell("Header1"), NewStringStreamCell("Header2")}},
			{{NewStyledStringStreamCell("Header3", StreamStyleBoldString), NewStringStreamCell("Header4")}},
		}

		colStyles0 := []StreamStyle{}
		for range workbookData[0][0] {
			colStyles0 = append(colStyles0, StreamStyleDefaultString)
		}

		err := file.AddSheetS(sheetNames[0], colStyles0)
		c.Assert(err, qt.ErrorMatches, "trying to make use of a style that has not been added")
	})
}

func TestStreamNoStylesWriteSError(t *testing.T) {
	c := qt.New(t)
	csRunO(c, "NoStylesWriteSError", func(c *qt.C, option FileOption) {
		buffer := bytes.NewBuffer(nil)
		var filePath string

		greenStyle := MakeStyle(GeneralFormat, DefaultFont(), FillGreen, DefaultAlignment(), DefaultBorder())

		sheetNames := []string{"Sheet1", "Sheet2"}
		workbookData := [][][]StreamCell{
			{{NewStringStreamCell("Header1"), NewStringStreamCell("Header2")}},
			{{NewStyledStringStreamCell("Header3", greenStyle), NewStringStreamCell("Header4")}},
		}

		err := writeStreamFileWithStyle(filePath, buffer, sheetNames, workbookData, StyleStreamTestsShouldMakeRealFiles, []StreamStyle{}, option)
		c.Assert(err, qt.ErrorMatches, "trying to make use of a style that has not been added")
	})
}

func checkForCorrectCellStyles(actualCells [][][]Cell, expectedCells [][][]StreamCell) error {
	for i := range actualCells {
		for j := range actualCells[i] {
			for k, actualCell := range actualCells[i][j] {
				expectedCell := expectedCells[i][j][k]
				if err := compareCellStyles(actualCell, expectedCell); err != nil {
					return err
				}
			}
		}
	}
	return nil
}

func compareCellStyles(cellA Cell, cellB StreamCell) error {
	fontA := cellA.style.Font
	fontB := cellB.cellStyle.style.Font

	if fontA != fontB {
		return errors.New("actual and expected font do not match")
	}

	numFmtA := cellA.NumFmt
	numFmtB := builtInNumFmt[cellB.cellStyle.xNumFmtId]
	if numFmtA != numFmtB {
		return errors.New("actual and expected NumFmt do not match")
	}

	return nil
}

func TestCustomNumberFormat(t *testing.T) {
	buffer := bytes.NewBuffer(nil)
	fileBuilder := NewStreamFileBuilder(buffer)
	customNumFmtId := fileBuilder.AddNewNumberFormat("0.00000")
	if customNumFmtId != 164 {
		t.Error("Unexpected value")
	}

	customDateNumFmtId := fileBuilder.AddNewNumberFormat("dd/mm/yy")
	if customDateNumFmtId != 165 {
		t.Error("Unexpected value")
	}

	if fileBuilder.AddNewNumberFormat("0.00000") != 164 {
		t.Error("Unexpected value")
	}

	customStyles := []StreamStyle{
		MakeStyle(customNumFmtId, DefaultFont(), DefaultFill(), DefaultAlignment(), DefaultBorder()),
		MakeStyle(customDateNumFmtId, DefaultFont(), DefaultFill(), DefaultAlignment(), DefaultBorder()),
	}
	err := fileBuilder.AddStreamStyleList(customStyles)
	if err != nil {
		t.Fatal(err)
	}

	err = fileBuilder.AddSheetS("Sheet1", customStyles)
	if err != nil {
		t.Fatal(err)
	}

	streamFile, err := fileBuilder.Build()
	if err != nil {
		t.Fatal(err)
	}
	if err = streamFile.Close(); err != nil {
		t.Fatal(err)
	}

	file, err := OpenBinary(buffer.Bytes())
	if err != nil {
		t.Fatal(err)
	}

	numberFormats := map[int]string{}
	for _, numFmt := range file.styles.NumFmts.NumFmt {
		numberFormats[numFmt.NumFmtId] = numFmt.FormatCode
	}

	formatCode, ok := numberFormats[164]
	if !ok {
		t.Error("Custom number format not found")
	}
	if formatCode != "0.00000" {
		t.Error("Incorrect format code")
	}

	formatCode, ok = numberFormats[165]
	if !ok {
		t.Error("Custom number format not found")
	}
	if formatCode != "dd/mm/yy" {
		t.Error("Incorrect format code")
	}
}
