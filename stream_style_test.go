package xlsx

import (
	"bytes"
	"fmt"
	. "gopkg.in/check.v1"
	"io"
	"reflect"
	"time"
)

const (
	StyleStreamTestsShouldMakeRealFiles = false
)

type StreamStyleSuite struct{}

var _ = Suite(&StreamStyleSuite{})

func (s *StreamSuite) TestStreamTestsShouldMakeRealFilesShouldBeFalse(t *C) {
	if StyleStreamTestsShouldMakeRealFiles {
		t.Fatal("TestsShouldMakeRealFiles should only be true for local debugging. Don't forget to switch back before commiting.")
	}
}

func (s *StreamSuite) TestXlsxStreamWriteWithStyle(t *C) {
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
					{NewStyledStringStreamCell("1", UnderlinedStrings), NewStyledStringStreamCell("25", ItalicStrings),
						NewStyledStringStreamCell("A", BoldStrings), NewStringStreamCell("B")},
					{NewIntegerStreamCell(1234), NewStyledIntegerStreamCell(98, BoldIntegers),
						NewStyledIntegerStreamCell(34, ItalicIntegers), NewStyledIntegerStreamCell(26, UnderlinedIntegers)},
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

	for i, testCase := range testCases {
		var filePath string
		var buffer bytes.Buffer
		if StyleStreamTestsShouldMakeRealFiles {
			filePath = fmt.Sprintf("WorkbookWithStyle%d.xlsx", i)
		}

		err := writeStreamFileWithStyle(filePath, &buffer, testCase.sheetNames, testCase.workbookData, StyleStreamTestsShouldMakeRealFiles, []StreamStyle{})
		if err != testCase.expectedError && err.Error() != testCase.expectedError.Error() {
			t.Fatalf("Error differs from expected error. Error: %v, Expected Error: %v ", err, testCase.expectedError)
		}
		if testCase.expectedError != nil {
			return
		}
		// read the file back with the xlsx package
		var bufReader *bytes.Reader
		var size int64
		if !StyleStreamTestsShouldMakeRealFiles {
			bufReader = bytes.NewReader(buffer.Bytes())
			size = bufReader.Size()
		}
		actualSheetNames, actualWorkbookData := readXLSXFile(t, filePath, bufReader, size, StyleStreamTestsShouldMakeRealFiles)
		// check if data was able to be read correctly
		if !reflect.DeepEqual(actualSheetNames, testCase.sheetNames) {
			t.Fatal("Expected sheet names to be equal")
		}

		expectedWorkbookDataStrings := [][][]string{}
		for j, _ := range testCase.workbookData {
			expectedWorkbookDataStrings = append(expectedWorkbookDataStrings, [][]string{})
			for k, _ := range testCase.workbookData[j] {
				expectedWorkbookDataStrings[j] = append(expectedWorkbookDataStrings[j], []string{})
				for _, cell := range testCase.workbookData[j][k] {
					expectedWorkbookDataStrings[j][k] = append(expectedWorkbookDataStrings[j][k], cell.cellData)
				}
			}

		}
		if !reflect.DeepEqual(actualWorkbookData, expectedWorkbookDataStrings) {
			t.Fatal("Expected workbook data to be equal")
		}
	}
}

// writeStreamFile will write the file using this stream package
func writeStreamFileWithStyle(filePath string, fileBuffer io.Writer, sheetNames []string, workbookData [][][]StreamCell,
	shouldMakeRealFiles bool, customStyles []StreamStyle) error {
	var file *StreamFileBuilder
	var err error
	if shouldMakeRealFiles {
		file, err = NewStreamFileBuilderForPath(filePath)
		if err != nil {
			return err
		}
	} else {
		file = NewStreamFileBuilder(fileBuffer)
	}

	defaultStyles := []StreamStyle{Strings, BoldStrings, ItalicStrings, UnderlinedStrings,
		Integers, BoldIntegers, ItalicIntegers, UnderlinedIntegers,
		Dates}
	allStylesToBeAdded := append(defaultStyles, customStyles...)
	err = file.AddStreamStyleList(allStylesToBeAdded)
	if err != nil {
		return err
	}

	for i, sheetName := range sheetNames {
		colStyles := []StreamStyle{}
		for _, cell := range workbookData[i][0] {
			colStyles = append(colStyles, cell.cellStyle)
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
		for _, row := range sheetData {
			//if i == 0 {
			//	continue
			//}
			err = streamFile.WriteS(row)
			if err != nil {
				return err
			}
		}
	}
	err = streamFile.Close()
	if err != nil {
		return err
	}
	return nil
}

func (s *StreamSuite) TestDates(t *C) {
	var filePath string
	var buffer bytes.Buffer
	if StyleStreamTestsShouldMakeRealFiles {
		filePath = fmt.Sprintf("Workbook_Date_test.xlsx")
	}

	sheetNames := []string{"Sheet1"}
	workbookData := [][][]StreamCell{
		{
			{NewStringStreamCell("Date:")},
			{NewDateStreamCell(time.Now())},
		},
	}

	err := writeStreamFileWithStyle(filePath, &buffer, sheetNames, workbookData, StyleStreamTestsShouldMakeRealFiles, []StreamStyle{})
	if err != nil {
		t.Fatal("Error during writing")
	}
}

func (s *StreamSuite) TestMakeNewStylesAndUseIt(t *C) {
	var filePath string
	var buffer bytes.Buffer
	if StyleStreamTestsShouldMakeRealFiles {
		filePath = fmt.Sprintf("Workbook_newStyle.xlsx")
	}

	timesNewRoman12 := NewFont(12, TimesNewRoman)
	timesNewRoman12.Color = RGB_Dard_Green
	courier12 := NewFont(12, Courier)
	courier12.Color = RGB_Dark_Red

	greenFill := NewFill(Solid_Cell_Fill, RGB_Light_Green, RGB_White)
	redFill := NewFill(Solid_Cell_Fill, RGB_Light_Red, RGB_White)

	greenStyle := MakeStyle(0, timesNewRoman12, greenFill, DefaultAlignment(), DefaultBorder())
	redStyle := MakeStyle(0, courier12, redFill, DefaultAlignment(), DefaultBorder())

	sheetNames := []string{"Sheet1"}
	workbookData := [][][]StreamCell{
		{
			{NewStringStreamCell("Header1"), NewStringStreamCell("Header2")},
			{NewStyledStringStreamCell("Good", greenStyle), NewStyledStringStreamCell("Bad", redStyle)},
		},
	}

	err := writeStreamFileWithStyle(filePath, &buffer, sheetNames, workbookData, StyleStreamTestsShouldMakeRealFiles, []StreamStyle{greenStyle, redStyle})

	if err != nil {
		t.Fatal("Error during writing")
	}

	// read the file back with the xlsx package
	var bufReader *bytes.Reader
	var size int64
	if !StyleStreamTestsShouldMakeRealFiles {
		bufReader = bytes.NewReader(buffer.Bytes())
		size = bufReader.Size()
	}
	actualSheetNames, actualWorkbookData := readXLSXFile(t, filePath, bufReader, size, StyleStreamTestsShouldMakeRealFiles)
	// check if data was able to be read correctly
	if !reflect.DeepEqual(actualSheetNames, sheetNames) {
		t.Fatal("Expected sheet names to be equal")
	}

	expectedWorkbookDataStrings := [][][]string{}
	for j, _ := range workbookData {
		expectedWorkbookDataStrings = append(expectedWorkbookDataStrings, [][]string{})
		for k, _ := range workbookData[j] {
			expectedWorkbookDataStrings[j] = append(expectedWorkbookDataStrings[j], []string{})
			for _, cell := range workbookData[j][k] {
				expectedWorkbookDataStrings[j][k] = append(expectedWorkbookDataStrings[j][k], cell.cellData)
			}
		}

	}
	if !reflect.DeepEqual(actualWorkbookData, expectedWorkbookDataStrings) {
		t.Fatal("Expected workbook data to be equal")
	}
}

func (s *StreamSuite) TestCloseWithNothingWrittenToSheetsWithStyle(t *C) {
	buffer := bytes.NewBuffer(nil)
	file := NewStreamFileBuilder(buffer)

	sheetNames := []string{"Sheet1", "Sheet2"}
	workbookData := [][][]StreamCell{
		{{NewStringStreamCell("Header1"), NewStringStreamCell("Header2")}},
		{{NewStringStreamCell("Header3"), NewStringStreamCell("Header4")}},
	}

	defaultStyles := []StreamStyle{Strings, BoldStrings, ItalicIntegers, UnderlinedStrings,
		Integers, BoldIntegers, ItalicIntegers, UnderlinedIntegers,
		Dates}
	err := file.AddStreamStyleList(defaultStyles)
	if err != nil {
		t.Fatal(err)
	}

	colStyles0 := []StreamStyle{}
	for _, cell := range workbookData[0][0] {
		colStyles0 = append(colStyles0, cell.cellStyle)
	}

	colStyles1 := []StreamStyle{}
	for _, cell := range workbookData[1][0] {
		colStyles1 = append(colStyles1, cell.cellStyle)
	}

	err = file.AddSheetS(sheetNames[0], colStyles0)
	if err != nil {
		t.Fatal(err)
	}
	err = file.AddSheetS(sheetNames[1], colStyles1)
	if err != nil {
		t.Fatal(err)
	}

	stream, err := file.Build()
	if err != nil {
		t.Fatal(err)
	}
	err = stream.Close()
	if err != nil {
		t.Fatal(err)
	}
	bufReader := bytes.NewReader(buffer.Bytes())
	size := bufReader.Size()

	actualSheetNames, actualWorkbookData := readXLSXFile(t, "", bufReader, size, false)
	// check if data was able to be read correctly
	if !reflect.DeepEqual(actualSheetNames, sheetNames) {
		t.Fatal("Expected sheet names to be equal")
	}
	expectedWorkbookDataStrings := [][][]string{}
	for range workbookData {
		expectedWorkbookDataStrings = append(expectedWorkbookDataStrings, [][]string{})
	}
	if !reflect.DeepEqual(actualWorkbookData, expectedWorkbookDataStrings) {
		t.Fatal("Expected workbook data to be equal")
	}
}

func (s *StreamSuite) TestBuildErrorsAfterBuildWithStyle(t *C) {
	file := NewStreamFileBuilder(bytes.NewBuffer(nil))

	defaultStyles := []StreamStyle{Strings, BoldStrings, ItalicIntegers, UnderlinedStrings,
		Integers, BoldIntegers, ItalicIntegers, UnderlinedIntegers,
		Dates}
	err := file.AddStreamStyleList(defaultStyles)
	if err != nil {
		t.Fatal(err)
	}

	err = file.AddSheetS("Sheet1", []StreamStyle{Strings})
	if err != nil {
		t.Fatal(err)
	}
	err = file.AddSheetS("Sheet2", []StreamStyle{Strings})
	if err != nil {
		t.Fatal(err)
	}

	_, err = file.Build()
	if err != nil {
		t.Fatal(err)
	}
	_, err = file.Build()
	if err != BuiltStreamFileBuilderError {
		t.Fatal(err)
	}
}

func (s *StreamSuite) TestAddSheetWithStyleErrorsAfterBuild(t *C) {
	file := NewStreamFileBuilder(bytes.NewBuffer(nil))

	defaultStyles := []StreamStyle{Strings, BoldStrings, ItalicIntegers, UnderlinedStrings,
		Integers, BoldIntegers, ItalicIntegers, UnderlinedIntegers,
		Dates}
	err := file.AddStreamStyleList(defaultStyles)
	if err != nil {
		t.Fatal(err)
	}

	err = file.AddSheetS("Sheet1", []StreamStyle{Strings})
	if err != nil {
		t.Fatal(err)
	}
	err = file.AddSheetS("Sheet2", []StreamStyle{Strings})
	if err != nil {
		t.Fatal(err)
	}

	_, err = file.Build()
	if err != nil {
		t.Fatal(err)
	}
	err = file.AddSheetS("Sheet3", []StreamStyle{Strings})
	if err != BuiltStreamFileBuilderError {
		t.Fatal(err)
	}
}
