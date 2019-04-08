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
					{MakeStyledStringStreamCell("1", UnderlinedStrings), MakeStyledStringStreamCell("25", ItalicStrings),
						MakeStyledStringStreamCell("A", BoldStrings), MakeStringStreamCell("B")},
					{MakeIntegerStreamCell(1234), MakeStyledIntegerStreamCell(98, BoldIntegers),
						MakeStyledIntegerStreamCell(34, ItalicIntegers), MakeStyledIntegerStreamCell(26, UnderlinedIntegers)},
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
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU")},
					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123)},
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
					{MakeStringStreamCell("Token")},
					{MakeIntegerStreamCell(123)},
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
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU")},

					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123)},
				},
				{
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU"),
						MakeStringStreamCell("Stock")},

					{MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(1)},

					{MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(3)},
				},
				{
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price")},

					{MakeIntegerStreamCell(9853), MakeStringStreamCell("Guacamole"),
						MakeIntegerStreamCell(500)},

					{MakeIntegerStreamCell(2357), MakeStringStreamCell("Margarita"),
						MakeIntegerStreamCell(700)},
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
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU")},

					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123)},
				},
				{
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU"),
						MakeStringStreamCell("Stock")},

					{MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(1)},

					{MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(3)},
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
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU")},

					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123)},
				},
				{
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU")},

					{MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346)},
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
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU")},

					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeStringStreamCell("asdf")},
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
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU")},

					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300)},
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
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU")},

					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123)},
				},
				{{}},
				{{MakeStringStreamCell("Id"), MakeStringStreamCell("Unit Cost")}},
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
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU")},

					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123)},
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
					{MakeStringStreamCell("Token"), MakeStringStreamCell(endSheetDataTag),
						MakeStringStreamCell("Price"), MakeStringStreamCell(fmt.Sprintf(dimensionTag, "A1:D1"))},
					// Japanese and emojis
					{MakeIntegerStreamCell(123), MakeStringStreamCell("パーティーへ行かないか"),
						MakeIntegerStreamCell(300), MakeStringStreamCell("🍕🐵 🙈 🙉 🙊")},
					// XML encoder/parser test strings
					{MakeIntegerStreamCell(123), MakeStringStreamCell(`<?xml version="1.0" encoding="ISO-8859-1"?>`),
						MakeIntegerStreamCell(300), MakeStringStreamCell(`<?xml version="1.0" encoding="ISO-8859-1"?><!DOCTYPE foo [ <!ELEMENT foo ANY ><!ENTITY xxe SYSTEM "file:///etc/passwd" >]><foo>&xxe;</foo>`)},
					// Upside down text and Right to Left Arabic text
					{MakeIntegerStreamCell(123), MakeStringStreamCell(`˙ɐnbᴉlɐ ɐuƃɐɯ ǝɹolop ʇǝ ǝɹoqɐl ʇn ʇunpᴉpᴉɔuᴉ ɹodɯǝʇ poɯsnᴉǝ op pǝs 'ʇᴉlǝ ƃuᴉɔsᴉdᴉpɐ ɹnʇǝʇɔǝsuoɔ 'ʇǝɯɐ ʇᴉs ɹolop ɯnsdᴉ ɯǝɹo˥
					00˙Ɩ$-`), MakeIntegerStreamCell(300), MakeStringStreamCell(`ﷺ`)} ,
					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123)},
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
		for j,_ := range testCase.workbookData {
			expectedWorkbookDataStrings = append(expectedWorkbookDataStrings, [][]string{})
			for k,_ := range testCase.workbookData[j]{
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
		header := workbookData[i][0]
		err := file.AddSheetWithStyle(sheetName, header)
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
		for i, row := range sheetData {
			if i == 0 {
				continue
			}
			err = streamFile.WriteWithStyle(row)
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
			{MakeStringStreamCell("Date:")},
			{MakeDateStreamCell(time.Now())},
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
			{MakeStringStreamCell("Header1"), MakeStringStreamCell("Header2")},
			{MakeStyledStringStreamCell("Good", greenStyle), MakeStyledStringStreamCell("Bad", redStyle)},
		},
	}

	err := writeStreamFileWithStyle(filePath, &buffer, sheetNames, workbookData, StyleStreamTestsShouldMakeRealFiles, []StreamStyle{greenStyle, redStyle})

	if err != nil {
		t.Fatal("Error during writing")
	}

	// read the file back with the xlsx package
	var bufReader *bytes.Reader
	var size int64
	if !TestsShouldMakeRealFiles {
		bufReader = bytes.NewReader(buffer.Bytes())
		size = bufReader.Size()
	}
	actualSheetNames, actualWorkbookData := readXLSXFile(t, filePath, bufReader, size, TestsShouldMakeRealFiles)
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
		{{MakeStringStreamCell("Header1"), MakeStringStreamCell("Header2")}},
		{{MakeStringStreamCell("Header3"), MakeStringStreamCell("Header4")}},
	}

	defaultStyles := []StreamStyle{Strings,BoldStrings,ItalicIntegers,UnderlinedStrings,
		Integers, BoldIntegers, ItalicIntegers, UnderlinedIntegers,
		Dates}
	err := file.AddStreamStyleList(defaultStyles)
	if err != nil {
		t.Fatal(err)
	}

	err = file.AddSheetWithStyle(sheetNames[0], workbookData[0][0])
	if err != nil {
		t.Fatal(err)
	}
	err = file.AddSheetWithStyle(sheetNames[1], workbookData[1][0])
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
	for j,_ := range workbookData {
		expectedWorkbookDataStrings = append(expectedWorkbookDataStrings, [][]string{})
		for k,_ := range workbookData[j]{
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

func (s *StreamSuite) TestBuildErrorsAfterBuildWithStyle(t *C) {
	file := NewStreamFileBuilder(bytes.NewBuffer(nil))

	defaultStyles := []StreamStyle{Strings,BoldStrings,ItalicIntegers,UnderlinedStrings,
		Integers, BoldIntegers, ItalicIntegers, UnderlinedIntegers,
		Dates}
	err := file.AddStreamStyleList(defaultStyles)
	if err != nil {
		t.Fatal(err)
	}

	err = file.AddSheetWithStyle("Sheet1", []StreamCell{MakeStringStreamCell("Header")})
	if err != nil {
		t.Fatal(err)
	}
	err = file.AddSheetWithStyle("Sheet2", []StreamCell{MakeStringStreamCell("Header")})
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

	defaultStyles := []StreamStyle{Strings,BoldStrings,ItalicIntegers,UnderlinedStrings,
		Integers, BoldIntegers, ItalicIntegers, UnderlinedIntegers,
		Dates}
	err := file.AddStreamStyleList(defaultStyles)
	if err != nil {
		t.Fatal(err)
	}

	err = file.AddSheetWithStyle("Sheet1", []StreamCell{MakeStringStreamCell("Header")})
	if err != nil {
		t.Fatal(err)
	}
	err = file.AddSheetWithStyle("Sheet2", []StreamCell{MakeStringStreamCell("Header2")})
	if err != nil {
		t.Fatal(err)
	}

	_, err = file.Build()
	if err != nil {
		t.Fatal(err)
	}
	err = file.AddSheetWithStyle("Sheet3", []StreamCell{MakeStringStreamCell("Header3")})
	if err != BuiltStreamFileBuilderError {
		t.Fatal(err)
	}
}





