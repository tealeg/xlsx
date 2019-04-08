package xlsx

import (
	"bytes"
	"fmt"
	"io"
	"reflect"
	"strings"

	. "gopkg.in/check.v1"
)

const (
	TestsShouldMakeRealFiles = false
)

type StreamSuite struct{}

var _ = Suite(&StreamSuite{})

func (s *StreamSuite) TestTestsShouldMakeRealFilesShouldBeFalse(t *C) {
	if TestsShouldMakeRealFiles {
		t.Fatal("TestsShouldMakeRealFiles should only be true for local debugging. Don't forget to switch back before commiting.")
	}
}

func (s *StreamSuite) TestXlsxStreamWrite(t *C) {
	// When shouldMakeRealFiles is set to true this test will make actual XLSX files in the file system.
	// This is useful to ensure files open in Excel, Numbers, Google Docs, etc.
	// In case of issues you can use "Open XML SDK 2.5" to diagnose issues in generated XLSX files:
	// https://www.microsoft.com/en-us/download/details.aspx?id=30425
	testCases := []struct {
		testName      string
		sheetNames    []string
		workbookData  [][][]string
		headerTypes   [][]*CellType
		expectedError error
	}{
		{
			testName: "One Sheet",
			sheetNames: []string{
				"Sheet1",
			},
			workbookData: [][][]string{
				{
					{"Token", "Name", "Price", "SKU"},
					{"123", "Taco", "300", "0000000123"},
				},
			},
			headerTypes: [][]*CellType{
				{nil, CellTypeString.Ptr(), nil, CellTypeString.Ptr()},
			},
		},
		{
			testName: "One Column",
			sheetNames: []string{
				"Sheet1",
			},
			workbookData: [][][]string{
				{
					{"Token"},
					{"123"},
				},
			},
		},
		{
			testName: "Several Sheets, with different numbers of columns and rows",
			sheetNames: []string{
				"Sheet 1", "Sheet 2", "Sheet3",
			},
			workbookData: [][][]string{
				{
					{"Token", "Name", "Price", "SKU"},
					{"123", "Taco", "300", "0000000123"},
				},
				{
					{"Token", "Name", "Price", "SKU", "Stock"},
					{"456", "Salsa", "200", "0346", "1"},
					{"789", "Burritos", "400", "754", "3"},
				},
				{
					{"Token", "Name", "Price"},
					{"9853", "Guacamole", "500"},
					{"2357", "Margarita", "700"},
				},
			},
		},
		{
			testName: "Two Sheets with same the name",
			sheetNames: []string{
				"Sheet 1", "Sheet 1",
			},
			workbookData: [][][]string{
				{
					{"Token", "Name", "Price", "SKU"},
					{"123", "Taco", "300", "0000000123"},
				},
				{
					{"Token", "Name", "Price", "SKU", "Stock"},
					{"456", "Salsa", "200", "0346", "1"},
					{"789", "Burritos", "400", "754", "3"},
				},
			},
			expectedError: fmt.Errorf("duplicate sheet name '%s'.", "Sheet 1"),
		},
		{
			testName: "One Sheet Registered, tries to write to two",
			sheetNames: []string{
				"Sheet 1",
			},
			workbookData: [][][]string{
				{
					{"Token", "Name", "Price", "SKU"},
					{"123", "Taco", "300", "0000000123"},
				},
				{
					{"Token", "Name", "Price", "SKU"},
					{"456", "Salsa", "200", "0346"},
				},
			},
			expectedError: AlreadyOnLastSheetError,
		},
		{
			testName: "One Sheet, too many columns in row 1",
			sheetNames: []string{
				"Sheet 1",
			},
			workbookData: [][][]string{
				{
					{"Token", "Name", "Price", "SKU"},
					{"123", "Taco", "300", "0000000123", "asdf"},
				},
			},
			expectedError: WrongNumberOfRowsError,
		},
		{
			testName: "One Sheet, too few columns in row 1",
			sheetNames: []string{
				"Sheet 1",
			},
			workbookData: [][][]string{
				{
					{"Token", "Name", "Price", "SKU"},
					{"123", "Taco", "300"},
				},
			},
			expectedError: WrongNumberOfRowsError,
		},
		{
			testName: "Lots of Sheets, only writes rows to one, only writes headers to one, should not error and should still create a valid file",
			sheetNames: []string{
				"Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4", "Sheet 5", "Sheet 6",
			},
			workbookData: [][][]string{
				{
					{"Token", "Name", "Price", "SKU"},
					{"123", "Taco", "300", "0000000123"},
				},
				{{}},
				{{"Id", "Unit Cost"}},
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
			workbookData: [][][]string{
				{
					{"Token", "Name", "Price", "SKU"},
					{"123", "Taco", "300", "0000000123"},
				},
				{{}},
			},
		},
		{
			testName: "Larger Sheet",
			sheetNames: []string{
				"Sheet 1",
			},
			workbookData: [][][]string{
				{
					{"Token", "Name", "Price", "SKU", "Token", "Name", "Price", "SKU", "Token", "Name", "Price", "SKU", "Token", "Name", "Price", "SKU", "Token", "Name", "Price", "SKU", "Token", "Name", "Price", "SKU"},
					{"123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123"},
					{"456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346"},
					{"789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754"},
					{"123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123"},
					{"456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346"},
					{"789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754"},
					{"123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123"},
					{"456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346"},
					{"789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754"},
					{"123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123"},
					{"456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346"},
					{"789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754"},
					{"123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123"},
					{"456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346"},
					{"789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754"},
					{"123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123"},
					{"456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346"},
					{"789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754"},
					{"123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123"},
					{"456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346"},
					{"789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754"},
					{"123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123", "123", "Taco", "300", "0000000123"},
					{"456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346", "456", "Salsa", "200", "0346"},
					{"789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754", "789", "Burritos", "400", "754"},
				},
			},
		},
		{
			testName: "UTF-8 Characters. This XLSX File loads correctly with Excel, Numbers, and Google Docs. It also passes Microsoft's Office File Format Validator.",
			sheetNames: []string{
				"Sheet1",
			},
			workbookData: [][][]string{
				{
					// String courtesy of https://github.com/minimaxir/big-list-of-naughty-strings/
					// Header row contains the tags that I am filtering on
					{"Token", endSheetDataTag, "Price", fmt.Sprintf(dimensionTag, "A1:D1")},
					// Japanese and emojis
					{"123", "パーティーへ行かないか", "300", "🍕🐵 🙈 🙉 🙊"},
					// XML encoder/parser test strings
					{"123", `<?xml version="1.0" encoding="ISO-8859-1"?>`, "300", `<?xml version="1.0" encoding="ISO-8859-1"?><!DOCTYPE foo [ <!ELEMENT foo ANY ><!ENTITY xxe SYSTEM "file:///etc/passwd" >]><foo>&xxe;</foo>`},
					// Upside down text and Right to Left Arabic text
					{"123", `˙ɐnbᴉlɐ ɐuƃɐɯ ǝɹolop ʇǝ ǝɹoqɐl ʇn ʇunpᴉpᴉɔuᴉ ɹodɯǝʇ poɯsnᴉǝ op pǝs 'ʇᴉlǝ ƃuᴉɔsᴉdᴉpɐ ɹnʇǝʇɔǝsuoɔ 'ʇǝɯɐ ʇᴉs ɹolop ɯnsdᴉ ɯǝɹo˥
					00˙Ɩ$-`, "300", `ﷺ`},
					{"123", "Taco", "300", "0000000123"},
				},
			},
		},
	}
	for i, testCase := range testCases {
		var filePath string
		var buffer bytes.Buffer
		if TestsShouldMakeRealFiles {
			filePath = fmt.Sprintf("Workbook%d.xlsx", i)
		}
		err := writeStreamFile(filePath, &buffer, testCase.sheetNames, testCase.workbookData, testCase.headerTypes, TestsShouldMakeRealFiles)
		if err != testCase.expectedError && err.Error() != testCase.expectedError.Error() {
			t.Fatalf("Error differs from expected error. Error: %v, Expected Error: %v ", err, testCase.expectedError)
		}
		if testCase.expectedError != nil {
			return
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
		if !reflect.DeepEqual(actualSheetNames, testCase.sheetNames) {
			t.Fatal("Expected sheet names to be equal")
		}
		if !reflect.DeepEqual(actualWorkbookData, testCase.workbookData) {
			t.Fatal("Expected workbook data to be equal")
		}
	}
}

// The purpose of TestXlsxStyleBehavior is to ensure that initMaxStyleId has the correct starting value
// and that the logic in AddSheet() that predicts Style IDs is correct.
func (s *StreamSuite) TestXlsxStyleBehavior(t *C) {
	file := NewFile()
	sheet, err := file.AddSheet("Sheet 1")
	if err != nil {
		t.Fatal(err)
	}
	row := sheet.AddRow()
	rowData := []string{"testing", "1", "2", "3"}
	if count := row.WriteSlice(&rowData, -1); count != len(rowData) {
		t.Fatal("not enough cells written")
	}
	parts, err := file.MarshallParts()
	styleSheet, ok := parts["xl/styles.xml"]
	if !ok {
		t.Fatal("no style sheet")
	}
	// Created an XLSX file with only the default style.
	// We expect that the number of styles is one more than our max index constant.
	// This means the library adds two styles by default.
	if !strings.Contains(styleSheet, fmt.Sprintf(`<cellXfs count="%d">`, initMaxStyleId+1)) {
		t.Fatal("Expected sheet to have two styles")
	}

	file = NewFile()
	sheet, err = file.AddSheet("Sheet 1")
	if err != nil {
		t.Fatal(err)
	}
	row = sheet.AddRow()
	rowData = []string{"testing", "1", "2", "3", "4"}
	if count := row.WriteSlice(&rowData, -1); count != len(rowData) {
		t.Fatal("not enough cells written")
	}
	sheet.Cols[0].SetType(CellTypeString)
	sheet.Cols[1].SetType(CellTypeString)
	sheet.Cols[3].SetType(CellTypeNumeric)
	sheet.Cols[4].SetType(CellTypeString)
	parts, err = file.MarshallParts()
	styleSheet, ok = parts["xl/styles.xml"]
	if !ok {
		t.Fatal("no style sheet")
	}
	// Created an XLSX file with two distinct cell types, which should create two new styles.
	// The same cell type was added three times, this should be coalesced into the same style rather than
	// recreating the style. This XLSX stream library depends on this behavior when predicting the next style id.
	if !strings.Contains(styleSheet, fmt.Sprintf(`<cellXfs count="%d">`, initMaxStyleId+1+2)) {
		t.Fatal("Expected sheet to have four styles")
	}
}

// writeStreamFile will write the file using this stream package
func writeStreamFile(filePath string, fileBuffer io.Writer, sheetNames []string, workbookData [][][]string, headerTypes [][]*CellType, shouldMakeRealFiles bool) error {
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
	for i, sheetName := range sheetNames {
		header := workbookData[i][0]
		var sheetHeaderTypes []*CellType
		if i < len(headerTypes) {
			sheetHeaderTypes = headerTypes[i]
		}
		err := file.AddSheet(sheetName, header, sheetHeaderTypes)
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
			err = streamFile.Write(row)
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

// readXLSXFile will read the file using the xlsx package.
func readXLSXFile(t *C, filePath string, fileBuffer io.ReaderAt, size int64, shouldMakeRealFiles bool) ([]string, [][][]string) {
	var readFile *File
	var err error
	if shouldMakeRealFiles {
		readFile, err = OpenFile(filePath)
		if err != nil {
			t.Fatal(err)
		}
	} else {
		readFile, err = OpenReaderAt(fileBuffer, size)
		if err != nil {
			t.Fatal(err)
		}
	}
	var actualWorkbookData [][][]string
	var sheetNames []string
	for _, sheet := range readFile.Sheets {
		sheetData := [][]string{}
		for _, row := range sheet.Rows {
			data := []string{}
			for _, cell := range row.Cells {
				str, err := cell.FormattedValue()
				if err != nil {
					t.Fatal(err)
				}
				data = append(data, str)
			}
			sheetData = append(sheetData, data)
		}
		sheetNames = append(sheetNames, sheet.Name)
		actualWorkbookData = append(actualWorkbookData, sheetData)
	}
	return sheetNames, actualWorkbookData
}

func (s *StreamSuite) TestAddSheetErrorsAfterBuild(t *C) {
	file := NewStreamFileBuilder(bytes.NewBuffer(nil))

	err := file.AddSheet("Sheet1", []string{"Header"}, nil)
	if err != nil {
		t.Fatal(err)
	}
	err = file.AddSheet("Sheet2", []string{"Header2"}, nil)
	if err != nil {
		t.Fatal(err)
	}

	_, err = file.Build()
	if err != nil {
		t.Fatal(err)
	}
	err = file.AddSheet("Sheet3", []string{"Header3"}, nil)
	if err != BuiltStreamFileBuilderError {
		t.Fatal(err)
	}
}

func (s *StreamSuite) TestBuildErrorsAfterBuild(t *C) {
	file := NewStreamFileBuilder(bytes.NewBuffer(nil))

	err := file.AddSheet("Sheet1", []string{"Header"}, nil)
	if err != nil {
		t.Fatal(err)
	}
	err = file.AddSheet("Sheet2", []string{"Header2"}, nil)
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

func (s *StreamSuite) TestCloseWithNothingWrittenToSheets(t *C) {
	buffer := bytes.NewBuffer(nil)
	file := NewStreamFileBuilder(buffer)

	sheetNames := []string{"Sheet1", "Sheet2"}
	workbookData := [][][]string{
		{{"Header1", "Header2"}},
		{{"Header3", "Header4"}},
	}
	err := file.AddSheet(sheetNames[0], workbookData[0][0], nil)
	if err != nil {
		t.Fatal(err)
	}
	err = file.AddSheet(sheetNames[1], workbookData[1][0], nil)
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
	if !reflect.DeepEqual(actualWorkbookData, workbookData) {
		t.Fatal("Expected workbook data to be equal")
	}
}
