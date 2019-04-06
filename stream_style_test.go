package xlsx

import (
	"bytes"
	"fmt"
	. "gopkg.in/check.v1"
	"io"
	"reflect"
)

const (
	StreamTestsShouldMakeRealFiles = false
)

type StreamStyleSuite struct{}

var _ = Suite(&StreamStyleSuite{})

func (s *StreamSuite) TestStreamTestsShouldMakeRealFilesShouldBeFalse(t *C) {
	if StreamTestsShouldMakeRealFiles {
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
			testName: "Number Row",
			sheetNames: []string{
				"Sheet1",
			},
			workbookData: [][][]StreamCell{
				{
					{MakeStringStreamCell("1"), MakeStringStreamCell("25"),
						MakeStringStreamCell("A"), MakeStringStreamCell("B")},
					{MakeIntegerStreamCell(1234), MakeStyledIntegerStreamCell(98, BoldIntegers),
						MakeStyledIntegerStreamCell(34, ItalicIntegers), MakeStyledIntegerStreamCell(26, UnderlinedIntegers)},
				},
			},
		},
	}

	for i, testCase := range testCases {
		var filePath string
		var buffer bytes.Buffer
		if StreamTestsShouldMakeRealFiles {
			filePath = fmt.Sprintf("WorkbookWithStyle%d.xlsx", i)
		}

		err := writeStreamFileWithStyle(filePath, &buffer, testCase.sheetNames, testCase.workbookData, StreamTestsShouldMakeRealFiles)
		if err != testCase.expectedError && err.Error() != testCase.expectedError.Error() {
			t.Fatalf("Error differs from expected error. Error: %v, Expected Error: %v ", err, testCase.expectedError)
		}
		if testCase.expectedError != nil {
			return
		}
		// read the file back with the xlsx package
		var bufReader *bytes.Reader
		var size int64
		if !StreamTestsShouldMakeRealFiles {
			bufReader = bytes.NewReader(buffer.Bytes())
			size = bufReader.Size()
		}
		actualSheetNames, actualWorkbookData := readXLSXFile(t, filePath, bufReader, size, StreamTestsShouldMakeRealFiles)
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
func writeStreamFileWithStyle(filePath string, fileBuffer io.Writer, sheetNames []string, workbookData [][][]StreamCell, shouldMakeRealFiles bool) error {
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

	err = file.AddStreamStyle(Strings)
	err = file.AddStreamStyle(BoldStrings)
	err = file.AddStreamStyle(ItalicStrings)
	err = file.AddStreamStyle(UnderlinedStrings)
	err = file.AddStreamStyle(Integers)
	err = file.AddStreamStyle(BoldIntegers)
	err = file.AddStreamStyle(ItalicIntegers)
	err = file.AddStreamStyle(UnderlinedIntegers)
	if err != nil { // TODO handle all errors not just one
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













