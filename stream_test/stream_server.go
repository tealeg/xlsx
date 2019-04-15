package main

import (
	. "github.com/damianszkuat/xlsx"
	"io"
	"net/http"
	"strconv"
	"math/rand"
)

func StreamFileWithDate(w http.ResponseWriter, r *http.Request) {
	w.Header().Add("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	sheetNames, workbookData := generateExcelFile(1, 1000000, 10)
	writeStreamFileWithStyle(w, sheetNames, workbookData, []StreamStyle{})
}

// writeStreamFile will write the file using this stream package
func writeStreamFileWithStyle(fileBuffer io.Writer, sheetNames []string,
	workbookData [][][]StreamCell, customStyles []StreamStyle) error {

	var file *StreamFileBuilder
	var err error

	file = NewStreamFileBuilder(fileBuffer)

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
		for _, row := range sheetData {
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

func generateExcelFile(numOfSheets int, numOfRows int, numOfCols int) ([]string, [][][]StreamCell) {
	var sheetNames []string
	var workbookData [][][]StreamCell
	for i := 0; i<numOfSheets; i++{
		sheetNames = append(sheetNames, strconv.Itoa(i))
		workbookData = append(workbookData, [][]StreamCell{})
		for j := 0; j<numOfRows; j++ {
			workbookData[i] = append(workbookData[i], []StreamCell{})
			for k := 0; k<numOfCols; k++ {
				var style StreamStyle

				if k%2==0 {
					style = StreamStyleDefaultInteger
				} else if k%3 == 0 {
					style = StreamStyleBoldInteger
				} else if k%5 == 0 {
					style = StreamStyleItalicInteger
				} else {
					style = StreamStyleUnderlinedInteger
				}

				workbookData[i][j] = append(workbookData[i][j], NewStyledIntegerStreamCell(rand.Intn(100),style))
			}
		}
	}

	return sheetNames, workbookData
}

func main() {
	http.HandleFunc("/", StreamFileWithDate)
	if err := http.ListenAndServe(":8080", nil); err != nil {
		panic(err)
	}
}
