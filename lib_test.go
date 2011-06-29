package xlsx


import (
	"fmt"
	"os"
	"testing"
)


func TestOpenFile(t *testing.T) {
	var xlsxFile *File
	var error os.Error
	xlsxFile, error = OpenFile("testfile.xlsx")
	if error != nil {
		t.Error(error.String())
		return
	}
	if xlsxFile == nil {
		t.Error("OpenFile returned nil FileInterface without generating an os.Error")
		return
	}

}

func TestExtractSheets(t *testing.T) {
	var xlsxFile *File
	var sheets []*Sheet
	xlsxFile, _ = OpenFile("testfile.xlsx")
	sheets = xlsxFile.Sheets
	if len(sheets) == 0 {
		t.Error("No sheets read from XLSX file")
		return
	}
	fmt.Printf("%v\n", len(sheets))
}

func TestCreateSheet(t *testing.T) {
	var xlsxFile *File
	var error os.Error
	var sheet *Sheet
	xlsxFile, error = OpenFile("testfile.xlsx")
	if error != nil {
		t.Error(error.String())
		return
	}
	if xlsxFile == nil {
		t.Error("OpenFile returned a nil File pointer but did not generate an error.")
		return
	}
	if len(xlsxFile.Sheets) == 0 {
		t.Error("Expected len(xlsxFile.Sheets) > 0")
		return
	}
	sheet = xlsxFile.Sheets[0]
	if len(sheet.Cells) == 0 {
		t.Error("Expected len(sheet.Cells) == 4")
	}
}

