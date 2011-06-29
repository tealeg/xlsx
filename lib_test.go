package xlsx


import (
	"fmt"
	"os"
	"testing"
)


// Test we can correctly open a XSLX file and return a xlsx.File
// struct.
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

// Test that when we open a real XLSX file we create xlsx.Sheet
// objects for the sheets inside the file and that these sheets are
// themselves correct.
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

