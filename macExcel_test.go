package xlsx

import (
	"testing"
)

func TestMacExcel(t *testing.T) {
	xlsxFile, error := OpenFile("macExcelTest.xlsx")
	if error != nil {
		t.Error(error.Error())
		return
	}
	if xlsxFile == nil {
		t.Error("OpenFile returned nil FileInterface without generating an os.Error")
		return
	}
	s := xlsxFile.Sheets[0].Cell(0, 0).String()
	if s != "编号" {
		t.Errorf("[TestMacExcel] xlsxFile.Sheets[0].Cell(0,0).String():'%s'", s)
		return
	}
}
