package xlsx

import (
	"fmt"
	"testing"
)

func TestDataValidation(t *testing.T) {
	var file *File
	var sheet *Sheet
	var row *Row
	var cell *Cell
	var err error
	var title string = "cell"
	var msg string = "cell msg"

	file = NewFile()
	sheet, err = file.AddSheet("Sheet1")
	if err != nil {
		fmt.Printf(err.Error())
	}
	row = sheet.AddRow()
	cell = row.AddCell()
	cell.Value = "b"

	dd := NewXlsxCellDataValidation(true, true, true)
	dd.SetDropList([]string{"a", "b", "b"})

	dd.SetInput(&title, &msg)
	cell.SetDataValidation(dd)

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetDropList([]string{"a", "b", "b"})
	title = "col b"
	dd.SetInput(&title, &msg)
	sheet.Col(2).SetDataValidation(dd, 0, 0)

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetDropList([]string{"a", "b", "b"})
	title = "col c range"
	dd.SetInput(&title, &msg)
	sheet.Col(3).SetDataValidation(dd, 3, 7)

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetDropList([]string{"a", "b", "b"})
	title = "col d start 3"
	dd.SetInput(&title, &msg)
	sheet.Col(4).SetDataValidationWithStart(dd, 1)

	if err != nil {
		fmt.Printf(err.Error())
	}

	file.Save("datavalidation.xlsx")

}
