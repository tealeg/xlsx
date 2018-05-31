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
	cell.Value = "a1"

	dd := NewXlsxCellDataValidation(true, true, true)
	dd.SetDropList([]string{"a1", "a2", "a3"})

	dd.SetInput(&title, &msg)
	cell.SetDataValidation(dd)

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetDropList([]string{"c1", "c2", "c3"})
	title = "col c"
	dd.SetInput(&title, &msg)
	sheet.Col(2).SetDataValidation(dd, 0, 0)

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetDropList([]string{"d", "d1", "d2"})
	title = "col d range"
	dd.SetInput(&title, &msg)
	sheet.Col(3).SetDataValidation(dd, 3, 7)

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetDropList([]string{"e1", "e2", "e3"})
	title = "col e start 3"
	dd.SetInput(&title, &msg)
	sheet.Col(4).SetDataValidationWithStart(dd, 1)

	index := 5
	rowIndex := 1
	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetRange(15, 4, DataValidationTypeTextLeng, DataValidationOperatorBetween)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorEqual)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorGreaterThanOrEqual)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorGreaterThan)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorLessThan)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorLessThanOrEqual)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorNotEqual)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorNotBetween)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	rowIndex++
	index = 5

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetRange(4, 15, DataValidationTypeWhole, DataValidationOperatorBetween)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorEqual)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorGreaterThanOrEqual)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorGreaterThan)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorLessThan)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorLessThanOrEqual)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorNotEqual)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true, true, true)
	dd.SetRange(10, 50, DataValidationTypeWhole, DataValidationOperatorNotBetween)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	if err != nil {
		fmt.Printf(err.Error())
	}

	file.Save("datavalidation.xlsx")

}

func TestReadDataValidation(t *testing.T) {
	file, err := OpenFile("datavalidation.xlsx")
	if nil != err {
		t.Errorf(err.Error())
		return
	}
	err = file.Save("datavalidation_read.xlsx")
	if nil != err {
		t.Errorf(err.Error())
		return
	}
}
