package xlsx

import (
	"bytes"
	"fmt"
	"testing"

	qt "github.com/frankban/quicktest"
)

func TestDataValidation(t *testing.T) {
	var file *File
	var sheet *Sheet
	var row *Row
	var cell *Cell
	var err error
	var title = "cell"
	var msg = "cell msg"

	c := qt.New(t)

	file = NewFile()
	sheet, err = file.AddSheet("Sheet1")
	if err != nil {
		fmt.Printf(err.Error())
	}
	row = sheet.AddRow()
	cell = row.AddCell()
	cell.Value = "a1"

	dd := NewXlsxCellDataValidation(true)
	err = dd.SetDropList([]string{"a1", "a2", "a3"})
	c.Assert(err, qt.IsNil)

	dd.SetInput(&title, &msg)
	cell.SetDataValidation(dd)

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetDropList([]string{"c1", "c2", "c3"})
	c.Assert(err, qt.IsNil)
	title = "col c"
	dd.SetInput(&title, &msg)
	sheet.SetDataValidation(2, 2, dd, 0, 0)

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetDropList([]string{"d", "d1", "d2"})
	c.Assert(err, qt.IsNil)
	title = "col d range"
	dd.SetInput(&title, &msg)
	sheet.SetDataValidation(3, 3, dd, 3, 7)

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetDropList([]string{"e1", "e2", "e3"})
	c.Assert(err, qt.IsNil)
	title = "col e start 3"
	dd.SetInput(&title, &msg)
	sheet.SetDataValidationWithStart(4, 4, dd, 1)

	index := 5
	rowIndex := 1
	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(15, 4, DataValidationTypeTextLeng, DataValidationOperatorBetween)
	c.Assert(err, qt.IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorEqual)
	c.Assert(err, qt.IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorGreaterThanOrEqual)
	c.Assert(err, qt.IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorGreaterThan)
	c.Assert(err, qt.IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorLessThan)
	c.Assert(err, qt.IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorLessThanOrEqual)
	c.Assert(err, qt.IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorNotEqual)
	c.Assert(err, qt.IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorNotBetween)
	c.Assert(err, qt.IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	rowIndex++
	index = 5

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(4, 15, DataValidationTypeWhole, DataValidationOperatorBetween)
	c.Assert(err, qt.IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorEqual)
	c.Assert(err, qt.IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorGreaterThanOrEqual)
	c.Assert(err, qt.IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorGreaterThan)
	c.Assert(err, qt.IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorLessThan)
	c.Assert(err, qt.IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorLessThanOrEqual)
	c.Assert(err, qt.IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorNotEqual)
	c.Assert(err, qt.IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 50, DataValidationTypeWhole, DataValidationOperatorNotBetween)
	if err != nil {
		t.Fatal(err)
	}
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetDropList([]string{"1", "2", "4"})
	c.Assert(err, qt.IsNil)
	dd1 := NewXlsxCellDataValidation(true)
	err = dd1.SetDropList([]string{"11", "22", "44"})
	c.Assert(err, qt.IsNil)
	dd2 := NewXlsxCellDataValidation(true)
	err = dd2.SetDropList([]string{"111", "222", "444"})
	c.Assert(err, qt.IsNil)
	sheet.SetDataValidation(12, 12, dd, 2, 10)
	sheet.SetDataValidation(12, 12, dd1, 3, 4)
	sheet.SetDataValidation(12, 12, dd2, 5, 7)

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetDropList([]string{"1", "2", "4"})
	c.Assert(err, qt.IsNil)
	dd1 = NewXlsxCellDataValidation(true)
	err = dd1.SetDropList([]string{"11", "22", "44"})
	c.Assert(err, qt.IsNil)
	sheet.SetDataValidation(13, 13, dd, 2, 10)
	sheet.SetDataValidation(13, 13, dd1, 1, 2)

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetDropList([]string{"1", "2", "4"})
	c.Assert(err, qt.IsNil)
	dd1 = NewXlsxCellDataValidation(true)
	err = dd1.SetDropList([]string{"11", "22", "44"})
	c.Assert(err, qt.IsNil)
	sheet.SetDataValidation(14, 14, dd, 2, 10)
	sheet.SetDataValidation(14, 14, dd1, 1, 5)

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetDropList([]string{"1", "2", "4"})
	if err != nil {
		t.Fatal(err)
	}
	dd1 = NewXlsxCellDataValidation(true)
	err = dd1.SetDropList([]string{"11", "22", "44"})
	c.Assert(err, qt.IsNil)
	sheet.SetDataValidation(15, 15, dd, 2, 10)
	sheet.SetDataValidation(15, 15, dd1, 1, 10)

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetDropList([]string{"1", "2", "4"})
	c.Assert(err, qt.IsNil)
	dd1 = NewXlsxCellDataValidation(true)
	err = dd1.SetDropList([]string{"11", "22", "44"})
	c.Assert(err, qt.IsNil)
	dd2 = NewXlsxCellDataValidation(true)
	err = dd2.SetDropList([]string{"111", "222", "444"})
	c.Assert(err, qt.IsNil)
	sheet.SetDataValidation(16, 16, dd, 10, 20)
	sheet.SetDataValidation(16, 16, dd1, 2, 4)
	sheet.SetDataValidation(16, 16, dd2, 21, 30)

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetDropList([]string{"d", "d1", "d2"})
	c.Assert(err, qt.IsNil)
	title = "col d range"
	dd.SetInput(&title, &msg)
	sheet.SetDataValidation(3, 3, dd, 3, Excel2006MaxRowIndex)

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetDropList([]string{"d", "d1", "d2"})
	c.Assert(err, qt.IsNil)
	title = "col d range"
	dd.SetInput(&title, &msg)
	sheet.SetDataValidation(3, 3, dd, 4, -1)
	maxRow := sheet.Col(3).DataValidation[len(sheet.Col(3).DataValidation)-1].maxRow
	c.Assert(maxRow, qt.Equals, Excel2006MaxRowIndex)

	dest := &bytes.Buffer{}
	err = file.Write(dest)
	c.Assert(err, qt.IsNil)
	// Read and write the file that was just saved.
	file, err = OpenBinary(dest.Bytes())
	c.Assert(err, qt.IsNil)
	dest = &bytes.Buffer{}
	err = file.Write(dest)
	c.Assert(err, qt.IsNil)
}

func TestDataValidation2(t *testing.T) {
	c := qt.New(t)
	// Show error and show info start disabled, but automatically get enabled when setting a message
	dd := NewXlsxCellDataValidation(true)
	c.Assert(dd.ShowErrorMessage, qt.Equals, false)
	c.Assert(dd.ShowInputMessage, qt.Equals, false)

	str := "you got an error"
	dd.SetError(StyleStop, &str, &str)
	c.Assert(dd.ShowErrorMessage, qt.Equals, true)
	c.Assert(dd.ShowInputMessage, qt.Equals, false)

	str = "hello"
	dd.SetInput(&str, &str)
	c.Assert(dd.ShowInputMessage, qt.Equals, true)

	// Check the formula created by this function
	// The sheet name needs single quotes, the single quote in the name gets escaped,
	// and all references are fixed.
	err := dd.SetInFileList("Sheet ' 2", 2, 1, 3, 10)
	c.Assert(err, qt.IsNil)
	expectedFormula := "'Sheet '' 2'!$C$2:$D$11"
	c.Assert(dd.Formula1, qt.Equals, expectedFormula)
	c.Assert(dd.Type, qt.Equals, "list")
}
