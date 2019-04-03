package xlsx

import (
	"bytes"
	"fmt"

	. "gopkg.in/check.v1"
)

type DataValidationSuite struct{}

var _ = Suite(&DataValidationSuite{})

func (d *DataValidationSuite) TestDataValidation(t *C) {
	var file *File
	var sheet *Sheet
	var row *Row
	var cell *Cell
	var err error
	var title = "cell"
	var msg = "cell msg"

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
	t.Assert(err, IsNil)

	dd.SetInput(&title, &msg)
	cell.SetDataValidation(dd)

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetDropList([]string{"c1", "c2", "c3"})
	t.Assert(err, IsNil)
	title = "col c"
	dd.SetInput(&title, &msg)
	sheet.Col(2).SetDataValidation(dd, 0, 0)

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetDropList([]string{"d", "d1", "d2"})
	t.Assert(err, IsNil)
	title = "col d range"
	dd.SetInput(&title, &msg)
	sheet.Col(3).SetDataValidation(dd, 3, 7)

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetDropList([]string{"e1", "e2", "e3"})
	t.Assert(err, IsNil)
	title = "col e start 3"
	dd.SetInput(&title, &msg)
	sheet.Col(4).SetDataValidationWithStart(dd, 1)

	index := 5
	rowIndex := 1
	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(15, 4, DataValidationTypeTextLeng, DataValidationOperatorBetween)
	t.Assert(err, IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorEqual)
	t.Assert(err, IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorGreaterThanOrEqual)
	t.Assert(err, IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorGreaterThan)
	t.Assert(err, IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorLessThan)
	t.Assert(err, IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorLessThanOrEqual)
	t.Assert(err, IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorNotEqual)
	t.Assert(err, IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorNotBetween)
	t.Assert(err, IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	rowIndex++
	index = 5

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(4, 15, DataValidationTypeWhole, DataValidationOperatorBetween)
	t.Assert(err, IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorEqual)
	t.Assert(err, IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorGreaterThanOrEqual)
	t.Assert(err, IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorGreaterThan)
	t.Assert(err, IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorLessThan)
	t.Assert(err, IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorLessThanOrEqual)
	t.Assert(err, IsNil)
	sheet.Cell(rowIndex, index).SetDataValidation(dd)
	index++

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorNotEqual)
	t.Assert(err, IsNil)
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
	t.Assert(err, IsNil)
	dd1 := NewXlsxCellDataValidation(true)
	err = dd1.SetDropList([]string{"11", "22", "44"})
	t.Assert(err, IsNil)
	dd2 := NewXlsxCellDataValidation(true)
	err = dd2.SetDropList([]string{"111", "222", "444"})
	t.Assert(err, IsNil)
	sheet.Col(12).SetDataValidation(dd, 2, 10)
	sheet.Col(12).SetDataValidation(dd1, 3, 4)
	sheet.Col(12).SetDataValidation(dd2, 5, 7)

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetDropList([]string{"1", "2", "4"})
	t.Assert(err, IsNil)
	dd1 = NewXlsxCellDataValidation(true)
	err = dd1.SetDropList([]string{"11", "22", "44"})
	t.Assert(err, IsNil)
	sheet.Col(13).SetDataValidation(dd, 2, 10)
	sheet.Col(13).SetDataValidation(dd1, 1, 2)

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetDropList([]string{"1", "2", "4"})
	t.Assert(err, IsNil)
	dd1 = NewXlsxCellDataValidation(true)
	err = dd1.SetDropList([]string{"11", "22", "44"})
	t.Assert(err, IsNil)
	sheet.Col(14).SetDataValidation(dd, 2, 10)
	sheet.Col(14).SetDataValidation(dd1, 1, 5)

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetDropList([]string{"1", "2", "4"})
	if err != nil {
		t.Fatal(err)
	}
	dd1 = NewXlsxCellDataValidation(true)
	err = dd1.SetDropList([]string{"11", "22", "44"})
	t.Assert(err, IsNil)
	sheet.Col(15).SetDataValidation(dd, 2, 10)
	sheet.Col(15).SetDataValidation(dd1, 1, 10)

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetDropList([]string{"1", "2", "4"})
	t.Assert(err, IsNil)
	dd1 = NewXlsxCellDataValidation(true)
	err = dd1.SetDropList([]string{"11", "22", "44"})
	t.Assert(err, IsNil)
	dd2 = NewXlsxCellDataValidation(true)
	err = dd2.SetDropList([]string{"111", "222", "444"})
	t.Assert(err, IsNil)
	sheet.Col(16).SetDataValidation(dd, 10, 20)
	sheet.Col(16).SetDataValidation(dd1, 2, 4)
	sheet.Col(16).SetDataValidation(dd2, 21, 30)

	dd = NewXlsxCellDataValidation(true)
	err = dd.SetDropList([]string{"d", "d1", "d2"})
	t.Assert(err, IsNil)
	title = "col d range"
	dd.SetInput(&title, &msg)
	sheet.Col(3).SetDataValidation(dd, 3, Excel2006MaxRowCount)

	dest := &bytes.Buffer{}
	err = file.Write(dest)
	t.Assert(err, IsNil)
	// Read and write the file that was just saved.
	file, err = OpenBinary(dest.Bytes())
	t.Assert(err, IsNil)
	dest = &bytes.Buffer{}
	err = file.Write(dest)
	t.Assert(err, IsNil)
}

func (d *DataValidationSuite) TestDataValidation2(t *C) {
	// Show error and show info start disabled, but automatically get enabled when setting a message
	dd := NewXlsxCellDataValidation(true)
	t.Assert(dd.ShowErrorMessage, Equals, false)
	t.Assert(dd.ShowInputMessage, Equals, false)

	str := "you got an error"
	dd.SetError(StyleStop, &str, &str)
	t.Assert(dd.ShowErrorMessage, Equals, true)
	t.Assert(dd.ShowInputMessage, Equals, false)

	str = "hello"
	dd.SetInput(&str, &str)
	t.Assert(dd.ShowInputMessage, Equals, true)

	// Check the formula created by this function
	// The sheet name needs single quotes, the single quote in the name gets escaped,
	// and all references are fixed.
	err := dd.SetInFileList("Sheet ' 2", 2, 1, 3, 10)
	t.Assert(err, IsNil)
	expectedFormula := "'Sheet '' 2'!$C$2:$D$11"
	t.Assert(dd.Formula1, Equals, expectedFormula)
	t.Assert(dd.Type, Equals, "list")
}
