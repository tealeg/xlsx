package xlsx

import (
	"bytes"
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

	csRunO(c, "DataValidation", func(c *qt.C, option FileOption) {
		file = NewFile(option)
		sheet, err = file.AddSheet("Sheet1")
		c.Assert(err, qt.Equals, nil)
		row = sheet.AddRow()
		cell = row.AddCell()
		cell.Value = "a1"

		dd := NewDataValidation(0, 0, 0, 0, true)
		err = dd.SetDropList([]string{"a1", "a2", "a3"})
		c.Assert(err, qt.IsNil)

		dd.SetInput(&title, &msg)
		cell.SetDataValidation(dd)

		dd = NewDataValidation(2, 0, 2, 0, true)
		err = dd.SetDropList([]string{"c1", "c2", "c3"})
		c.Assert(err, qt.IsNil)
		title = "col c"
		dd.SetInput(&title, &msg)
		sheet.AddDataValidation(dd)

		dd = NewDataValidation(3, 3, 3, 7, true)
		err = dd.SetDropList([]string{"d", "d1", "d2"})
		c.Assert(err, qt.IsNil)
		title = "col d range"
		dd.SetInput(&title, &msg)
		sheet.AddDataValidation(dd)

		dd = NewDataValidation(4, 1, 4, Excel2006MaxRowIndex, true)
		err = dd.SetDropList([]string{"e1", "e2", "e3"})
		c.Assert(err, qt.IsNil)
		title = "col e start 3"
		dd.SetInput(&title, &msg)
		sheet.AddDataValidation(dd)

		index := 5
		rowIndex := 1
		dd = NewDataValidation(rowIndex, index, rowIndex, index, true)
		err = dd.SetRange(15, 4, DataValidationTypeTextLeng, DataValidationOperatorBetween)
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		index++

		dd = NewDataValidation(rowIndex, index, rowIndex, index, true)
		err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorEqual)
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		index++

		dd = NewDataValidation(rowIndex, index, rowIndex, index, true)
		err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorGreaterThanOrEqual)
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		index++

		dd = NewDataValidation(rowIndex, index, rowIndex, index, true)
		err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorGreaterThan)
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		index++

		dd = NewDataValidation(rowIndex, index, rowIndex, index, true)
		err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorLessThan)
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		index++

		dd = NewDataValidation(rowIndex, index, rowIndex, index, true)
		err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorLessThanOrEqual)
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		index++

		dd = NewDataValidation(rowIndex, index, rowIndex, index, true)
		err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorNotEqual)
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		index++

		dd = NewDataValidation(rowIndex, index, rowIndex, index, true)
		err = dd.SetRange(10, 1, DataValidationTypeTextLeng, DataValidationOperatorNotBetween)
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		index++

		rowIndex++
		index = 5

		dd = NewDataValidation(rowIndex, index, rowIndex, index, true)
		err = dd.SetRange(4, 15, DataValidationTypeWhole, DataValidationOperatorBetween)
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		index++

		dd = NewDataValidation(rowIndex, index, rowIndex, index, true)
		err = dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorEqual)
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		index++

		dd = NewDataValidation(rowIndex, index, rowIndex, index, true)
		err = dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorGreaterThanOrEqual)
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		index++

		dd = NewDataValidation(rowIndex, index, rowIndex, index, true)
		err = dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorGreaterThan)
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		index++

		dd = NewDataValidation(rowIndex, index, rowIndex, index, true)
		err = dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorLessThan)
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		index++

		dd = NewDataValidation(rowIndex, index, rowIndex, index, true)
		err = dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorLessThanOrEqual)
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		index++

		dd = NewDataValidation(rowIndex, index, rowIndex, index, true)
		err = dd.SetRange(10, 1, DataValidationTypeWhole, DataValidationOperatorNotEqual)
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		index++

		dd = NewDataValidation(rowIndex, index, rowIndex, index, true)
		err = dd.SetRange(10, 50, DataValidationTypeWhole, DataValidationOperatorNotBetween)
		if err != nil {
			t.Fatal(err)
		}
		sheet.AddDataValidation(dd)
		index++

		dd = NewDataValidation(12, 2, 12, 10, true)
		err = dd.SetDropList([]string{"1", "2", "4"})
		c.Assert(err, qt.IsNil)
		dd1 := NewDataValidation(12, 3, 12, 4, true)
		err = dd1.SetDropList([]string{"11", "22", "44"})
		c.Assert(err, qt.IsNil)
		dd2 := NewDataValidation(12, 5, 12, 7, true)
		err = dd2.SetDropList([]string{"111", "222", "444"})
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		sheet.AddDataValidation(dd1)
		sheet.AddDataValidation(dd2)

		dd = NewDataValidation(13, 2, 13, 10, true)
		err = dd.SetDropList([]string{"1", "2", "4"})
		c.Assert(err, qt.IsNil)
		dd1 = NewDataValidation(13, 1, 13, 2, true)
		err = dd1.SetDropList([]string{"11", "22", "44"})
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		sheet.AddDataValidation(dd1)

		dd = NewDataValidation(14, 2, 14, 10, true)
		err = dd.SetDropList([]string{"1", "2", "4"})
		c.Assert(err, qt.IsNil)
		dd1 = NewDataValidation(14, 1, 14, 5, true)
		err = dd1.SetDropList([]string{"11", "22", "44"})
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		sheet.AddDataValidation(dd1)

		dd = NewDataValidation(15, 2, 15, 10, true)
		err = dd.SetDropList([]string{"1", "2", "4"})
		if err != nil {
			t.Fatal(err)
		}
		dd1 = NewDataValidation(15, 1, 15, 10, true)
		err = dd1.SetDropList([]string{"11", "22", "44"})
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		sheet.AddDataValidation(dd1)

		dd = NewDataValidation(16, 10, 16, 20, true)
		err = dd.SetDropList([]string{"1", "2", "4"})
		c.Assert(err, qt.IsNil)
		dd1 = NewDataValidation(16, 2, 16, 4, true)
		err = dd1.SetDropList([]string{"11", "22", "44"})
		c.Assert(err, qt.IsNil)
		dd2 = NewDataValidation(16, 12, 16, 30, true)
		err = dd2.SetDropList([]string{"111", "222", "444"})
		c.Assert(err, qt.IsNil)
		sheet.AddDataValidation(dd)
		sheet.AddDataValidation(dd1)
		sheet.AddDataValidation(dd2)

		dd = NewDataValidation(3, 3, 3, Excel2006MaxRowIndex, true)
		err = dd.SetDropList([]string{"d", "d1", "d2"})
		c.Assert(err, qt.IsNil)
		title = "col d range"
		dd.SetInput(&title, &msg)
		sheet.AddDataValidation(dd)

		dd = NewDataValidation(3, 4, 3, Excel2006MaxRowIndex, true)
		err = dd.SetDropList([]string{"d", "d1", "d2"})
		c.Assert(err, qt.IsNil)
		title = "col d range"
		dd.SetInput(&title, &msg)
		sheet.AddDataValidation(dd)

		dest := &bytes.Buffer{}
		err = file.Write(dest)
		c.Assert(err, qt.IsNil)
		// Read and write the file that was just saved.
		file, err = OpenBinary(dest.Bytes())
		c.Assert(err, qt.IsNil)
		dest = &bytes.Buffer{}
		err = file.Write(dest)
		c.Assert(err, qt.IsNil)
	})

	c.Run("DataValidation2", func(c *qt.C) {
		// Show error and show info start disabled, but automatically get enabled when setting a message
		dd := NewDataValidation(0, 0, 0, 0, true)
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
	})
}
