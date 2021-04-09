package xlsx

import (
	"testing"

	qt "github.com/frankban/quicktest"
)

func TestRow(t *testing.T) {
	c := qt.New(t)
	// Test we can add a new Cell to a Row
	csRunO(c, "TestAddCell", func(c *qt.C, option FileOption) {
		var f *File
		f = NewFile(option)
		sheet, _ := f.AddSheet("MySheet")
		row := sheet.AddRow()
		cell := row.AddCell()
		c.Assert(cell, qt.Not(qt.IsNil))
		c.Assert(row.Sheet.MaxCol, qt.Equals, 1)
		c.Assert(row.cellStoreRow.CellCount(), qt.Equals, 1)
	})

	csRunO(c, "TestGetCell", func(c *qt.C, option FileOption) {
		var f *File
		f = NewFile(option)
		sheet, _ := f.AddSheet("MySheet")
		row := sheet.AddRow()
		cell := row.AddCell()
		cell.SetValue("foo")
		cell1 := row.AddCell()
		cell1.SetValue("bar")

		cell2 := row.GetCell(0)
		c.Assert(cell.Value, qt.Equals, cell2.Value)
	})

	csRunO(c, "TestForEachCell", func(c *qt.C, option FileOption) {
		var f *File
		f, err := OpenFile("./testdocs/empty_cells.xlsx", option)
		c.Assert(err, qt.Equals, nil)
		sheet := f.Sheets[0]
		c.Run("NoOptions", func(c *qt.C) {
			output := [][]string{}
			err := sheet.ForEachRow(func(r *Row) error {
				cells := []string{}
				err := r.ForEachCell(func(c *Cell) error {
					cells = append(cells, c.Value)
					return nil
				})
				if err != nil {
					return err
				}
				output = append(output, cells)
				return nil
			})
			c.Assert(err, qt.Equals, nil)
			c.Assert(output, qt.DeepEquals, [][]string{
				{"", "B1", "C1", "D1"},
				{"A2", "", "C2", "D2"},
				{"A3", "B3", "", "D3"},
				{"A4", "B4", "C4", ""},
			})
		})

		c.Run("SkipEmptyCells", func(c *qt.C) {
			output := [][]string{}
			err := sheet.ForEachRow(func(r *Row) error {
				cells := []string{}
				err := r.ForEachCell(func(c *Cell) error {
					cells = append(cells, c.Value)
					return nil
				}, SkipEmptyCells)
				if err != nil {
					return err
				}
				output = append(output, cells)
				return nil
			})
			c.Assert(err, qt.Equals, nil)
			c.Assert(output, qt.DeepEquals,
				[][]string{
					{"B1", "C1", "D1"},
					{"A2", "C2", "D2"},
					{"A3", "B3", "D3"},
					{"A4", "B4", "C4"},
				})
		})

	})

	csRunO(c, "Test Set Height", func(c *qt.C, option FileOption) {
		var f *File
		f = NewFile(option)
		sheet, _ := f.AddSheet("MySheet")
		row := sheet.AddRow()
		c.Assert(row.height, qt.Equals, 0.0)
		c.Assert(row.GetHeight(), qt.Equals, 0.0)
		c.Assert(row.customHeight, qt.IsFalse)
		c.Assert(row.isCustom, qt.IsFalse)

		var heightToSet float64 = 30.0
		row.SetHeight(heightToSet)
		c.Assert(row.GetHeight(), qt.Equals, heightToSet)
		c.Assert(row.customHeight, qt.IsTrue)
		c.Assert(row.isCustom, qt.IsTrue)

		row.SetHeightCM(heightToSet)
		c.Assert(row.GetHeight(), qt.Equals, heightToSet*cmToPs)
		c.Assert(row.customHeight, qt.IsTrue)
		c.Assert(row.isCustom, qt.IsTrue)

	})
}
