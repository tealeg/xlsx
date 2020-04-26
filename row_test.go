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
		c.Assert(row.cellCount, qt.Equals, 1)
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
				c.Assert(cells, qt.HasLen, 4)
				output = append(output, cells)
				return nil
			})
			c.Assert(err, qt.Equals, nil)
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
}
