package xlsx

import (
	"testing"

	qt "github.com/frankban/quicktest"
)

func TestMemoryCellStore(t *testing.T) {
	c := qt.New(t)

	c.Run("RowNotFoundError", func(c *qt.C) {
		memoryCs, err := NewMemoryCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := memoryCs.(*MemoryCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		row, err := cs.ReadRow("I don't exist")
		c.Assert(err, qt.Not(qt.IsNil))
		c.Assert(row, qt.IsNil)
		_, ok = err.(*RowNotFoundError)
		c.Assert(ok, qt.Equals, true)
	})

	c.Run("Write and Read Row", func(c *qt.C) {
		mCs, err := NewMemoryCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := mCs.(*MemoryCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		s := &Style{
			Border: Border{
				Left:        "left",
				LeftColor:   "leftColor",
				Right:       "right",
				RightColor:  "rightColor",
				Top:         "top",
				TopColor:    "topColor",
				Bottom:      "bottom",
				BottomColor: "bottomColor",
			},
			Fill: Fill{
				PatternType: "PatternType",
				BgColor:     "BgColor",
				FgColor:     "FgColor",
			},
			Font: Font{
				Size:      1,
				Name:      "Font",
				Family:    2,
				Charset:   3,
				Color:     "Red",
				Bold:      true,
				Italic:    true,
				Underline: true,
			},
			Alignment: Alignment{
				Horizontal:   "left",
				Indent:       1,
				ShrinkToFit:  true,
				TextRotation: 90,
				Vertical:     "top",
				WrapText:     true,
			},
			ApplyBorder:    true,
			ApplyFill:      true,
			ApplyFont:      true,
			ApplyAlignment: true,
		}

		dv := &xlsxDataValidation{
			AllowBlank:       true,
			ShowInputMessage: true,
			ShowErrorMessage: true,
			Type:             "type",
			Sqref:            "sqref",
			Formula1:         "formula1",
			Formula2:         "formula1",
			Operator:         "operator",
		}

		rt := []RichTextRun{
			RichTextRun{
				Font: &RichTextFont{Bold: true},
				Text: "bold",
			},
			RichTextRun{
				Text: "normal",
			},
		}

		file := NewFile()
		sheet, _ := file.AddSheet("Test")
		row := sheet.AddRow()
		cell := row.AddCell()

		cell.Value = "value"
		cell.RichText = rt
		cell.formula = "formula"
		cell.style = s
		cell.NumFmt = "numFmt"
		cell.date1904 = true
		cell.Hidden = true
		cell.HMerge = 49
		cell.VMerge = 50
		cell.cellType = CellType(2)
		cell.DataValidation = dv
		cell.Hyperlink = Hyperlink{
			DisplayString: "displaystring",
			Link:          "link",
			Tooltip:       "tooltip",
		}

		err = cs.WriteRow(row)
		c.Assert(err, qt.IsNil)
		row2, err := cs.ReadRow(row.key())
		c.Assert(err, qt.IsNil)

		c.Assert(row2, qt.Not(qt.IsNil))
		c.Assert(row.Hidden, qt.Equals, row2.Hidden)
		c.Assert(row.GetHeight(), qt.Equals, row2.GetHeight())
		c.Assert(row.GetOutlineLevel(), qt.Equals, row2.GetOutlineLevel())
		c.Assert(row.isCustom, qt.Equals, row2.isCustom)
		c.Assert(row.num, qt.Equals, row2.num)
		c.Assert(row.cellCount, qt.Equals, row2.cellCount)

		cell2 := row.GetCell(0)
		c.Assert(err, qt.IsNil)

		c.Assert(cell.Value, qt.Equals, cell2.Value)
		c.Assert(cell.RichText, qt.DeepEquals, cell2.RichText)
		c.Assert(cell.formula, qt.Equals, cell2.formula)
		c.Assert(cell.NumFmt, qt.Equals, cell2.NumFmt)
		c.Assert(cell.date1904, qt.Equals, cell2.date1904)
		c.Assert(cell.Hidden, qt.Equals, cell2.Hidden)
		c.Assert(cell.HMerge, qt.Equals, cell2.HMerge)
		c.Assert(cell.VMerge, qt.Equals, cell2.VMerge)
		c.Assert(cell.cellType, qt.Equals, cell2.cellType)
		c.Assert(*cell.DataValidation, qt.DeepEquals, *cell2.DataValidation)
		c.Assert(cell.Hyperlink, qt.DeepEquals, cell2.Hyperlink)
		c.Assert(cell.num, qt.Equals, cell2.num)

		s2 := cell2.style
		c.Assert(s2.Border, qt.DeepEquals, s.Border)
		c.Assert(s2.Fill, qt.DeepEquals, s.Fill)
		c.Assert(s2.Font, qt.DeepEquals, s.Font)
		c.Assert(s2.Alignment, qt.DeepEquals, s.Alignment)
		c.Assert(s2.ApplyBorder, qt.Equals, s.ApplyBorder)
		c.Assert(s2.ApplyFill, qt.Equals, s.ApplyFill)
		c.Assert(s2.ApplyFont, qt.Equals, s.ApplyFont)
		c.Assert(s2.ApplyAlignment, qt.Equals, s.ApplyAlignment)

	})

}
