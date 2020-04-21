package xlsx

import (
	"bytes"
	"math"
	"testing"

	qt "github.com/frankban/quicktest"
)

func TestDiskVCellStore(t *testing.T) {
	c := qt.New(t)

	c.Run("RowNotFoundError", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		_, err = cs.ReadRow("I don't exist")
		c.Assert(err, qt.Not(qt.IsNil))
		_, ok = err.(*RowNotFoundError)
		c.Assert(ok, qt.Equals, true)
	})

	c.Run("Write and Read Empty Row", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		file := NewFile(UseDiskVCellStore)
		sheet, _ := file.AddSheet("Test")
		row := sheet.AddRow()

		row.Hidden = true
		row.SetHeight(40.4)
		row.SetOutlineLevel(2)
		row.isCustom = true
		row.num = 3
		row.cellCount = 0

		err = cs.WriteRow(row)
		c.Assert(err, qt.IsNil)
		row2, err := cs.ReadRow(row.key())
		c.Assert(err, qt.IsNil)
		c.Assert(row2, qt.Not(qt.IsNil))
		c.Assert(row.Hidden, qt.Equals, row2.Hidden)
		// We shouldn't have a sheet set here
		c.Assert(row2.Sheet, qt.IsNil)
		c.Assert(row.GetHeight(), qt.Equals, row2.GetHeight())
		c.Assert(row.GetOutlineLevel(), qt.Equals, row2.GetOutlineLevel())
		c.Assert(row.isCustom, qt.Equals, row2.isCustom)
		c.Assert(row.num, qt.Equals, row2.num)
		c.Assert(row.cellCount, qt.Equals, row2.cellCount)
	})

	c.Run("Write and Read Row with Cells", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		file := NewFile(UseDiskVCellStore)
		sheet, _ := file.AddSheet("Test")
		row := sheet.AddRow()

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

		dv.ErrorStyle = sPtr("errorstyle")
		dv.ErrorTitle = sPtr("errortitle")
		dv.Error = sPtr("error")
		dv.PromptTitle = sPtr("prompttitle")
		dv.Prompt = sPtr("prompt")

		cell := row.AddCell()
		cell.Value = "value"
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

		cell2 := row2.GetCell(0)

		c.Assert(cell.Value, qt.Equals, cell2.Value)
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

	c.Run("Write and Read Bool", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		cs.writeBool(true)
		cs.writeBool(false)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		v, err := cs.readBool()
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.Equals, true)
		v, err = cs.readBool()
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.Equals, false)
		v, err = cs.readBool()
		c.Assert(err, qt.Not(qt.IsNil))
	})

	c.Run("Write and Read unit separator", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		cs.writeUnitSeparator()
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		err = cs.readUnitSeparator()
		c.Assert(err, qt.IsNil)
		err = cs.readUnitSeparator()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read String", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		cs.writeString("simple")
		cs.writeString(`multi
line!`)
		cs.writeString("")
		cs.writeString("Scheiß encoding")
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		v, err := cs.readString()
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.Equals, "simple")
		v, err = cs.readString()
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.Equals, `multi
line!`)
		v, err = cs.readString()
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.Equals, "")
		v, err = cs.readString()
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.Equals, "Scheiß encoding")
		v, err = cs.readString()
		c.Assert(err, qt.Not(qt.IsNil))
	})

	c.Run("Write and Read Int", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		cs.writeInt(math.MinInt64)
		cs.writeInt(0)
		cs.writeInt(math.MaxInt64)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		v, err := cs.readInt()
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.Equals, math.MinInt64)
		v, err = cs.readInt()
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.Equals, 0)
		v, err = cs.readInt()
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.Equals, math.MaxInt64)
		v, err = cs.readInt()
		c.Assert(err, qt.Not(qt.IsNil))
	})

	c.Run("Write and Read String Pointer", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		s := "foo"
		cs.writeStringPointer(nil)
		cs.writeStringPointer(&s)
		s = "bar"
		cs.writeStringPointer(&s)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		v, err := cs.readStringPointer()
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.IsNil)
		v, err = cs.readStringPointer()
		c.Assert(err, qt.IsNil)
		c.Assert(*v, qt.Equals, "foo")
		v, err = cs.readStringPointer()
		c.Assert(err, qt.IsNil)
		c.Assert(*v, qt.Equals, "bar")
		v, err = cs.readStringPointer()
		c.Assert(err, qt.Not(qt.IsNil))
	})

	c.Run("Write and Read end of record", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		cs.writeEndOfRecord()
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		err = cs.readEndOfRecord()
		c.Assert(err, qt.IsNil)
		err = cs.readEndOfRecord()
		c.Assert(err, qt.Not(qt.IsNil))
	})

	c.Run("Write and Read Border", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		b := Border{
			Left:        "left",
			LeftColor:   "leftColor",
			Right:       "right",
			RightColor:  "rightColor",
			Top:         "top",
			TopColor:    "topColor",
			Bottom:      "bottom",
			BottomColor: "bottomColor",
		}
		cs.writeBorder(b)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		b2, err := cs.readBorder()
		c.Assert(err, qt.IsNil)
		c.Assert(b2, qt.DeepEquals, b)
		_, err = cs.readBorder()
		c.Assert(err, qt.Not(qt.IsNil))
	})

	c.Run("Write and Read Fill", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		b := Fill{
			PatternType: "PatternType",
			BgColor:     "BgColor",
			FgColor:     "FgColor",
		}
		cs.writeFill(b)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		b2, err := cs.readFill()
		c.Assert(err, qt.IsNil)
		c.Assert(b2, qt.DeepEquals, b)
		_, err = cs.readFill()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read Font", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		b := Font{
			Size:      1,
			Name:      "Font",
			Family:    2,
			Charset:   3,
			Color:     "Red",
			Bold:      true,
			Italic:    true,
			Underline: true,
		}
		cs.writeFont(b)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		b2, err := cs.readFont()
		c.Assert(err, qt.IsNil)
		c.Assert(b2, qt.DeepEquals, b)
		_, err = cs.readFont()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read Alignment", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		b := Alignment{
			Horizontal:   "left",
			Indent:       1,
			ShrinkToFit:  true,
			TextRotation: 90,
			Vertical:     "top",
			WrapText:     true,
		}
		cs.writeAlignment(b)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		b2, err := cs.readAlignment()
		c.Assert(err, qt.IsNil)
		c.Assert(b2, qt.DeepEquals, b2)
		_, err = cs.readAlignment()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read Style", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		s := Style{
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
		err = cs.writeStyle(&s)
		c.Assert(err, qt.IsNil)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		s2, err := cs.readStyle()
		c.Assert(err, qt.IsNil)
		// We can't just DeepEquals style because we can't
		// compare the nil pointer in the NamedStyle field.
		c.Assert(s2.Border, qt.DeepEquals, s.Border)
		c.Assert(s2.Fill, qt.DeepEquals, s.Fill)
		c.Assert(s2.Font, qt.DeepEquals, s.Font)
		c.Assert(s2.Alignment, qt.DeepEquals, s.Alignment)
		c.Assert(s2.ApplyBorder, qt.Equals, s.ApplyBorder)
		c.Assert(s2.ApplyFill, qt.Equals, s.ApplyFill)
		c.Assert(s2.ApplyFont, qt.Equals, s.ApplyFont)
		c.Assert(s2.ApplyAlignment, qt.Equals, s.ApplyAlignment)
		_, err = cs.readStyle()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read DataValidation", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

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

		dv.ErrorStyle = sPtr("errorstyle")
		dv.ErrorTitle = sPtr("errortitle")
		dv.Error = sPtr("error")
		dv.PromptTitle = sPtr("prompttitle")
		dv.Prompt = sPtr("prompt")

		cs.writeDataValidation(dv)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		dv2, err := cs.readDataValidation()
		c.Assert(err, qt.IsNil)
		c.Assert(dv2, qt.DeepEquals, dv)
		_, err = cs.readDataValidation()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read Cell", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		cell := &Cell{
			Value:          "value",
			formula:        "formula",
			style:          nil,
			NumFmt:         "numFmt",
			date1904:       true,
			Hidden:         true,
			HMerge:         49,
			VMerge:         50,
			cellType:       CellType(2),
			DataValidation: nil,
			Hyperlink: Hyperlink{
				DisplayString: "displaystring",
				Link:          "link",
				Tooltip:       "tooltip",
			},
			num: 1,
		}

		cs.writeCell(cell)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		cell2, err := cs.readCell()
		c.Assert(err, qt.IsNil)
		c.Assert(cell.Value, qt.Equals, cell2.Value)
		c.Assert(cell.RichText, qt.HasLen, 0)
		c.Assert(cell.formula, qt.Equals, cell2.formula)
		c.Assert(cell.style, qt.Equals, cell2.style)
		c.Assert(cell.NumFmt, qt.Equals, cell2.NumFmt)
		c.Assert(cell.date1904, qt.Equals, cell2.date1904)
		c.Assert(cell.Hidden, qt.Equals, cell2.Hidden)
		c.Assert(cell.HMerge, qt.Equals, cell2.HMerge)
		c.Assert(cell.VMerge, qt.Equals, cell2.VMerge)
		c.Assert(cell.cellType, qt.Equals, cell2.cellType)
		c.Assert(cell.DataValidation, qt.Equals, cell2.DataValidation)
		c.Assert(cell.Hyperlink, qt.DeepEquals, cell2.Hyperlink)
		c.Assert(cell.num, qt.Equals, cell2.num)
		_, err = cs.readCell()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read Cell with style", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		s := Style{
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

		cell := &Cell{
			Value:          "value",
			formula:        "formula",
			style:          &s,
			NumFmt:         "numFmt",
			date1904:       true,
			Hidden:         true,
			HMerge:         49,
			VMerge:         50,
			cellType:       CellType(2),
			DataValidation: nil,
			Hyperlink: Hyperlink{
				DisplayString: "displaystring",
				Link:          "link",
				Tooltip:       "tooltip",
			},
			num: 1,
		}

		cs.writeCell(cell)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		cell2, err := cs.readCell()
		c.Assert(err, qt.IsNil)
		c.Assert(cell.Value, qt.Equals, cell2.Value)
		c.Assert(cell.RichText, qt.HasLen, 0)
		c.Assert(cell.formula, qt.Equals, cell2.formula)
		c.Assert(cell.NumFmt, qt.Equals, cell2.NumFmt)
		c.Assert(cell.date1904, qt.Equals, cell2.date1904)
		c.Assert(cell.Hidden, qt.Equals, cell2.Hidden)
		c.Assert(cell.HMerge, qt.Equals, cell2.HMerge)
		c.Assert(cell.VMerge, qt.Equals, cell2.VMerge)
		c.Assert(cell.cellType, qt.Equals, cell2.cellType)
		c.Assert(cell.DataValidation, qt.Equals, cell2.DataValidation)
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

		_, err = cs.readCell()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read Cell with DataValidation", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

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
		sPtr := func(s string) *string {
			return &s
		}

		dv.ErrorStyle = sPtr("errorstyle")
		dv.ErrorTitle = sPtr("errortitle")
		dv.Error = sPtr("error")
		dv.PromptTitle = sPtr("prompttitle")
		dv.Prompt = sPtr("prompt")

		cell := &Cell{
			Value:          "value",
			formula:        "formula",
			style:          nil,
			NumFmt:         "numFmt",
			date1904:       true,
			Hidden:         true,
			HMerge:         49,
			VMerge:         50,
			cellType:       CellType(2),
			DataValidation: dv,
			Hyperlink: Hyperlink{
				DisplayString: "displaystring",
				Link:          "link",
				Tooltip:       "tooltip",
			},
			num: 1,
		}

		cs.writeCell(cell)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		cell2, err := cs.readCell()
		c.Assert(err, qt.IsNil)
		c.Assert(cell.Value, qt.Equals, cell2.Value)
		c.Assert(cell.RichText, qt.HasLen, 0)
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
		c.Assert(cell.style, qt.Equals, cell2.style)

		_, err = cs.readCell()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read Cell with RichText", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		cell := &Cell{
			RichText: []RichTextRun{
				RichTextRun{
					Font: &RichTextFont{Bold: true},
					Text: "rich text",
				},
			},
			formula:        "formula",
			style:          nil,
			NumFmt:         "numFmt",
			date1904:       true,
			Hidden:         true,
			HMerge:         49,
			VMerge:         50,
			cellType:       CellType(2),
			DataValidation: nil,
			Hyperlink: Hyperlink{
				DisplayString: "displaystring",
				Link:          "link",
				Tooltip:       "tooltip",
			},
			num: 1,
		}

		cs.writeCell(cell)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		cell2, err := cs.readCell()
		c.Assert(err, qt.IsNil)
		c.Assert(cell.Value, qt.Equals, cell2.Value)
		c.Assert(cell.RichText, qt.DeepEquals, cell2.RichText)
		c.Assert(cell.formula, qt.Equals, cell2.formula)
		c.Assert(cell.style, qt.Equals, cell2.style)
		c.Assert(cell.NumFmt, qt.Equals, cell2.NumFmt)
		c.Assert(cell.date1904, qt.Equals, cell2.date1904)
		c.Assert(cell.Hidden, qt.Equals, cell2.Hidden)
		c.Assert(cell.HMerge, qt.Equals, cell2.HMerge)
		c.Assert(cell.VMerge, qt.Equals, cell2.VMerge)
		c.Assert(cell.cellType, qt.Equals, cell2.cellType)
		c.Assert(cell.DataValidation, qt.Equals, cell2.DataValidation)
		c.Assert(cell.Hyperlink, qt.DeepEquals, cell2.Hyperlink)
		c.Assert(cell.num, qt.Equals, cell2.num)
		_, err = cs.readCell()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read RichTextColor RGB", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		c1 := RichTextColor{
			coreColor: xlsxColor{
				RGB:  "01234567",
				Tint: -0.3,
			},
		}

		err = cs.writeRichTextColor(&c1)
		c.Assert(err, qt.IsNil)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		c2, err := cs.readRichTextColor()
		c.Assert(err, qt.IsNil)
		c.Assert(c2.coreColor.RGB, qt.Equals, c1.coreColor.RGB)
		c.Assert(c2.coreColor.Tint, qt.Equals, c1.coreColor.Tint)
		c.Assert(c2.coreColor.Indexed, qt.Equals, c1.coreColor.Indexed)
		c.Assert(c2.coreColor.Theme, qt.Equals, c1.coreColor.Theme)
		_, err = cs.readBool()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read RichTextColor Indexed", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		indexed := 7

		c1 := RichTextColor{
			coreColor: xlsxColor{
				Indexed: &indexed,
				Tint:    0.4,
			},
		}

		err = cs.writeRichTextColor(&c1)
		c.Assert(err, qt.IsNil)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		c2, err := cs.readRichTextColor()
		c.Assert(err, qt.IsNil)
		c.Assert(c2.coreColor.RGB, qt.Equals, c1.coreColor.RGB)
		c.Assert(c2.coreColor.Tint, qt.Equals, c1.coreColor.Tint)
		c.Assert(*c2.coreColor.Indexed, qt.Equals, *c1.coreColor.Indexed)
		c.Assert(c2.coreColor.Theme, qt.Equals, c1.coreColor.Theme)
		_, err = cs.readBool()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read RichTextColor Theme", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		theme := 8

		c1 := RichTextColor{
			coreColor: xlsxColor{
				Theme: &theme,
			},
		}

		err = cs.writeRichTextColor(&c1)
		c.Assert(err, qt.IsNil)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		c2, err := cs.readRichTextColor()
		c.Assert(err, qt.IsNil)
		c.Assert(c2.coreColor.RGB, qt.Equals, c1.coreColor.RGB)
		c.Assert(c2.coreColor.Tint, qt.Equals, c1.coreColor.Tint)
		c.Assert(c2.coreColor.Indexed, qt.Equals, c1.coreColor.Indexed)
		c.Assert(*c2.coreColor.Theme, qt.Equals, *c1.coreColor.Theme)
		_, err = cs.readBool()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read RichTextFont Bold", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		f1 := RichTextFont{
			Name:      "Font1",
			Size:      12.5,
			Family:    RichTextFontFamilyScript,
			Charset:   RichTextCharsetGreek,
			Color:     &RichTextColor{coreColor: xlsxColor{RGB: "12345678"}},
			Bold:      true,
			Italic:    false,
			Strike:    false,
			VertAlign: RichTextVertAlignSuperscript,
			Underline: RichTextUnderlineSingle,
		}

		err = cs.writeRichTextFont(&f1)
		c.Assert(err, qt.IsNil)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		f2, err := cs.readRichTextFont()
		c.Assert(err, qt.IsNil)
		c.Assert(f2.Name, qt.Equals, f1.Name)
		c.Assert(f2.Size, qt.Equals, f1.Size)
		c.Assert(f2.Family, qt.Equals, f1.Family)
		c.Assert(f2.Charset, qt.Equals, f1.Charset)
		c.Assert(f2.Color.coreColor.RGB, qt.Equals, f1.Color.coreColor.RGB)
		c.Assert(f2.Bold, qt.Equals, f1.Bold)
		c.Assert(f2.Italic, qt.Equals, f1.Italic)
		c.Assert(f2.Strike, qt.Equals, f1.Strike)
		c.Assert(f2.VertAlign, qt.Equals, f1.VertAlign)
		c.Assert(f2.Underline, qt.Equals, f1.Underline)
		_, err = cs.readBool()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read RichTextFont Italic", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		f1 := RichTextFont{
			Italic: true,
		}

		err = cs.writeRichTextFont(&f1)
		c.Assert(err, qt.IsNil)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		f2, err := cs.readRichTextFont()
		c.Assert(err, qt.IsNil)
		c.Assert(f2.Name, qt.Equals, f1.Name)
		c.Assert(f2.Size, qt.Equals, f1.Size)
		c.Assert(f2.Family, qt.Equals, f1.Family)
		c.Assert(f2.Charset, qt.Equals, f1.Charset)
		c.Assert(f2.Color, qt.Equals, f1.Color)
		c.Assert(f2.Bold, qt.Equals, f1.Bold)
		c.Assert(f2.Italic, qt.Equals, f1.Italic)
		c.Assert(f2.Strike, qt.Equals, f1.Strike)
		c.Assert(f2.VertAlign, qt.Equals, f1.VertAlign)
		c.Assert(f2.Underline, qt.Equals, f1.Underline)
		_, err = cs.readBool()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read RichTextFont Strike", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		f1 := RichTextFont{
			Strike: true,
		}

		err = cs.writeRichTextFont(&f1)
		c.Assert(err, qt.IsNil)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		f2, err := cs.readRichTextFont()
		c.Assert(err, qt.IsNil)
		c.Assert(f2.Name, qt.Equals, f1.Name)
		c.Assert(f2.Size, qt.Equals, f1.Size)
		c.Assert(f2.Family, qt.Equals, f1.Family)
		c.Assert(f2.Charset, qt.Equals, f1.Charset)
		c.Assert(f2.Color, qt.Equals, f1.Color)
		c.Assert(f2.Bold, qt.Equals, f1.Bold)
		c.Assert(f2.Italic, qt.Equals, f1.Italic)
		c.Assert(f2.Strike, qt.Equals, f1.Strike)
		c.Assert(f2.VertAlign, qt.Equals, f1.VertAlign)
		c.Assert(f2.Underline, qt.Equals, f1.Underline)
		_, err = cs.readBool()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read RichTextRun with Font", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		r1 := RichTextRun{
			Font: &RichTextFont{
				Bold: true,
			},
			Text: "Text1",
		}

		err = cs.writeRichTextRun(&r1)
		c.Assert(err, qt.IsNil)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		r2, err := cs.readRichTextRun()
		c.Assert(err, qt.IsNil)
		c.Assert(r2.Font, qt.DeepEquals, r1.Font)
		c.Assert(r2.Text, qt.Equals, r1.Text)
		_, err = cs.readBool()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read RichTextRun without Font", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		r1 := RichTextRun{
			Text: "Text1",
		}

		err = cs.writeRichTextRun(&r1)
		c.Assert(err, qt.IsNil)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		r2, err := cs.readRichTextRun()
		c.Assert(err, qt.IsNil)
		c.Assert(r2.Font, qt.Equals, r1.Font)
		c.Assert(r2.Text, qt.Equals, r1.Text)
		_, err = cs.readBool()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read RichText", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		rt1 := []RichTextRun{
			RichTextRun{
				Text: "Text1",
			},
			RichTextRun{
				Font: &RichTextFont{
					Italic: true,
				},
				Text: "Text2",
			},
		}

		err = cs.writeRichText(rt1)
		c.Assert(err, qt.IsNil)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		rt2, err := cs.readRichText()
		c.Assert(err, qt.IsNil)
		c.Assert(rt2, qt.DeepEquals, rt1)
		_, err = cs.readBool()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read Nil RichText", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		var rt1 []RichTextRun = nil

		err = cs.writeRichText(rt1)
		c.Assert(err, qt.IsNil)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		rt2, err := cs.readRichText()
		c.Assert(err, qt.IsNil)
		c.Assert(rt2, qt.HasLen, 0)
		_, err = cs.readBool()
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read Empty RichText", func(c *qt.C) {
		diskvCs, err := NewDiskVCellStore()
		c.Assert(err, qt.IsNil)
		cs, ok := diskvCs.(*DiskVCellStore)
		c.Assert(ok, qt.Equals, true)
		defer cs.Close()

		rt1 := []RichTextRun{}

		err = cs.writeRichText(rt1)
		c.Assert(err, qt.IsNil)
		cs.reader = bytes.NewReader(cs.buf.Bytes())
		rt2, err := cs.readRichText()
		c.Assert(err, qt.IsNil)
		c.Assert(rt2, qt.HasLen, 0)
		_, err = cs.readBool()
		c.Assert(err, qt.Not(qt.IsNil))

	})

}
