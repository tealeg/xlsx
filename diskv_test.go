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

		_, err = cs.ReadRow("I don't exist", nil)
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

		err = cs.WriteRow(row)
		c.Assert(err, qt.IsNil)
		row2, err := cs.ReadRow(row.key(), sheet)
		c.Assert(err, qt.IsNil)
		c.Assert(row2, qt.Not(qt.IsNil))
		c.Assert(row.Hidden, qt.Equals, row2.Hidden)
		c.Assert(row.GetHeight(), qt.Equals, row2.GetHeight())
		c.Assert(row.GetOutlineLevel(), qt.Equals, row2.GetOutlineLevel())
		c.Assert(row.isCustom, qt.Equals, row2.isCustom)
		c.Assert(row.num, qt.Equals, row2.num)
		c.Assert(row.cellStoreRow.CellCount(), qt.Equals, row2.cellStoreRow.CellCount())
	})

	c.Run("Write and Read Row with Cells", func(c *qt.C) {
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

		cs := sheet.cellStore
		err := cs.WriteRow(row)
		c.Assert(err, qt.IsNil)
		row2, err := cs.ReadRow(row.key(), sheet)
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
		buf := bytes.NewBufferString("")
		writeBool(buf, true)
		writeBool(buf, false)
		reader := bytes.NewReader(buf.Bytes())
		v, err := readBool(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.Equals, true)
		v, err = readBool(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.Equals, false)
		v, err = readBool(reader)
		c.Assert(err, qt.Not(qt.IsNil))
	})

	c.Run("Write and Read unit separator", func(c *qt.C) {
		buf := bytes.NewBufferString("")
		writeUnitSeparator(buf)
		reader := bytes.NewReader(buf.Bytes())
		err := readUnitSeparator(reader)
		c.Assert(err, qt.IsNil)
		err = readUnitSeparator(reader)
		c.Assert(err, qt.Not(qt.IsNil))
	})

	c.Run("Write and Read String", func(c *qt.C) {
		buf := bytes.NewBufferString("")
		writeString(buf, "simple")
		writeString(buf, `multi
line!`)
		writeString(buf, "")
		writeString(buf, "Scheiß encoding")
		reader := bytes.NewReader(buf.Bytes())
		v, err := readString(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.Equals, "simple")
		v, err = readString(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.Equals, `multi
line!`)
		v, err = readString(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.Equals, "")
		v, err = readString(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.Equals, "Scheiß encoding")
		v, err = readString(reader)
		c.Assert(err, qt.Not(qt.IsNil))
	})

	c.Run("Write and Read Int", func(c *qt.C) {
		buf := bytes.NewBufferString("")
		writeInt(buf, math.MinInt64)
		writeInt(buf, 0)
		writeInt(buf, math.MaxInt64)
		reader := bytes.NewReader(buf.Bytes())
		v, err := readInt(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.Equals, math.MinInt64)
		v, err = readInt(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.Equals, 0)
		v, err = readInt(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.Equals, math.MaxInt64)
		v, err = readInt(reader)
		c.Assert(err, qt.Not(qt.IsNil))
	})

	c.Run("Write and Read String Pointer", func(c *qt.C) {
		buf := bytes.NewBufferString("")

		s := "foo"
		writeStringPointer(buf, nil)
		writeStringPointer(buf, &s)
		s = "bar"
		writeStringPointer(buf, &s)
		reader := bytes.NewReader(buf.Bytes())
		v, err := readStringPointer(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(v, qt.IsNil)
		v, err = readStringPointer(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(*v, qt.Equals, "foo")
		v, err = readStringPointer(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(*v, qt.Equals, "bar")
		v, err = readStringPointer(reader)
		c.Assert(err, qt.Not(qt.IsNil))
	})

	c.Run("Write and Read end of record", func(c *qt.C) {
		buf := bytes.NewBufferString("")
		writeEndOfRecord(buf)
		reader := bytes.NewReader(buf.Bytes())
		err := readEndOfRecord(reader)
		c.Assert(err, qt.IsNil)
		err = readEndOfRecord(reader)
		c.Assert(err, qt.Not(qt.IsNil))
	})

	c.Run("Write and Read Border", func(c *qt.C) {
		buf := bytes.NewBufferString("")

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
		writeBorder(buf, b)
		reader := bytes.NewReader(buf.Bytes())
		b2, err := readBorder(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(b2, qt.DeepEquals, b)
		_, err = readBorder(reader)
		c.Assert(err, qt.Not(qt.IsNil))
	})

	c.Run("Write and Read Fill", func(c *qt.C) {
		buf := bytes.NewBufferString("")
		b := Fill{
			PatternType: "PatternType",
			BgColor:     "BgColor",
			FgColor:     "FgColor",
		}
		writeFill(buf, b)
		reader := bytes.NewReader(buf.Bytes())
		b2, err := readFill(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(b2, qt.DeepEquals, b)
		_, err = readFill(reader)
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read Font", func(c *qt.C) {
		buf := bytes.NewBufferString("")
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
		writeFont(buf, b)
		reader := bytes.NewReader(buf.Bytes())
		b2, err := readFont(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(b2, qt.DeepEquals, b)
		_, err = readFont(reader)
		c.Assert(err, qt.Not(qt.IsNil))
	})

	c.Run("Write and Read Alignment", func(c *qt.C) {
		buf := bytes.NewBufferString("")
		b := Alignment{
			Horizontal:   "left",
			Indent:       1,
			ShrinkToFit:  true,
			TextRotation: 90,
			Vertical:     "top",
			WrapText:     true,
		}
		writeAlignment(buf, b)
		reader := bytes.NewReader(buf.Bytes())
		b2, err := readAlignment(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(b2, qt.DeepEquals, b2)
		_, err = readAlignment(reader)
		c.Assert(err, qt.Not(qt.IsNil))
	})

	c.Run("Write and Read Style", func(c *qt.C) {
		buf := bytes.NewBufferString("")
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
		err := writeStyle(buf, &s)
		c.Assert(err, qt.IsNil)
		reader := bytes.NewReader(buf.Bytes())
		s2, err := readStyle(reader)
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
		_, err = readStyle(reader)
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read DataValidation", func(c *qt.C) {
		buf := bytes.NewBufferString("")

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

		writeDataValidation(buf, dv)
		reader := bytes.NewReader(buf.Bytes())
		dv2, err := readDataValidation(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(dv2, qt.DeepEquals, dv)
		_, err = readDataValidation(reader)
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read Cell", func(c *qt.C) {
		buf := bytes.NewBufferString("")

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

		writeCell(buf, cell)
		reader := bytes.NewReader(buf.Bytes())
		cell2, err := readCell(reader)
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
		_, err = readCell(reader)
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read Cell with style", func(c *qt.C) {
		buf := bytes.NewBufferString("")

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

		writeCell(buf, cell)
		reader := bytes.NewReader(buf.Bytes())
		cell2, err := readCell(reader)
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

		_, err = readCell(reader)
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read Cell with DataValidation", func(c *qt.C) {
		buf := bytes.NewBufferString("")

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

		writeCell(buf, cell)
		reader := bytes.NewReader(buf.Bytes())
		cell2, err := readCell(reader)
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

		_, err = readCell(reader)
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read Cell with RichText", func(c *qt.C) {
		buf := bytes.NewBufferString("")

		cell := &Cell{
			RichText: []RichTextRun{
				{
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

		writeCell(buf, cell)
		reader := bytes.NewReader(buf.Bytes())
		cell2, err := readCell(reader)
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
		_, err = readCell(reader)
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read RichTextColor RGB", func(c *qt.C) {
		buf := bytes.NewBufferString("")

		c1 := RichTextColor{
			coreColor: xlsxColor{
				RGB:  "01234567",
				Tint: -0.3,
			},
		}

		err := writeRichTextColor(buf, &c1)
		c.Assert(err, qt.IsNil)
		reader := bytes.NewReader(buf.Bytes())
		c2, err := readRichTextColor(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(c2.coreColor.RGB, qt.Equals, c1.coreColor.RGB)
		c.Assert(c2.coreColor.Tint, qt.Equals, c1.coreColor.Tint)
		c.Assert(c2.coreColor.Indexed, qt.Equals, c1.coreColor.Indexed)
		c.Assert(c2.coreColor.Theme, qt.Equals, c1.coreColor.Theme)
		_, err = readBool(reader)
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read RichTextColor Indexed", func(c *qt.C) {
		buf := bytes.NewBufferString("")

		indexed := 7

		c1 := RichTextColor{
			coreColor: xlsxColor{
				Indexed: &indexed,
				Tint:    0.4,
			},
		}

		err := writeRichTextColor(buf, &c1)
		c.Assert(err, qt.IsNil)
		reader := bytes.NewReader(buf.Bytes())
		c2, err := readRichTextColor(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(c2.coreColor.RGB, qt.Equals, c1.coreColor.RGB)
		c.Assert(c2.coreColor.Tint, qt.Equals, c1.coreColor.Tint)
		c.Assert(*c2.coreColor.Indexed, qt.Equals, *c1.coreColor.Indexed)
		c.Assert(c2.coreColor.Theme, qt.Equals, c1.coreColor.Theme)
		_, err = readBool(reader)
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read RichTextColor Theme", func(c *qt.C) {
		buf := bytes.NewBufferString("")

		theme := 8

		c1 := RichTextColor{
			coreColor: xlsxColor{
				Theme: &theme,
			},
		}

		err := writeRichTextColor(buf, &c1)
		c.Assert(err, qt.IsNil)
		reader := bytes.NewReader(buf.Bytes())
		c2, err := readRichTextColor(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(c2.coreColor.RGB, qt.Equals, c1.coreColor.RGB)
		c.Assert(c2.coreColor.Tint, qt.Equals, c1.coreColor.Tint)
		c.Assert(c2.coreColor.Indexed, qt.Equals, c1.coreColor.Indexed)
		c.Assert(*c2.coreColor.Theme, qt.Equals, *c1.coreColor.Theme)
		_, err = readBool(reader)
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read RichTextFont Bold", func(c *qt.C) {
		buf := bytes.NewBufferString("")

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

		err := writeRichTextFont(buf, &f1)
		c.Assert(err, qt.IsNil)
		reader := bytes.NewReader(buf.Bytes())
		f2, err := readRichTextFont(reader)
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
		_, err = readBool(reader)
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read RichTextFont Italic", func(c *qt.C) {
		buf := bytes.NewBufferString("")

		f1 := RichTextFont{
			Italic: true,
		}

		err := writeRichTextFont(buf, &f1)
		c.Assert(err, qt.IsNil)
		reader := bytes.NewReader(buf.Bytes())
		f2, err := readRichTextFont(reader)
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
		_, err = readBool(reader)
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read RichTextFont Strike", func(c *qt.C) {
		buf := bytes.NewBufferString("")

		f1 := RichTextFont{
			Strike: true,
		}

		err := writeRichTextFont(buf, &f1)
		c.Assert(err, qt.IsNil)
		reader := bytes.NewReader(buf.Bytes())
		f2, err := readRichTextFont(reader)
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
		_, err = readBool(reader)
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read RichTextRun with Font", func(c *qt.C) {
		buf := bytes.NewBufferString("")

		r1 := RichTextRun{
			Font: &RichTextFont{
				Bold: true,
			},
			Text: "Text1",
		}

		err := writeRichTextRun(buf, &r1)
		c.Assert(err, qt.IsNil)
		reader := bytes.NewReader(buf.Bytes())
		r2, err := readRichTextRun(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(r2.Font, qt.DeepEquals, r1.Font)
		c.Assert(r2.Text, qt.Equals, r1.Text)
		_, err = readBool(reader)
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read RichTextRun without Font", func(c *qt.C) {
		buf := bytes.NewBufferString("")

		r1 := RichTextRun{
			Text: "Text1",
		}

		err := writeRichTextRun(buf, &r1)
		c.Assert(err, qt.IsNil)
		reader := bytes.NewReader(buf.Bytes())
		r2, err := readRichTextRun(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(r2.Font, qt.Equals, r1.Font)
		c.Assert(r2.Text, qt.Equals, r1.Text)
		_, err = readBool(reader)
		c.Assert(err, qt.Not(qt.IsNil))
	})

	c.Run("Write and Read RichText", func(c *qt.C) {
		buf := bytes.NewBufferString("")

		rt1 := []RichTextRun{
			{
				Text: "Text1",
			},
			{
				Font: &RichTextFont{
					Italic: true,
				},
				Text: "Text2",
			},
		}

		err := writeRichText(buf, rt1)
		c.Assert(err, qt.IsNil)
		reader := bytes.NewReader(buf.Bytes())
		rt2, err := readRichText(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(rt2, qt.DeepEquals, rt1)
		_, err = readBool(reader)
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read Nil RichText", func(c *qt.C) {
		buf := bytes.NewBufferString("")
		var rt1 []RichTextRun = nil

		err := writeRichText(buf, rt1)
		c.Assert(err, qt.IsNil)
		reader := bytes.NewReader(buf.Bytes())
		rt2, err := readRichText(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(rt2, qt.HasLen, 0)
		_, err = readBool(reader)
		c.Assert(err, qt.Not(qt.IsNil))

	})

	c.Run("Write and Read Empty RichText", func(c *qt.C) {
		buf := bytes.NewBufferString("")

		rt1 := []RichTextRun{}

		err := writeRichText(buf, rt1)
		c.Assert(err, qt.IsNil)
		reader := bytes.NewReader(buf.Bytes())
		rt2, err := readRichText(reader)
		c.Assert(err, qt.IsNil)
		c.Assert(rt2, qt.HasLen, 0)
		_, err = readBool(reader)
		c.Assert(err, qt.Not(qt.IsNil))
	})

}
