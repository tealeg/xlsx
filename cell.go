package xlsx

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

// Style is a high level structure intended to provide user access to
// the contents of Style within an XLSX file.
type Style struct {
	Border Border
	Fill   Fill
	Font   Font
}

func NewStyle() *Style {
	return &Style{}
}

// Border is a high level structure intended to provide user access to
// the contents of Border Style within an Sheet.
type Border struct {
	Left   string
	Right  string
	Top    string
	Bottom string
}

// Fill is a high level structure intended to provide user access to
// the contents of background and foreground color index within an Sheet.
type Fill struct {
	PatternType string
	BgColor     string
	FgColor     string
}

func NewFill(patternType, fgColor, bgColor string) *Fill {
	return &Fill{PatternType: patternType, FgColor: fgColor, BgColor: bgColor}
}

type Font struct {
	Size    int
	Name    string
	Family  int
	Charset int
}

func NewFont(size int, name string) *Font {
	return &Font{Size: size, Name: name}
}

// Cell is a high level structure intended to provide user access to
// the contents of Cell within an xlsx.Row.
type Cell struct {
	Value          string
	styleIndex     int
	styles         *xlsxStyles
	numFmtRefTable map[int]xlsxNumFmt
	date1904       bool
}

// CellInterface defines the public API of the Cell.
type CellInterface interface {
	String() string
}

// String returns the value of a Cell as a string.
func (c *Cell) String() string {
	return c.Value
}

// GetStyle returns the Style associated with a Cell
func (c *Cell) GetStyle() Style {
	var err error
	style := Style{}
	style.Border = Border{}
	style.Fill = Fill{}
	style.Font = Font{}

	if c.styleIndex > -1 && c.styleIndex <= len(c.styles.CellXfs) {
		xf := c.styles.CellXfs[c.styleIndex]

		if xf.ApplyBorder {
			style.Border.Left = c.styles.Borders[xf.BorderId].Left.Style
			style.Border.Right = c.styles.Borders[xf.BorderId].Right.Style
			style.Border.Top = c.styles.Borders[xf.BorderId].Top.Style
			style.Border.Bottom = c.styles.Borders[xf.BorderId].Bottom.Style
		}
		// TODO - consider how to handle the fact that
		// ApplyFill can be missing.  At the moment the XML
		// unmarshaller simply sets it to false, which creates
		// a bug.

		// if xf.ApplyFill {
		if xf.FillId > -1 && xf.FillId < len(c.styles.Fills) {
			xFill := c.styles.Fills[xf.FillId]
			style.Fill.PatternType = xFill.PatternFill.PatternType
			style.Fill.FgColor = xFill.PatternFill.FgColor.RGB
			style.Fill.BgColor = xFill.PatternFill.BgColor.RGB
		}
		// }
		if xf.ApplyFont {
			xfont := c.styles.Fonts[xf.FontId]
			style.Font.Size, err = strconv.Atoi(xfont.Sz.Val)
			if err != nil {
				panic(err.Error())
			}
			style.Font.Name = xfont.Name.Val
			style.Font.Family, _ = strconv.Atoi(xfont.Family.Val)
			style.Font.Charset, _ = strconv.Atoi(xfont.Charset.Val)
		}
	}
	return style
}

// The number format string is returnable from a cell.
func (c *Cell) GetNumberFormat() string {
	var numberFormat string = ""
	if c.styleIndex > -1 && c.styleIndex <= len(c.styles.CellXfs) {
		xf := c.styles.CellXfs[c.styleIndex]
		numFmt := c.numFmtRefTable[xf.NumFmtId]
		numberFormat = numFmt.FormatCode
	}
	return strings.ToLower(numberFormat)
}

func (c *Cell) formatToTime(format string) string {
	f, err := strconv.ParseFloat(c.Value, 64)
	if err != nil {
		return err.Error()
	}
	return TimeFromExcelTime(f, c.date1904).Format(format)
}

func (c *Cell) formatToFloat(format string) string {
	f, err := strconv.ParseFloat(c.Value, 64)
	if err != nil {
		return err.Error()
	}
	return fmt.Sprintf(format, f)
}

func (c *Cell) formatToInt(format string) string {
	f, err := strconv.ParseFloat(c.Value, 64)
	if err != nil {
		return err.Error()
	}
	return fmt.Sprintf(format, int(f))
}

// Return the formatted version of the value.
func (c *Cell) FormattedValue() string {
	var numberFormat string = c.GetNumberFormat()
	switch numberFormat {
	case "general":
		return c.Value
	case "0", "#,##0":
		return c.formatToInt("%d")
	case "0.00", "#,##0.00", "@":
		return c.formatToFloat("%.2f")
	case "#,##0 ;(#,##0)", "#,##0 ;[red](#,##0)":
		f, err := strconv.ParseFloat(c.Value, 64)
		if err != nil {
			return err.Error()
		}
		if f < 0 {
			i := int(math.Abs(f))
			return fmt.Sprintf("(%d)", i)
		}
		i := int(f)
		return fmt.Sprintf("%d", i)
	case "#,##0.00;(#,##0.00)", "#,##0.00;[red](#,##0.00)":
		f, err := strconv.ParseFloat(c.Value, 64)
		if err != nil {
			return err.Error()
		}
		if f < 0 {
			return fmt.Sprintf("(%.2f)", f)
		}
		return fmt.Sprintf("%.2f", f)
	case "0%":
		f, err := strconv.ParseFloat(c.Value, 64)
		if err != nil {
			return err.Error()
		}
		f = f * 100
		return fmt.Sprintf("%d%%", int(f))
	case "0.00%":
		f, err := strconv.ParseFloat(c.Value, 64)
		if err != nil {
			return err.Error()
		}
		f = f * 100
		return fmt.Sprintf("%.2f%%", f)
	case "0.00e+00", "##0.0e+0":
		return c.formatToFloat("%e")
	case "mm-dd-yy":
		return c.formatToTime("01-02-06")
	case "d-mmm-yy":
		return c.formatToTime("2-Jan-06")
	case "d-mmm":
		return c.formatToTime("2-Jan")
	case "mmm-yy":
		return c.formatToTime("Jan-06")
	case "h:mm am/pm":
		return c.formatToTime("3:04 pm")
	case "h:mm:ss am/pm":
		return c.formatToTime("3:04:05 pm")
	case "h:mm":
		return c.formatToTime("15:04")
	case "h:mm:ss":
		return c.formatToTime("15:04:05")
	case "m/d/yy h:mm":
		return c.formatToTime("1/2/06 15:04")
	case "mm:ss":
		return c.formatToTime("04:05")
	case "[h]:mm:ss":
		f, err := strconv.ParseFloat(c.Value, 64)
		if err != nil {
			return err.Error()
		}
		t := TimeFromExcelTime(f, c.date1904)
		if t.Hour() > 0 {
			return t.Format("15:04:05")
		}
		return t.Format("04:05")
	case "mmss.0":
		f, err := strconv.ParseFloat(c.Value, 64)
		if err != nil {
			return err.Error()
		}
		t := TimeFromExcelTime(f, c.date1904)
		return fmt.Sprintf("%0d%0d.%d", t.Minute(), t.Second(), t.Nanosecond()/1000)

	case "yyyy\\-mm\\-dd":
		return c.formatToTime("2006\\-01\\-02")
	case "dd/mm/yy":
		return c.formatToTime("02/01/06")
	case "hh:mm:ss":
		return c.formatToTime("15:04:05")
	case "dd/mm/yy\\ hh:mm":
		return c.formatToTime("02/01/06\\ 15:04")
	case "dd/mm/yyyy hh:mm:ss":
		return c.formatToTime("02/01/2006 15:04:05")
	case "yy-mm-dd":
		return c.formatToTime("06-01-02")
	case "d-mmm-yyyy":
		return c.formatToTime("2-Jan-2006")
	case "m/d/yy":
		return c.formatToTime("1/2/06")
	case "m/d/yyyy":
		return c.formatToTime("1/2/2006")
	case "dd-mmm-yyyy":
		return c.formatToTime("02-Jan-2006")
	case "dd/mm/yyyy":
		return c.formatToTime("02/01/2006")
	case "mm/dd/yy hh:mm am/pm":
		return c.formatToTime("01/02/06 03:04 pm")
	case "mm/dd/yyyy hh:mm:ss":
		return c.formatToTime("01/02/2006 15:04:05")
	case "yyyy-mm-dd hh:mm:ss":
		return c.formatToTime("2006-01-02 15:04:05")
	}
	return c.Value
}

func (c *Cell) SetStyle(style *Style) int {
	if c.styles == nil {
		c.styles = &xlsxStyles{}
	}
	index := len(c.styles.Fonts)
	xFont := xlsxFont{}
	xFill := xlsxFill{}
	xBorder := xlsxBorder{}
	xCellStyleXf := xlsxXf{}
	xCellXf := xlsxXf{}
	xFont.Sz.Val = strconv.Itoa(style.Font.Size)
	xFont.Name.Val = style.Font.Name
	xFont.Family.Val = strconv.Itoa(style.Font.Family)
	xFont.Charset.Val = strconv.Itoa(style.Font.Charset)
	xPatternFill := xlsxPatternFill{}
	xPatternFill.PatternType = style.Fill.PatternType
	xPatternFill.FgColor.RGB = style.Fill.FgColor
	xPatternFill.BgColor.RGB = style.Fill.BgColor
	xFill.PatternFill = xPatternFill
	c.styles.Fonts = append(c.styles.Fonts, xFont)
	c.styles.Fills = append(c.styles.Fills, xFill)
	c.styles.Borders = append(c.styles.Borders, xBorder)
	c.styles.CellStyleXfs = append(c.styles.CellStyleXfs, xCellStyleXf)
	c.styles.CellXfs = append(c.styles.CellXfs, xCellXf)
	return index
}
