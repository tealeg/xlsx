package xlsx

import (
	"fmt"
	"math"
	"strconv"
)

// Cell is a high level structure intended to provide user access to
// the contents of Cell within an xlsx.Row.
type Cell struct {
	Row      Row
	Value    string
	style    Style
	numFmt   string
	date1904 bool
	Hidden   bool
}

// CellInterface defines the public API of the Cell.
type CellInterface interface {
	String() string
	FormattedValue() string
}

func NewCell(r Row) *Cell {
	return &Cell{style: *NewStyle(), Row: r}
}

// String returns the value of a Cell as a string.
func (c *Cell) String() string {
	return c.FormattedValue()
}

// GetStyle returns the Style associated with a Cell
func (c *Cell) GetStyle() Style {
	return c.style
}

// SetStyle sets the style of a cell.
func (c *Cell) SetStyle(style Style) {
	c.style = style
}

// The number format string is returnable from a cell.
func (c *Cell) GetNumberFormat() string {
	return c.numFmt
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
