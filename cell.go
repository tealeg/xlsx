package xlsx

import (
	"fmt"
	"math"
	"strconv"
)

// CellType is an int type for storing metadata about the data type in the cell.
type CellType int

// Known types for cell values.
const (
	CellTypeString CellType = iota
	CellTypeFormula
	CellTypeNumeric
	CellTypeBool
	CellTypeInline
	CellTypeError
)

// Cell is a high level structure intended to provide user access to
// the contents of Cell within an xlsx.Row.
type Cell struct {
	Row      *Row
	Value    string
	formula  string
	style    *Style
	numFmt   string
	date1904 bool
	Hidden   bool
	cellType CellType
}

// CellInterface defines the public API of the Cell.
type CellInterface interface {
	String() string
	FormattedValue() string
}

// NewCell creates a cell and adds it to a row.
func NewCell(r *Row) *Cell {
	return &Cell{style: NewStyle(), Row: r}
}

// Type returns the CellType of a cell. See CellType constants for more details.
func (c *Cell) Type() CellType {
	return c.cellType
}

// SetString sets the value of a cell to a string.
func (c *Cell) SetString(s string) {
	c.Value = s
	c.formula = ""
	c.cellType = CellTypeString
}

// String returns the value of a Cell as a string.
func (c *Cell) String() string {
	return c.FormattedValue()
}

// SetFloat sets the value of a cell to a float.
func (c *Cell) SetFloat(n float64) {
	c.SetFloatWithFormat(n, "0.00e+00")
}

/*
	The following are samples of format samples.

	* "0.00e+00"
	* "0", "#,##0"
	* "0.00", "#,##0.00", "@"
	* "#,##0 ;(#,##0)", "#,##0 ;[red](#,##0)"
	* "#,##0.00;(#,##0.00)", "#,##0.00;[red](#,##0.00)"
	* "0%", "0.00%"
	* "0.00e+00", "##0.0e+0"
*/

// SetFloatWithFormat sets the value of a cell to a float and applies
// formatting to the cell.
func (c *Cell) SetFloatWithFormat(n float64, format string) {
	// tmp value. final value is formatted by FormattedValue() method
	c.Value = fmt.Sprintf("%e", n)
	c.numFmt = format
	c.Value = c.FormattedValue()
	c.formula = ""
	c.cellType = CellTypeNumeric
}

// Float returns the value of cell as a number.
func (c *Cell) Float() (float64, error) {
	f, err := strconv.ParseFloat(c.Value, 64)
	if err != nil {
		return math.NaN(), err
	}
	return f, nil
}

// SetInt64 sets a cell's value to a 64-bit integer.
func (c *Cell) SetInt64(n int64) {
	c.Value = fmt.Sprintf("%d", n)
	c.numFmt = "0"
	c.formula = ""
	c.cellType = CellTypeNumeric
}

// Int64 returns the value of cell as 64-bit integer.
func (c *Cell) Int64() (int64, error) {
	f, err := strconv.ParseInt(c.Value, 10, 64)
	if err != nil {
		return -1, err
	}
	return f, nil
}

// SetInt sets a cell's value to an integer.
func (c *Cell) SetInt(n int) {
	c.Value = fmt.Sprintf("%d", n)
	c.numFmt = "0"
	c.formula = ""
	c.cellType = CellTypeNumeric
}

// Int returns the value of cell as integer.
// Has max 53 bits of precision
// See: float64(int64(math.MaxInt))
func (c *Cell) Int() (int, error) {
	f, err := strconv.ParseFloat(c.Value, 64)
	if err != nil {
		return -1, err
	}
	return int(f), nil
}

// SetBool sets a cell's value to a boolean.
func (c *Cell) SetBool(b bool) {
	if b {
		c.Value = "1"
	} else {
		c.Value = "0"
	}
	c.cellType = CellTypeBool
}

// Bool returns a boolean from a cell's value.
// TODO: Determine if the current return value is
// appropriate for types other than CellTypeBool.
func (c *Cell) Bool() bool {
	return c.Value == "1"
}

// SetFormula sets the format string for a cell.
func (c *Cell) SetFormula(formula string) {
	c.formula = formula
	c.cellType = CellTypeFormula
}

// Formula returns the formula string for the cell.
func (c *Cell) Formula() string {
	return c.formula
}

// GetStyle returns the Style associated with a Cell
func (c *Cell) GetStyle() *Style {
	return c.style
}

// SetStyle sets the style of a cell.
func (c *Cell) SetStyle(style *Style) {
	c.style = style
}

// GetNumberFormat returns the number format string for a cell.
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

// FormattedValue returns the formatted version of the value.
// If it's a string type, c.Value will just be returned. Otherwise,
// it will attempt to apply Excel formatting to the value.
func (c *Cell) FormattedValue() string {
	var numberFormat = c.GetNumberFormat()
	switch numberFormat {
	case "general", "@":
		return c.Value
	case "0", "#,##0":
		return c.formatToInt("%d")
	case "0.00", "#,##0.00":
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
