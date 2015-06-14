package xlsx

import (
	"fmt"
	"math"
	"strconv"
	"strings"
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
	HMerge   int
	VMerge   int
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

// Merge with other cells, horizontally and/or vertically.
func (c *Cell) Merge(hcells, vcells int) {
	c.HMerge = hcells
	c.VMerge = vcells
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
	c.SetFloatWithFormat(n, "general")
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
	// beauty the output when the float is small enough
	if n != 0 && n < 0.00001 {
		c.Value = strconv.FormatFloat(n, 'e', -1, 64)
	} else {
		c.Value = strconv.FormatFloat(n, 'f', -1, 64)
	}
	c.numFmt = format
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
	// If bool, just return the value.
	if c.cellType == CellTypeBool {
		return c.Value == "1"
	}
	// If numeric, base it on a non-zero.
	if c.cellType == CellTypeNumeric {
		return c.Value != "0"
	}
	// Return whether there's an empty string.
	return c.Value != ""
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
	if isTimeFormat(numberFormat) {
		return parseTime(c)
	}
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
	}
	return c.Value
}

// parseTime returns a string parsed using time.Time
func parseTime(c *Cell) string {
	f, err := strconv.ParseFloat(c.Value, 64)
	if err != nil {
		return err.Error()
	}
	val := TimeFromExcelTime(f, c.date1904)
	format := c.GetNumberFormat()
	// Replace Excel placeholders with Go time placeholders.
	// For example, replace yyyy with 2006. These are in a specific order,
    // due to the fact that m is used in month, minute, and am/pm. It would
    // be easier to fix that with regular expressions, but if it's possible
    // to keep this simple it would be easier to maintain.
	replacements := []struct{ xltime, gotime string }{
		{"yyyy", "2006"},
		{"yy", "06"},
		{"dd", "02"},
		{"d", "2"},
		{"mmm", "Jan"},
		{"mmss", "0405"},
		{"ss", "05"},
		{"hh", "15"},
		{"h", "3"},
		{"mm:", "04:"},
		{":mm", ":04"},
		{"mm", "01"},
		{"am/pm", "pm"},
		{"m/", "1/"},
		{".0", ".9999"},
	}
	for _, repl := range replacements {
		format = strings.Replace(format, repl.xltime, repl.gotime, 1)
	}
	// If the hour is optional, strip it out, along with the
	// possible dangling colon that would remain.
	if val.Hour() < 1 {
		format = strings.Replace(format, "]:", "]", 1)
		format = strings.Replace(format, "[3]", "", 1)
		format = strings.Replace(format, "[15]", "", 1)
	} else {
		format = strings.Replace(format, "[3]", "3", 1)
		format = strings.Replace(format, "[15]", "15", 1)
	}
	return val.Format(format)
}

// isTimeFormat checks whether an Excel format string represents
// a time.Time.
func isTimeFormat(format string) bool {
	dateParts := []string{
		"yy", "hh", "am", "pm", "ss", "mm", ":",
	}
	for _, part := range dateParts {
		if strings.Contains(format, part) {
			return true
		}
	}
	return false
}
