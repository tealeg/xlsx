package xlsx

// Default column width in excel
const ColWidth = 9.5
const Excel2006MaxRowIndex = 1048576
const Excel2006MinRowIndex = 1

type Col struct {
	Min                 int
	Max                 int
	Hidden              bool
	Width               float64
	Collapsed           bool
	OutlineLevel        uint8
	numFmt              string
	parsedNumFmt        *parsedNumberFormat
	style               *Style
	DataValidation      *xlsxCellDataValidation
	DataValidationStart int
	DataValidationEnd   int
}

// SetType will set the format string of a column based on the type that you want to set it to.
// This function does not really make a lot of sense.
func (c *Col) SetType(cellType CellType) {
	switch cellType {
	case CellTypeString:
		c.numFmt = builtInNumFmt[builtInNumFmtIndex_STRING]
	case CellTypeNumeric:
		c.numFmt = builtInNumFmt[builtInNumFmtIndex_INT]
	case CellTypeBool:
		c.numFmt = builtInNumFmt[builtInNumFmtIndex_GENERAL] //TEMP
	case CellTypeInline:
		c.numFmt = builtInNumFmt[builtInNumFmtIndex_STRING]
	case CellTypeError:
		c.numFmt = builtInNumFmt[builtInNumFmtIndex_GENERAL] //TEMP
	case CellTypeDate:
		// Cells that are stored as dates are not properly supported in this library.
		// They should instead be stored as a Numeric with a date format.
		c.numFmt = builtInNumFmt[builtInNumFmtIndex_GENERAL]
	case CellTypeStringFormula:
		c.numFmt = builtInNumFmt[builtInNumFmtIndex_STRING]
	}
}

// GetStyle returns the Style associated with a Col
func (c *Col) GetStyle() *Style {
	return c.style
}

// SetStyle sets the style of a Col
func (c *Col) SetStyle(style *Style) {
	c.style = style
}

// SetDataValidation set data validation with start,end ; start or end  = 0  equal all column
func (c *Col) SetDataValidation(dd *xlsxCellDataValidation, start, end int) {
	//2006 excel all row 1048576
	if 0 == start {
		c.DataValidationStart = Excel2006MinRowIndex
	} else {
		c.DataValidationStart = start
	}

	if 0 == end || c.DataValidationEnd > Excel2006MaxRowIndex {
		c.DataValidationEnd = Excel2006MaxRowIndex
	} else {
		c.DataValidationEnd = end
	}
	c.DataValidation = dd

}

// SetDataValidationWithStart set data validation with start
func (c *Col) SetDataValidationWithStart(dd *xlsxCellDataValidation, start int) {
	//2006 excel all row 1048576
	if 0 == start {
		c.DataValidationStart = Excel2006MinRowIndex
	} else {
		c.DataValidationStart = start
	}

	c.DataValidationEnd = Excel2006MaxRowIndex
	c.DataValidation = dd

}
