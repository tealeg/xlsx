package xlsx

// Default column width in excel
const ColWidth = 9.5
const Excel2006MaxRowCount = 1048576
const Excel2006MaxRowIndex = Excel2006MaxRowCount - 1

type Col struct {
	Min            int
	Max            int
	Hidden         bool
	Width          float64
	Collapsed      bool
	OutlineLevel   uint8
	numFmt         string
	parsedNumFmt   *parsedNumberFormat
	style          *Style
	DataValidation []*xlsxCellDataValidation
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

// SetDataValidation set data validation with zero based start and end.
// Set end to -1 for all rows.
func (c *Col) SetDataValidation(dd *xlsxCellDataValidation, start, end int) {
	if end < 0 {
		end = Excel2006MaxRowIndex
	}

	dd.minRow = start
	dd.maxRow = end

	tmpDD := make([]*xlsxCellDataValidation, 0)
	for _, item := range c.DataValidation {
		if item.maxRow < dd.minRow {
			tmpDD = append(tmpDD, item) //No intersection
		} else if item.minRow > dd.maxRow {
			tmpDD = append(tmpDD, item) //No intersection
		} else if dd.minRow <= item.minRow && dd.maxRow >= item.maxRow {
			continue //union , item can be ignored
		} else if dd.minRow >= item.minRow {
			//Split into three or two, Newly added object, intersect with the current object in the lower half
			tmpSplit := new(xlsxCellDataValidation)
			*tmpSplit = *item

			if dd.minRow > item.minRow { //header whetherneed to split
				item.maxRow = dd.minRow - 1
				tmpDD = append(tmpDD, item)
			}
			if dd.maxRow < tmpSplit.maxRow { //footer whetherneed to split
				tmpSplit.minRow = dd.maxRow + 1
				tmpDD = append(tmpDD, tmpSplit)
			}

		} else {
			item.minRow = dd.maxRow + 1
			tmpDD = append(tmpDD, item)
		}
	}
	tmpDD = append(tmpDD, dd)
	c.DataValidation = tmpDD
}

// SetDataValidationWithStart set data validation with a zero basd start row.
// This will apply to the rest of the rest of the column.
func (c *Col) SetDataValidationWithStart(dd *xlsxCellDataValidation, start int) {
	c.SetDataValidation(dd, start, -1)
}
