package xlsx

import "strconv"

// StreamCell holds the data, style and type of cell for streaming
type StreamCell struct {
	cellData  string
	cellStyle StreamStyle
	cellType  CellType
}

// NewStreamCell creates a new StreamCell
func NewStreamCell(cellData string, cellStyle StreamStyle, cellType CellType) StreamCell{
	return StreamCell{
		cellData:  cellData,
		cellStyle: cellStyle,
		cellType:  cellType,
	}
}

// MakeStringStreamCell creates a new cell that holds string data, is of type string and uses general formatting
func MakeStringStreamCell(cellData string) StreamCell{
	return NewStreamCell(cellData, Strings, CellTypeString)
}

// MakeStyledStringStreamCell creates a new cell that holds a string and is styled according to the given style
func MakeStyledStringStreamCell(cellData string, cellStyle StreamStyle) StreamCell {
	return NewStreamCell(cellData, cellStyle, CellTypeString)
}

// MakeIntegerStreamCell creates a new cell that holds an integer value (represented as string),
// is formatted as a standard integer and is of type numeric.
func MakeIntegerStreamCell(cellData int) StreamCell {
	return NewStreamCell(strconv.Itoa(cellData), Integers, CellTypeNumeric)
}

// MakeStyledIntegerStreamCell created a new cell that holds an integer value (represented as string)
// and is styled according to the given style.
func MakeStyledIntegerStreamCell(cellData int, cellStyle StreamStyle) StreamCell {
	return NewStreamCell(strconv.Itoa(cellData), cellStyle, CellTypeNumeric)
}