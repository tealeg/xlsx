package xlsx

import (
	"fmt"
	"strings"
	"unicode/utf8"
)

type DataValidationType int

// Data validation types
const (
	_DataValidationType = iota
	typeNone            //inline use
	DataValidationTypeCustom
	DataValidationTypeDate
	DataValidationTypeDecimal
	dataValidationTypeList //inline use
	DataValidationTypeTextLeng
	DataValidationTypeTime
	// DataValidationTypeWhole Integer
	DataValidationTypeWhole
)

const (
	// dataValidationFormulaStrLen 255 runes + 2 quotes
	dataValidationFormulaStrLen = 257
	// dataValidationFormulaStrLenErr
	dataValidationFormulaStrLenErr = "data validation must be 0-255 runes"
)

type DataValidationErrorStyle int

// Data validation error styles
const (
	_ DataValidationErrorStyle = iota
	StyleStop
	StyleWarning
	StyleInformation
)

// Data validation error styles
const (
	styleStop        = "stop"
	styleWarning     = "warning"
	styleInformation = "information"
)

// DataValidationOperator operator enum
type DataValidationOperator int

// Data validation operators
const (
	_DataValidationOperator = iota
	DataValidationOperatorBetween
	DataValidationOperatorEqual
	DataValidationOperatorGreaterThan
	DataValidationOperatorGreaterThanOrEqual
	DataValidationOperatorLessThan
	DataValidationOperatorLessThanOrEqual
	DataValidationOperatorNotBetween
	DataValidationOperatorNotEqual
)

// NewDataValidation return data validation struct
func NewDataValidation(startRow, startCol, endRow, endCol int, allowBlank bool) *xlsxDataValidation {
	startX := ColIndexToLetters(startCol)
	startY := RowIndexToString(startRow)
	endX := ColIndexToLetters(endCol)
	endY := RowIndexToString(endRow)

	sqref := startX + startY
	if startX != endX || startY != endY {
		sqref += ":" + endX + endY
	}
	return &xlsxDataValidation{
		AllowBlank: allowBlank,
		Sqref:      sqref,
	}
}

// SetError set error notice
func (dd *xlsxDataValidation) SetError(style DataValidationErrorStyle, title, msg *string) {
	dd.ShowErrorMessage = true
	dd.Error = msg
	dd.ErrorTitle = title
	strStyle := styleStop
	switch style {
	case StyleStop:
		strStyle = styleStop
	case StyleWarning:
		strStyle = styleWarning
	case StyleInformation:
		strStyle = styleInformation

	}
	dd.ErrorStyle = &strStyle
}

// SetInput set prompt notice
func (dd *xlsxDataValidation) SetInput(title, msg *string) {
	dd.ShowInputMessage = true
	dd.PromptTitle = title
	dd.Prompt = msg
}

// SetDropList sets a hard coded list of values that the drop down will choose from.
// List validations do not work in Apple Numbers.
func (dd *xlsxDataValidation) SetDropList(keys []string) error {
	formula := "\"" + strings.Join(keys, ",") + "\""
	if dataValidationFormulaStrLen < utf8.RuneCountInString(formula) {
		return fmt.Errorf(dataValidationFormulaStrLenErr)
	}
	dd.Formula1 = formula
	dd.Type = convDataValidationType(dataValidationTypeList)
	return nil
}

// SetInFileList is like SetDropList, excel that instead of having a hard coded list,
// a reference to a part of the file is accepted and the list is automatically taken from there.
// Setting y2 to -1 will select all the way to the end of the column. Selecting to the end of the
// column will cause Google Sheets to spin indefinitely while trying to load the possible drop down
// values (more than 5 minutes).
// List validations do not work in Apple Numbers.
func (dd *xlsxDataValidation) SetInFileList(sheet string, x1, y1, x2, y2 int) error {
	start := GetCellIDStringFromCoordsWithFixed(x1, y1, true, true)
	if y2 < 0 {
		y2 = Excel2006MaxRowIndex
	}

	end := GetCellIDStringFromCoordsWithFixed(x2, y2, true, true)
	// Escape single quotes in the file name.
	// Single quotes are escaped by replacing them with two single quotes.
	sheet = strings.Replace(sheet, "'", "''", -1)
	formula := "'" + sheet + "'" + externalSheetBangChar + start + cellRangeChar + end
	dd.Formula1 = formula
	dd.Type = convDataValidationType(dataValidationTypeList)
	return nil
}

// SetDropList data validation range
func (dd *xlsxDataValidation) SetRange(f1, f2 int, t DataValidationType, o DataValidationOperator) error {
	formula1 := fmt.Sprintf("%d", f1)
	formula2 := fmt.Sprintf("%d", f2)

	switch o {
	case DataValidationOperatorBetween:
		if f1 > f2 {
			formula1, formula2 = formula2, formula1
		}
	case DataValidationOperatorNotBetween:
		if f1 > f2 {
			formula1, formula2 = formula2, formula1
		}
	}

	dd.Formula1 = formula1
	dd.Formula2 = formula2
	dd.Type = convDataValidationType(t)
	dd.Operator = convDataValidationOperatior(o)
	return nil
}

// convDataValidationType get excel data validation type
func convDataValidationType(t DataValidationType) string {
	typeMap := map[DataValidationType]string{
		typeNone:                   "none",
		DataValidationTypeCustom:   "custom",
		DataValidationTypeDate:     "date",
		DataValidationTypeDecimal:  "decimal",
		dataValidationTypeList:     "list",
		DataValidationTypeTextLeng: "textLength",
		DataValidationTypeTime:     "time",
		DataValidationTypeWhole:    "whole",
	}

	return typeMap[t]

}

// convDataValidationOperatior get excel data validation operator
func convDataValidationOperatior(o DataValidationOperator) string {
	typeMap := map[DataValidationOperator]string{
		DataValidationOperatorBetween:            "between",
		DataValidationOperatorEqual:              "equal",
		DataValidationOperatorGreaterThan:        "greaterThan",
		DataValidationOperatorGreaterThanOrEqual: "greaterThanOrEqual",
		DataValidationOperatorLessThan:           "lessThan",
		DataValidationOperatorLessThanOrEqual:    "lessThanOrEqual",
		DataValidationOperatorNotBetween:         "notBetween",
		DataValidationOperatorNotEqual:           "notEqual",
	}

	return typeMap[o]

}
