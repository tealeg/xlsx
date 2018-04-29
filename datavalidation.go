package xlsx

import (
	"fmt"
	"strings"
)

type DataValidationType int

// Data validation types
const (
	_DataValidationType = iota
	typeNone            //inline use
	DataValidationTypeCustom
	DataValidationTypeDate
	DataValidationTypeDecimal
	typeList //inline use
	DataValidationTypeTextLeng
	DataValidationTypeTime
	// DataValidationTypeWhole Integer
	DataValidationTypeWhole
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

// NewXlsxCellDataValidation return data validation struct
func NewXlsxCellDataValidation(allowBlank, ShowInputMessage, showErrorMessage bool) *xlsxCellDataValidation {
	return &xlsxCellDataValidation{
		AllowBlank:       convBoolToStr(allowBlank),
		ShowErrorMessage: convBoolToStr(showErrorMessage),
		ShowInputMessage: convBoolToStr(ShowInputMessage),
	}
}

// SetError set error notice
func (dd *xlsxCellDataValidation) SetError(style DataValidationErrorStyle, title, msg *string) {
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
func (dd *xlsxCellDataValidation) SetInput(title, msg *string) {
	dd.PromptTitle = title
	dd.Prompt = msg
}

// SetDropList data validation list
func (dd *xlsxCellDataValidation) SetDropList(keys []string) {
	dd.Formula1 = "\"" + strings.Join(keys, ",") + "\""
	dd.Type = convDataValidationType(typeList)
}

// SetDropList data validation range
func (dd *xlsxCellDataValidation) SetRange(f1, f2 int, t DataValidationType, o DataValidationOperator) {
	formula1 := fmt.Sprintf("%d", f1)
	formula2 := fmt.Sprintf("%d", f2)
	switch o {
	case DataValidationOperatorBetween:
		if f1 > f2 {
			tmp := formula1
			formula1 = formula2
			formula2 = tmp
		}
	case DataValidationOperatorNotBetween:
		if f1 > f2 {
			tmp := formula1
			formula1 = formula2
			formula2 = tmp
		}
	}

	dd.Formula1 = formula1
	dd.Formula2 = formula2
	dd.Type = convDataValidationType(t)
	dd.Operator = convDataValidationOperatior(o)
}

// convBoolToStr  convert boolean to string , false to 0, true to 1
func convBoolToStr(bl bool) string {
	if bl {
		return "1"
	}
	return "0"
}

// convDataValidationType get excel data validation type
func convDataValidationType(t DataValidationType) string {
	typeMap := map[DataValidationType]string{
		typeNone:                   "none",
		DataValidationTypeCustom:   "custom",
		DataValidationTypeDate:     "date",
		DataValidationTypeDecimal:  "decimal",
		typeList:                   "list",
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
