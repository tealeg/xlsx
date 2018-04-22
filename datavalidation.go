package xlsx

import (
	"strings"
)

type DataValidationType int

// Data validation types
/*const (
	_ DataValidationType = iota
	TypeNone
	TypeCustom
	TypeDate
	TypeDecimal
	TypeList
	TypeTextLeng
	TypeTime
	TypeWhole
)*/

// Data validation types
const (
	typeNone     = "none"
	typeCustom   = "custom"
	typeDate     = "date"
	typeDecimal  = "decimal"
	typeList     = "list"
	typeTextLeng = "textLength"
	typeTime     = "time"
	typeWhole    = "whole"
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

/*
type DataValidationOperator int

// Data validation operators
const (
	_DataValidationOperator     = iota
	OperatorBetween             = "between"
	OperatorEqual               = "equal"
	OperatorGreaterThan         = "greaterThan"
	OperatorGgreaterThanOrEqual = "greaterThanOrEqual"
	OperatorLessThan            = "lessThan"
	OperatorLessThanOrEqual     = "lessThanOrEqual"
	OperatorNotBetween          = "notBetween"
	OperatorNotEqual            = "notEqual"
)

// Data validation operators
const (
	operatorBetween            = "between"
	operatorEQUAL              = "equal"
	operatorGREATERTHAN        = "greaterThan"
	operatorGREATERTHANOREQUAL = "greaterThanOrEqual"
	operatorLESSTHAN           = "lessThan"
	operatorLESSTHANOREQUAL    = "lessThanOrEqual"
	operatorNOTBETWEEN         = "notBetween"
	operatorNOTEQUAL           = "notEqual"
)
*/

// NewDataValidation return data validation struct
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

func (dd *xlsxCellDataValidation) SetDropList(keys []string) {
	dd.Formula1 = "\"" + strings.Join(keys, ",") + "\""
	dd.Type = typeList
}

// convBoolToStr  convert boolean to string , false to 0, true to 1
func convBoolToStr(bl bool) string {
	if bl {
		return "1"
	}
	return "0"
}
