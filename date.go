package xlsx

import (
	"fmt"
	"math"
	"strconv"
)

//# Pre-calculate the datetime epochs for efficiency.
var (
	_JDN_delta = []int{2415080 - 61, 2416482 - 1}
	//epoch_1904         = time.Date(1904, 1, 1, 0, 0, 0, 0, time.Local)
	//epoch_1900         = time.Date(1899, 12, 31, 0, 0, 0, 0, time.Local)
	//epoch_1900_minus_1 = time.Date(1899, 12, 30, 0, 0, 0, 0, time.Local)
	_XLDAYS_TOO_LARGE = []int{2958466, 2958466 - 1462} //# This is equivalent to 10000-01-01
)

var (
//ErrXLDateBadTuple = errors.New("XLDate is bad tuple")
//ErrXLDateError    = errors.New("XLDateError")
)

func XLDateTooLarge(d float64) error {
	return fmt.Errorf("XLDate %v is too large", d)
}

func XLDateAmbiguous(d float64) error {
	return fmt.Errorf("XLDate %v is ambiguous", d)
}

func XLDateNegative(d float64) error {
	return fmt.Errorf("XLDate %v is Negative", d)
}

func XLDateBadDatemode(datemode int) error {
	return fmt.Errorf("XLDate is bad datemode %d", datemode)
}

func divmod(a, b int) (int, int) {
	c := a % b
	return (a - c) / b, c
}

func div(a, b int) int {
	return (a - a%b) / b
}

func max(a, b int) int {
	if a > b {
		return a
	}
	return b
}

// this func provide a method to convert date cell string to
// a slice []int. the []int means []int{year, month, day, hour, minute, second}
func StrToDate(data string, datemode int) ([]int, error) {
	xldate, err := strconv.ParseFloat(data, 64)
	if err != nil {
		return nil, err
	}

	if datemode != 0 && datemode != 1 {
		return nil, XLDateBadDatemode(datemode)
	}
	if xldate == 0.00 {
		return []int{0, 0, 0, 0, 0, 0}, nil
	}
	if xldate < 0.00 {
		return nil, XLDateNegative(xldate)
	}
	xldays := int(xldate)
	frac := xldate - float64(xldays)
	seconds := int(math.Floor(frac * 86400.0))
	hour, minute, second := 0, 0, 0
	//assert 0 <= seconds <= 86400
	if seconds == 86400 {
		xldays += 1
	} else {
		//# second = seconds % 60; minutes = seconds // 60
		var minutes int
		minutes, second = divmod(seconds, 60)
		//# minute = minutes % 60; hour    = minutes // 60
		hour, minute = divmod(minutes, 60)
	}
	if xldays >= _XLDAYS_TOO_LARGE[datemode] {
		return nil, XLDateTooLarge(xldate)
	}

	if xldays == 0 {
		return []int{0, 0, 0, hour, minute, second}, nil
	}

	if xldays < 61 && datemode == 0 {
		return nil, XLDateAmbiguous(xldate)
	}

	jdn := xldays + _JDN_delta[datemode]
	yreg := ((((jdn*4+274277)/146097)*3/4)+jdn+1363)*4 + 3
	mp := ((yreg%1461)/4)*535 + 333
	d := ((mp % 16384) / 535) + 1
	//# mp /= 16384
	mp >>= 14
	if mp >= 10 {
		return []int{(yreg / 1461) - 4715, mp - 9, d, hour, minute, second}, nil
	}
	return []int{(yreg / 1461) - 4716, mp + 3, d, hour, minute, second}, nil
}
