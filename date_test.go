package xlsx

import (
	"testing"
	"time"

	qt "github.com/frankban/quicktest"
)

func TestFractionOfADay(t *testing.T) {
	c := qt.New(t)
	var h, m, s, n int
	h, m, s, n = fractionOfADay(0)
	c.Assert(h, qt.Equals, 0)
	c.Assert(m, qt.Equals, 0)
	c.Assert(s, qt.Equals, 0)
	c.Assert(n, qt.Equals, 0)
	h, m, s, n = fractionOfADay(1.0 / 24.0)
	c.Assert(h, qt.Equals, 1)
	c.Assert(m, qt.Equals, 0)
	c.Assert(s, qt.Equals, 0)
	c.Assert(n, qt.Equals, 0)
}

func TestJulianDateToGregorianTime(t *testing.T) {
	c := qt.New(t)
	c.Assert(julianDateToGregorianTime(2400000.5, 51544.0),
		qt.Equals, time.Date(2000, 1, 1, 0, 0, 0, 0, time.UTC))
	c.Assert(julianDateToGregorianTime(2400000.5, 51544.5),
		qt.Equals, time.Date(2000, 1, 1, 12, 0, 0, 0, time.UTC))
	c.Assert(julianDateToGregorianTime(2400000.5, 51544.245),
		qt.Equals, time.Date(2000, 1, 1, 5, 52, 48, 0, time.UTC))
	c.Assert(julianDateToGregorianTime(2400000.5, 51544.2456),
		qt.Equals, time.Date(2000, 1, 1, 5, 53, 39, 840000000, time.UTC))
	/* test rounding: 0.24560789123*24*3600 = 21220.521802272 */
	c.Assert(julianDateToGregorianTime(2400000.5, 51544.24560789123),
		qt.Equals, time.Date(2000, 1, 1, 5, 53, 40, 521802000, time.UTC))
	c.Assert(julianDateToGregorianTime(2400000.5, 51544.1),
		qt.Equals, time.Date(2000, 1, 1, 2, 24, 00, 0, time.UTC))
	c.Assert(julianDateToGregorianTime(2400000.5, 51544.75),
		qt.Equals, time.Date(2000, 1, 1, 18, 0, 0, 0, time.UTC))
}

func TestTimeFromExcelTime(t *testing.T) {
	c := qt.New(t)
	date := TimeFromExcelTime(0, false)
	c.Assert(date, qt.Equals, time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC))
	date = TimeFromExcelTime(60, false)
	c.Assert(date, qt.Equals, time.Date(1900, 2, 28, 0, 0, 0, 0, time.UTC))
	date = TimeFromExcelTime(61, false)
	c.Assert(date, qt.Equals, time.Date(1900, 3, 1, 0, 0, 0, 0, time.UTC))
	date = TimeFromExcelTime(41275.0, false)
	c.Assert(date, qt.Equals, time.Date(2013, 1, 1, 0, 0, 0, 0, time.UTC))
	date = TimeFromExcelTime(401769, false)
	c.Assert(date, qt.Equals, time.Date(3000, 1, 1, 0, 0, 0, 0, time.UTC))
}

func TestTimeFromExcelTimeWithFractionalPart(t *testing.T) {
	c := qt.New(t)
	date := TimeFromExcelTime(0.114583333333333, false)
	c.Assert(date.Round(time.Second), qt.Equals, time.Date(1899, 12, 30, 2, 45, 0, 0, time.UTC))

	date = TimeFromExcelTime(60.1145833333333, false)
	c.Assert(date.Round(time.Second), qt.Equals, time.Date(1900, 2, 28, 2, 45, 0, 0, time.UTC))

	date = TimeFromExcelTime(61.3986111111111, false)
	c.Assert(date.Round(time.Second), qt.Equals, time.Date(1900, 3, 1, 9, 34, 0, 0, time.UTC))

	date = TimeFromExcelTime(37947.75, false)
	c.Assert(date.Round(time.Second), qt.Equals, time.Date(2003, 11, 22, 18, 0, 0, 0, time.UTC))

	date = TimeFromExcelTime(41275.1145833333, false)
	c.Assert(date.Round(time.Second), qt.Equals, time.Date(2013, 1, 1, 2, 45, 0, 0, time.UTC))
}

func TestTimeFromExcelTimeWith1904Offest(t *testing.T) {
	c := qt.New(t)
	date1904Offset := TimeFromExcelTime(39813.0, true)
	c.Assert(date1904Offset, qt.Equals, time.Date(2013, 1, 1, 0, 0, 0, 0, time.UTC))

}
