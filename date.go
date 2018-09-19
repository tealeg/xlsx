package xlsx

import (
	"math"
	"time"
)

const (
	MJD_0 float64 = 2400000.5
	MJD_JD2000 float64 = 51544.5

	secondsInADay = float64((24*time.Hour)/time.Second)
	nanosInADay = float64((24*time.Hour)/time.Nanosecond)
)

var (
	timeLocationUTC, _ = time.LoadLocation("UTC")

	unixEpoc = time.Date(1970, time.January, 1, 0, 0, 0, 0, time.UTC)
	// In 1900 mode, Excel takes dates in floating point numbers of days starting with Jan 1 1900.
	// The days are not zero indexed, so Jan 1 1900 would be 1.
	// Except that Excel pretends that Feb 29, 1900 occurred to be compatible with a bug in Lotus 123.
	// So, this constant uses Dec 30, 1899 instead of Jan 1, 1900, so the diff will be correct.
	// http://www.cpearson.com/excel/datetime.htm
	excel1900Epoc = time.Date(1899, time.December, 30, 0, 0, 0, 0, time.UTC)
	excel1904Epoc = time.Date(1904, time.January, 1, 0, 0, 0, 0, time.UTC)
	// Days between epocs, including both off by one errors for 1900.
	daysBetween1970And1900 = float64(unixEpoc.Sub(excel1900Epoc)/(24 * time.Hour))
	daysBetween1970And1904 = float64(unixEpoc.Sub(excel1904Epoc)/(24 * time.Hour))
)

func TimeToUTCTime(t time.Time) time.Time {
	return time.Date(t.Year(), t.Month(), t.Day(), t.Hour(), t.Minute(), t.Second(), t.Nanosecond(), timeLocationUTC)
}

func shiftJulianToNoon(julianDays, julianFraction float64) (float64, float64) {
	switch {
	case -0.5 < julianFraction && julianFraction < 0.5:
		julianFraction += 0.5
	case julianFraction >= 0.5:
		julianDays += 1
		julianFraction -= 0.5
	case julianFraction <= -0.5:
		julianDays -= 1
		julianFraction += 1.5
	}
	return julianDays, julianFraction
}

// Return the integer values for hour, minutes, seconds and
// nanoseconds that comprised a given fraction of a day.
// values would round to 1 us.
func fractionOfADay(fraction float64) (hours, minutes, seconds, nanoseconds int) {

	const (
		c1us  = 1e3
		c1s   = 1e9
		c1day = 24 * 60 * 60 * c1s
	)

	frac := int64(c1day*fraction + c1us/2)
	nanoseconds = int((frac%c1s)/c1us) * c1us
	frac /= c1s
	seconds = int(frac % 60)
	frac /= 60
	minutes = int(frac % 60)
	hours = int(frac / 60)
	return
}

func julianDateToGregorianTime(part1, part2 float64) time.Time {
	part1I, part1F := math.Modf(part1)
	part2I, part2F := math.Modf(part2)
	julianDays := part1I + part2I
	julianFraction := part1F + part2F
	julianDays, julianFraction = shiftJulianToNoon(julianDays, julianFraction)
	day, month, year := doTheFliegelAndVanFlandernAlgorithm(int(julianDays))
	hours, minutes, seconds, nanoseconds := fractionOfADay(julianFraction)
	return time.Date(year, time.Month(month), day, hours, minutes, seconds, nanoseconds, time.UTC)
}

// By this point generations of programmers have repeated the
// algorithm sent to the editor of "Communications of the ACM" in 1968
// (published in CACM, volume 11, number 10, October 1968, p.657).
// None of those programmers seems to have found it necessary to
// explain the constants or variable names set out by Henry F. Fliegel
// and Thomas C. Van Flandern.  Maybe one day I'll buy that jounal and
// expand an explanation here - that day is not today.
func doTheFliegelAndVanFlandernAlgorithm(jd int) (day, month, year int) {
	l := jd + 68569
	n := (4 * l) / 146097
	l = l - (146097*n+3)/4
	i := (4000 * (l + 1)) / 1461001
	l = l - (1461*i)/4 + 31
	j := (80 * l) / 2447
	d := l - (2447*j)/80
	l = j / 11
	m := j + 2 - (12 * l)
	y := 100*(n-49) + i + l
	return d, m, y
}

// Convert an excelTime representation (stored as a floating point number) to a time.Time.
func TimeFromExcelTime(excelTime float64, date1904 bool) time.Time {
	var date time.Time
	var wholeDaysPart = int(excelTime)
	// Excel uses Julian dates prior to March 1st 1900, and
	// Gregorian thereafter.
	if wholeDaysPart <= 61 {
		const OFFSET1900 = 15018.0
		const OFFSET1904 = 16480.0
		var date time.Time
		if date1904 {
			date = julianDateToGregorianTime(MJD_0, excelTime+OFFSET1904)
		} else {
			date = julianDateToGregorianTime(MJD_0, excelTime+OFFSET1900)
		}
		return date
	}
	var floatPart = excelTime - float64(wholeDaysPart)
	if date1904 {
		date = excel1904Epoc
	} else {
		date = excel1900Epoc
	}
	durationPart := time.Duration(nanosInADay * floatPart)
	return date.AddDate(0,0, wholeDaysPart).Add(durationPart)
}

// TimeToExcelTime will convert a time.Time into Excel's float representation, in either 1900 or 1904
// mode. If you don't know which to use, set date1904 to false.
// TODO should this should handle Julian dates?
func TimeToExcelTime(t time.Time, date1904 bool) float64 {
	// Get the number of days since the unix epoc
	daysSinceUnixEpoc := float64(t.Unix())/secondsInADay
	// Get the number of nanoseconds in days since Unix() is in seconds.
	nanosPart := float64(t.Nanosecond())/nanosInADay
	// Add both together plus the number of days difference between unix and Excel epocs.
	var offsetDays float64
	if date1904 {
		offsetDays = daysBetween1970And1904
	} else {
		offsetDays = daysBetween1970And1900
	}
	daysSinceExcelEpoc := daysSinceUnixEpoc + offsetDays + nanosPart
	return daysSinceExcelEpoc
}
