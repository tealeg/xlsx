package xlsx


// XLSXSST directly maps the sst element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main currently
// I have not checked this for completeness - it does as much as need.
type XLSXSST struct {
	Count       string    `xml:"count,attr"`
	UniqueCount string    `xml:"uniqueCount,attr"`
	SI          []XLSXSI  `xml:"si"`
}


// XLSXSI directly maps the si element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked this for completeness - it does as
// much as I need.
type XLSXSI struct {
	T string `xml:"t"`
}

// // XLSXT directly maps the t element from the namespace
// // http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// // currently I have not checked this for completeness - it does as
// // much as I need.
// type XLSXT struct {
// 	Data string `xml:"chardata"`
// }


// MakeSharedStringRefTable() takes an XLSXSST struct and converts
// it's contents to an slice of strings used to refer to string values
// by numeric index - this is the model used within XLSX worksheet (a
// numeric reference is stored to a shared cell value).
func MakeSharedStringRefTable(source *XLSXSST) []string {
	reftable := make([]string, len(source.SI))
	for i, si := range source.SI {
		reftable[i] = si.T
	}
	return reftable
}

// ResolveSharedString() looks up a string value by numeric index from
// a provided reference table (just a slice of strings in the correct
// order).  This function only exists to provide clarity or purpose
// via it's name.
func ResolveSharedString(reftable []string, index int) string {
	return reftable[index]
}


