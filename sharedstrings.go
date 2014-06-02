package xlsx

// xlsxSST directly maps the sst element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main currently
// I have not checked this for completeness - it does as much as need.
type xlsxSST struct {
	Count       int   `xml:"count,attr"`
	UniqueCount int   `xml:"uniqueCount,attr"`
	SI          []xlsxSI `xml:"si"`
}

// xlsxSI directly maps the si element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked this for completeness - it does as
// much as I need.
type xlsxSI struct {
	T string  `xml:"t"`
	R []xlsxR `xml:"r"`
}

// xlsxR directly maps the r element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked this for completeness - it does as
// much as I need.
type xlsxR struct {
	T string `xml:"t"`
}

// // xlsxT directly maps the t element from the namespace
// // http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// // currently I have not checked this for completeness - it does as
// // much as I need.
// type xlsxT struct {
// 	Data string `xml:"chardata"`
// }


type RefTable []string

// NewSharedStringRefTable() creates a new, empty RefTable.
func NewSharedStringRefTable() RefTable {
	return RefTable{}
}

// MakeSharedStringRefTable() takes an xlsxSST struct and converts
// it's contents to an slice of strings used to refer to string values
// by numeric index - this is the model used within XLSX worksheet (a
// numeric reference is stored to a shared cell value).
func MakeSharedStringRefTable(source *xlsxSST) RefTable {
	reftable := make(RefTable, len(source.SI))
	for i, si := range source.SI {
		if len(si.R) > 0 {
			for j := 0; j < len(si.R); j++ {
				reftable[i] = reftable[i] + si.R[j].T
			}
		} else {
			reftable[i] = si.T
		}
	}
	return reftable
}

// makeXlsxSST() takes a RefTable and returns and
// equivalent xlsxSST representation.
func (rt RefTable) makeXlsxSST() xlsxSST {
	sst := xlsxSST{}
	sst.Count = len(rt)
	sst.UniqueCount = sst.Count
	for _, ref := range rt {
		si := xlsxSI{}
		si.T = ref
		sst.SI = append(sst.SI, si)
	}
	return sst
}

// Resolvesharedstring() looks up a string value by numeric index from
// a provided reference table (just a slice of strings in the correct
// order).  This function only exists to provide clarity or purpose
// via it's name.
func (rt RefTable) ResolveSharedString(index int) string {
	return rt[index]
}
