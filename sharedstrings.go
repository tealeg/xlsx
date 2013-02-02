package xlsx

import (
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
)

// xlsxSST directly maps the sst element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main currently
// I have not checked this for completeness - it does as much as need.
type xlsxSST struct {
	XMLName     xml.Name `xml:"sst"`
	Xmlns       string   `xml:"xmlns,attr"`
	Count       string   `xml:"count,attr"`
	UniqueCount string   `xml:"uniqueCount,attr"`
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

// MakeSharedStringRefTable() takes an xlsxSST struct and converts
// it's contents to an slice of strings used to refer to string values
// by numeric index - this is the model used within XLSX worksheet (a
// numeric reference is stored to a shared cell value).
func MakeSharedStringRefTable(source *xlsxSST) []string {
	reftable := make([]string, len(source.SI))
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

//GetStringIndex loop the string table to find the index
//if not found, add a new one and return the index
func (sst *xlsxSST) getIndex(str string) (int, error) {
	for index, sharedString := range sst.SI {
		if str == sharedString.T {
			return index, nil
		}
	}
	oldLen := len(sst.SI)
	sst.SI = append(sst.SI, xlsxSI{T: str})
	count, _ := strconv.Atoi(sst.Count)
	uniqueCount, _ := strconv.Atoi(sst.UniqueCount)

	count++
	uniqueCount++
	sst.Count = fmt.Sprintf("%d", count)
	sst.UniqueCount = fmt.Sprintf("%d", uniqueCount)

	return oldLen, nil
}

// write shared string table to xml file
func (sst *xlsxSST) WriteTo(w io.Writer) error {
	data, err := xml.Marshal(sst)
	if err != nil {
		return err
	}
	content := string(data)
	_, err = w.Write([]byte(Header))
	_, err = w.Write([]byte(content))
	return err
}

// ResolveSharedString() looks up a string value by numeric index from
// a provided reference table (just a slice of strings in the correct
// order).  This function only exists to provide clarity or purpose
// via it's name.
func ResolveSharedString(reftable []string, index int) string {
	return reftable[index]
}
