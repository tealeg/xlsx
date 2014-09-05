package xlsx

import (
	"archive/zip"
	"encoding/xml"
)

// File is a high level structure providing a slice of Sheet structs
// to the user.
type File struct {
	worksheets     map[string]*zip.File
	referenceTable *RefTable
	styles         *xlsxStyles
	Sheets         []*Sheet          // sheet access by index
	Sheet          map[string]*Sheet // sheet access by name
}


// Create a new File
func NewFile() (file *File) {
	file = &File{};
	file.Sheets = make([]*Sheet, 0, 100)
	file.Sheet = make(map[string]*Sheet)
	return
}

// OpenFile() take the name of an XLSX file and returns a populated
// xlsx.File struct for it.
func OpenFile(filename string) (*File, error) {
	var f *zip.ReadCloser
	f, err := zip.OpenReader(filename)
	if err != nil {
		return nil, err
	}
	return ReadZip(f)
}

// Add a new Sheet, with the provided name, to a File
func (f *File) AddSheet(sheetName string) (sheet *Sheet) {
	sheet = &Sheet{}
	f.Sheets = append(f.Sheets, sheet)
	f.Sheet[sheetName] = sheet
	return sheet
}


func (f *File) MarshallParts() ([]string, error) {
	var parts []string
	var refTable *RefTable = NewSharedStringRefTable()
	var err error
	var sheetCount int = len(f.Sheets)

	marshal := func(thing interface{}) (string, error) {
		body, err := xml.MarshalIndent(thing, "  ", "  ")
		if err != nil {
			return "", err
		}
		return xml.Header + string(body), nil
	}

	parts = make([]string, sheetCount + 5)
	for i, sheet := range f.Sheets {
		xSheet := sheet.makeXLSXSheet(refTable)
		parts[i], err = marshal(xSheet)
		if err != nil {
			return parts, err
		}
	}
	xSST := refTable.makeXLSXSST()
	parts[sheetCount], err = marshal(xSST)
	if err != nil {
		return parts, err
	}
	return parts, nil
}
