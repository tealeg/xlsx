package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"strconv"
)

// File is a high level structure providing a slice of Sheet structs
// to the user.
type File struct {
	worksheets     map[string]*zip.File
	referenceTable *RefTable
	styles         *xlsxStyles
	Sheets        map[string]*Sheet          // sheet access by index
}


// Create a new File
func NewFile() (file *File) {
	file = &File{};
	file.Sheets = make(map[string]*Sheet)
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
	f.Sheets[sheetName] = sheet
	return sheet
}

func (f *File) makeWorkbook() xlsxWorkbook {
	var workbook xlsxWorkbook
	workbook = xlsxWorkbook{}
	workbook.FileVersion = xlsxFileVersion{}
	workbook.FileVersion.AppName = "Go XLSX"
	workbook.WorkbookPr = xlsxWorkbookPr{BackupFile: false}
	workbook.BookViews = xlsxBookViews{}
	workbook.BookViews.WorkBookView = make([]xlsxWorkBookView, 1)
	workbook.BookViews.WorkBookView[0] = xlsxWorkBookView{}
	workbook.Sheets = xlsxSheets{}
	workbook.Sheets.Sheet = make([]xlsxSheet, len(f.Sheets))
	return workbook
}


func (f *File) MarshallParts() (map[string]string, error) {
	var parts map[string]string
	var refTable *RefTable = NewSharedStringRefTable()
	var workbookRels WorkBookRels = make(WorkBookRels)
	var err error
	var workbook xlsxWorkbook
	var types xlsxTypes = MakeDefaultContentTypes()

	marshal := func(thing interface{}) (string, error) {
		body, err := xml.MarshalIndent(thing, "  ", "  ")
		if err != nil {
			return "", err
		}
		return xml.Header + string(body), nil
	}


	parts = make(map[string]string)
	workbook = f.makeWorkbook()
	sheetIndex := 1

	for sheetName, sheet := range f.Sheets {
		xSheet := sheet.makeXLSXSheet(refTable)
		rId := fmt.Sprintf("rId%d", sheetIndex)
		sheetId := strconv.Itoa(sheetIndex)
		sheetPath := fmt.Sprintf("worksheets/sheet%d.xml", sheetIndex)
		partName := "xl/" + sheetPath
		types.Overrides = append(
			types.Overrides,
			xlsxOverride{
				PartName: partName,
				ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"})
		workbookRels[rId] = sheetPath
		workbook.Sheets.Sheet[sheetIndex - 1] = xlsxSheet{
			Name: sheetName,
			SheetId: sheetId,
			Id: rId}
		parts[sheetPath], err = marshal(xSheet)
		if err != nil {
			return parts, err
		}
		sheetIndex++
	}

	parts["xl/workbook.xml"], err = marshal(workbook)
	if err != nil {
		return parts, err
	}

	parts[".rels"] = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`

	parts["docProps/app.xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <TotalTime>0</TotalTime>
  <Application>Go XLSX</Application>
</Properties>`

	// TODO - do this properly, modification and revision information
	parts["docProps/core.xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></cp:coreProperties>`

	xSST := refTable.makeXLSXSST()
	parts["xl/sharedStrings.xml"], err = marshal(xSST)
	if err != nil {
		return parts, err
	}
	sheetId := fmt.Sprintf("rId%d", sheetIndex)
	sheetPath := "sharedStrings.xml"
	workbookRels[sheetId] = sheetPath
	sheetIndex++
	xWRel := workbookRels.MakeXLSXWorkbookRels()

	parts["xl/_rels/workbook.xml.rels"], err = marshal(xWRel)
	if err != nil {
		return parts, err
	}

	parts["[Content_Types].xml"], err = marshal(types)
	if err != nil {
		return parts, err
	}

	return parts, nil
}
