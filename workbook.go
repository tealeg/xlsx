package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strings"
)

// xlsxWorkbook directly maps the workbook element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxWorkbook struct {
	FileVersion  xlsxFileVersion  `xml:"fileVersion"`
	WorkbookPr   xlsxWorkbookPr   `xml:"workbookPr"`
	BookViews    xlsxBookViews    `xml:"bookViews"`
	Sheets       xlsxSheets       `xml:"sheets"`
	DefinedNames xlsxDefinedNames `xml:"definedNames"`
	CalcPr       xlsxCalcPr       `xml:"calcPr"`

	xlsxsheetinfo map[string]*xlsxWorksheet // sheet date access by name
	styleinfo     *xlsxStyles               // styles data in xml struct
	sstinfo       *xlsxSST                  // shared strings data in xml struct
	worksheets    map[string]*zip.File      // xml files
	xlsxName      string
	rc            *zip.ReadCloser
}

// xlsxFileVersion directly maps the fileVersion element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type xlsxFileVersion struct {
	AppName      string `xml:"appName,attr"`
	LastEdited   string `xml:"lastEdited,attr"`
	LowestEdited string `xml:"lowestEdited,attr"`
	RupBuild     string `xml:"rupBuild,attr"`
}

// xlsxWorkbookPr directly maps the workbookPr element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type xlsxWorkbookPr struct {
	DefaultThemeVersion string `xml:"defaultThemeVersion,attr"`
}

// xlsxBookViews directly maps the bookViews element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type xlsxBookViews struct {
	WorkBookView []xlsxWorkBookView `xml:"workbookView"`
}

// xlsxWorkBookView directly maps the workbookView element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type xlsxWorkBookView struct {
	XWindow      string `xml:"xWindow,attr"`
	YWindow      string `xml:"yWindow,attr"`
	WindowWidth  string `xml:"windowWidth,attr"`
	WindowHeight string `xml:"windowHeight,attr"`
}

// xlsxSheets directly maps the sheets element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheets struct {
	Sheet []xlsxSheet `xml:"sheet"`
}

// xlsxSheet directly maps the sheet element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheet struct {
	Name    string `xml:"name,attr"`
	SheetId string `xml:"sheetId,attr"`
	Id      string `xml:"id,attr"`
}

// xlsxDefinedNames directly maps the definedNames element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type xlsxDefinedNames struct {
	DefinedName []xlsxDefinedName `xml:"definedName"`
}

// xlsxDefinedName directly maps the definedName element from the
// namespace http://schemas.openxmlformats.org/spreadsheetml/2006/main
// - currently I have not checked it for completeness - it does as
// much as I need.
type xlsxDefinedName struct {
	Data         string `xml:",chardata"`
	Name         string `xml:"name,attr"`
	LocalSheetID string `xml:"localSheetId,attr"`
}

// xlsxCalcPr directly maps the calcPr element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxCalcPr struct {
	CalcId string `xml:"calcId,attr"`
}

// readWorkbookFromZipFile() is an internal helper function to
// extract a workbook from the workbook.xml file within
// the XLSX zip file.
func (book *xlsxWorkbook) readWorkbookFromZipFile(bookxml *zip.File) error {
	rc, error := bookxml.Open()
	if error != nil {
		return error
	}
	decoder := xml.NewDecoder(rc)
	error = decoder.Decode(book)
	if error != nil {
		return error
	}
	return nil
}

// read date stored in the sheet by name
// return error
func (book *xlsxWorkbook) readSheetFromZipFile(name string) error {
	var shxmlname string
	for _, sheet := range book.Sheets.Sheet {
		if sheet.Name == name {
			shxmlname = fmt.Sprintf("sheet%s", sheet.Id[3:])
		}
	}
	worksheet := new(xlsxWorksheet)
	shxml := book.worksheets[shxmlname]
	rc, error := shxml.Open()
	if error != nil {
		return error
	}
	decoder := xml.NewDecoder(rc)
	error = decoder.Decode(worksheet)
	if book.xlsxsheetinfo == nil {
		book.xlsxsheetinfo = make(map[string]*xlsxWorksheet)
	}
	worksheet.sst = book.sstinfo
	book.xlsxsheetinfo[name] = worksheet

	return nil
}

// readSharedStringsFromZipFile() is an internal helper function to
// extract a reference table from the sharedStrings.xml file within
// the XLSX zip file.
func (book *xlsxWorkbook) readSharedStringsFromZipFile(sstxml *zip.File) error {
	rc, error := sstxml.Open()
	if error != nil {
		return error
	}
	sst := new(xlsxSST)
	decoder := xml.NewDecoder(rc)
	error = decoder.Decode(sst)
	if error != nil {
		return error
	}
	book.sstinfo = sst

	return nil
}

// readStylesFromZipFile() is an internal helper function to
// extract a style table from the style.xml file within
// the XLSX zip file.
func (book *xlsxWorkbook) readStylesFromZipFile(stylexml *zip.File) error {
	rc, error := stylexml.Open()
	if error != nil {
		return error
	}
	style := new(xlsxStyles)
	decoder := xml.NewDecoder(rc)
	error = decoder.Decode(style)
	if error != nil {
		return error
	}
	book.styleinfo = style
	return nil
}

// close workbook
func (book *xlsxWorkbook) Close() {
	if book.rc != nil {
		book.rc.Close()
		book.rc = nil
	}
}

func (book *xlsxWorkbook) Sheet(name string) (sh *xlsxWorksheet) {
	if book.xlsxsheetinfo[name] == nil {
		error := book.readSheetFromZipFile(name)
		if error != nil {
			panic(error)
		}
	}
	return book.xlsxsheetinfo[name]
}

// save the workbook that has modified
func (book *xlsxWorkbook) Save(xlsxPath string) error {
	xlsxFile, err := os.Create(xlsxPath)
	if err != nil {
		return err
	}
	defer xlsxFile.Close()
	newXlsxZip := zip.NewWriter(xlsxFile)
	defer newXlsxZip.Close()
	oldZip, err := zip.OpenReader(book.xlsxName)
	if err != nil {
		return err
	}

	for _, oldf := range oldZip.File {
		if oldf.Name != "xl/sharedStrings.xml" {
			newf, err := newXlsxZip.Create(oldf.Name)
			if err != nil {
				return err
			}
			oldfrd, err := oldf.Open()
			if err != nil {
				return err
			}
			if !strings.HasPrefix(oldf.Name, "xl/worksheets/sheet") {
				_, err = io.Copy(newf, oldfrd)
				if err != nil {
					return err
				}
			} else {
				changed, sh := book.isSheetChanged(oldf.Name)
				if changed {
					err = sh.WriteTo(newf)
					if err != nil {
						return err
					}
				} else {
					_, err = io.Copy(newf, oldfrd)
					if err != nil {
						return err
					}
				}
			}
		}
	}

	newSstFile, err := newXlsxZip.Create("xl/sharedStrings.xml")
	if err != nil {
		return err
	}
	err = book.sstinfo.WriteTo(newSstFile)
	if err != nil {
		return err
	}
	return nil
}

// determin if one sheet is changed
// if changed return the sheet
func (book *xlsxWorkbook) isSheetChanged(shxmlname string) (bool, *xlsxWorksheet) {
	if book.xlsxsheetinfo == nil {
		return false, nil
	}
	for _, v := range book.Sheets.Sheet {
		if v.Id[3:] == shxmlname[19:len(shxmlname)-4] {
			sh := book.xlsxsheetinfo[v.Name]
			if sh == nil || !sh.changed {
				return false, nil
			} else if sh.changed {
				return true, sh
			}
		}
	}
	return false, nil
}
