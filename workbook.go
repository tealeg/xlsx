// wb := OpenWorkbook(filepath)
// wb.Sheets[0].SetCell(0, 0, "Foo")
// wb.Save


package xlsx

import (
	"archive/zip"
	"errors"
	"fmt"
	//"log"
	"encoding/xml"
	"io"
	"io/ioutil"
)

type Workbook struct{
	Sheets []*Sheet
	sst    *sharedStringTable
	file   *zip.Writer
}

//marshal data from workbook.xml
type workbookInfo struct{
	XMLName xml.Name `xml:"workbook"`
	SheetsInfo []sheetInfo `xml:"sheets>sheet"`
}
type sheetInfo struct{
	SheetId int `xml:"sheetId,attr"`
	Name    string `xml:"name,attr"`
}


func getWorkbookInfo(rd io.Reader)(*workbookInfo, error){
	data, err := ioutil.ReadAll(rd)
	if err != nil{
		return nil, err
	}
	wi := workbookInfo{}
	err = xml.Unmarshal(data, &wi)
	if err != nil{
		return nil, err
	}
	return &wi, nil
}

func getFileFromZipReader(rd *zip.ReadCloser, fileName string)(io.Reader, error){
	for _, f :=range rd.File{
		if f.Name == fileName{
			frd, err := f.Open()
			if err != nil{
				return nil, err
			}
			return frd, nil
		}
	}
	return nil, errors.New(fmt.Sprintf("getFileFromZipReader: can't find the file %s", fileName))
}

//OpenWorkbook open the xlsx file, and return the Workbook object
func OpenWorkbook(xlsxPath string)(*Workbook, error){
	//open the zip
	rd, err := zip.OpenReader(xlsxPath)
	if err != nil{
		return nil, err
	}
	defer rd.Close()

	workbookInfoReader, err := getFileFromZipReader(rd, "xl/workbook.xml")
	if err != nil{
		return nil, err
	}

	wbi, err := getWorkbookInfo(workbookInfoReader)
	if err != nil{
		return nil, err
	}
	
	var sst *sharedStringTable
	sstrd, err := getFileFromZipReader(rd, "xl/sharedStrings.xml")
	if err != nil{
		return nil, err
	}
	sst, err = newSharedStringsTable(sstrd)
	if err != nil{
		return nil, err
	}

	sheets := make([]*Sheet, len(wbi.SheetsInfo))
	for i, si := range wbi.SheetsInfo{
		fileName := fmt.Sprintf("xl/worksheets/sheet%d.xml", si.SheetId)
		sheetRd, err := getFileFromZipReader(rd, fileName)
		if err != nil{
			return nil, err
		}
		sheets[i], err = NewSheet(sheetRd, sst)
		if err != nil{
			return nil, err
		}

	}
	return &Workbook{Sheets:sheets, sst:sst}, nil
}

func (this *Workbook) Save() (error){
	return nil

}
