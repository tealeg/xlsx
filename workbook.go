// wb := OpenWorkbook(filepath)
// wb.Sheets[0].SetCell(0, 0, "Foo")
// wb.Save


package xlsx

import (
	"archive/zip"
	"errors"
	"fmt"
	"encoding/xml"
	"io"
	"io/ioutil"
	"os"
	"strings"
	
)

const (
	sharedStringFileName = "xl/sharedStrings.xml"
	sheetFileNamePrefix = "xl/worksheets/sheet"
	)
type Workbook struct{
	Sheets []*Sheet
	sst    *sharedStringTable
	xlsxName string
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
	return &Workbook{Sheets:sheets, sst:sst, xlsxName:xlsxPath}, nil
}

func (this *Workbook) Save(xlsxPath string) (error){
	xlsxFile, err := os.Create(xlsxPath)
	if err != nil{
		return err
	}
	defer xlsxFile.Close()
	newXlsxZip := zip.NewWriter(xlsxFile)
	defer newXlsxZip.Close()
	oldZip, err := zip.OpenReader(this.xlsxName)
	if err != nil{
		return err
	}
	
	for _, oldf := range oldZip.File{
		if oldf.Name != sharedStringFileName && !strings.HasPrefix(oldf.Name, sheetFileNamePrefix){
			newf, err := newXlsxZip.Create(oldf.Name)	
			if err != nil{
				return err
			}
			oldfrd, err := oldf.Open()
			if err != nil{
				return err
			}
			_, err = io.Copy(newf, oldfrd)
			if err != nil{
				return err
			}
		}
	}

	newSstFile, err := newXlsxZip.Create(sharedStringFileName)
	if err != nil{
		return err
	}
	err = this.sst.WriteTo(newSstFile)
	if err != nil{
		return err
	}
	for i, sheet := range this.Sheets{
		fileName := fmt.Sprintf("xl/worksheets/sheet%d.xml", i+1)
		sheetXml, err := newXlsxZip.Create(fileName)
		if err != nil{
			return nil
		}
		err = sheet.WriteTo(sheetXml)
		if err != nil{
			return nil
		}
	}
	return nil

}
