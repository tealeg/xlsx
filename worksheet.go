package xlsx
import (
	"io/ioutil"
	"encoding/xml"
	"io"
	"errors"
	"strconv"
	"strings"
	"fmt"
	)
type Sheet struct{
	Row []row `xml:"row"`
	sst *sharedStringTable
	head, tail string
}

type row struct {
	R     string `xml:"r,attr"`
	Spans string `xml:"spans,attr"`
	Ht    string `xml:"ht,attr"`
	Cht   string `xml:"customHeight,attr"`
	X14ac string `xml:"dyDescent,attr"`//how to deal attr namespace x14ac
	C     []column `xml:"c"`
}

type column struct {
	V string `xml:"v,omitempty"`
	R string `xml:"r,attr"`
	S string `xml:"s,attr"`
	T string `xml:"t,omitempty,attr"`

}
//NewSheet marshal the reader's content, the sst can be nil
//using setSharedStringTable set is later
func NewSheet(r io.Reader, sst *sharedStringTable)(*Sheet, error){

	data, err := ioutil.ReadAll(r)
	content := string(data)
	index1 := strings.Index(content, `<sheetData>`)
	index2 := strings.Index(content, `</sheetData>`)
	if index1 == -1 {
		return nil, errors.New(fmt.Sprintf("Can't find the sheetData tag, %s", content))
	}
	if index2 == -1{
		return nil, errors.New(fmt.Sprintf("Can't find the </sheetData> %s", content))
	}
	head := content[0:index1]
	tail := content[index2 + len(`</sheetData>`):]
	sheetData  := content[index1: index2+len(`</sheetData>`)]
	if err != nil{
		return nil, err
	}
	sheet := &Sheet{head:head, tail:tail, sst:sst}
	err = xml.Unmarshal([]byte(sheetData), sheet)
	if err != nil{
		return nil, err
	}
	return sheet, nil
}

func (this *Sheet) WriteTo(w io.Writer)(error){
	data, err := xml.MarshalIndent(this, " ", "    ")
	if err != nil{
		return err
	}
	_, err = w.Write([]byte(this.head))
	_, err = w.Write(data)
	_, err = w.Write([]byte(this.tail))
	return err
}

func (this *Sheet) Cells(row, column int) (string, error){
	if this.Row == nil || len(this.Row) == 0{
		return "", errors.New("Illegal sheet, row = nil")
	}
	if row >= len(this.Row){
		return "", errors.New("Row is Out of range")
	}
	if column >= len(this.Row[row].C){
		return "",errors.New("Column is out of range")
	}
	colomnData := this.Row[row].C[column]

	if colomnData.T == "s"{
		if this.sst == nil{
			return "", errors.New("Sheet::Cells, sst is nil. Invalid shared string")
		}
		index, err := strconv.Atoi(colomnData.V)
		if err != nil{
			return "", err
		}
		ret, err := this.sst.getString(index)
		if err != nil{
			return "", err
		}
		return ret, nil
	}
			
	return colomnData.V, nil
}

func (this *Sheet)SetCell(row int, column int, content string)(error){
	if row >= len(this.Row){
		return errors.New("Row is Out of range")
	}
	if column >= len(this.Row[row].C){
		return errors.New("Column is out of range")
	}

	if _, err1 := strconv.ParseInt(content, 10, 64); err1 != nil{
		this.Row[row].C[column].V = content
		this.Row[row].C[column].T = ""
	}
	if 	_, err2 := strconv.ParseFloat(content, 64); err2 != nil{
		this.Row[row].C[column].V = content
		this.Row[row].C[column].T = ""
		return nil
	}
	if this.sst ==nil{
		return errors.New("The shared string table is nil")
	}
	index, _ := this.sst.getIndex(content)
	this.Row[row].C[column].V = index
	this.Row[row].C[column].T = "s"
	return nil
}

func (this *Sheet)setSharedStringTable(sst* sharedStringTable){
	this.sst = sst
}