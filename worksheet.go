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
	XMLName xml.Name `xml:"sheetData"`
	Row []row `xml:"row"`
	sst *sharedStringTable
	head, tail string
}

type row struct {
	R     string `xml:"r,attr"`
	Spans string `xml:"spans,attr"`
	Ht    string `xml:"ht,attr"`
	Cht   string `xml:"customHeight,attr"`
	ThickBot string `xml:"thickBot,attr,omitempty"`
	X14ac string `xml:"dyDescent,attr"`//how to deal attr namespace x14ac
	C     []cell `xml:"c"`
}

type cell struct {
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
	data, err := xml.MarshalIndent(this, "", "    ")
	if err != nil{
		return err
	}
	content := string(data)
	content = strings.Replace(content, "dyDescent=","x14ac:dyDescent=", -1)
	_, err = w.Write([]byte(this.head))
	_, err = w.Write([]byte(content))
	_, err = w.Write([]byte(this.tail))
	return err
}

//
//Can't add a new row yet.....
//
func (this *Sheet) Cells(rowIndex, colIndex int) (string, error){

	colomnData, _ := this.getCell(rowIndex, colIndex)

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

func (this *Sheet)getCell(rowIndex int, colIndex int)(*cell, error){
	cellName := getCellName(rowIndex, colIndex)
	for i, row := range this.Row{
		if row.R == fmt.Sprintf("%d", rowIndex + 1){
			for j, col := range row.C{
				if col.R == cellName{
					return &this.Row[i].C[j], nil
				}
			}
			newCell :=  cell{R:cellName}
			lenEnd := len(this.Row[i].C)
			this.Row[i].C = append(this.Row[i].C, newCell)
			return &this.Row[i].C[lenEnd + 1], nil
		}
	}
	//didn't find the row
	newRow := row{R:fmt.Sprintf("%d", rowIndex + 1)}
	newRow.C = make([]cell, 1)
	newRow.C[0] = cell{R:cellName}
	this.Row = append(this.Row, newRow)
	return &newRow.C[0], nil
}
	


func (this *Sheet)SetCell(row int, col int, value interface{})(error){

	colomnData, _ := this.getCell(row, col)

	if this.sst ==nil{
		return errors.New("The shared string table is nil")
	}
	if strValue, ok := value.(string); ok{
		index, err := this.sst.getIndex(strValue)
		if err != nil{
			return err
		}
		colomnData.V = fmt.Sprintf("%d",index)
		colomnData.T = "s"
	}else if intValue, ok := value.(int); ok{
		colomnData.V = fmt.Sprintf("%d",intValue)
		colomnData.T = ""
	}else if floatValue, ok := value.(float32); ok{
		colomnData.V = fmt.Sprintf("%f", floatValue)
		colomnData.T = ""
	}else{
		return errors.New("Unknow type")
	}
	return nil
}

func (this *Sheet)setSharedStringTable(sst* sharedStringTable){
	this.sst = sst
}

func getColumnName(columnIndex int) string{
	intOfA := int([]byte("A")[0])
	colName := string(intOfA + columnIndex % 26)
	columnIndex = columnIndex / 26
	if columnIndex != 0{
		return	string(intOfA - 1 + columnIndex) + colName
	}
	return colName
}

func getCellName(row, column int) string{
	return getColumnName(column) + fmt.Sprintf("%d", row + 1)
}
	
	