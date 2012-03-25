package xlsx
import (
	"encoding/xml"
	"io"
	"io/ioutil"
	"strconv"
	"fmt"
	"errors"
	)

//TODO: 1. What's the meaning of Count and UniqueCount, how to update
// it, if , for ex, one item added
type sharedStringTable struct{
	XMLName xml.Name `xml:"sst"`
	Xmlns       string         `xml:"xmlns,attr"`
	Count       string         `xml:"count,attr"`
	UniqueCount string         `xml:"uniqueCount,attr"`
	SI          []si   `xml:"si"`
}

type si struct {
	T string `xml:"t"`
	PhoneticPr pp `xml:"phoneticPr,omitempty"`
	
}
//no sure the meaning, just load to marshal back
type pp struct {
	FontId string `xml:"fontId,attr,omitempty"`
	Type   string `xml:"type,attr,omitempty"`
}

func newSharedStringsTable(r io.Reader)(*sharedStringTable, error){
	data, err := ioutil.ReadAll(r)
	if err != nil{
		return nil, err
	}
	sst := &sharedStringTable{}
	err = xml.Unmarshal(data, sst)
	if err != nil{
		return nil, err
	}
	return sst, nil
}

//GetStringIndex loop the string table to find the index
//if not found, add a new one and return the index
func (this* sharedStringTable) getIndex(str string)(int, error){
	for index, sharedString := range this.SI{
		if str == sharedString.T{
			return index, nil
		}
	}
	oldLen := len(this.SI)
	this.SI = append(this.SI, si{T:str})
	count, _ := strconv.Atoi(this.Count)
	uniqueCount, _ := strconv.Atoi(this.UniqueCount)

	count++
	uniqueCount++
	this.Count = fmt.Sprintf("%d", count)
	this.UniqueCount = fmt.Sprintf("%d", uniqueCount)

	return oldLen , nil
}

func (this *sharedStringTable) WriteTo(w io.Writer)(error){
	data, err := xml.MarshalIndent(this, " ", "    ")
	if err != nil{
		return err
	}
	_, err = w.Write([]byte(xml.Header))
	_, err = w.Write(data)
	return err
}

func(this *sharedStringTable) getString(index int)(string, error){
	if index >= len(this.SI){
		return "", errors.New("Out of range")
	}
	return this.SI[index].T, nil
}
//func (this *SharedStringsTable)