package xlsx
import (
	"testing"
	"bytes"
	//"os"
	)

var SST_DATA = `
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t>bar</t><phoneticPr fontId="1" type="noConversion"/></si><si><t>foo</t><phoneticPr fontId="1" type="noConversion"/></si></sst>`

var SHEET_DATA = `
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><dimension ref="A1:B2"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="13.5" x14ac:dyDescent="0.15"/><sheetData><row r="1" spans="1:2" x14ac:dyDescent="0.15"><c r="A1" t="s"><v>1</v></c><c r="B1" t="s"><v>0</v></c></row><row r="2" spans="1:2" x14ac:dyDescent="0.15"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c></row></sheetData><phoneticPr fontId="1" type="noConversion"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`

func TestSheet(t *testing.T){

	sst, err := newSharedStringsTable(bytes.NewBufferString(SST_DATA))
	if err != nil{
		t.Error("ERR= %s", err)
	}
	data := bytes.NewBufferString(SHEET_DATA)
	sheet, err  := NewSheet(data, sst)
	if err != nil{
		t.Error(err)
	}
	if sheet.Row[0].X14ac != "0.15"{
		t.Errorf("Exptexted 0.15 get %s", sheet.Row[0].X14ac)
	}
	value, err := sheet.Cells(0,0)
	if err != nil{
		t.Error("ERR=", err)
	}
	if value != "foo"{
		t.Errorf("Expected foo, get %s", value)
	}

	value, err = sheet.Cells(0,1)
	if err != nil{
		t.Error("ERR=", err)
	}
	if value != "bar"{
		t.Errorf("Expected bar, get %s", value)
	}

	value, err = sheet.Cells(1,0)
	if err != nil{
		t.Error("ERR=", err)
	}
	if value != "1"{
		t.Errorf("Expected 1, get %s", value)
	}

	value, err = sheet.Cells(1,1)
	if err != nil{
		t.Error("ERR=", err)
	}
	if value != "2"{
		t.Errorf("Expected 2, get %s", value)
	}
	//sheet.SetCell(0, 1, "444")
	//sheet.WriteTo(os.Stdout)
}

func TestGetColumnName(t *testing.T){
	col := getColumnName(0)
	if col != "A"{
		t.Errorf("Expected A, get %s", col)
	}

	col = getColumnName(1)
	if col != "B"{
		t.Errorf("Expected B, get %s", col)
	}

	col = getColumnName(25)
	if col != "Z"{
		t.Errorf("Expected Z, get %s", col)
	}

	col = getColumnName(26)
	if col != "AA"{
		t.Errorf("Expected AA, get %s", col)
	}

	cellName := getCellName(0,0)
	if cellName != "A1"{
		t.Errorf("Expected A1, get %s", cellName)
	}

	cellName = getCellName(0,25)
	if cellName != "Z1"{
		t.Errorf("Expected Z1, get %s", cellName)
	}
	cellName = getCellName(26,3)
	if cellName != "D27"{
		t.Errorf("Expected D27, get %s", cellName)
	}

	cellName = getCellName(0, 0)
	if cellName != "A1"{
		t.Errorf("Expected A1, get %s", cellName)
	}

}
