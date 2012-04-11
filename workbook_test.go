package xlsx

import (
	"testing"
	"bytes"
)


func TestWorkbookInfo(t *testing.T){
	data := bytes.NewBufferString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9302"/><workbookPr defaultThemeVersion="124226"/><bookViews><workbookView xWindow="120" yWindow="60" windowWidth="13995" windowHeight="4905"/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId2"/><sheet name="Sheet3" sheetId="3" r:id="rId3"/></sheets><calcPr calcId="144525"/></workbook>`)
	wbi, err := getWorkbookInfo(data)
	if err != nil{
		t.Error("ERR=", err)
	}
	if len(wbi.SheetsInfo) != 3{
		t.Error("Expected 3, get ", len(wbi.SheetsInfo))
	}
	if wbi.SheetsInfo[0].SheetId != 1{
		t.Error("Expected 1, get ", wbi.SheetsInfo[0].SheetId)
	}
	if wbi.SheetsInfo[1].SheetId != 2{
		t.Error("Expected 2, get ", wbi.SheetsInfo[1].SheetId)
	}
	if wbi.SheetsInfo[2].SheetId != 3{
		t.Error("Expected 3, get ", wbi.SheetsInfo[2].SheetId)
	}

	if wbi.SheetsInfo[0].Name != "Sheet1"{
		t.Error("Expected Sheet1, get ", wbi.SheetsInfo[0].Name)
	}
	if wbi.SheetsInfo[1].Name != "Sheet2"{
		t.Error("Expected Sheet2, get ", wbi.SheetsInfo[1].Name)
	}
	if wbi.SheetsInfo[2].Name != "Sheet3"{
		t.Error("Expected Sheet3, get ", wbi.SheetsInfo[2].Name)
	}

}

func TestWorkbook(t *testing.T){
	_ = t
	book, err := OpenWorkbook("test.xlsx")
	if err != nil{
		t.Error("Can't open the file, ERR=", err)
	}
	if len(book.Sheets) != 3{
		t.Error("the number of sheets is not 3, get", len(book.Sheets))
	}
	value, err := book.Sheets[0].Cells(0,0)
	if err != nil{
		t.Error("ERR=", err)
	}
	if value != "foo"{
		t.Errorf("expected foo, get %s", value)
	}

	//sheet0
	value, err = book.Sheets[0].Cells(0,1)
	if err != nil{
		t.Error("ERR=", err)
	}
	if value != "bar"{
		t.Errorf("expected bar, get %s", value)
	}

	value, err = book.Sheets[0].Cells(1,0)
	if err != nil{
		t.Error("ERR=", err)
	}
	if value != "1"{
		t.Errorf("expected 1, get %s", value)
	}

	value, err = book.Sheets[0].Cells(1,1)
	if err != nil{
		t.Error("ERR=", err)
	}
	if value != "2"{
		t.Errorf("expected 2, get %s", value)
	}
	//sheet1

	value, err = book.Sheets[1].Cells(0,0)
	if err != nil{
		t.Error("ERR=", err)
	}
	if value != "sheet1"{
		t.Errorf("expected sheet1, get %s", value)
	}

	value, err = book.Sheets[1].Cells(1,0)
	if err != nil{
		t.Error("ERR=", err)
	}
	if value != "3.3"{
		t.Errorf("expected 3.3, get %s", value)
	}

	value, err = book.Sheets[1].Cells(0,1)
	if err != nil{
		t.Error("ERR=", err)
	}
	if value != "zhu"{
		t.Errorf("expected zhu, get %s", value)
	}
	
	value, err = book.Sheets[1].Cells(1,1)
	if err != nil{
		t.Error("ERR=", err)
	}
	if value != "100"{
		t.Errorf("expected 100, get %s", value)
	}

	// err = book.Save("test_readonly.xlsx")
	// if err != nil{
	// 	t.Errorf("Fail to save the xlsx, ERR=%s", err)
	// }

	//change the sheet
	if err = book.Sheets[0].SetCell(0, 0, "朱碧岑"); err != nil{
		t.Errorf("ERR=%s", err)
	}
	//dimesnsion outof

	if err = book.Sheets[0].SetCell(26, 0, "朱碧岑"); err != nil{
		t.Errorf("ERR=%s", err)
	}
	err = book.Save("test_save.xlsx")
	if err != nil{
		t.Errorf("Fail to save the xlsx, ERR=%s", err)
	}
}
	

func TestWorkbook_template(t *testing.T){
	book, err := OpenWorkbook("storedata-template.xlsx")
	if err != nil{
		t.Errorf("Can't open the file, err= %s", err)
	}
	if err = book.Save("storedata-template-save.xlsx"); err != nil{
		t.Error("Fail to save the tmpplae.xlsx")
	}
}
 