package xlsx
import (
	"testing"
	"bytes"
	//"os"
	)

func TestNewSharedStringsTable(t *testing.T){
	data := bytes.NewBufferString(`
      <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
			 count="1448" uniqueCount="463">
      <si><t>查看总人次</t> </si>
	  <si><t>3日</t>     </si>
      <si><t>4日</t>     </si>
      <si>
          <t>分析</t> 
          <phoneticPr fontId="9" type="noConversion" /> 
		</si></sst>`)

	sst, err := newSharedStringsTable(data)
	if err != nil{
		t.Fatalf("Can't New the shared string table, ERR=%s", err)
	}
	if sst.Count != "1448"{
		t.Errorf("Expected sst.count = 1448, get %s", sst.Count)
	}
	
	if sst.UniqueCount != "463"{
		t.Errorf("Expected sst.unqueCount == 463, get %s", sst.UniqueCount)
	}
	if sst.SI[0].T != "查看总人次"{
		t.Errorf("Expected 查看总人次, get %s", sst.SI[0].T)
	}
	if sst.SI[1].T != "3日"{
		t.Errorf("Expected 3日, get %s", sst.SI[1].T)
	}
	if sst.SI[2].T != "4日"{
		t.Errorf("Expected 4日, get %s", sst.SI[2].T)
	}
	if sst.SI[3].T != "分析"{
		t.Errorf("Expected 分析, get %s", sst.SI[3].T)
	}

	if sst.SI[3].PhoneticPr.FontId != "9"{
		t.Errorf("Exptected sst.SI[4].PhoneticPr.FontId == 9, get %s", sst.SI[3].PhoneticPr.FontId)
	}

	if sst.SI[3].PhoneticPr.Type != "noConversion"{
		t.Errorf("Exptected sst.SI[4].PhoneticPr.Type == noConversion, get %s", sst.SI[3].PhoneticPr.Type)
	}

	//sst.Save(os.Stdout)

	index, _ := sst.getIndex("朱碧岑")
	if index != 4{
		t.Errorf("Expected 4, get %s", index)
	}

	if sst.Count != "1449"{
		t.Errorf("Expected sst.count = 1449, get %s", sst.Count)
	}
	
	if sst.UniqueCount != "464"{
		t.Errorf("Expected sst.unqueCount == 464, get %s", sst.UniqueCount)
	}
	//sst.WriteTo(os.Stdout)
} 


		
		
		

