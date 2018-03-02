package xlsx

import "fmt"

func ExampleRow_ReadStruct() {
	//example type
	type structTest struct {
		IntVal     int     `xlsx:"0"`
		StringVal  string  `xlsx:"1"`
		FloatVal   float64 `xlsx:"2"`
		IgnoredVal int     `xlsx:"-"`
		BoolVal    bool    `xlsx:"4"`
	}
	structVal := structTest{
		IntVal:     16,
		StringVal:  "heyheyhey :)!",
		FloatVal:   3.14159216,
		IgnoredVal: 7,
		BoolVal:    true,
	}
	//create a new xlsx file and write a struct
	//in a new row
	f := NewFile()
	sheet, _ := f.AddSheet("TestRead")
	row := sheet.AddRow()
	row.WriteStruct(&structVal, -1)

	//read the struct from the same row
	readStruct := &structTest{}
	err := row.ReadStruct(readStruct)
	if err != nil {
		panic(err)
	} else {
		fmt.Println(readStruct)
	}
}
