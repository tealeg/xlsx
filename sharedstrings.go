package xlsx


type XLSXSST struct {
	Count       string "attr"
	UniqueCount string "attr"
	SI          []XLSXSI
}


type XLSXSI struct {
	T XLSXT
}


type XLSXT struct {
	Data string "chardata"
}
