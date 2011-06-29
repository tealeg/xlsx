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


func MakeSharedStringRefTable(source *XLSXSST) []string {
	reftable := make([]string, len(source.SI))
	for i, si := range source.SI {
		reftable[i] = si.T.Data
	}
	return reftable
}

func ResolveSharedString(reftable []string, index int) string {
	return reftable[index]
}


