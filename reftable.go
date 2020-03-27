package xlsx

type plainTextOrRichText struct {
	plainText  string
	isRichText bool
	richText   []RichTextRun
}

type RefTable struct {
	indexedStrings []plainTextOrRichText
	knownStrings   map[string]int
	knownRichTexts map[string][]int
	isWrite        bool
}

// NewSharedStringRefTable creates a new, empty RefTable.
func NewSharedStringRefTable() *RefTable {
	rt := RefTable{}
	rt.knownStrings = make(map[string]int)
	rt.knownRichTexts = make(map[string][]int)
	return &rt
}

// MakeSharedStringRefTable takes an xlsxSST struct and converts
// it's contents to an slice of strings used to refer to string values
// by numeric index - this is the model used within XLSX worksheet (a
// numeric reference is stored to a shared cell value).
func MakeSharedStringRefTable(source *xlsxSST) *RefTable {
	reftable := NewSharedStringRefTable()
	reftable.isWrite = false
	for _, si := range source.SI {
		if len(si.R) > 0 {
			richText := xmlToRichText(si.R)
			reftable.AddRichText(richText)
		} else {
			reftable.AddString(si.T.getText())
		}
	}
	return reftable
}

// makeXlsxSST takes a RefTable and returns and
// equivalent xlsxSST representation.
func (rt *RefTable) makeXLSXSST() xlsxSST {
	sst := xlsxSST{}
	sst.Count = len(rt.indexedStrings)
	sst.UniqueCount = sst.Count
	for _, ref := range rt.indexedStrings {
		si := xlsxSI{}
		if ref.isRichText {
			si.R = richTextToXml(ref.richText)
		} else {
			si.T = &xlsxT{Text: ref.plainText}
		}
		sst.SI = append(sst.SI, si)
	}
	return sst
}

// ResolveSharedString looks up a string value or the rich text by numeric index from
// a provided reference table (just a slice of strings in the correct order).
// If the rich text was found, non-empty slice will be returned in richText.
// This function only exists to provide clarity of purpose via it's name.
func (rt *RefTable) ResolveSharedString(index int) (plainText string, richText []RichTextRun) {
	ptrt := rt.indexedStrings[index]
	if ptrt.isRichText {
		richText = ptrt.richText
	} else {
		plainText = ptrt.plainText
	}
	return
}

// AddString adds a string to the reference table and return it's
// numeric index.  If the string already exists then it simply returns
// the existing index.
func (rt *RefTable) AddString(str string) int {
	if rt.isWrite {
		index, ok := rt.knownStrings[str]
		if ok {
			return index
		}
	}
	ptrt := plainTextOrRichText{plainText: str, isRichText: false}
	rt.indexedStrings = append(rt.indexedStrings, ptrt)
	index := len(rt.indexedStrings) - 1
	rt.knownStrings[str] = index
	return index
}

// AddRichText adds a set of rich text to the reference table and return it's
// numeric index.  If a set of rich text already exists then it simply returns
// the existing index.
func (rt *RefTable) AddRichText(r []RichTextRun) int {
	plain := richTextToPlainText(r)
	if rt.isWrite {
		indices, ok := rt.knownRichTexts[plain]
		if ok {
			for _, index := range indices {
				if areRichTextsEqual(rt.indexedStrings[index].richText, r) {
					return index
				}
			}
		}
	}
	ptrt := plainTextOrRichText{isRichText: true}
	ptrt.richText = append(ptrt.richText, r...)
	rt.indexedStrings = append(rt.indexedStrings, ptrt)
	index := len(rt.indexedStrings) - 1
	rt.knownRichTexts[plain] = append(rt.knownRichTexts[plain], index)
	return index
}

func areRichTextsEqual(r1 []RichTextRun, r2 []RichTextRun) bool {
	if len(r1) != len(r2) {
		return false
	}
	for i, rt1 := range r1 {
		rt2 := r2[i]
		if !rt1.Equals(&rt2) {
			return false
		}
	}
	return true
}

func (rt *RefTable) Length() int {
	return len(rt.indexedStrings)
}
