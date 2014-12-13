package xlsx

// Default column width in excel
const ColWidth = 9.5

type Col struct {
	Min       int
	Max       int
	Hidden    bool
	Width     float64
	Collapsed bool
	//	Style     int
}
