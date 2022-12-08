package xlsx

type StreamRow struct {
	Cells    []StreamCell
	Height   float64
	isCustom bool
}

func (r *StreamRow) SetHeight(ht float64) {
	r.Height = ht
	r.isCustom = true
}

func (r *StreamRow) SetHeightCM(ht float64) {
	r.Height = ht * cmToPostscriptPts
	r.isCustom = true
}

func NewStreamRow(cells []StreamCell) StreamRow {
	return StreamRow{Cells: cells}
}
