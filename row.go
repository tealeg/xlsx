package xlsx

import "strconv"

type Row struct {
	Cells        []*Cell
	Hidden       bool
	Sheet        *Sheet
	Height       float64
	OutlineLevel uint8
	isCustom     bool
}

func (r *Row) SetHeight(ht float64) {
	r.Height = ht
	r.isCustom = true
}

func (r *Row) SetHeightCM(ht float64) {
	r.Height = ht * 28.3464567 // Convert CM to postscript points
	r.isCustom = true
}

func (r *Row) AddCell() *Cell {
	cell := NewCell(r)
	r.Cells = append(r.Cells, cell)
	r.Sheet.maybeAddCol(len(r.Cells))
	return cell
}

// AddStreamCell takes as input a StreamCell, creates a new Cell from it,
// and appends the new cell to the row.
func (r *Row) AddStreamCell(streamCell StreamCell) {
	cell := NewCell(r)
	cell.Value = streamCell.cellData
	cell.style = streamCell.cellStyle.style
	cell.NumFmt = strconv.Itoa(streamCell.cellStyle.xNumFmtId)
	cell.cellType = streamCell.cellType
	r.Cells = append(r.Cells, cell)
	r.Sheet.maybeAddCol(len(r.Cells))
}


