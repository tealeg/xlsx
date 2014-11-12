package xlsx

type Row struct {
	Cells  []*Cell
	Hidden bool
	sheet  *Sheet
}

func (r *Row) AddCell() *Cell {
	cell := &Cell{}
	r.Cells = append(r.Cells, cell)
	r.sheet.maybeAddCol(len(r.Cells))
	return cell
}
