package xlsx

type Row struct {
	Cells  []*Cell
	Hidden bool
}

func (r *Row) AddCell() *Cell {
	cell := &Cell{}
	r.Cells = append(r.Cells, cell)
	return cell
}
