package xlsx

type Row struct {
	Cells []*Cell
}

func (r *Row) AddCell() *Cell {
	cell := &Cell{}
	r.Cells = append(r.Cells, cell)
	return cell
}
