package xlsx

import (
	"fmt"
	"strings"
)

type MemoryRow struct {
	row    *Row
	maxCol int
	cells  []*Cell
}

func makeMemoryRow(sheet *Sheet) *MemoryRow {
	mr := &MemoryRow{
		row:    new(Row),
		maxCol: -1,
	}
	mr.row.Sheet = sheet
	mr.row.cellStoreRow = mr
	sheet.setCurrentRow(mr.row)
	return mr
}

func (mr *MemoryRow) Updatable() {
	// Do nothing
}

func (mr *MemoryRow) CellUpdatable(c *Cell) {
	// Do nothing
}

func (mr *MemoryRow) AddCell() *Cell {
	cell := newCell(mr.row, mr.maxCol+1)
	mr.PushCell(cell)
	return cell
}

func (mr *MemoryRow) PushCell(c *Cell) {
	if c.num > mr.maxCol {
		mr.maxCol = c.num
	}
	mr.growCellsSlice(c.num + 1)
	mr.cells[c.num] = c
}

func (mr *MemoryRow) growCellsSlice(newSize int) {
	capacity := cap(mr.cells)
	if newSize > capacity {
		newCap := 2 * capacity
		if newSize > newCap {
			newCap = newSize
		}
		capacity = newCap
	}
	newSlice := make([]*Cell, newSize, capacity)
	copy(newSlice, mr.cells)
	mr.cells = newSlice
}

func (mr *MemoryRow) GetCell(colIdx int) *Cell {
	if colIdx >= len(mr.cells) {
		cell := newCell(mr.row, colIdx)
		mr.growCellsSlice(colIdx + 1)

		mr.cells[colIdx] = cell
		return cell
	}

	cell := mr.cells[colIdx]
	if cell == nil {
		cell = newCell(mr.row, colIdx)
		mr.cells[colIdx] = cell
	}
	return cell
}

func (mr *MemoryRow) ForEachCell(cvf CellVisitorFunc, option ...CellVisitorOption) error {
	flags := &cellVisitorFlags{}
	for _, opt := range option {
		opt(flags)
	}
	fn := func(ci int, c *Cell) error {
		if c == nil {
			if flags.skipEmptyCells {
				return nil
			}
			c = mr.GetCell(ci)
		}
		if !c.Modified() && flags.skipEmptyCells {
			return nil
		}
		c.Row = mr.row
		return cvf(c)
	}

	for ci, cell := range mr.cells {
		err := fn(ci, cell)
		if err != nil {
			return err
		}
	}
	cellCount := len(mr.cells)
	if !flags.skipEmptyCells {
		for ci := cellCount; ci < mr.row.Sheet.MaxCol; ci++ {
			c := mr.GetCell(ci)
			err := cvf(c)
			if err != nil {
				return err
			}

		}
	}

	return nil
}

// MaxCol returns the index of the rightmost cell in the row's column.
func (mr *MemoryRow) MaxCol() int {
	return mr.maxCol
}

// CellCount returns the total number of cells in the row.
func (mr *MemoryRow) CellCount() int {
	return mr.maxCol + 1
}

// MemoryCellStore is the default CellStore - it holds all rows and
// cells in system memory.  This is fast, right up until you run out
// of memory ;-)
type MemoryCellStore struct {
	rows map[string]*Row
}

// UseMemoryCellStore is a FileOption that makes all Sheet instances
// for a File use memory as their backing store.  This is the default
// backing store.  You can use this option when you are comfortable
// keeping the contents of each Sheet in memory.  This is faster than
// using a disk backed store, but can easily use a large amount of
// memory and, if you exhaust the available system memory, it'll
// actualy be slower than using a disk backed store (e.g. DiskV).
func UseMemoryCellStore(f *File) {
	f.cellStoreConstructor = NewMemoryCellStore
}

// NewMemoryCellStore returns a pointer to a newly allocated MemoryCellStore
func NewMemoryCellStore() (CellStore, error) {
	cs := &MemoryCellStore{
		rows: make(map[string]*Row),
	}
	return cs, nil
}

// Close is nullOp for the MemoryCellStore, but we have to comply with
// the interface.
func (mcs *MemoryCellStore) Close() error {
	return nil
}

// ReadRow returns a Row identfied by the given key.
func (mcs *MemoryCellStore) ReadRow(key string, s *Sheet) (*Row, error) {
	r, ok := mcs.rows[key]
	if !ok {
		return nil, NewRowNotFoundError(key, "No such row")
	}
	return r, nil
}

// WriteRow pushes the Row to the MemoryCellStore.
func (mcs *MemoryCellStore) WriteRow(r *Row) error {
	if r != nil {
		key := r.key()
		mcs.rows[key] = r
	}
	return nil
}

// MoveRow moves the persisted Row's position in the sheet.
func (mcs *MemoryCellStore) MoveRow(r *Row, index int) error {
	oldKey := r.key()
	r.num = index
	newKey := r.key()
	if _, exists := mcs.rows[newKey]; exists {
		return fmt.Errorf("Target index for row (%d) would overwrite a row already exists", index)
	}
	mcs.rows[newKey] = r
	delete(mcs.rows, oldKey)
	return nil
}

// RemoveRow removes a row from the sheet, it doesn't specifically
// move any following rows, leaving this decision to the user.
func (mcs *MemoryCellStore) RemoveRow(key string) error {
	r, ok := mcs.rows[key]
	if ok {
		r.Sheet.setCurrentRow(nil)
		delete(mcs.rows, key)
	}
	return nil
}

// MakeRowWithLen returns an empty Row, with a preconfigured starting length.
func (mcs *MemoryCellStore) MakeRowWithLen(sheet *Sheet, len int) *Row {
	mr := makeMemoryRow(sheet)
	mr.maxCol = len - 1
	mr.growCellsSlice(len)
	return mr.row
}

// MakeRow returns an empty Row
func (mcs *MemoryCellStore) MakeRow(sheet *Sheet) *Row {
	return makeMemoryRow(sheet).row
}

// Extract the row key from a provided cell key
func keyToRowKey(key string) string {
	parts := strings.Split(key, ":")
	return parts[0] + ":" + parts[1]
}
