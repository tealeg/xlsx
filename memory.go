package xlsx

import (
	"fmt"
	"strings"
)

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
func (mcs *MemoryCellStore) ReadRow(key string) (*Row, error) {
	r, ok := mcs.rows[key]
	if !ok {
		return nil, NewRowNotFoundError(key, "No such row")
	}
	return r, nil
}

// WriteRow pushes the Row to the MemoryCellStore.
func (mcs *MemoryCellStore) WriteRow(r *Row) error {
	if r != nil {
		mcs.rows[r.key()] = r
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
	delete(mcs.rows, oldKey)
	mcs.rows[newKey] = r
	return nil
}

// RemoveRow removes a row from the sheet, it doesn't specifically
// move any following rows, leaving this decision to the user.
func (mcs *MemoryCellStore) RemoveRow(key string) error {
	delete(mcs.rows, key)
	return nil
}

// Extract the row key from a provided cell key
func keyToRowKey(key string) string {
	parts := strings.Split(key, ":")
	return parts[0] + ":" + parts[1]
}
