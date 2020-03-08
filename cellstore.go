package xlsx

import "fmt"

// CellStore provides an interface for interacting with backend cell
// storage. For example, this allows us, as required, to persist cells
// to some store instead of holding them in memory.  This tactic
// allows us a degree of control around the characteristics of our
// programs when handling large spreadsheets - we can choose to run
// more slowly, but without exhausting system memory.
//
// If you wish to implement a custom CellStore you must not only
// support this interface, but also a CellStoreConstructor and a
// FileOption that set's the File's cellStoreConstructor to the right
// constructor.
type CellStore interface {
	ReadRow(key string) (*Row, error)
	WriteRow(r *Row) error
	MoveRow(r *Row, newIndex int) error
	RemoveRow(key string) error
	Close() error
}

// CellStoreConstructor defines the signature of a function that will
// be used to return a new instance of the CellStore implmentation,
// you must pass this into
type CellStoreConstructor func() (CellStore, error)

// CellVisitorFunc defines the signature of a function that will be
// called when visiting a Cell using CellStore.ForEachInRow.
type CellVisitorFunc func(c *Cell) error

// RowNotFoundError is an Error that should be returned by a
// RowStore implementation if a call to ReadRow is made with a key
// that doesn't correspond to any persisted Row.
type RowNotFoundError struct {
	key    string
	reason string
}

// NewRowNotFoundError creates a new RowNotFoundError, capturing the Row key and the reason this key could not be found.
//
func NewRowNotFoundError(key, reason string) *RowNotFoundError {
	return &RowNotFoundError{key, reason}
}

// Error returns a human-readable description of the failure to find a Row.  It makes RowNotFoundError comply with the Error interface.
func (cnfe RowNotFoundError) Error() string {
	return fmt.Sprintf("Row %q not found. %s", cnfe.key, cnfe.reason)
}
