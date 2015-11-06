package xlsx

import (
	"fmt"
	"reflect"
)

// Writes an array to row r. Accepts a pointer to array type 'e',
// and writes the number of columns to write, 'cols'. If 'cols' is < 0,
// the entire array will be written if possible. Returns -1 if the 'e'
// doesn't point to an array, otherwise the number of columns written.
func (r *Row) WriteSlice(e interface{}, cols int) int {
	if cols == 0 {
		return cols
	}

	// it's a slice, so open up its values
	v := reflect.ValueOf(e).Elem()
	if v.Kind() != reflect.Slice { // is 'e' even a slice?
		return -1
	}

	n := v.Len()
	if cols < n && cols > 0 {
		n = cols
	}

	var setCell func(reflect.Value)
	setCell = func(val reflect.Value) {
		switch t := val.Interface().(type) {
		case int, int8, int16, int32:
			cell := r.AddCell()
			cell.SetInt(t.(int))
		case int64:
			cell := r.AddCell()
			cell.SetInt64(t)
		case string:
			cell := r.AddCell()
			cell.SetString(t)
		case float32, float64:
			cell := r.AddCell()
			cell.SetFloat(t.(float64))
		case bool:
			cell := r.AddCell()
			cell.SetBool(t)
		case fmt.Stringer:
			cell := r.AddCell()
			cell.SetString(t.String())
		default:
			if val.Kind() == reflect.Interface {
				setCell(reflect.ValueOf(t))
			}
		}
	}

	var i int
	for i = 0; i < n; i++ {
		setCell(v.Index(i))
	}
	return i
}

// Writes a struct to row r. Accepts a pointer to struct type 'e',
// and the number of columns to write, `cols`. If 'cols' is < 0,
// the entire struct will be written if possible. Returns -1 if the 'e'
// doesn't point to a struct, otherwise the number of columns written
func (r *Row) WriteStruct(e interface{}, cols int) int {
	if cols == 0 {
		return cols
	}

	v := reflect.ValueOf(e).Elem()
	if v.Kind() != reflect.Struct {
		return -1 // bail if it's not a struct
	}

	n := v.NumField() // number of fields in struct
	if cols < n && cols > 0 {
		n = cols
	}

	var k int
	for i := 0; i < n; i, k = i+1, k+1 {
		cell := r.AddCell()

		switch t := v.Field(i).Interface().(type) {
		case int, int8, int16, int32:
			cell.SetInt(t.(int))
		case int64:
			cell.SetInt64(t)
		case string:
			cell.SetString(t)
		case float32, float64:
			cell.SetFloat(t.(float64))
		case bool:
			cell.SetBool(t)
		case fmt.Stringer:
			cell.SetString(t.String())
		default:
			k-- // nothing set so reset to previous
		}
	}

	return k
}
