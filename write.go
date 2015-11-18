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
		case fmt.Stringer: // check Stringer first
			cell := r.AddCell()
			cell.SetString(t.String())
		default:
			switch val.Kind() { // underlying type of slice
			case reflect.String:
				cell := r.AddCell()
				cell.SetString(t.(string))
			case reflect.Int, reflect.Int8,
				reflect.Int16, reflect.Int32:
				cell := r.AddCell()
				cell.SetInt(t.(int))
			case reflect.Int64:
				cell := r.AddCell()
				cell.SetInt64(t.(int64))
			case reflect.Bool:
				cell := r.AddCell()
				cell.SetBool(t.(bool))
			case reflect.Float64, reflect.Float32:
				cell := r.AddCell()
				cell.SetFloat(t.(float64))
			case reflect.Interface:
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
		f := v.Field(i)
		cell := r.AddCell()

		switch t := f.Interface().(type) {
		case fmt.Stringer: // check Stringer first
			cell.SetString(t.String())
		default:
			switch f.Kind() {
			case reflect.Int, reflect.Int8,
				reflect.Int16, reflect.Int32:
				cell.SetInt(t.(int))
			case reflect.Int64:
				cell.SetInt64(t.(int64))
			case reflect.String:
				cell.SetString(t.(string))
			case reflect.Float64, reflect.Float32:
				cell.SetFloat(t.(float64))
			case reflect.Bool:
				cell.SetBool(t.(bool))
			default:
				k-- // nothing set so reset to previous
			}
		}
	}

	return k
}
