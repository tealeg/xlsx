package xlsx

import "reflect"

// Writes an array to row r. Accepts a pointer to array type 'e',
// and writes the number of columns to write, 'cols'. If 'cols' is < 0,
// the entire array will be written if possible. Returns -1 if the 'e'
// doesn't point to an array, otherwise the number of columns written.
func (r *Row) WriteSlice(e interface{}, cols int) int {
	if cols == 0 {
		return cols
	}

	t := reflect.TypeOf(e).Elem()
	if t.Kind() != reflect.Slice { // is 'e' even a slice?
		return -1
	}

	// it's a slice, so open up its values
	v := reflect.ValueOf(e).Elem()

	n := v.Len()
	if cols < n && cols > 0 {
		n = cols
	}

	var i int
	switch t.Elem().Kind() { // underlying type of slice
	case reflect.String:
		for i = 0; i < n; i++ {
			cell := r.AddCell()
			cell.SetString(v.Index(i).Interface().(string))
		}
	case reflect.Int, reflect.Int8,
		reflect.Int16, reflect.Int32:
		for i = 0; i < n; i++ {
			cell := r.AddCell()
			cell.SetInt(v.Index(i).Interface().(int))
		}
	case reflect.Int64:
		for i = 0; i < n; i++ {
			cell := r.AddCell()
			cell.SetInt64(v.Index(i).Interface().(int64))
		}
	case reflect.Bool:
		for i = 0; i < n; i++ {
			cell := r.AddCell()
			cell.SetBool(v.Index(i).Interface().(bool))
		}
	case reflect.Float64, reflect.Float32:
		for i = 0; i < n; i++ {
			cell := r.AddCell()
			cell.SetFloat(v.Index(i).Interface().(float64))
		}
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
		f := v.Field(i).Kind()
		cell := r.AddCell()

		switch f {
		case reflect.Int, reflect.Int8,
			reflect.Int16, reflect.Int32:
			cell.SetInt(v.Field(i).Interface().(int))
		case reflect.Int64:
			cell.SetInt64(v.Field(i).Interface().(int64))
		case reflect.String:
			cell.SetString(v.Field(i).Interface().(string))
		case reflect.Float64, reflect.Float32:
			cell.SetFloat(v.Field(i).Interface().(float64))
		case reflect.Bool:
			cell.SetBool(v.Field(i).Interface().(bool))
		default:
			k-- // nothing set so reset to previous
		}
	}

	return k
}
