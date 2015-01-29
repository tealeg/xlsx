package xlsx

import "reflect"

// Writes a struct to row r. Accepts a pointer to struct type 'e',
// and the number of columns to write, `cols`. Returns -1 if the 'e'
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
	if cols < n {
		n = cols
	}

	var k int
	for i := 0; i < n; i, k = i+1, k+1 {
		f := v.Field(i).Kind()
		cell := r.AddCell()

		switch f {
		case reflect.Int, reflect.Int8, reflect.Int16,
			reflect.Int32, reflect.Int64:
			cell.SetInt(v.Field(i).Interface().(int))
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
