// +build gofuzz

package xlsx

// Fuzz tests parsing and cell processing
func Fuzz(fuzz []byte) int {
	file, err := OpenBinary(fuzz)
	if err != nil {
		return 0
	}
	for _, sheet := range file.Sheets {
		sheet.ForEachRow(func(r *Row) error {
			return r.ForEachCell(func(c *Cell) error {
				c.String()
				return nil
			})
		})
	}
	return 1
}
