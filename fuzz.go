// +build gofuzz

package xlsx

// Fuzz tests parsing and cell processing
func Fuzz(fuzz []byte) int {
	file, err := OpenBinary(fuzz)
	if err != nil {
		return 0
	}
	for _, sheet := range file.Sheets {
		for _, row := range sheet.Rows {
			for _, cell := range row.Cells {
				cell.String()
			}
		}
	}
	return 1
}
