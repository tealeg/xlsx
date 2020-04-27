package xlsx

// sPtr simply returns a pointer to the provided string.
func sPtr(s string) *string {
	return &s
}

func iPtr(i int) *int {
	return &i
}

func fPtr(f float64) *float64 {
	return &f
}

func bPtr(b bool) *bool {
	return &b
}

func u8Ptr(u uint8) *uint8 {
	return &u
}
