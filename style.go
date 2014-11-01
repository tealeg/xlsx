package xlsx

// Style is a high level structure intended to provide user access to
// the contents of Style within an XLSX file.
type Style struct {
	Border Border
	Fill   Fill
	Font   Font
}

func NewStyle() *Style {
	return &Style{}
}

// Border is a high level structure intended to provide user access to
// the contents of Border Style within an Sheet.
type Border struct {
	Left   string
	Right  string
	Top    string
	Bottom string
}

// Fill is a high level structure intended to provide user access to
// the contents of background and foreground color index within an Sheet.
type Fill struct {
	PatternType string
	BgColor     string
	FgColor     string
}

func NewFill(patternType, fgColor, bgColor string) *Fill {
	return &Fill{PatternType: patternType, FgColor: fgColor, BgColor: bgColor}
}

type Font struct {
	Size    int
	Name    string
	Family  int
	Charset int
}

func NewFont(size int, name string) *Font {
	return &Font{Size: size, Name: name}
}
