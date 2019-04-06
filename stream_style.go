package xlsx

// StreamStyle has style and formatting information.
// Used to store a style for streaming
type StreamStyle struct {
	xNumFmtId	int
	style 		*Style
}

const (
	GeneralFormat				= 0
	IntegerFormat				= 1
	DecimalFormat 				= 2
	DateFormat_dd_mm_yy 		= 14
	DateTimeFormat_d_m_yy_h_mm 	= 22
)

var Strings StreamStyle
var BoldStrings StreamStyle
var ItalicStrings StreamStyle
var UnderlinedStrings StreamStyle

var Integers StreamStyle
var BoldIntegers StreamStyle
var ItalicIntegers StreamStyle
var UnderlinedIntegers StreamStyle

// var DefaultStyles []StreamStyle


var Bold *Font
var Italic *Font
var Underlined *Font

var GreenCell *Fill
var RedCell *Fill
var WhiteCel *Fill

func init(){
	// Init Fonts
	Bold = DefaultFont()
	Bold.Bold = true

	Italic = DefaultFont()
	Italic.Italic = true

	Underlined = DefaultFont()
	Underlined.Underline = true

	// Init Fills
	GreenCell = NewFill(Solid_Cell_Fill, RGB_Light_Green, RGB_White)
	RedCell = NewFill(Solid_Cell_Fill, RGB_Light_Red, RGB_White)
	WhiteCel = DefaultFill()

	// Init default string styles
	Strings = MakeStringStyle(DefaultFont(), DefaultFill(), DefaultAlignment(), DefaultBorder())
	BoldStrings = MakeStringStyle(Bold, DefaultFill(), DefaultAlignment(), DefaultBorder())
	ItalicStrings = MakeStringStyle(Italic, DefaultFill(), DefaultAlignment(), DefaultBorder())
	UnderlinedStrings = MakeStringStyle(Underlined, DefaultFill(), DefaultAlignment(), DefaultBorder())

	//DefaultStyles = append(DefaultStyles, Strings)
	//DefaultStyles = append(DefaultStyles, BoldStrings)
	//DefaultStyles = append(DefaultStyles, ItalicStrings)
	//DefaultStyles = append(DefaultStyles, UnderlinedStrings)

	// Init default Integer styles
	Integers = MakeIntegerStyle(DefaultFont(), DefaultFill(), DefaultAlignment(), DefaultBorder())
	BoldIntegers = MakeIntegerStyle(Bold, DefaultFill(), DefaultAlignment(), DefaultBorder())
	ItalicIntegers = MakeIntegerStyle(Italic, DefaultFill(), DefaultAlignment(), DefaultBorder())
	UnderlinedIntegers = MakeIntegerStyle(Underlined, DefaultFill(), DefaultAlignment(), DefaultBorder())

	//DefaultStyles = append(DefaultStyles, Integers)
	//DefaultStyles = append(DefaultStyles, BoldIntegers)
	//DefaultStyles = append(DefaultStyles, ItalicIntegers)
	//DefaultStyles = append(DefaultStyles, UnderlinedIntegers)
}

// MakeStyle creates a new StreamStyle and add it to the styles that will be streamed
// This function returns a reference to the created StreamStyle
func MakeStyle(formatStyleId int, font *Font, fill *Fill, alignment *Alignment, border *Border) StreamStyle {
	newStyle := NewStyle()

	newStyle.Font = *font
	newStyle.Fill = *fill
	newStyle.Alignment = *alignment
	newStyle.Border = *border

	newStyle.ApplyFont = true
	newStyle.ApplyFill = true
	newStyle.ApplyAlignment = true
	newStyle.ApplyBorder = true

	newStreamStyle := StreamStyle{
		xNumFmtId: 	formatStyleId,
		style: 		newStyle,
	}

	// DefaultStyles = append(DefaultStyles, newStreamStyle)
	return newStreamStyle
}

// MakeStringStyle creates a new style that can be used on cells with string data.
// If used on other data the formatting might be wrong.
func MakeStringStyle(font *Font, fill *Fill, alignment *Alignment, border *Border) StreamStyle{
	return MakeStyle(GeneralFormat, font, fill, alignment, border)
}

// MakeIntegerStyle creates a new style that can be used on cells with integer data.
// If used on other data the formatting might be wrong.
func MakeIntegerStyle(font *Font, fill *Fill, alignment *Alignment, border *Border) StreamStyle{
	return MakeStyle(IntegerFormat, font, fill, alignment, border)
}

// MakeDecimalStyle creates a new style that can be used on cells with decimal numeric data.
// If used on other data the formatting might be wrong.
func MakeDecimalStyle(font *Font, fill *Fill, alignment *Alignment, border *Border) StreamStyle{
	return MakeStyle(DecimalFormat, font, fill, alignment, border)
}

// MakeDateStyle creates a new style that can be used on cells with Date data.
// The formatting used is: dd_mm_yy
// If used on other data the formatting might be wrong.
func MakeDateStyle(font *Font, fill *Fill, alignment *Alignment, border *Border) StreamStyle{
	return MakeStyle(DateFormat_dd_mm_yy, font, fill, alignment, border)
}