package xlsx

// StreamStyle has style and formatting information.
// Used to store a style for streaming
type StreamStyle struct {
	xNumFmtId int
	style     *Style
}

const (
	GeneralFormat              = 0
	IntegerFormat              = 1
	DecimalFormat              = 2
	DateFormat_dd_mm_yy        = 14
	DateTimeFormat_d_m_yy_h_mm = 22
)

var (
	StreamStyleFromColumn StreamStyle

	StreamStyleDefaultString    StreamStyle
	StreamStyleBoldString       StreamStyle
	StreamStyleItalicString     StreamStyle
	StreamStyleUnderlinedString StreamStyle

	StreamStyleDefaultInteger    StreamStyle
	StreamStyleBoldInteger       StreamStyle
	StreamStyleItalicInteger     StreamStyle
	StreamStyleUnderlinedInteger StreamStyle

	StreamStyleDefaultDate StreamStyle

	StreamStyleDefaultDecimal StreamStyle
)
var (
	FontBold       *Font
	FontItalic     *Font
	FontUnderlined *Font
)
var (
	FillGreen *Fill
	FillRed   *Fill
	FillWhite *Fill
)

func init() {
	// Init Fonts
	FontBold = DefaultFont()
	FontBold.Bold = true

	FontItalic = DefaultFont()
	FontItalic.Italic = true

	FontUnderlined = DefaultFont()
	FontUnderlined.Underline = true

	// Init Fills
	FillGreen = NewFill(Solid_Cell_Fill, RGB_Light_Green, RGB_White)
	FillRed = NewFill(Solid_Cell_Fill, RGB_Light_Red, RGB_White)
	FillWhite = DefaultFill()

	// Init default string styles
	StreamStyleDefaultString = MakeStringStyle(DefaultFont(), DefaultFill(), DefaultAlignment(), DefaultBorder())
	StreamStyleBoldString = MakeStringStyle(FontBold, DefaultFill(), DefaultAlignment(), DefaultBorder())
	StreamStyleItalicString = MakeStringStyle(FontItalic, DefaultFill(), DefaultAlignment(), DefaultBorder())
	StreamStyleUnderlinedString = MakeStringStyle(FontUnderlined, DefaultFill(), DefaultAlignment(), DefaultBorder())

	// Init default Integer styles
	StreamStyleDefaultInteger = MakeIntegerStyle(DefaultFont(), DefaultFill(), DefaultAlignment(), DefaultBorder())
	StreamStyleBoldInteger = MakeIntegerStyle(FontBold, DefaultFill(), DefaultAlignment(), DefaultBorder())
	StreamStyleItalicInteger = MakeIntegerStyle(FontItalic, DefaultFill(), DefaultAlignment(), DefaultBorder())
	StreamStyleUnderlinedInteger = MakeIntegerStyle(FontUnderlined, DefaultFill(), DefaultAlignment(), DefaultBorder())

	StreamStyleDefaultDate = MakeDateStyle(DefaultFont(), DefaultFill(), DefaultAlignment(), DefaultBorder())

	StreamStyleDefaultDecimal = MakeDecimalStyle(DefaultFont(), DefaultFill(), DefaultAlignment(), DefaultBorder())

	DefaultStringStreamingCellMetadata = StreamingCellMetadata{CellTypeString, StreamStyleDefaultString}
	DefaultNumericStreamingCellMetadata = StreamingCellMetadata{CellTypeNumeric, StreamStyleDefaultString}
	DefaultDecimalStreamingCellMetadata = StreamingCellMetadata{CellTypeNumeric, StreamStyleDefaultDecimal}
	DefaultIntegerStreamingCellMetadata = StreamingCellMetadata{CellTypeNumeric, StreamStyleDefaultInteger}
	DefaultDateStreamingCellMetadata = StreamingCellMetadata{CellTypeDate, StreamStyleDefaultDate}
}

// MakeStyle creates a new StreamStyle and add it to the styles that will be streamed.
func MakeStyle(numFormatId int, font *Font, fill *Fill, alignment *Alignment, border *Border) StreamStyle {
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
		xNumFmtId: numFormatId,
		style:     newStyle,
	}

	return newStreamStyle
}

// MakeStringStyle creates a new style that can be used on cells with string data.
// If used on other data the formatting might be wrong.
func MakeStringStyle(font *Font, fill *Fill, alignment *Alignment, border *Border) StreamStyle {
	return MakeStyle(GeneralFormat, font, fill, alignment, border)
}

// MakeIntegerStyle creates a new style that can be used on cells with integer data.
// If used on other data the formatting might be wrong.
func MakeIntegerStyle(font *Font, fill *Fill, alignment *Alignment, border *Border) StreamStyle {
	return MakeStyle(IntegerFormat, font, fill, alignment, border)
}

// MakeDecimalStyle creates a new style that can be used on cells with decimal numeric data.
// If used on other data the formatting might be wrong.
func MakeDecimalStyle(font *Font, fill *Fill, alignment *Alignment, border *Border) StreamStyle {
	return MakeStyle(DecimalFormat, font, fill, alignment, border)
}

// MakeDateStyle creates a new style that can be used on cells with Date data.
// The formatting used is: dd_mm_yy
// If used on other data the formatting might be wrong.
func MakeDateStyle(font *Font, fill *Fill, alignment *Alignment, border *Border) StreamStyle {
	return MakeStyle(DateFormat_dd_mm_yy, font, fill, alignment, border)
}
