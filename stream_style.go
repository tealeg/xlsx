package xlsx

// StreamStyle has style and formatting information.
// Used to store a style for streaming
type StreamStyle struct {
	xNumFmtId	int
	style 		*Style
}

var DefaultStringStyle *StreamStyle
var DefaultStringBoldStyle *StreamStyle
var DefaultStringItalicStyle *StreamStyle
var DefaultStringUnderlinedStyle *StreamStyle

var DefaultNumericStyle *StreamStyle
var DefaultNumericBoldStyle *StreamStyle
var DefaultNumericItalicStyle *StreamStyle
var DefaultNumericUnderlinedStyle *StreamStyle

var DefaultStyles []*StreamStyle

func init(){
	// default string styles
	DefaultStringStyle = &StreamStyle{
		xNumFmtId: 0,
		style: NewStyle(),
	}
	DefaultStringBoldStyle = &StreamStyle{
		xNumFmtId: 0,
		style: NewStyle(),
	}
	DefaultStringBoldStyle.style.Font.Bold = true
	DefaultStringItalicStyle = &StreamStyle{
		xNumFmtId: 0,
		style: NewStyle(),
	}
	DefaultStringItalicStyle.style.Font.Italic = true
	DefaultStringUnderlinedStyle = &StreamStyle{
		xNumFmtId: 0,
		style: NewStyle(),
	}
	DefaultStringUnderlinedStyle.style.Font.Underline = true

	DefaultStyles = append(DefaultStyles, DefaultStringStyle)
	DefaultStyles = append(DefaultStyles, DefaultStringBoldStyle)
	DefaultStyles = append(DefaultStyles, DefaultStringItalicStyle)
	DefaultStyles = append(DefaultStyles, DefaultStringUnderlinedStyle)

	// default string styles
	DefaultNumericStyle = &StreamStyle{
		xNumFmtId: 1,
		style: NewStyle(),
	}
	DefaultNumericBoldStyle = &StreamStyle{
		xNumFmtId: 1,
		style: NewStyle(),
	}
	DefaultNumericBoldStyle.style.Font.Bold = true
	DefaultNumericItalicStyle = &StreamStyle{
		xNumFmtId: 1,
		style: NewStyle(),
	}
	DefaultNumericItalicStyle.style.Font.Italic = true
	DefaultNumericUnderlinedStyle = &StreamStyle{
		xNumFmtId: 1,
		style: NewStyle(),
	}
	DefaultNumericUnderlinedStyle.style.Font.Underline = true

	DefaultStyles = append(DefaultStyles, DefaultNumericStyle)
	DefaultStyles = append(DefaultStyles, DefaultNumericBoldStyle)
	DefaultStyles = append(DefaultStyles, DefaultNumericItalicStyle)
	DefaultStyles = append(DefaultStyles, DefaultNumericUnderlinedStyle)


}

// MakeStyle creates a new StreamStyle and add it to the styles that will be streamed
// This function returns a reference to the created StreamStyle
func MakeStyle(formatStyleId int, font Font, fill Fill, alignment Alignment, border Border) *StreamStyle {
	newStyle := NewStyle()

	newStyle.Font = font
	newStyle.Fill = fill
	newStyle.Alignment = alignment
	newStyle.Border = border

	newStyle.ApplyFont = true
	newStyle.ApplyFill = true
	newStyle.ApplyAlignment = true
	newStyle.ApplyBorder = true

	newStreamStyle := &StreamStyle{
		xNumFmtId: 	formatStyleId,
		style: 		newStyle,
	}

	DefaultStyles = append(DefaultStyles, newStreamStyle)
	return newStreamStyle
}

