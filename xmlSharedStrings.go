package xlsx

import (
	"encoding/xml"
	"errors"
	"strings"
)

// xlsxSST directly maps the sst element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main currently
// I have not checked this for completeness - it does as much as I need.
type xlsxSST struct {
	XMLName     xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main sst"`
	Count       int      `xml:"count,attr"`
	UniqueCount int      `xml:"uniqueCount,attr"`
	SI          []xlsxSI `xml:"si"`
}

// xlsxSI directly maps the si element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked this for completeness - it does as
// much as I need.
type xlsxSI struct {
	T *xlsxT  `xml:"t"`
	R []xlsxR `xml:"r"`
}

// xlsxR directly maps the r element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked this for completeness - it does as
// much as I need.
type xlsxR struct {
	RPr *xlsxRunProperties `xml:"rPr"`
	T   xlsxT              `xml:"t"`
}

// xlsxRunProperties directly maps the rPr element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main
type xlsxRunProperties struct {
	RFont     *xlsxVal      `xml:"rFont"`
	Charset   *xlsxIntVal   `xml:"charset"`
	Family    *xlsxIntVal   `xml:"family"`
	B         xlsxBoolProp  `xml:"b"`
	I         xlsxBoolProp  `xml:"i"`
	Strike    xlsxBoolProp  `xml:"strike"`
	Outline   xlsxBoolProp  `xml:"outline"`
	Shadow    xlsxBoolProp  `xml:"shadow"`
	Condense  xlsxBoolProp  `xml:"condense"`
	Extend    xlsxBoolProp  `xml:"extend"`
	Color     *xlsxColor    `xml:"color"`
	Sz        *xlsxFloatVal `xml:"sz"`
	U         *xlsxVal      `xml:"u"`
	VertAlign *xlsxVal      `xml:"vertAlign"`
	Scheme    *xlsxVal      `xml:"scheme"`
}

// xlsxBoolProp handles "CT_BooleanProperty" type which is declared in the XML Schema of Office Open XML.
// XML attribute "val" is optional. If "val" was omitted, the property value becomes "true".
// On the serialization, the struct which has "true" will be serialized an empty XML tag without "val" attributes,
// and the struct which has "false" will not be serialized.
type xlsxBoolProp struct {
	Val bool `xml:"val,attr"`
}

// MarshalXML implements xml.Marshaler interface for xlsxBoolProp
func (b *xlsxBoolProp) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if b.Val {
		if err := e.EncodeToken(start); err != nil {
			return err
		}
		if err := e.EncodeToken(xml.EndElement{Name: start.Name}); err != nil {
			return err
		}
	}
	return nil
}

// UnmarshalXML implements xml.Unmarshaler interface for xlsxBoolProp
func (b *xlsxBoolProp) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	boolVal := true
	for _, attr := range start.Attr {
		if attr.Name.Space == "" && attr.Name.Local == "val" {
			// supports xsd:boolean
			switch attr.Value {
			case "true", "1":
				boolVal = true
			case "false", "0":
				boolVal = false
			default:
				return errors.New(
					"Cannot unmarshal into xlsxBoolProp: \"" +
						attr.Value + "\" is not a valid boolean value")
			}
		}
	}
	b.Val = boolVal
	return d.Skip()
}

// xlsxIntVal is like xlsxVal, except it has an int value
type xlsxIntVal struct {
	Val int `xml:"val,attr"`
}

// xlsxFloatVal is like xlsxVal, except it has a float value
type xlsxFloatVal struct {
	Val float64 `xml:"val,attr"`
}

// xlsxT represents a text. It will be serialized as a XML tag which has character data.
// Attribute xml:space="preserve" will be added to the XML tag if needed.
type xlsxT struct {
	Text string `xml:",chardata"`
}

// MarshalXML implements xml.Marshaler interface for xlsxT
func (t *xlsxT) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if needPreserve(t.Text) {
		attr := xml.Attr{
			Name:  xml.Name{Local: "xml:space"},
			Value: "preserve",
		}
		start.Attr = append(start.Attr, attr)
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if err := e.EncodeToken(xml.CharData(t.Text)); err != nil {
		return err
	}
	if err := e.EncodeToken(xml.EndElement{Name: start.Name}); err != nil {
		return err
	}
	return nil
}

// getText is a nil-safe utility function that gets a string from xlsxT.
// If the pointer of xlsxT was nil, returns an empty string.
func (t *xlsxT) getText() string {
	if t == nil {
		return ""
	}
	return t.Text
}

// needPreserve determines whether xml:space="preserve" is needed.
func needPreserve(s string) bool {
	if len(s) == 0 {
		return false
	}
	// Note:
	// xml:space="preserve" is not needed for CR and TAB
	// because they are serialized as "&#xD;" and "&#x9;".
	c := s[0]
	if c <= 32 && c != 9 && c != 13 {
		return true
	}
	c = s[len(s)-1]
	if c <= 32 && c != 9 && c != 13 {
		return true
	}
	return strings.ContainsRune(s, '\u000a')
}
