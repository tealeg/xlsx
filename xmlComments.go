package xlsx

import (
	"encoding/xml"
)

// xlsxComments directly maps the comments element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main
type xlsxComments struct {
	XMLName     xml.Name      `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main comments"`
	Authors     []string      `xml:"authors>author"`
	CommentList []xlsxComment `xml:"commentList>comment"`
}

// xlsxComment directly maps the comment element from the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main
// Partially implemented for reading comment text values
type xlsxComment struct {
	XMLName  xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main comment"`
	AuthorID string   `xml:"authorId,attr"`
	Ref      string   `xml:"ref,attr"`
	Value    string   `xml:"text>r>t"`
}
