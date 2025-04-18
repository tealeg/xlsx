package xlsx

import "encoding/xml"

type RelationshipType string

const (
	RelationshipTypeHyperlink RelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)

type RelationshipTargetMode string

const (
	RelationshipTargetModeExternal RelationshipTargetMode = "External"
)

type xlsxRels struct {
	XMLName       xml.Name       `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Relationships []xlsxRelation `xml:"Relationship"`
}

type xlsxRelation struct {
	Id         string                 `xml:"Id,attr"`
	Type       RelationshipType       `xml:"Type,attr"`
	Target     string                 `xml:"Target,attr"`
	TargetMode RelationshipTargetMode `xml:"TargetMode,attr,omitempty"`
}
