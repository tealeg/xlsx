package xlsx

import (
	"encoding/xml"
)

type xlsxExternalLink struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`

	Relationship xlsxRelationship `xml:"Relationship"`
}

type xlsxRelationship struct {
	Target string `xml:"Target,attr"`
}

