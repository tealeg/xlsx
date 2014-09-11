package xlsx

import (
	"encoding/xml"
)

type xlsxTypes struct {
	XMLName		xml.Name `xml:"http://schemas.openxmlformats.org/package/2006/content-types Types"`

	Overrides []xlsxOverride `xml:"Override"`
}

type xlsxOverride struct {
	PartName string	`xml:",attr"`
	ContentType string `xml:",attr"`
}
