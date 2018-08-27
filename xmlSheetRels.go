package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"io/ioutil"

	"github.com/pkg/errors"
)

// Implementation of worksheet relationships, primarily with the focus of adding support for hyperlinks, which are located in the xl/worksheets/_rels path location for each workbook.

// XmlSheetRels type for holding data for sheet relationships in the workbook.
type XmlSheetRels struct {
	XMLName   xml.Name  `xml:"Relationships"`
	SheetName string    `xml:",attr,omitempty"`
	Rels      []*xmlRel `xml:"Relationship,omitempty"`
}

type xmlRel struct {
	XMLName    xml.Name `xml:"Relationship"`
	Rid        string   `xml:"Id,attr,omitempty"`
	RelType    string   `xml:"Type,attr,omitempty"`
	Target     string   `xml:"Target,attr,omitempty"`
	TargetMode string   `xml:"TargetMode,attr,omitempty"`
}

// Check directory if any files exist in xl/worksheets/_rels
func readSheetRelsFromZipFile(worksheetRels map[string]*zip.File) (map[string]*XmlSheetRels, error) {
	sheetRels := make(map[string]*XmlSheetRels)

	for sheetName, f := range worksheetRels {

		xmlSheetRels := new(XmlSheetRels)

		rc, err := f.Open()
		if err != nil {
			return nil, errors.WithStack(err)
		}

		bytes, err := ioutil.ReadAll(rc)
		if err != nil {
			return nil, errors.WithStack(err)
		}

		xmlSheetRels.SheetName = sheetName

		err = xml.Unmarshal(bytes, &xmlSheetRels)
		if err != nil {
			return nil, errors.WithStack(err)
		}

		sheetRels[sheetName] = xmlSheetRels
	}

	return sheetRels, nil
}
