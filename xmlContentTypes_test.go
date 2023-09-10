package xlsx

import (
	"encoding/xml"
	"testing"

	qt "github.com/frankban/quicktest"
)

func TestMarshalContentTypes(t *testing.T) {
	c := qt.New(t)
	var types xlsxTypes = xlsxTypes{}
	types.Overrides = make([]xlsxOverride, 1)
	types.Overrides[0] = xlsxOverride{PartName: "/_rels/.rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml"}
	output, err := xml.Marshal(types)
	stringOutput := xml.Header + string(output)
	c.Assert(err, qt.IsNil)
	expectedContentTypes := `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Override></Types>`
	c.Assert(stringOutput, qt.Equals, expectedContentTypes)
}

func TestMakeDefaultContentTypes(t *testing.T) {
	c := qt.New(t)
	var types xlsxTypes = MakeDefaultContentTypes()
	c.Assert(len(types.Overrides), qt.Equals, 8)
	c.Assert(types.Overrides[0].PartName, qt.Equals, "/_rels/.rels")
	c.Assert(types.Overrides[0].ContentType, qt.Equals, "application/vnd.openxmlformats-package.relationships+xml")
	c.Assert(types.Overrides[1].PartName, qt.Equals, "/docProps/app.xml")
	c.Assert(types.Overrides[1].ContentType, qt.Equals, "application/vnd.openxmlformats-officedocument.extended-properties+xml")
	c.Assert(types.Overrides[2].PartName, qt.Equals, "/docProps/core.xml")
	c.Assert(types.Overrides[2].ContentType, qt.Equals, "application/vnd.openxmlformats-package.core-properties+xml")
	c.Assert(types.Overrides[3].PartName, qt.Equals, "/xl/_rels/workbook.xml.rels")
	c.Assert(types.Overrides[3].ContentType, qt.Equals, "application/vnd.openxmlformats-package.relationships+xml")
	c.Assert(types.Overrides[4].PartName, qt.Equals, "/xl/sharedStrings.xml")
	c.Assert(types.Overrides[4].ContentType, qt.Equals, "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
	c.Assert(types.Overrides[5].PartName, qt.Equals, "/xl/styles.xml")
	c.Assert(types.Overrides[5].ContentType, qt.Equals, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
	c.Assert(types.Overrides[6].PartName, qt.Equals, "/xl/workbook.xml")
	c.Assert(types.Overrides[6].ContentType, qt.Equals, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
	c.Assert(types.Overrides[7].PartName, qt.Equals, "/xl/theme/theme1.xml")
	c.Assert(types.Overrides[7].ContentType, qt.Equals, "application/vnd.openxmlformats-officedocument.theme+xml")

	c.Assert(types.Defaults[0].Extension, qt.Equals, "rels")
	c.Assert(types.Defaults[0].ContentType, qt.Equals, "application/vnd.openxmlformats-package.relationships+xml")
	c.Assert(types.Defaults[1].Extension, qt.Equals, "xml")
	c.Assert(types.Defaults[1].ContentType, qt.Equals, "application/xml")

}
