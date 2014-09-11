package xlsx

import (
	"encoding/xml"
	. "gopkg.in/check.v1"
)

type ContentTypesSuite struct {}

var _ = Suite(&ContentTypesSuite{})


func (l *ContentTypesSuite) TestMarshalContentTypes(c *C) {
	var types xlsxTypes = xlsxTypes{}
	types.Overrides = make([]xlsxOverride, 1)
	types.Overrides[0] = xlsxOverride{PartName: "/_rels/.rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml"}
	output, err := xml.MarshalIndent(types, "   ", "   ")
	stringOutput := xml.Header + string(output)
	c.Assert(err, IsNil)
	expectedContentTypes := `<?xml version="1.0" encoding="UTF-8"?>
   <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Override>
   </Types>`
	c.Assert(stringOutput, Equals, expectedContentTypes)
}

