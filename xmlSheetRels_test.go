package xlsx

import (
	"archive/zip"
	"bytes"
	"io/ioutil"

	. "gopkg.in/check.v1"
)

// Implementation of worksheet relationships, primarily with the focus of adding support for hyperlinks, which are located in the xl/worksheets/_rels path location for each workbook.
// Each sheet that has hyperlinks has a matching sheetN.xml.rels file in this directory, with the structure of the file matching XmlSheetRels and xmlRel. The content of these sheets are marshalled into these structs and saved as a field onto the File.
// "sheetN.xml"s contain <hyperlinks> tags that contain "rIDn" reference IDs that refer to the relationship objects in these relationship files in this directory, where the raw URL of any hyperlinks are stored. The xml.rels files are the only location of raw hyperlink values, and can only be traced with these rIDs matching the rID from a sheet's <hyperlink> tag.
func (l *LibSuite) TestReadSheetRelsFromZipFile(c *C) {

	expectedMap := map[string]*XmlSheetRels{
		"sheet1": &XmlSheetRels{
			SheetName: "sheet1",
			Rels: []*xmlRel{
				&xmlRel{
					Rid:        "rId3",
					RelType:    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
					Target:     "https://github.com/",
					TargetMode: "External",
				},
				&xmlRel{
					Rid:        "rId2",
					RelType:    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
					Target:     "https://www.wikipedia.org/",
					TargetMode: "External",
				},
				&xmlRel{
					Rid:        "rId1",
					RelType:    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
					Target:     "https://www.google.com/",
					TargetMode: "External",
				},
			},
		},
		"sheet2": &XmlSheetRels{
			SheetName: "sheet2",
			Rels: []*xmlRel{
				&xmlRel{
					Rid:        "rId1",
					RelType:    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
					Target:     "http://facebook.com/",
					TargetMode: "External",
				},
			},
		},
	}

	// Test setup with test file.
	testFile := "testdocs/testHyperlinks.xlsx"
	fileBytes, err := ioutil.ReadFile(testFile)
	if err != nil {
		c.Fatal(err)
	}

	r := bytes.NewReader(fileBytes)
	size := int64(r.Len())

	zipReader, err := zip.NewReader(r, size)
	if err != nil {
		c.Fatal(err)
	}

	worksheetRels := make(map[string]*zip.File)

	// Make the worksheetRels map that will be passed to the testable function.
	// Its key is the sheetName (without any extensions) and the value is the pointer to the zipFile.
	for _, sheetRel := range zipReader.File {
		if len(sheetRel.Name) > 17 {
			if sheetRel.Name[0:19] == "xl/worksheets/_rels" {
				worksheetRels[sheetRel.Name[20:len(sheetRel.Name)-9]] = sheetRel
			}
		}

	}

	// Call the actual function to be tested, compare values.
	actualRelationshipMap, actualErr := readSheetRelsFromZipFile(worksheetRels)

	c.Assert(actualErr, Equals, nil)

	for sheetName, expectedSheetRels := range expectedMap {

		actualSheetRels := actualRelationshipMap[sheetName]

		for relIndex, expectedRelFile := range expectedSheetRels.Rels {

			actualRelFile := actualSheetRels.Rels[relIndex]

			c.Assert(actualRelFile.Rid, Equals, expectedRelFile.Rid)
			c.Assert(actualRelFile.RelType, Equals, expectedRelFile.RelType)
			c.Assert(actualRelFile.Target, Equals, expectedRelFile.Target)
			c.Assert(actualRelFile.TargetMode, Equals, expectedRelFile.TargetMode)

		}

	}

}
