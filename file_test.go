package xlsx

import (
	"encoding/xml"
	. "gopkg.in/check.v1"
)

type FileSuite struct {}

var _ = Suite(&FileSuite{})

// Test we can correctly open a XSLX file and return a xlsx.File
// struct.
func (l *FileSuite) TestOpenFile(c *C) {
	var xlsxFile *File
	var error error

	xlsxFile, error = OpenFile("testfile.xlsx")
	c.Assert(error, IsNil)
	c.Assert(xlsxFile, NotNil)
}

// Test we can create a File object from scratch
func (l *FileSuite) TestCreateFile(c *C) {
	var xlsxFile *File

	xlsxFile = NewFile()
	c.Assert(xlsxFile, NotNil)
}

// Test that when we open a real XLSX file we create xlsx.Sheet
// objects for the sheets inside the file and that these sheets are
// themselves correct.
func (l *FileSuite) TestCreateSheet(c *C) {
	var xlsxFile *File
	var err error
	var sheet *Sheet
	var row *Row
	xlsxFile, err = OpenFile("testfile.xlsx")
	c.Assert(err, IsNil)
	c.Assert(xlsxFile, NotNil)
	sheetLen := len(xlsxFile.Sheets)
	c.Assert(sheetLen, Equals, 3)
	sheet = xlsxFile.Sheets["Tabelle1"]
	rowLen := len(sheet.Rows)
	c.Assert(rowLen, Equals, 2)
	row = sheet.Rows[0]
	c.Assert(len(row.Cells), Equals, 2)
	cell := row.Cells[0]
	cellstring := cell.String()
	c.Assert(cellstring, Equals, "Foo")
}

// Test that we can add a sheet to a File
func (l *FileSuite) TestAddSheet(c *C) {
	var f *File

	f = NewFile()
	sheet := f.AddSheet("MySheet")
	c.Assert(sheet, NotNil)
	c.Assert(len(f.Sheets), Equals, 1)
	c.Assert(f.Sheets["MySheet"], Equals, sheet)
}

func (l *FileSuite) TestMarshalWorkbook(c *C) {
	var f *File

	f = NewFile()

	f.AddSheet("MyFirstSheet")
	f.AddSheet("MySecondSheet")
	workbook := f.makeWorkbook()
	workbook.Sheets.Sheet[0] = xlsxSheet{
		Name: "MyFirstSheet",
		SheetId: "1",
		Id: "rId1"}

	workbook.Sheets.Sheet[1] = xlsxSheet{
		Name: "MySecondSheet",
		SheetId: "2",
		Id: "rId2"}

	expectedWorkbook := `<?xml version="1.0" encoding="UTF-8"?>
   <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <fileVersion appName="Go XLSX"></fileVersion>
      <workbookPr date1904="false"></workbookPr>
      <bookViews>
         <workbookView></workbookView>
      </bookViews>
      <sheets>
         <sheet name="MyFirstSheet" sheetId="1" xmlns:relationships="http://schemas.openxmlformats.org/officeDocument/2006/relationships" relationships:id="rId1"></sheet>
         <sheet name="MySecondSheet" sheetId="2" xmlns:relationships="http://schemas.openxmlformats.org/officeDocument/2006/relationships" relationships:id="rId2"></sheet>
      </sheets>
      <definedNames></definedNames>
      <calcPr></calcPr>
   </workbook>`
	output, err := xml.MarshalIndent(workbook, "   ", "   ")
	c.Assert(err, IsNil)
	stringOutput := xml.Header + string(output)
	c.Assert(stringOutput, Equals, expectedWorkbook)
}


// Test that we can marshall a File to a collection of xml files
func (l *FileSuite) TestMarshalFile(c *C) {
	var f *File
	f = NewFile()
	sheet1 := f.AddSheet("MySheet")
	row1 := sheet1.AddRow()
	cell1 := row1.AddCell()
	cell1.Value = "A cell!"
	sheet2 := f.AddSheet("AnotherSheet")
	row2 := sheet2.AddRow()
	cell2 := row2.AddCell()
	cell2.Value = "A cell!"
	parts, err := f.MarshallParts()
	c.Assert(err, IsNil)
	c.Assert(len(parts), Equals, 10)

	// sheets
	expectedSheet := `<?xml version="1.0" encoding="UTF-8"?>
  <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <dimension ref="A1:A1"></dimension>
    <sheetData>
      <row r="1">
        <c r="A1" t="s">
          <v>0</v>
        </c>
      </row>
    </sheetData>
  </worksheet>`
	c.Assert(parts["worksheets/sheet1.xml"], Equals, expectedSheet)
	c.Assert(parts["worksheets/sheet2.xml"], Equals, expectedSheet)

	// .rels.xml
	expectedRels := `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`
	c.Assert(parts[".rels"], Equals, expectedRels)

	// app.xml
	expectedApp := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <TotalTime>0</TotalTime>
  <Application>Go XLSX</Application>
</Properties>`
	c.Assert(parts["docProps/app.xml"], Equals, expectedApp)

	// core.xml
	expectedCore := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></cp:coreProperties>`
	c.Assert(parts["docProps/core.xml"], Equals, expectedCore)

	// sharedStrings.xml
	expectedXLSXSST := `<?xml version="1.0" encoding="UTF-8"?>
  <sst count="1" uniqueCount="1">
    <si>
      <t>A cell!</t>
    </si>
  </sst>`
	c.Assert(parts["xl/sharedStrings.xml"], Equals, expectedXLSXSST)

	// workbook.xml.rels
	expectedXLSXWorkbookRels := `<?xml version="1.0" encoding="UTF-8"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Target="worksheets/sheet1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"></Relationship>
    <Relationship Id="rId2" Target="worksheets/sheet2.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"></Relationship>
    <Relationship Id="rId3" Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"></Relationship>
  </Relationships>`
	c.Assert(parts["xl/_rels/workbook.xml.rels"], Equals, expectedXLSXWorkbookRels)

	// workbook.xml
	expectedWorkbook := `<?xml version="1.0" encoding="UTF-8"?>
  <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fileVersion appName="Go XLSX"></fileVersion>
    <workbookPr date1904="false"></workbookPr>
    <bookViews>
      <workbookView></workbookView>
    </bookViews>
    <sheets>
      <sheet name="MySheet" sheetId="1" xmlns:relationships="http://schemas.openxmlformats.org/officeDocument/2006/relationships" relationships:id="rId1"></sheet>
      <sheet name="AnotherSheet" sheetId="2" xmlns:relationships="http://schemas.openxmlformats.org/officeDocument/2006/relationships" relationships:id="rId2"></sheet>
    </sheets>
    <definedNames></definedNames>
    <calcPr></calcPr>
  </workbook>`
	c.Assert(parts["xl/workbook.xml"], Equals, expectedWorkbook)

	// [Content_Types].xml
	expectedContentTypes := `<?xml version="1.0" encoding="UTF-8"?>
  <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Override>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"></Override>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"></Override>
    <Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Override>
    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"></Override>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"></Override>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"></Override>
    <Override PartName="xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"></Override>
    <Override PartName="xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"></Override>
  </Types>`
	c.Assert(parts["[Content_Types].xml"], Equals, expectedContentTypes)


	// styles.xml
	//
	// For now we only allow simple string data in the
	// spreadsheet.  Style support will follow.
	expectedStyles := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
</styleSheet>`
	c.Assert(parts["xl/styles.xml"], Equals, expectedStyles)
}
