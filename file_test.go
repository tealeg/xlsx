package xlsx

import (
	"encoding/xml"
	"path/filepath"

	. "gopkg.in/check.v1"
)

type FileSuite struct{}

var _ = Suite(&FileSuite{})

// Test we can correctly open a XSLX file and return a xlsx.File
// struct.
func (l *FileSuite) TestOpenFile(c *C) {
	var xlsxFile *File
	var error error

	xlsxFile, error = OpenFile("./testdocs/testfile.xlsx")
	c.Assert(error, IsNil)
	c.Assert(xlsxFile, NotNil)
}

func (l *FileSuite) TestOpenFileWithoutStyleAndSharedStrings(c *C) {
	var xlsxFile *File
	var error error

	xlsxFile, error = OpenFile("./testdocs/noStylesAndSharedStringsTest.xlsx")
	c.Assert(error, IsNil)
	c.Assert(xlsxFile, NotNil)
}

// Test that we can correctly extract a reference table from the
// sharedStrings.xml file embedded in the XLSX file and return a
// reference table of string values from it.
func (l *FileSuite) TestReadSharedStringsFromZipFile(c *C) {
	var xlsxFile *File
	var err error
	xlsxFile, err = OpenFile("./testdocs/testfile.xlsx")
	c.Assert(err, IsNil)
	c.Assert(xlsxFile.referenceTable, NotNil)
}

// Helper function used to test contents of a given xlsxXf against
// expectations.
func testXf(c *C, result, expected *xlsxXf) {
	c.Assert(result.ApplyAlignment, Equals, expected.ApplyAlignment)
	c.Assert(result.ApplyBorder, Equals, expected.ApplyBorder)
	c.Assert(result.ApplyFont, Equals, expected.ApplyFont)
	c.Assert(result.ApplyFill, Equals, expected.ApplyFill)
	c.Assert(result.ApplyProtection, Equals, expected.ApplyProtection)
	c.Assert(result.BorderId, Equals, expected.BorderId)
	c.Assert(result.FillId, Equals, expected.FillId)
	c.Assert(result.FontId, Equals, expected.FontId)
	c.Assert(result.NumFmtId, Equals, expected.NumFmtId)
}

// We can correctly extract a style table from the style.xml file
// embedded in the XLSX file and return a styles struct from it.
func (l *FileSuite) TestReadStylesFromZipFile(c *C) {
	var xlsxFile *File
	var err error
	var fontCount, fillCount, borderCount, cellStyleXfCount, cellXfCount int
	var font xlsxFont
	var fill xlsxFill
	var border xlsxBorder
	var xf xlsxXf

	xlsxFile, err = OpenFile("./testdocs/testfile.xlsx")
	c.Assert(err, IsNil)
	c.Assert(xlsxFile.styles, NotNil)

	fontCount = len(xlsxFile.styles.Fonts)
	c.Assert(fontCount, Equals, 4)

	font = xlsxFile.styles.Fonts[0]
	c.Assert(font.Sz.Val, Equals, "11")
	c.Assert(font.Name.Val, Equals, "Calibri")

	fillCount = len(xlsxFile.styles.Fills)
	c.Assert(fillCount, Equals, 3)

	fill = xlsxFile.styles.Fills[2]
	c.Assert(fill.PatternFill.PatternType, Equals, "solid")

	borderCount = len(xlsxFile.styles.Borders)
	c.Assert(borderCount, Equals, 2)

	border = xlsxFile.styles.Borders[1]
	c.Assert(border.Left.Style, Equals, "thin")
	c.Assert(border.Right.Style, Equals, "thin")
	c.Assert(border.Top.Style, Equals, "thin")
	c.Assert(border.Bottom.Style, Equals, "thin")

	cellStyleXfCount = len(xlsxFile.styles.CellStyleXfs)
	c.Assert(cellStyleXfCount, Equals, 20)

	xf = xlsxFile.styles.CellStyleXfs[0]
	expectedXf := &xlsxXf{
		ApplyAlignment:  true,
		ApplyBorder:     true,
		ApplyFont:       true,
		ApplyFill:       false,
		ApplyProtection: true,
		BorderId:        0,
		FillId:          0,
		FontId:          0,
		NumFmtId:        164}
	testXf(c, &xf, expectedXf)

	cellXfCount = len(xlsxFile.styles.CellXfs)
	c.Assert(cellXfCount, Equals, 3)

	xf = xlsxFile.styles.CellXfs[0]
	expectedXf = &xlsxXf{
		ApplyAlignment:  false,
		ApplyBorder:     false,
		ApplyFont:       false,
		ApplyFill:       false,
		ApplyProtection: false,
		BorderId:        0,
		FillId:          0,
		FontId:          0,
		NumFmtId:        164}
	testXf(c, &xf, expectedXf)
}

// We can correctly extract a map of relationship Ids to the worksheet files in
// which they are contained from the XLSX file.
func (l *FileSuite) TestReadWorkbookRelationsFromZipFile(c *C) {
	var xlsxFile *File
	var err error

	xlsxFile, err = OpenFile("./testdocs/testfile.xlsx")
	c.Assert(err, IsNil)
	c.Assert(len(xlsxFile.Sheets), Equals, 3)
	sheet, ok := xlsxFile.Sheet["Tabelle1"]
	c.Assert(ok, Equals, true)
	c.Assert(sheet, NotNil)
}

// +build fudge
func (l *FileSuite) TestGetStyleFromZipFile(c *C) {
	var xlsxFile *File
	var err error
	var style Style

	xlsxFile, err = OpenFile("./testdocs/testfile.xlsx")
	c.Assert(err, IsNil)
	sheetCount := len(xlsxFile.Sheets)
	c.Assert(sheetCount, Equals, 3)

	tabelle1 := xlsxFile.Sheet["Tabelle1"]

	row0 := tabelle1.Rows[0]
	cellFoo := row0.Cells[0]
	style = cellFoo.GetStyle()
	c.Assert(cellFoo.String(), Equals, "Foo")
	c.Assert(style.Fill.BgColor, Equals, "FF33CCCC")

	row1 := tabelle1.Rows[1]
	cellQuuk := row1.Cells[1]
	style = cellQuuk.GetStyle()
	c.Assert(cellQuuk.String(), Equals, "Quuk")
	c.Assert(style.Border.Left, Equals, "thin")

	cellBar := row0.Cells[1]
	c.Assert(cellBar.String(), Equals, "Bar")
	c.Assert(cellBar.GetStyle().Fill.BgColor, Equals, "")
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
	xlsxFile, err = OpenFile("./testdocs/testfile.xlsx")
	c.Assert(err, IsNil)
	c.Assert(xlsxFile, NotNil)
	sheetLen := len(xlsxFile.Sheets)
	c.Assert(sheetLen, Equals, 3)
	sheet = xlsxFile.Sheet["Tabelle1"]
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
	c.Assert(f.Sheet["MySheet"], Equals, sheet)
}

// Test that we can get the Nth sheet
func (l *FileSuite) TestNthSheet(c *C) {
	var f *File

	f = NewFile()
	sheet := f.AddSheet("MySheet")
	sheetByIndex := f.Sheets[0]
	sheetByName := f.Sheet["MySheet"]
	c.Assert(sheetByIndex, NotNil)
	c.Assert(sheetByIndex, Equals, sheet)
	c.Assert(sheetByIndex, Equals, sheetByName)
}

// Test that we can create a Workbook and marshal it to XML.
func (l *FileSuite) TestMarshalWorkbook(c *C) {
	var f *File

	f = NewFile()

	f.AddSheet("MyFirstSheet")
	f.AddSheet("MySecondSheet")
	workbook := f.makeWorkbook()
	workbook.Sheets.Sheet[0] = xlsxSheet{
		Name:    "MyFirstSheet",
		SheetId: "1",
		Id:      "rId1"}

	workbook.Sheets.Sheet[1] = xlsxSheet{
		Name:    "MySecondSheet",
		SheetId: "2",
		Id:      "rId2"}

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
    <cols>
      <col min="0" max="0"></col>
    </cols>
    <sheetData>
      <row r="1">
        <c r="A1" t="s">
          <v>0</v>
        </c>
      </row>
    </sheetData>
  </worksheet>`
	c.Assert(parts["xl/worksheets/sheet1.xml"], Equals, expectedSheet)
	c.Assert(parts["xl/worksheets/sheet2.xml"], Equals, expectedSheet)

	// .rels.xml
	expectedRels := `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`
	c.Assert(parts["_rels/.rels"], Equals, expectedRels)

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
  <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
    <si>
      <t>A cell!</t>
    </si>
  </sst>`
	c.Assert(parts["xl/sharedStrings.xml"], Equals, expectedXLSXSST)

	// workbook.xml.rels
	expectedXLSXWorkbookRels := `<?xml version="1.0" encoding="UTF-8"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Target="worksheets/sheet1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"></Relationship>
    <Relationship Id="rId2" Target="worksheets/sheet2.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"></Relationship>
    <Relationship Id="rId3" Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"></Relationship>
    <Relationship Id="rId4" Target="styles.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"></Relationship>
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
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"></Override>
    <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"></Override>
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Default>
    <Default Extension="xml" ContentType="application/xml"></Default>
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

// We can save a File as a valid XLSX file at a given path.
func (l *FileSuite) TestSaveFile(c *C) {
	var tmpPath string = c.MkDir()
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
	xlsxPath := filepath.Join(tmpPath, "TestSaveFile.xlsx")
	err := f.Save(xlsxPath)
	c.Assert(err, IsNil)

	// Let's eat our own dog food
	xlsxFile, err := OpenFile(xlsxPath)
	c.Assert(err, IsNil)
	c.Assert(xlsxFile, NotNil)
	c.Assert(len(xlsxFile.Sheets), Equals, 2)

	sheet1, ok := xlsxFile.Sheet["MySheet"]
	c.Assert(ok, Equals, true)
	c.Assert(len(sheet1.Rows), Equals, 1)
	row1 = sheet1.Rows[0]
	c.Assert(len(row1.Cells), Equals, 1)
	cell1 = row1.Cells[0]
	c.Assert(cell1.Value, Equals, "A cell!")
}

type SliceReaderSuite struct{}

var _ = Suite(&SliceReaderSuite{})

func (s *SliceReaderSuite) TestFileToSlice(c *C) {
	output, err := FileToSlice("./testdocs/testfile.xlsx")
	c.Assert(err, IsNil)
	fileToSliceCheckOutput(c, output)
}

func (s *SliceReaderSuite) TestFileObjToSlice(c *C) {
	f, err := OpenFile("./testdocs/testfile.xlsx")
	output, err := f.ToSlice()
	c.Assert(err, IsNil)
	fileToSliceCheckOutput(c, output)
}

func fileToSliceCheckOutput(c *C, output [][][]string) {
	c.Assert(len(output), Equals, 3)
	c.Assert(len(output[0]), Equals, 2)
	c.Assert(len(output[0][0]), Equals, 2)
	c.Assert(output[0][0][0], Equals, "Foo")
	c.Assert(output[0][0][1], Equals, "Bar")
	c.Assert(len(output[0][1]), Equals, 2)
	c.Assert(output[0][1][0], Equals, "Baz")
	c.Assert(output[0][1][1], Equals, "Quuk")
	c.Assert(len(output[1]), Equals, 0)
	c.Assert(len(output[2]), Equals, 0)
}
