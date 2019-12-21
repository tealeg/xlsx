package xlsx

import (
	"context"
	"encoding/xml"
	"io"
	"os"
	"path/filepath"
	"testing"
	"time"

	qt "github.com/frankban/quicktest"
	. "gopkg.in/check.v1"
)

// ReaderAtCounter wraps a ReaderAt and counts the number of bytes that are read out of it
type ReaderAtCounter struct {
	readerAt  io.ReaderAt
	bytesRead int
}

var _ io.ReaderAt = &ReaderAtCounter{}

// NewReaderAtCounter creates a ReaderAtCounter by opening the file name, and provides the size which is needed for
// opening as XLSX.
func NewReaderAtCounter(name string) (*ReaderAtCounter, int64, error) {
	f, err := os.Open(name)
	if err != nil {
		return nil, -1, err
	}
	fi, err := f.Stat()
	if err != nil {
		f.Close()
		return nil, -1, err
	}
	readerAtCounter := &ReaderAtCounter{
		readerAt: f,
	}
	return readerAtCounter, fi.Size(), nil
}

func (r *ReaderAtCounter) ReadAt(p []byte, off int64) (n int, err error) {
	n, err = r.readerAt.ReadAt(p, off)
	r.bytesRead += n
	return n, err
}

func (r *ReaderAtCounter) GetBytesRead() int {
	return r.bytesRead
}

type FileSuite struct{}

var _ = Suite(&FileSuite{})

// Test we can correctly open a XSLX file and return a xlsx.File
// struct.
func (l *FileSuite) TestOpenFile(c *C) {
	var xlsxFile *File
	var err error

	xlsxFile, err = OpenFile("./testdocs/testfile.xlsx")
	c.Assert(err, IsNil)
	c.Assert(xlsxFile, NotNil)
}

func (l *FileSuite) TestPartialReadsWithFewSharedStringsOnlyPartiallyReads(c *C) {
	// This test verifies that a large file is only partially read when using a small row limit.
	// This file is 11,228,530 bytes, but only 14,020 bytes get read out when using a row limit of 10.
	// I'm specifying a limit of 20,000 to prevent test flakiness if the bytes read fluctuates with small code changes.
	rowLimit := 10
	// It is possible that readLimit will need to be increased by a small amount in the future, but do not increase it
	// to anywhere near a significant amount of 11 million. We're testing that this number is low, to ensure that partial
	// reads are fast.
	readLimit := 20 * 1000
	reader, size, err := NewReaderAtCounter("testdocs/large_sheet_no_shared_strings_no_dimension_tag.xlsx")
	if err != nil {
		c.Fatal(err)
	}
	file, err := OpenReaderAtWithRowLimit(reader, size, rowLimit)
	if reader.bytesRead > readLimit {
		// If this test begins failing, do not increase readLimit dramatically. Instead investigate why the number of
		// bytes read went up and fix this issue.
		c.Errorf("Reading %v rows from a sheet with ~31,000 rows and few shared strings read %v bytes, must read less than %v bytes", rowLimit, reader.bytesRead, readLimit)
	}
	if len(file.Sheets[0].Rows) != rowLimit {
		c.Errorf("Expected sheet to have %v rows, but found %v rows", rowLimit, len(file.Sheets[0].Rows))
	}
}

func (l *FileSuite) TestPartialReadsWithLargeSharedStringsOnlyPartiallyReads(c *C) {
	// This test verifies that a large file is only partially read when using a small row limit.
	// This file is 7,055,632 bytes, but only 1,092,839 bytes get read out when using a row limit of 10.
	// I'm specifying a limit of 1.2 MB to prevent test flakiness if the bytes read fluctuates with small code changes.
	// The reason that this test has a much larger limit than TestPartialReadsWithFewSharedStringsOnlyPartiallyReads
	// is that this file has a Shared Strings file that is a little over 1 MB.
	rowLimit := 10
	// It is possible that readLimit will need to be increased by a small amount in the future, but do not increase it
	// to anywhere near a significant amount of 7 million. We're testing that this number is low, to ensure that partial
	// reads are fast.
	readLimit := int(1.2 * 1000 * 1000)
	reader, size, err := NewReaderAtCounter("testdocs/large_sheet_large_sharedstrings_dimension_tag.xlsx")
	if err != nil {
		c.Fatal(err)
	}
	file, err := OpenReaderAtWithRowLimit(reader, size, rowLimit)
	if reader.bytesRead > readLimit {
		// If this test begins failing, do not increase readLimit dramatically. Instead investigate why the number of
		// bytes read went up and fix this issue.
		c.Errorf("Reading %v rows from a sheet with ~31,000 rows and a large shared strings read %v bytes, must read less than %v bytes", rowLimit, reader.bytesRead, readLimit)
	}
	// This is testing that the sheet was truncated, but it is also testing that the dimension tag was ignored.
	// If the dimension tag is not correctly ignored, there will be 10 rows of the data, plus ~31k empty rows tacked on.
	if len(file.Sheets[0].Rows) != rowLimit {
		c.Errorf("Expected sheet to have %v rows, but found %v rows", rowLimit, len(file.Sheets[0].Rows))
	}
}

func (l *FileSuite) TestPartialReadsWithFewerRowsThanRequested(c *C) {
	rowLimit := 10
	file, err := OpenFileWithRowLimit("testdocs/testfile.xlsx", rowLimit)
	if err != nil {
		c.Fatal(err)
	}
	if len(file.Sheets[0].Rows) != 2 {
		c.Errorf("Expected sheet to have %v rows, but found %v rows", 2, len(file.Sheets[0].Rows))
	}
}

func (l *FileSuite) TestFileLargeTimeout(c *C) {
	rowLimit := 1000000
	ctx, cancel := context.WithTimeout(context.Background(), 200*time.Millisecond)
	defer cancel()

	_, err := OpenFileWithCtxRowLimit(ctx, "testdocs/testlargefile.xlsx", rowLimit)
	if err != context.DeadlineExceeded {
		c.Fatal(err)
	}
}

func (l *FileSuite) TestOpenFileWithoutStyleAndSharedStrings(c *C) {
	var xlsxFile *File
	var error error

	xlsxFile, error = OpenFile("./testdocs/noStylesAndSharedStringsTest.xlsx")
	c.Assert(error, IsNil)
	c.Assert(xlsxFile, NotNil)
}

func (l *FileSuite) TestOpenFileWithChartsheet(c *C) {
	xlsxFile, error := OpenFile("./testdocs/testchartsheet.xlsx")
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

	fontCount = len(xlsxFile.styles.Fonts.Font)
	c.Assert(fontCount, Equals, 4)

	font = xlsxFile.styles.Fonts.Font[0]
	c.Assert(font.Sz.Val, Equals, "11")
	c.Assert(font.Name.Val, Equals, "Calibri")

	fillCount = xlsxFile.styles.Fills.Count
	c.Assert(fillCount, Equals, 3)

	fill = xlsxFile.styles.Fills.Fill[2]
	c.Assert(fill.PatternFill.PatternType, Equals, "solid")

	borderCount = xlsxFile.styles.Borders.Count
	c.Assert(borderCount, Equals, 2)

	border = xlsxFile.styles.Borders.Border[1]
	c.Assert(border.Left.Style, Equals, "thin")
	c.Assert(border.Right.Style, Equals, "thin")
	c.Assert(border.Top.Style, Equals, "thin")
	c.Assert(border.Bottom.Style, Equals, "thin")

	cellStyleXfCount = xlsxFile.styles.CellStyleXfs.Count
	c.Assert(cellStyleXfCount, Equals, 20)

	xf = xlsxFile.styles.CellStyleXfs.Xf[0]
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

	c.Assert(xf.Alignment, NotNil)
	c.Assert(xf.Alignment.Horizontal, Equals, "general")
	c.Assert(xf.Alignment.Indent, Equals, 0)
	c.Assert(xf.Alignment.ShrinkToFit, Equals, false)
	c.Assert(xf.Alignment.TextRotation, Equals, 0)
	c.Assert(xf.Alignment.Vertical, Equals, "bottom")
	c.Assert(xf.Alignment.WrapText, Equals, false)

	cellXfCount = xlsxFile.styles.CellXfs.Count
	c.Assert(cellXfCount, Equals, 3)

	xf = xlsxFile.styles.CellXfs.Xf[0]
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

// Style information is correctly extracted from the zipped XLSX file.
func TestGetStyleFromZipFile(t *testing.T) {
	c := qt.New(t)
	var xlsxFile *File
	var err error
	var style *Style
	var val string

	xlsxFile, err = OpenFile("./testdocs/testfile.xlsx")
	c.Assert(err, qt.IsNil)
	sheetCount := len(xlsxFile.Sheets)
	c.Assert(sheetCount, qt.Equals, 3)

	tabelle1 := xlsxFile.Sheet["Tabelle1"]

	row0 := tabelle1.Rows[0]
	cellFoo := row0.Cells[0]
	style = cellFoo.GetStyle()
	if val, err = cellFoo.FormattedValue(); err != nil {
		c.Error(err)
	}
	c.Assert(val, qt.Equals, "Foo")
	c.Assert(style.Fill.BgColor, qt.Equals, "FF33CCCC")
	c.Assert(style.ApplyFill, qt.Equals, false)
	c.Assert(style.ApplyFont, qt.Equals, true)

	row1 := tabelle1.Rows[1]
	cellQuuk := row1.Cells[1]
	style = cellQuuk.GetStyle()
	if val, err = cellQuuk.FormattedValue(); err != nil {
		c.Error(err)
	}
	c.Assert(val, qt.Equals, "Quuk")
	c.Assert(style.Border.Left, qt.Equals, "thin")
	c.Assert(style.ApplyBorder, qt.Equals, true)

	cellBar := row0.Cells[1]
	if val, err = cellBar.FormattedValue(); err != nil {
		c.Error(err)
	}
	c.Assert(val, qt.Equals, "Bar")
	c.Assert(cellBar.GetStyle().Fill.BgColor, qt.Equals, "")
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
	if val, err := cell.FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "Foo")
	}
}

// Test that we can add a sheet to a File
func (l *FileSuite) TestAddSheet(c *C) {
	var f *File

	f = NewFile()
	sheet, err := f.AddSheet("MySheet")
	c.Assert(err, IsNil)
	c.Assert(sheet, NotNil)
	c.Assert(len(f.Sheets), Equals, 1)
	c.Assert(f.Sheet["MySheet"], Equals, sheet)
}

// Test that AddSheet returns an error if you try to add two sheets with the same name
func (l *FileSuite) TestAddSheetWithDuplicateName(c *C) {
	f := NewFile()
	_, err := f.AddSheet("MySheet")
	c.Assert(err, IsNil)
	_, err = f.AddSheet("MySheet")
	c.Assert(err, ErrorMatches, "duplicate sheet name 'MySheet'.")
}

// Test that AddSheet returns an error if you try to add sheet with name as empty string
func (l *FileSuite) TestAddSheetWithEmptyName(c *C) {
	f := NewFile()
	_, err := f.AddSheet("")
	c.Assert(err, ErrorMatches, "sheet name must be 31 or fewer characters long.  It is currently '0' characters long")
}

// Test that we can append a sheet to a File
func (l *FileSuite) TestAppendSheet(c *C) {
	var f *File

	f = NewFile()
	s := Sheet{}
	sheet, err := f.AppendSheet(s, "MySheet")
	c.Assert(err, IsNil)
	c.Assert(sheet, NotNil)
	c.Assert(len(f.Sheets), Equals, 1)
	c.Assert(f.Sheet["MySheet"], Equals, sheet)
}

// Test that AppendSheet returns an error if you try to add two sheets with the same name
func (l *FileSuite) TestAppendSheetWithDuplicateName(c *C) {
	f := NewFile()
	s := Sheet{}
	_, err := f.AppendSheet(s, "MySheet")
	c.Assert(err, IsNil)
	_, err = f.AppendSheet(s, "MySheet")
	c.Assert(err, ErrorMatches, "duplicate sheet name 'MySheet'.")
}

// Test that we can read & create a 31 rune sheet name
func (l *FileSuite) TestMaxSheetNameLength(c *C) {
	// Open a genuine xlsx created by Microsoft Excel 2007
	xlsxFile, err := OpenFile("./testdocs/max_sheet_name_length.xlsx")
	c.Assert(err, IsNil)
	c.Assert(xlsxFile, NotNil)
	c.Assert(xlsxFile.Sheets[0].Name, Equals, "αααααβββββγγγγγδδδδδεεεεεζζζζζη")
	// Create a new file with the same sheet name
	f := NewFile()
	s, err := f.AddSheet(xlsxFile.Sheets[0].Name)
	c.Assert(err, IsNil)
	c.Assert(s.Name, Equals, "αααααβββββγγγγγδδδδδεεεεεζζζζζη")
}

// Test that we can get the Nth sheet
func (l *FileSuite) TestNthSheet(c *C) {
	var f *File

	f = NewFile()
	sheet, _ := f.AddSheet("MySheet")
	sheetByIndex := f.Sheets[0]
	sheetByName := f.Sheet["MySheet"]
	c.Assert(sheetByIndex, NotNil)
	c.Assert(sheetByIndex, Equals, sheet)
	c.Assert(sheetByIndex, Equals, sheetByName)
}

// Test invalid sheet name characters
func (l *FileSuite) TestInvalidSheetNameCharacters(c *C) {
	f := NewFile()
	for _, invalidChar := range []string{":", "\\", "/", "?", "*", "[", "]"} {
		_, err := f.AddSheet(invalidChar)
		c.Assert(err, NotNil)
	}
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
		Id:      "rId1",
		State:   "visible"}

	workbook.Sheets.Sheet[1] = xlsxSheet{
		Name:    "MySecondSheet",
		SheetId: "2",
		Id:      "rId2",
		State:   "visible"}

	expectedWorkbook := `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="Go XLSX"></fileVersion><workbookPr showObjects="all" date1904="false"></workbookPr><workbookProtection></workbookProtection><bookViews><workbookView showHorizontalScroll="true" showVerticalScroll="true" showSheetTabs="true" tabRatio="204" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"></workbookView></bookViews><sheets><sheet name="MyFirstSheet" sheetId="1" r:id="rId1" state="visible"></sheet><sheet name="MySecondSheet" sheetId="2" r:id="rId2" state="visible"></sheet></sheets><definedNames></definedNames><calcPr iterateCount="100" refMode="A1" iterateDelta="0.001"></calcPr></workbook>`
	output, err := xml.Marshal(workbook)
	c.Assert(err, IsNil)
	outputStr := replaceRelationshipsNameSpace(string(output))
	stringOutput := xml.Header + outputStr
	c.Assert(stringOutput, Equals, expectedWorkbook)
}

// Test that we can marshall a File to a collection of xml files
func (l *FileSuite) TestMarshalFile(c *C) {
	var f *File
	f = NewFile()
	sheet1, _ := f.AddSheet("MySheet")
	row1 := sheet1.AddRow()
	cell1 := row1.AddCell()
	cell1.SetString("A cell!")
	sheet2, _ := f.AddSheet("AnotherSheet")
	row2 := sheet2.AddRow()
	cell2 := row2.AddCell()
	cell2.SetString("A cell!")
	parts, err := f.MarshallParts()
	c.Assert(err, IsNil)
	c.Assert(len(parts), Equals, 11)

	// sheets
	expectedSheet1 := `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr filterMode="false"><pageSetUpPr fitToPage="false"></pageSetUpPr></sheetPr><dimension ref="A1"></dimension><sheetViews><sheetView windowProtection="false" showFormulas="false" showGridLines="true" showRowColHeaders="true" showZeros="true" rightToLeft="false" tabSelected="true" showOutlineSymbols="true" defaultGridColor="true" view="normal" topLeftCell="A1" colorId="64" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100" workbookViewId="0"><selection pane="topLeft" activeCell="A1" activeCellId="0" sqref="A1"></selection></sheetView></sheetViews><sheetFormatPr defaultRowHeight="12.85"></sheetFormatPr><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData><printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"></printOptions><pageMargins left="0.7875" right="0.7875" top="1.05277777777778" bottom="1.05277777777778" header="0.7875" footer="0.7875"></pageMargins><pageSetup paperSize="9" scale="100" firstPageNumber="1" fitToWidth="1" fitToHeight="1" pageOrder="downThenOver" orientation="portrait" usePrinterDefaults="false" blackAndWhite="false" draft="false" cellComments="none" useFirstPageNumber="true" horizontalDpi="300" verticalDpi="300" copies="1"></pageSetup><headerFooter differentFirst="false" differentOddEven="false"><oddHeader>&amp;C&amp;&#34;Times New Roman,Regular&#34;&amp;12&amp;A</oddHeader><oddFooter>&amp;C&amp;&#34;Times New Roman,Regular&#34;&amp;12Page &amp;P</oddFooter></headerFooter></worksheet>`
	c.Assert(parts["xl/worksheets/sheet1.xml"], Equals, expectedSheet1)

	expectedSheet2 := `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr filterMode="false"><pageSetUpPr fitToPage="false"></pageSetUpPr></sheetPr><dimension ref="A1"></dimension><sheetViews><sheetView windowProtection="false" showFormulas="false" showGridLines="true" showRowColHeaders="true" showZeros="true" rightToLeft="false" tabSelected="false" showOutlineSymbols="true" defaultGridColor="true" view="normal" topLeftCell="A1" colorId="64" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100" workbookViewId="0"><selection pane="topLeft" activeCell="A1" activeCellId="0" sqref="A1"></selection></sheetView></sheetViews><sheetFormatPr defaultRowHeight="12.85"></sheetFormatPr><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData><printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"></printOptions><pageMargins left="0.7875" right="0.7875" top="1.05277777777778" bottom="1.05277777777778" header="0.7875" footer="0.7875"></pageMargins><pageSetup paperSize="9" scale="100" firstPageNumber="1" fitToWidth="1" fitToHeight="1" pageOrder="downThenOver" orientation="portrait" usePrinterDefaults="false" blackAndWhite="false" draft="false" cellComments="none" useFirstPageNumber="true" horizontalDpi="300" verticalDpi="300" copies="1"></pageSetup><headerFooter differentFirst="false" differentOddEven="false"><oddHeader>&amp;C&amp;&#34;Times New Roman,Regular&#34;&amp;12&amp;A</oddHeader><oddFooter>&amp;C&amp;&#34;Times New Roman,Regular&#34;&amp;12Page &amp;P</oddFooter></headerFooter></worksheet>`
	c.Assert(parts["xl/worksheets/sheet2.xml"], Equals, expectedSheet2)

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

	// theme1.xml
	expectedTheme := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office-Design">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1>
        <a:sysClr val="windowText" lastClr="000000"/>
      </a:dk1>
      <a:lt1>
        <a:sysClr val="window" lastClr="FFFFFF"/>
      </a:lt1>
      <a:dk2>
        <a:srgbClr val="1F497D"/>
      </a:dk2>
      <a:lt2>
        <a:srgbClr val="EEECE1"/>
      </a:lt2>
      <a:accent1>
        <a:srgbClr val="4F81BD"/>
      </a:accent1>
      <a:accent2>
        <a:srgbClr val="C0504D"/>
      </a:accent2>
      <a:accent3>
        <a:srgbClr val="9BBB59"/>
      </a:accent3>
      <a:accent4>
        <a:srgbClr val="8064A2"/>
      </a:accent4>
      <a:accent5>
        <a:srgbClr val="4BACC6"/>
      </a:accent5>
      <a:accent6>
        <a:srgbClr val="F79646"/>
      </a:accent6>
      <a:hlink>
        <a:srgbClr val="0000FF"/>
      </a:hlink>
      <a:folHlink>
        <a:srgbClr val="800080"/>
      </a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Cambria"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
        <a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
        <a:font script="Hang" typeface="맑은 고딕"/>
        <a:font script="Hans" typeface="宋体"/>
        <a:font script="Hant" typeface="新細明體"/>
        <a:font script="Arab" typeface="Times New Roman"/>
        <a:font script="Hebr" typeface="Times New Roman"/>
        <a:font script="Thai" typeface="Tahoma"/>
        <a:font script="Ethi" typeface="Nyala"/>
        <a:font script="Beng" typeface="Vrinda"/>
        <a:font script="Gujr" typeface="Shruti"/>
        <a:font script="Khmr" typeface="MoolBoran"/>
        <a:font script="Knda" typeface="Tunga"/>
        <a:font script="Guru" typeface="Raavi"/>
        <a:font script="Cans" typeface="Euphemia"/>
        <a:font script="Cher" typeface="Plantagenet Cherokee"/>
        <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
        <a:font script="Tibt" typeface="Microsoft Himalaya"/>
        <a:font script="Thaa" typeface="MV Boli"/>
        <a:font script="Deva" typeface="Mangal"/>
        <a:font script="Telu" typeface="Gautami"/>
        <a:font script="Taml" typeface="Latha"/>
        <a:font script="Syrc" typeface="Estrangelo Edessa"/>
        <a:font script="Orya" typeface="Kalinga"/>
        <a:font script="Mlym" typeface="Kartika"/>
        <a:font script="Laoo" typeface="DokChampa"/>
        <a:font script="Sinh" typeface="Iskoola Pota"/>
        <a:font script="Mong" typeface="Mongolian Baiti"/>
        <a:font script="Viet" typeface="Times New Roman"/>
        <a:font script="Uigh" typeface="Microsoft Uighur"/>
        <a:font script="Geor" typeface="Sylfaen"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Arial"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
        <a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
        <a:font script="Hang" typeface="맑은 고딕"/>
        <a:font script="Hans" typeface="宋体"/>
        <a:font script="Hant" typeface="新細明體"/>
        <a:font script="Arab" typeface="Arial"/>
        <a:font script="Hebr" typeface="Arial"/>
        <a:font script="Thai" typeface="Tahoma"/>
        <a:font script="Ethi" typeface="Nyala"/>
        <a:font script="Beng" typeface="Vrinda"/>
        <a:font script="Gujr" typeface="Shruti"/>
        <a:font script="Khmr" typeface="DaunPenh"/>
        <a:font script="Knda" typeface="Tunga"/>
        <a:font script="Guru" typeface="Raavi"/>
        <a:font script="Cans" typeface="Euphemia"/>
        <a:font script="Cher" typeface="Plantagenet Cherokee"/>
        <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
        <a:font script="Tibt" typeface="Microsoft Himalaya"/>
        <a:font script="Thaa" typeface="MV Boli"/>
        <a:font script="Deva" typeface="Mangal"/>
        <a:font script="Telu" typeface="Gautami"/>
        <a:font script="Taml" typeface="Latha"/>
        <a:font script="Syrc" typeface="Estrangelo Edessa"/>
        <a:font script="Orya" typeface="Kalinga"/>
        <a:font script="Mlym" typeface="Kartika"/>
        <a:font script="Laoo" typeface="DokChampa"/>
        <a:font script="Sinh" typeface="Iskoola Pota"/>
        <a:font script="Mong" typeface="Mongolian Baiti"/>
        <a:font script="Viet" typeface="Arial"/>
        <a:font script="Uigh" typeface="Microsoft Uighur"/>
        <a:font script="Geor" typeface="Sylfaen"/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="50000"/>
                <a:satMod val="300000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="35000">
              <a:schemeClr val="phClr">
                <a:tint val="37000"/>
                <a:satMod val="300000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:tint val="15000"/>
                <a:satMod val="350000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="1"/>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="100000"/>
                <a:shade val="100000"/>
                <a:satMod val="130000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:tint val="50000"/>
                <a:shade val="100000"/>
                <a:satMod val="350000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="0"/>
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr">
              <a:shade val="95000"/>
              <a:satMod val="105000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
        <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
        <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="38000"/>
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="35000"/>
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="35000"/>
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
          <a:scene3d>
            <a:camera prst="orthographicFront">
              <a:rot lat="0" lon="0" rev="0"/>
            </a:camera>
            <a:lightRig rig="threePt" dir="t">
              <a:rot lat="0" lon="0" rev="1200000"/>
            </a:lightRig>
          </a:scene3d>
          <a:sp3d>
            <a:bevelT w="63500" h="25400"/>
          </a:sp3d>
        </a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="40000"/>
                <a:satMod val="350000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="40000">
              <a:schemeClr val="phClr">
                <a:tint val="45000"/>
                <a:shade val="99000"/>
                <a:satMod val="350000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="20000"/>
                <a:satMod val="255000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:path path="circle">
            <a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
          </a:path>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="80000"/>
                <a:satMod val="300000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="30000"/>
                <a:satMod val="200000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:path path="circle">
            <a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
          </a:path>
        </a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults>
    <a:spDef>
      <a:spPr/>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:style>
        <a:lnRef idx="1">
          <a:schemeClr val="accent1"/>
        </a:lnRef>
        <a:fillRef idx="3">
          <a:schemeClr val="accent1"/>
        </a:fillRef>
        <a:effectRef idx="2">
          <a:schemeClr val="accent1"/>
        </a:effectRef>
        <a:fontRef idx="minor">
          <a:schemeClr val="lt1"/>
        </a:fontRef>
      </a:style>
    </a:spDef>
    <a:lnDef>
      <a:spPr/>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:style>
        <a:lnRef idx="2">
          <a:schemeClr val="accent1"/>
        </a:lnRef>
        <a:fillRef idx="0">
          <a:schemeClr val="accent1"/>
        </a:fillRef>
        <a:effectRef idx="1">
          <a:schemeClr val="accent1"/>
        </a:effectRef>
        <a:fontRef idx="minor">
          <a:schemeClr val="tx1"/>
        </a:fontRef>
      </a:style>
    </a:lnDef>
  </a:objectDefaults>
  <a:extraClrSchemeLst/>
</a:theme>`
	c.Assert(parts["xl/theme/theme1.xml"], Equals, expectedTheme)

	// sharedStrings.xml
	expectedXLSXSST := `<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>A cell!</t></si></sst>`
	c.Assert(parts["xl/sharedStrings.xml"], Equals, expectedXLSXSST)

	// workbook.xml.rels
	expectedXLSXWorkbookRels := `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"></Relationship><Relationship Id="rId2" Target="worksheets/sheet2.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"></Relationship><Relationship Id="rId3" Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"></Relationship><Relationship Id="rId4" Target="theme/theme1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"></Relationship><Relationship Id="rId5" Target="styles.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"></Relationship></Relationships>`
	c.Assert(parts["xl/_rels/workbook.xml.rels"], Equals, expectedXLSXWorkbookRels)

	// workbook.xml
	// Note that the following XML snippet is just pasted in here to correspond to the hack
	// added in file.go to support Apple Numbers so the test passes.
	// `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`
	expectedWorkbook := `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="Go XLSX"></fileVersion><workbookPr showObjects="all" date1904="false"></workbookPr><workbookProtection></workbookProtection><bookViews><workbookView showHorizontalScroll="true" showVerticalScroll="true" showSheetTabs="true" tabRatio="204" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"></workbookView></bookViews><sheets><sheet name="MySheet" sheetId="1" r:id="rId1" state="visible"></sheet><sheet name="AnotherSheet" sheetId="2" r:id="rId2" state="visible"></sheet></sheets><definedNames></definedNames><calcPr iterateCount="100" refMode="A1" iterateDelta="0.001"></calcPr></workbook>`
	c.Assert(parts["xl/workbook.xml"], Equals, expectedWorkbook)

	// [Content_Types].xml
	expectedContentTypes := `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Override><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"></Override><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"></Override><Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Override><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"></Override><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"></Override><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"></Override><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"></Override><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"></Override><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"></Override><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Default><Default Extension="xml" ContentType="application/xml"></Default></Types>`
	c.Assert(parts["[Content_Types].xml"], Equals, expectedContentTypes)

	// styles.xml
	//
	// For now we only allow simple string data in the
	// spreadsheet.  Style support will follow.
	expectedStyles := `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Arial"/><family val="2"/><color theme="1" /><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/></border></borders><cellStyleXfs count="1"><xf applyAlignment="0" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="0" fillId="0" fontId="0" numFmtId="0"><alignment horizontal="general" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf></cellStyleXfs><cellXfs count="1"><xf applyAlignment="0" applyBorder="0" applyFont="0" applyFill="0" applyNumberFormat="0" applyProtection="0" borderId="0" fillId="0" fontId="0" numFmtId="0"><alignment horizontal="general" indent="0" shrinkToFit="0" textRotation="0" vertical="bottom" wrapText="0"/></xf></cellXfs></styleSheet>`

	c.Assert(parts["xl/styles.xml"], Equals, expectedStyles)
}

// We can save a File as a valid XLSX file at a given path.
func (l *FileSuite) TestSaveFile(c *C) {
	var tmpPath string = c.MkDir()
	var f *File
	f = NewFile()
	sheet1, _ := f.AddSheet("MySheet")
	row1 := sheet1.AddRow()
	cell1 := row1.AddCell()
	cell1.Value = "A cell!"
	sheet2, _ := f.AddSheet("AnotherSheet")
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

func (l *FileSuite) TestMarshalFileWithHyperlinks(c *C) {
	var f *File
	f = NewFile()
	sheet1, _ := f.AddSheet("MySheet")
	row1 := sheet1.AddRow()
	cell1 := row1.AddCell()
	cell1.SetString("A cell!")
	cell1.SetHyperlink("http://www.google.com", "", "")
	c.Assert(cell1.Value, Equals, "http://www.google.com")
	sheet2, _ := f.AddSheet("AnotherSheet")
	row2 := sheet2.AddRow()
	cell2 := row2.AddCell()
	cell2.SetString("A cell!")
	cell2.SetHyperlink("http://www.google.com/index.html", "This is a hyperlink", "Click on the cell text to follow the hyperlink")
	c.Assert(cell2.Value, Equals, "This is a hyperlink")
	parts, err := f.MarshallParts()
	c.Assert(err, IsNil)
	c.Assert(len(parts), Equals, 13)
}

// We can save a File as a valid XLSX file at a given path.
func (l *FileSuite) TestSaveFileWithHyperlinks(c *C) {
	var tmpPath string = c.MkDir()
	var f *File
	f = NewFile()
	sheet1, _ := f.AddSheet("MySheet")
	row1 := sheet1.AddRow()
	cell1 := row1.AddCell()
	cell1.SetString("A cell!")
	cell1.SetHyperlink("http://www.google.com", "", "")
	c.Assert(cell1.Value, Equals, "http://www.google.com")
	sheet2, _ := f.AddSheet("AnotherSheet")
	row2 := sheet2.AddRow()
	cell2 := row2.AddCell()
	cell2.SetString("A cell!")
	cell2.SetHyperlink("http://www.google.com/index.html", "This is a hyperlink", "Click on the cell text to follow the hyperlink")
	c.Assert(cell2.Value, Equals, "This is a hyperlink")
	xlsxPath := filepath.Join(tmpPath, "TestSaveFile.xlsx")
	err := f.Save(xlsxPath)
	c.Assert(err, IsNil)

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
	c.Assert(cell1.Value, Equals, "http://www.google.com")
}

type SliceReaderSuite struct{}

var _ = Suite(&SliceReaderSuite{})

func (s *SliceReaderSuite) TestFileToSlice(c *C) {
	output, err := FileToSlice("./testdocs/testfile.xlsx")
	c.Assert(err, IsNil)
	fileToSliceCheckOutput(c, output)
}

func (s *SliceReaderSuite) TestFileToSliceMissingCol(c *C) {
	// Test xlsx file with the A column removed
	_, err := FileToSlice("./testdocs/testFileToSlice.xlsx")
	c.Assert(err, IsNil)
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

func (s *SliceReaderSuite) TestFileToSliceUnmerged(c *C) {
	output, err := FileToSliceUnmerged("./testdocs/testfile.xlsx")
	c.Assert(err, IsNil)
	fileToSliceCheckOutput(c, output)

	// merged cells
	output, err = FileToSliceUnmerged("./testdocs/merged_cells.xlsx")
	c.Assert(err, IsNil)
	c.Assert(output[0][6][2], Equals, "Happy New Year!")
	c.Assert(output[0][6][1], Equals, "Happy New Year!")
	c.Assert(output[0][1][0], Equals, "01.01.2016")
	c.Assert(output[0][2][0], Equals, "01.01.2016")
}

func (l *FileSuite) TestReadWorkbookWithTypes(c *C) {
	var xlsxFile *File
	var err error

	xlsxFile, err = OpenFile("./testdocs/testcelltypes.xlsx")
	c.Assert(err, IsNil)
	c.Assert(len(xlsxFile.Sheets), Equals, 1)
	sheet := xlsxFile.Sheet["Sheet1"]
	c.Assert(len(sheet.Rows), Equals, 8)
	c.Assert(len(sheet.Rows[0].Cells), Equals, 2)

	// string 1
	c.Assert(sheet.Rows[0].Cells[0].Type(), Equals, CellTypeString)
	if val, err := sheet.Rows[0].Cells[0].FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "hello world")
	}

	// string 2
	c.Assert(sheet.Rows[1].Cells[0].Type(), Equals, CellTypeString)
	if val, err := sheet.Rows[1].Cells[0].FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "日本語")
	}

	// integer
	c.Assert(sheet.Rows[2].Cells[0].Type(), Equals, CellTypeNumeric)
	intValue, _ := sheet.Rows[2].Cells[0].Int()
	c.Assert(intValue, Equals, 12345)

	// float
	c.Assert(sheet.Rows[3].Cells[0].Type(), Equals, CellTypeNumeric)
	floatValue, _ := sheet.Rows[3].Cells[0].Float()
	c.Assert(floatValue, Equals, 1.024)

	// Now it can't detect date
	c.Assert(sheet.Rows[4].Cells[0].Type(), Equals, CellTypeNumeric)
	intValue, _ = sheet.Rows[4].Cells[0].Int()
	c.Assert(intValue, Equals, 40543)

	// bool
	c.Assert(sheet.Rows[5].Cells[0].Type(), Equals, CellTypeBool)
	c.Assert(sheet.Rows[5].Cells[0].Bool(), Equals, true)

	// formula
	c.Assert(sheet.Rows[6].Cells[0].Type(), Equals, CellTypeNumeric)
	c.Assert(sheet.Rows[6].Cells[0].Formula(), Equals, "10+20")
	c.Assert(sheet.Rows[6].Cells[0].Value, Equals, "30")

	// error
	c.Assert(sheet.Rows[7].Cells[0].Type(), Equals, CellTypeError)
	c.Assert(sheet.Rows[7].Cells[0].Formula(), Equals, "10/0")
	c.Assert(sheet.Rows[7].Cells[0].Value, Equals, "#DIV/0!")
}

func (s *SliceReaderSuite) TestFileWithEmptyRows(c *C) {
	f, err := OpenFile("./testdocs/empty_rows.xlsx")
	c.Assert(err, IsNil)
	sheet, ok := f.Sheet["EmptyRows"]
	c.Assert(ok, Equals, true)

	if val, err := sheet.Cell(0, 0).FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "")
	}
	if val, err := sheet.Cell(2, 0).FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "A3")
	}
}

func (s *SliceReaderSuite) TestFileWithEmptyCols(c *C) {
	f, err := OpenFile("./testdocs/empty_rows.xlsx")
	c.Assert(err, IsNil)
	sheet, ok := f.Sheet["EmptyCols"]
	c.Assert(ok, Equals, true)

	if val, err := sheet.Cell(0, 0).FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "")
	}
	if val, err := sheet.Cell(0, 2).FormattedValue(); err != nil {
		c.Error(err)
	} else {
		c.Assert(val, Equals, "C1")
	}
}
