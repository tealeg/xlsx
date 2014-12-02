package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strconv"
)

// File is a high level structure providing a slice of Sheet structs
// to the user.
type File struct {
	worksheets     map[string]*zip.File
	numFmtRefTable map[int]xlsxNumFmt
	referenceTable *RefTable
	Date1904       bool
	styles         *xlsxStyles
	Sheets         []*Sheet
	Sheet          map[string]*Sheet
}

// Create a new File
func NewFile() (file *File) {
	file = &File{}
	file.Sheet = make(map[string]*Sheet)
	file.Sheets = make([]*Sheet, 0)
	return
}

// OpenFile() take the name of an XLSX file and returns a populated
// xlsx.File struct for it.
func OpenFile(filename string) (*File, error) {
	var f *zip.ReadCloser
	f, err := zip.OpenReader(filename)
	if err != nil {
		return nil, err
	}
	return ReadZip(f)
}

// A convenient wrapper around File.ToSlice, FileToSlice will
// return the raw data contained in an Excel XLSX file as three
// dimensional slice.  The first index represents the sheet number,
// the second the row number, and the third the cell number.
//
// For example:
//
//    var mySlice [][][]string
//    var value string
//    mySlice = xlsx.FileToSlice("myXLSX.xlsx")
//    value = mySlice[0][0][0]
//
// Here, value would be set to the raw value of the cell A1 in the
// first sheet in the XLSX file.
func FileToSlice(path string) ([][][]string, error) {
	f, err := OpenFile(path)
	if err != nil {
		return nil, err
	}
	return f.ToSlice()
}

// Save the File to an xlsx file at the provided path.
func (f *File) Save(path string) (err error) {
	var parts map[string]string
	var target *os.File
	var zipWriter *zip.Writer

	parts, err = f.MarshallParts()
	if err != nil {
		return
	}

	target, err = os.Create(path)
	if err != nil {
		return
	}

	zipWriter = zip.NewWriter(target)

	for partName, part := range parts {
		var writer io.Writer
		writer, err = zipWriter.Create(partName)
		if err != nil {
			return
		}
		_, err = writer.Write([]byte(part))
		if err != nil {
			return
		}
	}
	err = zipWriter.Close()
	if err != nil {
		return
	}

	return target.Close()
}

// Add a new Sheet, with the provided name, to a File
func (f *File) AddSheet(sheetName string) (sheet *Sheet) {
	sheet = &Sheet{Name: sheetName}
	f.Sheet[sheetName] = sheet
	f.Sheets = append(f.Sheets, sheet)
	return sheet
}

func (f *File) makeWorkbook() xlsxWorkbook {
	var workbook xlsxWorkbook
	workbook = xlsxWorkbook{}
	workbook.FileVersion = xlsxFileVersion{}
	workbook.FileVersion.AppName = "Go XLSX"
	workbook.WorkbookPr = xlsxWorkbookPr{BackupFile: false}
	workbook.BookViews = xlsxBookViews{}
	workbook.BookViews.WorkBookView = make([]xlsxWorkBookView, 1)
	workbook.BookViews.WorkBookView[0] = xlsxWorkBookView{}
	workbook.Sheets = xlsxSheets{}
	workbook.Sheets.Sheet = make([]xlsxSheet, len(f.Sheets))
	return workbook
}

// Construct a map of file name to XML content representing the file
// in terms of the structure of an XLSX file.
func (f *File) MarshallParts() (map[string]string, error) {
	var parts map[string]string
	var refTable *RefTable = NewSharedStringRefTable()
	refTable.isWrite = true
	var workbookRels WorkBookRels = make(WorkBookRels)
	var err error
	var workbook xlsxWorkbook
	var types xlsxTypes = MakeDefaultContentTypes()

	marshal := func(thing interface{}) (string, error) {
		body, err := xml.MarshalIndent(thing, "  ", "  ")
		if err != nil {
			return "", err
		}
		return xml.Header + string(body), nil
	}

	parts = make(map[string]string)
	workbook = f.makeWorkbook()
	sheetIndex := 1

	styles := &xlsxStyles{}
	for _, sheet := range f.Sheets {
		xSheet := sheet.makeXLSXSheet(refTable, styles)
		rId := fmt.Sprintf("rId%d", sheetIndex)
		sheetId := strconv.Itoa(sheetIndex)
		sheetPath := fmt.Sprintf("worksheets/sheet%d.xml", sheetIndex)
		partName := "xl/" + sheetPath
		types.Overrides = append(
			types.Overrides,
			xlsxOverride{
				PartName:    "/" + partName,
				ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"})
		workbookRels[rId] = sheetPath
		workbook.Sheets.Sheet[sheetIndex-1] = xlsxSheet{
			Name:    sheet.Name,
			SheetId: sheetId,
			Id:      rId}
		parts[partName], err = marshal(xSheet)
		if err != nil {
			return parts, err
		}
		sheetIndex++
	}

	parts["xl/workbook.xml"], err = marshal(workbook)
	if err != nil {
		return parts, err
	}

	parts["_rels/.rels"] = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`

	parts["docProps/app.xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <TotalTime>0</TotalTime>
  <Application>Go XLSX</Application>
</Properties>`

	// TODO - do this properly, modification and revision information
	parts["docProps/core.xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></cp:coreProperties>`

	parts["xl/theme/theme1.xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
        <a:latin typeface="Calibri"/>
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

	xSST := refTable.makeXLSXSST()
	parts["xl/sharedStrings.xml"], err = marshal(xSST)
	if err != nil {
		return parts, err
	}

	xWRel := workbookRels.MakeXLSXWorkbookRels()

	parts["xl/_rels/workbook.xml.rels"], err = marshal(xWRel)
	if err != nil {
		return parts, err
	}

	parts["[Content_Types].xml"], err = marshal(types)
	if err != nil {
		return parts, err
	}

	parts["xl/styles.xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
</styleSheet>`

	return parts, nil
}

// Return the raw data contained in the File as three
// dimensional slice.  The first index represents the sheet number,
// the second the row number, and the third the cell number.
//
// For example:
//
//    var mySlice [][][]string
//    var value string
//    mySlice = xlsx.FileToSlice("myXLSX.xlsx")
//    value = mySlice[0][0][0]
//
// Here, value would be set to the raw value of the cell A1 in the
// first sheet in the XLSX file.
func (file *File) ToSlice() (output [][][]string, err error) {
	output = [][][]string{}
	for _, sheet := range file.Sheets {
		s := [][]string{}
		for _, row := range sheet.Rows {
			if row == nil {
				continue
			}
			r := []string{}
			for _, cell := range row.Cells {
				r = append(r, cell.String())
			}
			s = append(s, r)
		}
		output = append(output, s)
	}
	return output, nil
}
