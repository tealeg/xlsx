package xlsx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"
	"unicode/utf8"
)

// File is a high level structure providing a slice of Sheet structs
// to the user.
type File struct {
	worksheets           map[string]*zip.File
	worksheetRels        map[string]*zip.File
	referenceTable       *RefTable
	Date1904             bool
	styles               *xlsxStyleSheet
	Sheets               []*Sheet
	Sheet                map[string]*Sheet
	theme                *theme
	DefinedNames         []*xlsxDefinedName
	cellStoreConstructor CellStoreConstructor
	rowLimit             int
}

const NoRowLimit int = -1

type FileOption func(f *File)

// RowLimit will limit the rows handled in any given sheet to the
// first n, where n is the number of rows.
func RowLimit(n int) FileOption {
	return func(f *File) {
		f.rowLimit = n
	}
}

// NewFile creates a new File struct. You may pass it zero, one or
// many FileOption functions that affect the behaviour of the file.
func NewFile(options ...FileOption) *File {
	f := &File{
		Sheet:                make(map[string]*Sheet),
		Sheets:               make([]*Sheet, 0),
		DefinedNames:         make([]*xlsxDefinedName, 0),
		rowLimit:             NoRowLimit,
		cellStoreConstructor: NewMemoryCellStore,
	}
	for _, opt := range options {
		opt(f)
	}
	return f
}

// OpenFile will take the name of an XLSX file and returns a populated
// xlsx.File struct for it.  You may pass it zero, one or
// many FileOption functions that affect the behaviour of the file.
func OpenFile(fileName string, options ...FileOption) (file *File, err error) {
	var z *zip.ReadCloser
	wrap := func(err error) (*File, error) {
		return nil, fmt.Errorf("OpenFile: %w", err)
	}

	z, err = zip.OpenReader(fileName)
	if err != nil {
		return wrap(err)
	}
	file, err = ReadZip(z, options...)
	if err != nil {
		return wrap(err)
	}
	return file, nil
}

// OpenBinary() take bytes of an XLSX file and returns a populated
// xlsx.File struct for it.
func OpenBinary(bs []byte, options ...FileOption) (*File, error) {
	r := bytes.NewReader(bs)
	return OpenReaderAt(r, int64(r.Len()), options...)

}

// OpenReaderAt() take io.ReaderAt of an XLSX file and returns a populated
// xlsx.File struct for it.
func OpenReaderAt(r io.ReaderAt, size int64, options ...FileOption) (*File, error) {
	file, err := zip.NewReader(r, size)
	if err != nil {
		return nil, err
	}
	return ReadZipReader(file, options...)
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
func FileToSlice(path string, options ...FileOption) ([][][]string, error) {
	f, err := OpenFile(path, options...)
	if err != nil {
		return nil, err
	}
	return f.ToSlice()
}

// FileToSliceUnmerged is a wrapper around File.ToSliceUnmerged.
// It returns the raw data contained in an Excel XLSX file as three
// dimensional slice. Merged cells will be unmerged. Covered cells become the
// values of theirs origins.
func FileToSliceUnmerged(path string, options ...FileOption) ([][][]string, error) {
	f, err := OpenFile(path, options...)
	if err != nil {
		return nil, err
	}
	return f.ToSliceUnmerged()
}

// Save the File to an xlsx file at the provided path.
func (f *File) Save(path string) (err error) {
	wrap := func(err error) error {
		return fmt.Errorf("File.Save(%s): %w", path, err)
	}
	target, err := os.Create(path)
	if err != nil {
		return wrap(err)
	}
	err = f.Write(target)
	if err != nil {
		return wrap(err)
	}
	err = target.Close()
	if err != nil {
		return wrap(err)
	}
	return nil
}

// Write the File to io.Writer as xlsx
func (f *File) Write(writer io.Writer) error {
	wrap := func(err error) error {
		return fmt.Errorf("File.Write: %w", err)
	}
	zipWriter := zip.NewWriter(writer)
	err := f.MarshallParts(zipWriter)
	if err != nil {
		return wrap(err)
	}
	err = zipWriter.Close()
	if err != nil {
		return wrap(err)
	}
	return nil
}

// AddSheet Add a new Sheet, with the provided name, to a File.
// The minimum sheet name length is 1 character. If the sheet name length is less an error is thrown.
// The maximum sheet name length is 31 characters. If the sheet name length is exceeded an error is thrown.
// These special characters are also not allowed: : \ / ? * [ ]
func (f *File) AddSheet(sheetName string) (*Sheet, error) {
	return f.AddSheetWithCellStore(sheetName, f.cellStoreConstructor)
}

func (f *File) AddSheetWithCellStore(sheetName string, constructor CellStoreConstructor) (*Sheet, error) {
	var err error
	if _, exists := f.Sheet[sheetName]; exists {
		return nil, fmt.Errorf("duplicate sheet name '%s'.", sheetName)
	}
	runeLength := utf8.RuneCountInString(sheetName)
	if runeLength > 31 || runeLength == 0 {
		return nil, fmt.Errorf("sheet name must be 31 or fewer characters long.  It is currently '%d' characters long", runeLength)
	}
	// Iterate over the runes
	for _, r := range sheetName {
		// Excel forbids : \ / ? * [ ]
		if r == ':' || r == '\\' || r == '/' || r == '?' || r == '*' || r == '[' || r == ']' {
			return nil, fmt.Errorf("sheet name must not contain any restricted characters : \\ / ? * [ ] but contains '%s'", string(r))
		}
	}
	sheet := &Sheet{
		Name:     sheetName,
		File:     f,
		Selected: len(f.Sheets) == 0,
		Cols:     &ColStore{},
	}

	sheet.cellStore, err = constructor()
	if err != nil {
		return nil, err
	}
	f.Sheet[sheetName] = sheet
	f.Sheets = append(f.Sheets, sheet)
	return sheet, nil
}

// Appends an existing Sheet, with the provided name, to a File
func (f *File) AppendSheet(sheet Sheet, sheetName string) (*Sheet, error) {
	if _, exists := f.Sheet[sheetName]; exists {
		return nil, fmt.Errorf("duplicate sheet name '%s'.", sheetName)
	}
	sheet.Name = sheetName
	sheet.File = f
	sheet.Selected = len(f.Sheets) == 0
	f.Sheet[sheetName] = &sheet
	f.Sheets = append(f.Sheets, &sheet)
	return &sheet, nil
}

func (f *File) makeWorkbook() xlsxWorkbook {
	return xlsxWorkbook{
		FileVersion: xlsxFileVersion{AppName: "Go XLSX"},
		WorkbookPr:  xlsxWorkbookPr{ShowObjects: "all"},
		BookViews: xlsxBookViews{
			WorkBookView: []xlsxWorkBookView{
				{
					ShowHorizontalScroll: true,
					ShowSheetTabs:        true,
					ShowVerticalScroll:   true,
					TabRatio:             204,
					WindowHeight:         8192,
					WindowWidth:          16384,
					XWindow:              "0",
					YWindow:              "0",
				},
			},
		},
		Sheets: xlsxSheets{Sheet: make([]xlsxSheet, len(f.Sheets))},
		CalcPr: xlsxCalcPr{
			IterateCount: 100,
			RefMode:      "A1",
			Iterate:      false,
			IterateDelta: 0.001,
		},
	}
}

// Some tools that read XLSX files have very strict requirements about
// the structure of the input XML.  In particular both Numbers on the Mac
// and SAS dislike inline XML namespace declarations, or namespace
// prefixes that don't match the ones that Excel itself uses.  This is a
// problem because the Go XML library doesn't multiple namespace
// declarations in a single element of a document.  This function is a
// horrible hack to fix that after the XML marshalling is completed.
func replaceRelationshipsNameSpace(workbookMarshal string) string {
	newWorkbook := strings.Replace(workbookMarshal, `xmlns:relationships="http://schemas.openxmlformats.org/officeDocument/2006/relationships" relationships:id`, `r:id`, -1)
	// Dirty hack to fix issues #63 and #91; encoding/xml currently
	// "doesn't allow for additional namespaces to be defined in the
	// root element of the document," as described by @tealeg in the
	// comments for #63.
	oldXmlns := `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`
	newXmlns := `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
	return strings.Replace(newWorkbook, oldXmlns, newXmlns, 1)
}

func addRelationshipNameSpaceToWorksheet(worksheetMarshal string) string {
	oldXmlns := `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`
	newXmlns := `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
	newSheetMarshall := strings.Replace(worksheetMarshal, oldXmlns, newXmlns, 1)

	oldHyperlink := `<hyperlink id=`
	newHyperlink := `<hyperlink r:id=`
	newSheetMarshall = strings.Replace(newSheetMarshall, oldHyperlink, newHyperlink, -1)
	return newSheetMarshall
}

// MakeStreamParts constructs a map of file name to XML content
// representing the file in terms of the structure of an XLSX file.
func (f *File) MakeStreamParts() (map[string]string, error) {
	var parts map[string]string
	var refTable *RefTable = NewSharedStringRefTable()
	refTable.isWrite = true
	var workbookRels WorkBookRels = make(WorkBookRels)
	var err error
	var workbook xlsxWorkbook
	var types xlsxTypes = MakeDefaultContentTypes()

	marshal := func(thing interface{}) (string, error) {
		body, err := xml.Marshal(thing)
		if err != nil {
			return "", err
		}
		return xml.Header + string(body), nil
	}

	parts = make(map[string]string)
	workbook = f.makeWorkbook()
	sheetIndex := 1

	if f.styles == nil {
		f.styles = newXlsxStyleSheet(f.theme)
	}
	f.styles.reset()
	if len(f.Sheets) == 0 {
		err := errors.New("Workbook must contains atleast one worksheet")
		return nil, err
	}
	for _, sheet := range f.Sheets {
		// Make sure we don't lose the current state!
		err := sheet.cellStore.WriteRow(sheet.currentRow)
		if err != nil {
			return nil, err
		}

		xSheetRels := sheet.makeXLSXSheetRelations()
		xSheet := sheet.makeXLSXSheet(refTable, f.styles, xSheetRels)
		rId := fmt.Sprintf("rId%d", sheetIndex)
		sheetId := strconv.Itoa(sheetIndex)
		sheetPath := fmt.Sprintf("worksheets/sheet%d.xml", sheetIndex)
		partName := "xl/" + sheetPath
		relPartName := fmt.Sprintf("xl/worksheets/_rels/sheet%d.xml.rels", sheetIndex)
		sheetState := sheetStateVisible
		if sheet.Hidden {
			sheetState = sheetStateHidden
		}
		types.Overrides = append(
			types.Overrides,
			xlsxOverride{
				PartName:    "/" + partName,
				ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"})
		workbookRels[rId] = sheetPath
		workbook.Sheets.Sheet[sheetIndex-1] = xlsxSheet{
			Name:    sheet.Name,
			SheetId: sheetId,
			Id:      rId,
			State:   sheetState}

		worksheetMarshal, err := marshal(xSheet)
		if err != nil {
			return parts, err
		}
		worksheetMarshal = addRelationshipNameSpaceToWorksheet(worksheetMarshal)
		parts[partName] = worksheetMarshal
		if xSheetRels != nil {
			parts[relPartName], err = marshal(xSheetRels)
			if err != nil {
				return parts, err
			}
		}
		sheetIndex++
	}

	workbookMarshal, err := marshal(workbook)
	if err != nil {
		return parts, err
	}
	workbookMarshal = replaceRelationshipsNameSpace(workbookMarshal)
	parts["xl/workbook.xml"] = workbookMarshal
	if err != nil {
		return parts, err
	}

	parts["_rels/.rels"] = TEMPLATE__RELS_DOT_RELS
	parts["docProps/app.xml"] = TEMPLATE_DOCPROPS_APP
	// TODO - do this properly, modification and revision information
	parts["docProps/core.xml"] = TEMPLATE_DOCPROPS_CORE
	parts["xl/theme/theme1.xml"] = TEMPLATE_XL_THEME_THEME

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

	parts["xl/styles.xml"], err = f.styles.Marshal()
	if err != nil {
		return parts, err
	}

	return parts, nil
}

// MarshallParts constructs a map of file name to XML content representing the file
// in terms of the structure of an XLSX file.
func (f *File) MarshallParts(zipWriter *zip.Writer) error {
	var refTable *RefTable = NewSharedStringRefTable()
	refTable.isWrite = true
	var workbookRels WorkBookRels = make(WorkBookRels)
	var err error
	var workbook xlsxWorkbook
	var types xlsxTypes = MakeDefaultContentTypes()

	wrap := func(err error) error {
		return fmt.Errorf("MarshallParts: %w", err)
	}

	marshal := func(thing interface{}) (string, error) {
		body, err := xml.Marshal(thing)
		if err != nil {
			return "", fmt.Errorf("xml.Marshal: %w", err)
		}
		return xml.Header + string(body), nil
	}

	writePart := func(partName, part string) error {
		w, err := zipWriter.Create(partName)
		if err != nil {
			return fmt.Errorf("zipwriter.Create(%s): %w", partName, err)
		}
		_, err = w.Write([]byte(part))
		if err != nil {
			return fmt.Errorf("zipwriter.Write(%s): %w", part, err)
		}
		return nil
	}

	// parts = make(map[string]string)
	workbook = f.makeWorkbook()
	sheetIndex := 1

	if f.styles == nil {
		f.styles = newXlsxStyleSheet(f.theme)
	}
	f.styles.reset()
	if len(f.Sheets) == 0 {
		err := errors.New("MarshalParts: Workbook must contain at least one worksheet")
		return wrap(err)
	}
	for _, sheet := range f.Sheets {
		if sheet.currentRow != nil {
			// Make sure we don't lose the current state!
			err := sheet.cellStore.WriteRow(sheet.currentRow)
			if err != nil {
				return wrap(err)
			}
		}

		xSheetRels := sheet.makeXLSXSheetRelations()
		rId := fmt.Sprintf("rId%d", sheetIndex)
		sheetId := strconv.Itoa(sheetIndex)
		sheetPath := fmt.Sprintf("worksheets/sheet%d.xml", sheetIndex)
		partName := "xl/" + sheetPath
		relPartName := fmt.Sprintf("xl/worksheets/_rels/sheet%d.xml.rels", sheetIndex)
		types.Overrides = append(
			types.Overrides,
			xlsxOverride{
				PartName:    "/" + partName,
				ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"})
		workbookRels[rId] = sheetPath
		workbook.Sheets.Sheet[sheetIndex-1] = xlsxSheet{
			Name:    sheet.Name,
			SheetId: sheetId,
			Id:      rId,
			State:   sheet.getState()}

		w, err := zipWriter.Create(partName)
		if err != nil {
			return wrap(err)
		}
		err = sheet.MarshalSheet(w, refTable, f.styles, xSheetRels)
		if err != nil {
			return wrap(err)
		}

		if xSheetRels != nil {
			relPart, err := marshal(xSheetRels)
			if err != nil {
				return wrap(err)
			}
			err = writePart(relPartName, relPart)
			if err != nil {
				return wrap(err)
			}
		}
		sheetIndex++
	}

	workbookMarshal, err := marshal(workbook)
	if err != nil {
		return err
	}
	workbookMarshal = replaceRelationshipsNameSpace(workbookMarshal)
	err = writePart("xl/workbook.xml", workbookMarshal)
	if err != nil {
		return err
	}

	err = writePart("_rels/.rels", TEMPLATE__RELS_DOT_RELS)
	if err != nil {
		return err
	}

	err = writePart("docProps/app.xml", TEMPLATE_DOCPROPS_APP)
	if err != nil {
		return err
	}
	// TODO - do this properly, modification and revision information
	err = writePart("docProps/core.xml", TEMPLATE_DOCPROPS_CORE)
	if err != nil {
		return err
	}
	err = writePart("xl/theme/theme1.xml", TEMPLATE_XL_THEME_THEME)
	if err != nil {
		return err
	}

	xSST := refTable.makeXLSXSST()
	sharedStrings, err := marshal(xSST)
	if err != nil {
		return err
	}
	err = writePart("xl/sharedStrings.xml", sharedStrings)
	if err != nil {
		return err
	}

	xWRel := workbookRels.MakeXLSXWorkbookRels()
	relPart, err := marshal(xWRel)
	if err != nil {
		return err
	}

	err = writePart("xl/_rels/workbook.xml.rels", relPart)
	if err != nil {
		return err
	}

	typesS, err := marshal(types)
	if err != nil {
		return err
	}
	err = writePart("[Content_Types].xml", typesS)
	if err != nil {
		return err
	}

	styles, err := f.styles.Marshal()
	if err != nil {
		return err
	}

	return writePart("xl/styles.xml", styles)
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
func (f *File) ToSlice() (output [][][]string, err error) {
	output = [][][]string{}
	for _, sheet := range f.Sheets {
		s := [][]string{}
		err := sheet.ForEachRow(func(row *Row) error {
			r := []string{}
			err := row.ForEachCell(func(cell *Cell) error {
				str, err := cell.FormattedValue()
				if err != nil {
					// Recover from strconv.NumError if the value is an empty string,
					// and insert an empty string in the output.
					if numErr, ok := err.(*strconv.NumError); ok && numErr.Num == "" {
						str = ""
					} else {
						return err
					}
				}
				r = append(r, str)
				return nil
			})
			if err != nil {
				return err
			}

			s = append(s, r)
			return nil
		})
		if err != nil {
			return output, err
		}
		output = append(output, s)
	}
	return output, nil
}

// ToSliceUnmerged returns the raw data contained in the File as three
// dimensional slice (s. method ToSlice).
// A covered cell become the value of its origin cell.
// Example: table where A1:A2 merged.
// | 01.01.2011 | Bread | 20 |
// |            | Fish  | 70 |
// This sheet will be converted to the slice:
// [  [01.01.2011 Bread 20]
// 		[01.01.2011 Fish 70] ]
func (f *File) ToSliceUnmerged() (output [][][]string, err error) {
	output, err = f.ToSlice()
	if err != nil {
		return nil, err
	}

	for s, sheet := range f.Sheets {
		err := sheet.ForEachRow(func(row *Row) error {
			r := row.num
			err := row.ForEachCell(func(cell *Cell) error {
				c := cell.num
				if cell.HMerge > 0 {
					for i := c + 1; i <= c+cell.HMerge; i++ {
						output[s][r][i] = output[s][r][c]
					}
				}

				if cell.VMerge > 0 {
					for i := r + 1; i <= r+cell.VMerge; i++ {
						output[s][i][c] = output[s][r][c]
					}
				}
				return nil
			})
			return err
		})
		if err != nil {
			return output, err
		}
	}

	return output, nil
}
