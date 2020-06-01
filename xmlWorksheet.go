package xlsx

import (
	"encoding/xml"
	"errors"
	"fmt"
	"reflect"
	"strconv"
	"strings"

	"github.com/shabbyrobe/xmlwriter"
)

type RelationshipType string

const (
	RelationshipTypeHyperlink RelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)

type RelationshipTargetMode string

const (
	RelationshipTargetModeExternal RelationshipTargetMode = "External"
)

// xlsxWorksheetRels contains xlsxWorksheetRelation
type xlsxWorksheetRels struct {
	XMLName       xml.Name                `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Relationships []xlsxWorksheetRelation `xml:"Relationship"`
}

type xlsxWorksheetRelation struct {
	Id         string                 `xml:"Id,attr"`
	Type       RelationshipType       `xml:"Type,attr"`
	Target     string                 `xml:"Target,attr"`
	TargetMode RelationshipTargetMode `xml:"TargetMode,attr"`
}

// xlsxWorksheet directly maps the worksheet element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxWorksheet struct {
	XMLName         xml.Name             `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main worksheet"`
	XMLNSR          string               `xml:"xmlns:r,attr"`
	SheetPr         xlsxSheetPr          `xml:"sheetPr"`
	Dimension       xlsxDimension        `xml:"dimension"`
	SheetViews      xlsxSheetViews       `xml:"sheetViews"`
	SheetFormatPr   xlsxSheetFormatPr    `xml:"sheetFormatPr"`
	Cols            *xlsxCols            `xml:"cols,omitempty"`
	SheetData       xlsxSheetData        `xml:"sheetData"`
	Hyperlinks      *xlsxHyperlinks      `xml:"hyperlinks,omitempty"`
	DataValidations *xlsxDataValidations `xml:"dataValidations"`
	AutoFilter      *xlsxAutoFilter      `xml:"autoFilter,omitempty"`
	MergeCells      *xlsxMergeCells      `xml:"mergeCells,omitempty"`
	PrintOptions    *xlsxPrintOptions    `xml:"printOptions,omitempty"`
	PageMargins     *xlsxPageMargins     `xml:"pageMargins,omitempty"`
	PageSetUp       *xlsxPageSetUp       `xml:"pageSetup,omitempty"`
	HeaderFooter    *xlsxHeaderFooter    `xml:"headerFooter,omitempty"`
}

// xlsxHeaderFooter directly maps the headerFooter element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxHeaderFooter struct {
	DifferentFirst   *bool           `xml:"differentFirst,attr,omitempty"`
	DifferentOddEven *bool           `xml:"differentOddEven,attr,omitempty"`
	OddHeader        []xlsxOddHeader `xml:"oddHeader"`
	OddFooter        []xlsxOddFooter `xml:"oddFooter"`
}

// xlsxOddHeader directly maps the oddHeader element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxOddHeader struct {
	Content string `xml:",chardata"`
}

// xlsxOddFooter directly maps the oddFooter element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxOddFooter struct {
	Content string `xml:",chardata"`
}

// xlsxPageSetUp directly maps the pageSetup element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxPageSetUp struct {
	PaperSize          string  `xml:"paperSize,attr"`
	Scale              int     `xml:"scale,attr"`
	FirstPageNumber    int     `xml:"firstPageNumber,attr"`
	FitToWidth         int     `xml:"fitToWidth,attr"`
	FitToHeight        int     `xml:"fitToHeight,attr"`
	PageOrder          string  `xml:"pageOrder,attr"`
	Orientation        string  `xml:"orientation,attr"`
	UsePrinterDefaults bool    `xml:"usePrinterDefaults,attr"`
	BlackAndWhite      bool    `xml:"blackAndWhite,attr"`
	Draft              bool    `xml:"draft,attr"`
	CellComments       string  `xml:"cellComments,attr"`
	UseFirstPageNumber bool    `xml:"useFirstPageNumber,attr"`
	HorizontalDPI      float32 `xml:"horizontalDpi,attr"`
	VerticalDPI        float32 `xml:"verticalDpi,attr"`
	Copies             int     `xml:"copies,attr"`
}

// xlsxPrintOptions directly maps the printOptions element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxPrintOptions struct {
	Headings           bool `xml:"headings,attr"`
	GridLines          bool `xml:"gridLines,attr"`
	GridLinesSet       bool `xml:"gridLinesSet,attr"`
	HorizontalCentered bool `xml:"horizontalCentered,attr"`
	VerticalCentered   bool `xml:"verticalCentered,attr"`
}

// xlsxPageMargins directly maps the pageMargins element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxPageMargins struct {
	Left   float64 `xml:"left,attr"`
	Right  float64 `xml:"right,attr"`
	Top    float64 `xml:"top,attr"`
	Bottom float64 `xml:"bottom,attr"`
	Header float64 `xml:"header,attr"`
	Footer float64 `xml:"footer,attr"`
}

// xlsxSheetFormatPr directly maps the sheetFormatPr element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheetFormatPr struct {
	DefaultColWidth  float64 `xml:"defaultColWidth,attr,omitempty"`
	DefaultRowHeight float64 `xml:"defaultRowHeight,attr"`
	OutlineLevelCol  uint8   `xml:"outlineLevelCol,attr,omitempty"`
	OutlineLevelRow  uint8   `xml:"outlineLevelRow,attr,omitempty"`
}

// xlsxSheetViews directly maps the sheetViews element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheetViews struct {
	SheetView []xlsxSheetView `xml:"sheetView"`
}

// xlsxSheetView directly maps the sheetView element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheetView struct {
	WindowProtection        bool            `xml:"windowProtection,attr"`
	ShowFormulas            bool            `xml:"showFormulas,attr"`
	ShowGridLines           bool            `xml:"showGridLines,attr"`
	ShowRowColHeaders       bool            `xml:"showRowColHeaders,attr"`
	ShowZeros               bool            `xml:"showZeros,attr"`
	RightToLeft             bool            `xml:"rightToLeft,attr"`
	TabSelected             bool            `xml:"tabSelected,attr"`
	ShowOutlineSymbols      bool            `xml:"showOutlineSymbols,attr"`
	DefaultGridColor        bool            `xml:"defaultGridColor,attr"`
	View                    string          `xml:"view,attr"`
	TopLeftCell             string          `xml:"topLeftCell,attr"`
	ColorId                 int             `xml:"colorId,attr"`
	ZoomScale               float64         `xml:"zoomScale,attr"`
	ZoomScaleNormal         float64         `xml:"zoomScaleNormal,attr"`
	ZoomScalePageLayoutView float64         `xml:"zoomScalePageLayoutView,attr"`
	WorkbookViewId          int             `xml:"workbookViewId,attr"`
	Pane                    *xlsxPane       `xml:"pane"`
	Selection               []xlsxSelection `xml:"selection"`
}

// xlsxSelection directly maps the selection element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSelection struct {
	Pane         string `xml:"pane,attr"`
	ActiveCell   string `xml:"activeCell,attr"`
	ActiveCellId int    `xml:"activeCellId,attr"`
	SQRef        string `xml:"sqref,attr"`
}

// xlsxSelection directly maps the selection element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxPane struct {
	XSplit      float64 `xml:"xSplit,attr"`
	YSplit      float64 `xml:"ySplit,attr"`
	TopLeftCell string  `xml:"topLeftCell,attr"`
	ActivePane  string  `xml:"activePane,attr"`
	State       string  `xml:"state,attr"` // Either "split" or "frozen"
}

// xlsxSheetPr directly maps the sheetPr element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheetPr struct {
	FilterMode  bool              `xml:"filterMode,attr"`
	PageSetUpPr []xlsxPageSetUpPr `xml:"pageSetUpPr"`
}

// xlsxPageSetUpPr directly maps the pageSetupPr element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxPageSetUpPr struct {
	FitToPage bool `xml:"fitToPage,attr"`
}

// xlsxCols directly maps the cols element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxCols struct {
	Col []xlsxCol `xml:"col"`
}

// xlsxCol directly maps the col element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxCol struct {
	Collapsed    *bool    `xml:"collapsed,attr,omitempty"`
	Hidden       *bool    `xml:"hidden,attr,omitempty"`
	Max          int      `xml:"max,attr"`
	Min          int      `xml:"min,attr"`
	Style        *int     `xml:"style,attr,omitempty"`
	Width        *float64 `xml:"width,attr,omitempty"`
	CustomWidth  *bool    `xml:"customWidth,attr,omitempty"`
	OutlineLevel *uint8   `xml:"outlineLevel,attr,omitempty"`
	BestFit      *bool    `xml:"bestFit,attr,omitempty"`
	Phonetic     *bool    `xml:"phonetic,attr,omitempty"`
}

// xlsxDimension directly maps the dimension element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxDimension struct {
	Ref string `xml:"ref,attr"`
}

// xlsxSheetData directly maps the sheetData element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheetData struct {
	XMLName xml.Name  `xml:"sheetData"`
	Row     []xlsxRow `xml:"row"`
}

// xlsxDataValidations  excel cell data validation
type xlsxDataValidations struct {
	DataValidation []*xlsxDataValidation `xml:"dataValidation"`
	Count          int                   `xml:"count,attr"`
}

// xlsxDataValidation
// A single item of data validation defined on a range of the worksheet.
// The list validation type would more commonly be called "a drop down box."
type xlsxDataValidation struct {
	// A boolean value indicating whether the data validation allows the use of empty or blank
	//entries. 1 means empty entries are OK and do not violate the validation constraints.
	AllowBlank bool `xml:"allowBlank,attr,omitempty"`
	// A boolean value indicating whether to display the input prompt message.
	ShowInputMessage bool `xml:"showInputMessage,attr,omitempty"`
	// A boolean value indicating whether to display the error alert message when an invalid
	// value has been entered, according to the criteria specified.
	ShowErrorMessage bool `xml:"showErrorMessage,attr,omitempty"`
	// The style of error alert used for this data validation.
	// warning, infomation, or stop
	// Stop will prevent the user from entering data that does not pass validation.
	ErrorStyle *string `xml:"errorStyle,attr"`
	// Title bar text of error alert.
	ErrorTitle *string `xml:"errorTitle,attr"`
	// The relational operator used with this data validation.
	// The possible values for this can be equal, notEqual, lessThan, etc.
	// This only applies to certain validation types.
	Operator string `xml:"operator,attr,omitempty"`
	// Message text of error alert.
	Error *string `xml:"error,attr"`
	// Title bar text of input prompt.
	PromptTitle *string `xml:"promptTitle,attr"`
	// Message text of input prompt.
	Prompt *string `xml:"prompt,attr"`
	// The type of data validation.
	// none, custom, date, decimal, list, textLength, time, whole
	Type string `xml:"type,attr"`
	// Range over which data validation is applied.
	// Cell or range, eg: A1 OR A1:A20
	Sqref string `xml:"sqref,attr,omitempty"`
	// The first formula in the Data Validation dropdown. It is used as a bounds for 'between' and
	// 'notBetween' relational operators, and the only formula used for other relational operators
	// (equal, notEqual, lessThan, lessThanOrEqual, greaterThan, greaterThanOrEqual), or for custom
	// or list type data validation. The content can be a formula or a constant or a list series (comma separated values).
	Formula1 string `xml:"formula1"`
	// The second formula in the DataValidation dropdown. It is used as a bounds for 'between' and
	// 'notBetween' relational operators only.
	Formula2 string `xml:"formula2,omitempty"`
}

// xlsxRow directly maps the row element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxRow struct {
	R            int     `xml:"r,attr"`
	Spans        string  `xml:"spans,attr,omitempty"`
	Hidden       bool    `xml:"hidden,attr,omitempty"`
	C            []xlsxC `xml:"c"`
	Ht           string  `xml:"ht,attr,omitempty"`
	CustomHeight bool    `xml:"customHeight,attr,omitempty"`
	OutlineLevel uint8   `xml:"outlineLevel,attr,omitempty"`
}

type xlsxAutoFilter struct {
	Ref string `xml:"ref,attr"`
}

type xlsxMergeCell struct {
	Ref string `xml:"ref,attr"` // ref: horiz "A1:C1", vert "B3:B6", both  "D3:G4"
}

type xlsxMergeCells struct {
	XMLName  xml.Name                 `xml:"mergeCells,omitempty"`
	Count    int                      `xml:"count,attr,omitempty"`
	Cells    []xlsxMergeCell          `xml:"mergeCell,omitempty"`
	CellsMap map[string]xlsxMergeCell `xml:"-"`
}

func (mc *xlsxMergeCells) addCell(cell xlsxMergeCell) {
	if mc.CellsMap == nil {
		mc.CellsMap = make(map[string]xlsxMergeCell)
	}
	cellRefs := strings.Split(cell.Ref, ":")
	mc.CellsMap[cellRefs[0]] = cell
}

type xlsxHyperlinks struct {
	HyperLinks []xlsxHyperlink `xml:"hyperlink"`
}

type xlsxHyperlink struct {
	RelationshipId string `xml:"id,attr"`
	Reference      string `xml:"ref,attr"`
	DisplayString  string `xml:"display,attr,omitempty"`
	Tooltip        string `xml:"tooltip,attr,omitempty"`
}

// Return the cartesian extent of a merged cell range from its origin
// cell (the closest merged cell to the to left of the sheet.
func (mc *xlsxMergeCells) getExtent(cellRef string) (int, int, error) {
	wrap := func(err error) (int, int, error) {
		return -1, -1, fmt.Errorf("getExtent: %w", err)
	}

	if mc == nil {
		return 0, 0, nil
	}
	if cell, ok := mc.CellsMap[cellRef]; ok {
		parts := strings.Split(cell.Ref, ":")
		startx, starty, err := GetCoordsFromCellIDString(parts[0])
		if err != nil {
			return wrap(err)
		}
		endx, endy, err := GetCoordsFromCellIDString(parts[1])
		if err != nil {
			return wrap(err)
		}
		return endx - startx, endy - starty, nil
	}
	return 0, 0, nil
}

// xlsxC directly maps the c element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxC struct {
	XMLName xml.Name
	R       string  `xml:"r,attr"`           // Cell ID, e.g. A1
	S       int     `xml:"s,attr,omitempty"` // Style reference.
	T       string  `xml:"t,attr,omitempty"` // Type.
	F       *xlsxF  `xml:"f,omitempty"`      // Formula
	V       string  `xml:"v,omitempty"`      // Value
	Is      *xlsxSI `xml:"is,omitempty"`     // Inline String.
}

// xlsxF directly maps the f element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxF struct {
	Content string `xml:",chardata"`
	T       string `xml:"t,attr,omitempty"`   // Formula type
	Ref     string `xml:"ref,attr,omitempty"` // Shared formula ref
	Si      int    `xml:"si,attr,omitempty"`  // Shared formula index
}

// Create a new XLSX Worksheet with default values populated.
// Strictly for internal use only!
func newXlsxWorksheet() (worksheet *xlsxWorksheet) {
	worksheet = &xlsxWorksheet{}
	worksheet.XMLNSR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	worksheet.SheetPr.FilterMode = false
	worksheet.SheetPr.PageSetUpPr = make([]xlsxPageSetUpPr, 1)
	worksheet.SheetPr.PageSetUpPr[0] = xlsxPageSetUpPr{FitToPage: false}
	worksheet.SheetViews.SheetView = make([]xlsxSheetView, 1)
	worksheet.SheetViews.SheetView[0] = xlsxSheetView{
		ColorId:                 64,
		DefaultGridColor:        true,
		RightToLeft:             false,
		Selection:               make([]xlsxSelection, 1),
		ShowFormulas:            false,
		ShowGridLines:           true,
		ShowOutlineSymbols:      true,
		ShowRowColHeaders:       true,
		ShowZeros:               true,
		TabSelected:             false,
		TopLeftCell:             "A1",
		View:                    "normal",
		WindowProtection:        false,
		WorkbookViewId:          0,
		ZoomScale:               100,
		ZoomScaleNormal:         100,
		ZoomScalePageLayoutView: 100}
	worksheet.SheetViews.SheetView[0].Selection[0] = xlsxSelection{
		Pane:         "topLeft",
		ActiveCell:   "A1",
		ActiveCellId: 0,
		SQRef:        "A1"}
	worksheet.SheetFormatPr.DefaultRowHeight = 12.85

	return
}

// setup the CellsMap so that we can rapidly calculate extents
func (worksheet *xlsxWorksheet) mapMergeCells() {

	if worksheet.MergeCells != nil {
		for _, cell := range worksheet.MergeCells.Cells {
			worksheet.MergeCells.addCell(cell)
		}
	}

}

func makeXMLAttr(fv reflect.Value, parentName, name string) (xmlwriter.Attr, error) {
	attr := xmlwriter.Attr{
		Name: name,
	}

	if fv.Kind() == reflect.Ptr {
		elm := fv.Elem()
		if elm.Kind() == reflect.Invalid {
			return attr, nil
		}
		return makeXMLAttr(elm, parentName, name)
	}

	switch fv.Kind() {
	case reflect.Bool:
		attr = attr.Bool(fv.Bool())
	case reflect.Int:
		attr = attr.Int(int(fv.Int()))
	case reflect.Int8:
		attr = attr.Int8(int8(fv.Int()))
	case reflect.Int16:
		attr = attr.Int16(int16(fv.Int()))
	case reflect.Int32:
		attr = attr.Int32(int32(fv.Int()))
	case reflect.Int64:
		attr = attr.Int64(fv.Int())
	case reflect.Uint:
		attr = attr.Uint(int(fv.Uint()))
	case reflect.Uint8:
		attr = attr.Uint8(uint8(fv.Uint()))
	case reflect.Uint16:
		attr = attr.Uint16(uint16(fv.Uint()))
	case reflect.Uint32:
		attr = attr.Uint32(uint32(fv.Uint()))
	case reflect.Uint64:
		attr = attr.Uint64(fv.Uint())
	case reflect.Float32:
		attr = attr.Float32(float32(fv.Float()))
	case reflect.Float64:
		attr = attr.Float64(fv.Float())
	case reflect.String:
		attr.Value = fv.String()
	default:
		return attr, fmt.Errorf("Not yet handled %s.%s (%s)", parentName, name, fv.Kind())

	}

	return attr, nil
}

func parseXMLTag(tag string) (string, string, bool, bool, bool) {
	var xmlNS string
	var name string
	var omitempty bool
	var isAttr bool
	var charData bool
	parts := strings.Split(tag, ",")
	partLen := len(parts)
	if partLen > 0 {
		nameParts := strings.Split(parts[0], " ")
		if len(nameParts) > 1 {
			xmlNS = nameParts[0]
			name = nameParts[1]
		} else {
			name = nameParts[0]
		}
	}
	if partLen > 1 {
		for _, p := range parts[1:] {
			omitempty = omitempty || p == "omitempty"
			isAttr = isAttr || p == "attr"
			charData = charData || p == "chardata"
		}
	}
	return xmlNS, name, omitempty, isAttr, charData
}

func emitStructAsXML(v reflect.Value, name, xmlNS string) (xmlwriter.Elem, error) {
	if v.Kind() == reflect.Ptr {
		return emitStructAsXML(v.Elem(), name, xmlNS)
	}
	output := xmlwriter.Elem{
		Name: name,
	}

	if xmlNS != "" {
		output.Attrs = append(output.Attrs, xmlwriter.Attr{
			Name:  "xmlns",
			Value: xmlNS,
		})
	}

	for i := 0; i < v.NumField(); i++ {
		var xmlNS string
		var name string
		var omitempty bool
		var isAttr bool
		var charData bool
		fv := v.Field(i)
		ft := v.Type().Field(i)
		tag := ft.Tag.Get("xml")
		if tag == "" {
			// This field is not intended for export!
			continue
		}

		xmlNS, name, omitempty, isAttr, charData = parseXMLTag(tag)
		if name == "-" {
			// This name means we shouldn't emit this element.
			continue
		}
		if isAttr {
			if omitempty && reflect.Zero(fv.Type()).Interface() == fv.Interface() {
				// The value is this types zero value
				continue
			}

			if output.Name == "hyperlink" && name == "id" {
				// Hack to respect the relationship namespace
				name = "r:id"
			}
			attr, err := makeXMLAttr(fv, output.Name, name)
			if err != nil {
				return output, err
			}
			output.Attrs = append(output.Attrs, attr)
			continue
		}
		if charData {
			output.Content = append(output.Content, xmlwriter.Text(fv.String()))
			continue
		}
		switch ft.Name {
		case "XMLName":
			output.Name = name
			output.Attrs = append(output.Attrs, xmlwriter.Attr{
				Name:  "xmlns",
				Value: xmlNS,
			})
		case "SheetData", "MergeCells":
			// Skip SheetData here, we explicitly generate this in writeXML below
			// Microsoft Excel considers a mergeCells element before a sheetData element to be
			// an error and will fail to open the document, so we'll be back with this data
			// from writeXml later (but we'll call it OutputMergeCells to make it past this case.)

			continue
		default:
			if fv.Kind() == reflect.Ptr {
				if fv.IsNil() {
					continue
				}
				fv = fv.Elem()
			}
			switch fv.Kind() {
			case reflect.Struct:
				elem, err := emitStructAsXML(fv, name, xmlNS)
				if err != nil {
					return output, err
				}
				output.Content = append(output.Content, elem)
			case reflect.Slice:
				for i := 0; i < fv.Len(); i++ {
					v := fv.Index(i)
					elem, err := emitStructAsXML(v, name, xmlNS)
					if err != nil {
						return output, err
					}
					output.Content = append(output.Content, elem)
				}
			case reflect.String:
				elem := xmlwriter.Elem{Name: name}
				if xmlNS != "" {
					elem.Attrs = append(elem.Attrs, xmlwriter.Attr{
						Name:  "xmlns",
						Value: xmlNS,
					})
				}
				elem.Content = append(elem.Content, xmlwriter.Text(fv.String()))
				output.Content = append(output.Content, elem)
			default:
				return output, fmt.Errorf("Todo with unhandled kind %s : %s", fv.Kind(), name)
			}
		}
	}
	return output, nil

}

func (worksheet *xlsxWorksheet) makeXlsxRowFromRow(row *Row, styles *xlsxStyleSheet, refTable *RefTable) (*xlsxRow, error) {
	xRow := &xlsxRow{}
	xRow.R = row.num + 1
	if row.isCustom {
		xRow.CustomHeight = true
		xRow.Ht = fmt.Sprintf("%g", row.GetHeight())
	}
	xRow.OutlineLevel = row.GetOutlineLevel()

	err := row.ForEachCell(func(cell *Cell) error {
		var XfId int

		col := row.Sheet.Col(cell.num)
		if col != nil {
			XfId = col.outXfID
		}

		// generate NumFmtId and add new NumFmt
		xNumFmt := styles.newNumFmt(cell.NumFmt)

		style := cell.style
		switch {
		case style != nil:
			XfId = handleStyleForXLSX(style, xNumFmt.NumFmtId, styles)
		case len(cell.NumFmt) == 0:
			// Do nothing
		case col == nil:
			XfId = handleNumFmtIdForXLSX(xNumFmt.NumFmtId, styles)
		case !compareFormatString(col.numFmt, cell.NumFmt):
			XfId = handleNumFmtIdForXLSX(xNumFmt.NumFmtId, styles)
		}
		xC := xlsxC{
			S: XfId,
			R: GetCellIDStringFromCoords(cell.num, row.num),
		}
		if cell.formula != "" {
			xC.F = &xlsxF{Content: cell.formula}
		}
		switch cell.cellType {
		case CellTypeInline:
			// Inline strings are turned into shared strings since they are more efficient.
			// This is what Excel does as well.
			fallthrough
		case CellTypeString:
			if len(cell.Value) > 0 {
				xC.V = strconv.Itoa(refTable.AddString(cell.Value))
			} else if len(cell.RichText) > 0 {
				xC.V = strconv.Itoa(refTable.AddRichText(cell.RichText))
			}
			xC.T = "s"
		case CellTypeNumeric:
			// Numeric is the default, so the type can be left blank
			xC.V = cell.Value
		case CellTypeBool:
			xC.V = cell.Value
			xC.T = "b"
		case CellTypeError:
			xC.V = cell.Value
			xC.T = "e"
		case CellTypeDate:
			xC.V = cell.Value
			xC.T = "d"
		case CellTypeStringFormula:
			xC.V = cell.Value
			xC.T = "str"
		default:
			return errors.New("unknown cell type cannot be marshaled")
		}
		xRow.C = append(xRow.C, xC)

		return nil
	})

	return xRow, err
}

func (worksheet *xlsxWorksheet) WriteXML(xw *xmlwriter.Writer, s *Sheet, styles *xlsxStyleSheet, refTable *RefTable) (err error) {
	var output xmlwriter.Elem
	worksheet.XMLNSR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	elem := reflect.ValueOf(worksheet)
	output, err = emitStructAsXML(elem, "", "")
	if err != nil {
		return
	}

	ec := xmlwriter.ErrCollector{}
	defer ec.Set(&err)
	ec.Do(
		xw.StartElem(output),
		xw.StartElem(xmlwriter.Elem{Name: "sheetData"}),
		s.ForEachRow(func(row *Row) error {
			xRow, err := worksheet.makeXlsxRowFromRow(row, styles, refTable)
			if err != nil {
				return err
			}
			elem := reflect.ValueOf(xRow)
			output, err := emitStructAsXML(elem, "row", "")
			if err != nil {
				return err
			}
			err = xw.Write(output)
			if err != nil {
				return err
			}
			return xw.Flush()

		}, SkipEmptyRows),
		xw.EndElem("sheetData"),
		func() error {
			if worksheet.MergeCells != nil {
				mergeCells, err := emitStructAsXML(reflect.ValueOf(worksheet.MergeCells), "OutputMergeCells", "")
				if err != nil {
					return err
				}
				return xw.Write(mergeCells)
			}
			return nil
		}(),
		xw.EndElem(output.Name),
		xw.Flush(),
	)
	return

}
