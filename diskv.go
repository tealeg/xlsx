package xlsx

import (
	"bytes"
	"encoding/binary"
	"errors"
	"fmt"
	"io/ioutil"
	"math"
	"os"
	"strings"

	"github.com/peterbourgon/diskv"
	"github.com/rogpeppe/fastuuid"
)

const (
	TRUE  = 0x01
	FALSE = 0x00
	US    = 0x1f // Unit Separator
	RS    = 0x1e // Record Separator
	GS    = 0x1d // Group Separator
)

var generator *fastuuid.Generator

func init() {
	generator = fastuuid.MustNewGenerator()
}

// DiskVCellStore is an implementation of the CellStore interface, backed by DiskV
type DiskVCellStore struct {
	baseDir string
	ibuf    []byte
	buf     *bytes.Buffer
	reader  *bytes.Reader
	store   *diskv.Diskv
}

// UseDiskVCellStore is a FileOption that makes all Sheet instances
// for a File use DiskV as their backing store.  You can use this
// option when handling very large Sheets that would otherwise riquire
// allocating vast amounts of memory.
func UseDiskVCellStore(f *File) {
	f.cellStoreConstructor = NewDiskVCellStore
}

// NewDiskVCellStore is a CellStoreConstructor than returns a
// CellStore in terms of DiskV.
func NewDiskVCellStore() (CellStore, error) {
	cs := &DiskVCellStore{
		buf: bytes.NewBuffer([]byte{}),
	}

	dir, err := ioutil.TempDir("", "cellstore"+generator.Hex128())
	if err != nil {
		return nil, err
	}
	cs.baseDir = dir
	cs.store = diskv.New(diskv.Options{
		BasePath:     dir,
		CacheSizeMax: 1024 * 1024, // 1MB for file. TODO make this configurable
	})
	cs.ibuf = make([]byte, binary.MaxVarintLen64)
	return cs, nil
}

// ReadRow reads a row from the persistant store, identified by key,
// into memory and returns it.
func (cs *DiskVCellStore) ReadRow(key string) (*Row, error) {
	b, err := cs.store.Read(key)
	if err != nil {
		if _, ok := err.(*os.PathError); ok {
			return nil, NewRowNotFoundError(key, err.Error())
		}
		return nil, err
	}
	cs.buf.Reset()
	_, err = cs.buf.Write(b)
	if err != nil {
		return nil, err
	}
	cs.reader = bytes.NewReader(cs.buf.Bytes())
	return cs.readRow()
}

// MoveRow moves a Row from one position in a Sheet (index) to another
// within the persistant store.
func (cs *DiskVCellStore) MoveRow(r *Row, index int) error {
	oldKey := r.key()
	r.num++
	newKey := r.key()
	if cs.store.Has(newKey) {
		return fmt.Errorf("Target index for row (%d) would overwrite a row already exists", index)
	}
	err := cs.store.Erase(oldKey)
	if err != nil {
		return err
	}
	cs.buf.Reset()
	err = cs.writeRow(r)
	if err != nil {
		return err
	}
	return cs.store.WriteStream(newKey, cs.buf, true)
}

// RemoveRow removes a Row from the Sheet's representation in the
// persistant store.
func (cs *DiskVCellStore) RemoveRow(key string) error {
	return cs.store.Erase(key)
}

// Close will remove the persisant storage for a given Sheet completely.
func (cs *DiskVCellStore) Close() error {
	return os.RemoveAll(cs.baseDir)

}

func (cs *DiskVCellStore) writeBool(b bool) error {
	if b {
		err := cs.buf.WriteByte(TRUE)
		if err != nil {
			return err
		}
	} else {
		err := cs.buf.WriteByte(FALSE)
		if err != nil {
			return err
		}
	}
	return cs.writeUnitSeparator()
}

//
func (cs *DiskVCellStore) writeUnitSeparator() error {
	return cs.buf.WriteByte(US)
}

//
func (cs *DiskVCellStore) readUnitSeparator() error {
	us, err := cs.reader.ReadByte()
	if err != nil {
		return err
	}
	if us != US {
		return errors.New("Invalid format in cellstore, not unit separator found")
	}
	return nil
}

func (cs *DiskVCellStore) writeGroupSeparator() error {
	return cs.buf.WriteByte(GS)
}

//
func (cs *DiskVCellStore) readGroupSeparator() error {
	gs, err := cs.reader.ReadByte()
	if err != nil {
		return err
	}
	if gs != GS {
		return errors.New("Invalid format in cellstore, not group separator found")
	}
	return nil
}

//
func (cs *DiskVCellStore) readBool() (bool, error) {
	b, err := cs.reader.ReadByte()
	if err != nil {
		return false, err
	}
	err = cs.readUnitSeparator()
	if err != nil {
		return false, err
	}
	if b == TRUE {
		return true, nil
	}
	return false, nil
}

//-
func (cs *DiskVCellStore) writeString(s string) error {
	_, err := cs.buf.WriteString(s)
	if err != nil {
		return err
	}
	return cs.writeUnitSeparator()
}

//
func (cs *DiskVCellStore) readString() (string, error) {
	var s strings.Builder
	for {
		b, err := cs.reader.ReadByte()
		if err != nil {
			return "", err
		}
		if b == US {
			return s.String(), nil
		}
		err = s.WriteByte(b)
		if err != nil {
			return s.String(), err
		}
	}
}

func (cs *DiskVCellStore) writeInt(i int) error {
	n := binary.PutVarint(cs.ibuf, int64(i))
	_, err := cs.buf.Write(cs.ibuf[:n])
	if err != nil {
		return err
	}
	return cs.writeUnitSeparator()
}

func (cs *DiskVCellStore) readFloat() (float64, error) {
	i, err := binary.ReadUvarint(cs.reader)
	if err != nil {
		return -1, err
	}
	err = cs.readUnitSeparator()
	if err != nil {
		return -2, err
	}
	return math.Float64frombits(i), nil

}

func (cs *DiskVCellStore) writeFloat(f float64) error {
	bits := math.Float64bits(f)
	n := binary.PutUvarint(cs.ibuf, bits)
	_, err := cs.buf.Write(cs.ibuf[:n])
	if err != nil {
		return err
	}
	return cs.writeUnitSeparator()
}

//
func (cs *DiskVCellStore) readInt() (int, error) {
	i, err := binary.ReadVarint(cs.reader)
	if err != nil {
		return -1, err
	}
	err = cs.readUnitSeparator()
	if err != nil {
		return -1, err
	}
	return int(i), nil
}

//
func (cs *DiskVCellStore) writeStringPointer(sp *string) error {
	err := cs.writeBool(sp == nil)
	if err != nil {
		return err
	}
	if sp != nil {
		_, err = cs.buf.WriteString(*sp)
		if err != nil {
			return err
		}
	}
	return cs.writeUnitSeparator()
}

//
func (cs *DiskVCellStore) readStringPointer() (*string, error) {
	isNil, err := cs.readBool()
	if err != nil {
		return nil, err
	}
	if isNil {
		err := cs.readUnitSeparator()
		return nil, err
	}
	s, err := cs.readString()
	return &s, err
}

//
func (cs *DiskVCellStore) writeEndOfRecord() error {
	return cs.buf.WriteByte(RS)
}

func (cs *DiskVCellStore) readEndOfRecord() error {
	b, err := cs.reader.ReadByte()
	if err != nil {
		return err
	}
	if b != RS {
		return errors.New("Expected end of record, but not found")
	}
	return nil
}

func (cs *DiskVCellStore) writeBorder(b Border) error {
	if err := cs.writeString(b.Left); err != nil {
		return err
	}
	if err := cs.writeString(b.LeftColor); err != nil {
		return err
	}
	if err := cs.writeString(b.Right); err != nil {
		return err
	}
	if err := cs.writeString(b.RightColor); err != nil {
		return err
	}
	if err := cs.writeString(b.Top); err != nil {
		return err
	}
	if err := cs.writeString(b.TopColor); err != nil {
		return err
	}
	if err := cs.writeString(b.Bottom); err != nil {
		return err
	}
	if err := cs.writeString(b.BottomColor); err != nil {
		return err
	}
	return nil
}

//
func (cs *DiskVCellStore) readBorder() (Border, error) {
	var err error
	b := Border{}
	if b.Left, err = cs.readString(); err != nil {
		return b, err
	}
	if b.LeftColor, err = cs.readString(); err != nil {
		return b, err
	}
	if b.Right, err = cs.readString(); err != nil {
		return b, err
	}
	if b.RightColor, err = cs.readString(); err != nil {
		return b, err
	}
	if b.Top, err = cs.readString(); err != nil {
		return b, err
	}
	if b.TopColor, err = cs.readString(); err != nil {
		return b, err
	}
	if b.Bottom, err = cs.readString(); err != nil {
		return b, err
	}
	if b.BottomColor, err = cs.readString(); err != nil {
		return b, err
	}
	return b, nil
}

func (cs *DiskVCellStore) writeFill(f Fill) error {
	if err := cs.writeString(f.PatternType); err != nil {
		return err
	}
	if err := cs.writeString(f.BgColor); err != nil {
		return err
	}
	if err := cs.writeString(f.FgColor); err != nil {
		return err
	}
	return nil
}

func (cs *DiskVCellStore) readFill() (Fill, error) {
	var err error
	f := Fill{}
	if f.PatternType, err = cs.readString(); err != nil {
		return f, err
	}
	if f.BgColor, err = cs.readString(); err != nil {
		return f, err
	}
	if f.FgColor, err = cs.readString(); err != nil {
		return f, err
	}
	return f, nil
}

func (cs *DiskVCellStore) writeFont(f Font) error {
	if err := cs.writeFloat(f.Size); err != nil {
		return err
	}
	if err := cs.writeString(f.Name); err != nil {
		return err
	}
	if err := cs.writeInt(f.Family); err != nil {
		return err
	}
	if err := cs.writeInt(f.Charset); err != nil {
		return err
	}
	if err := cs.writeString(f.Color); err != nil {
		return err
	}
	if err := cs.writeBool(f.Bold); err != nil {
		return err
	}
	if err := cs.writeBool(f.Italic); err != nil {
		return err
	}
	if err := cs.writeBool(f.Underline); err != nil {
		return err
	}
	return nil
}

func (cs *DiskVCellStore) readFont() (Font, error) {
	var err error
	f := Font{}
	if f.Size, err = cs.readFloat(); err != nil {
		return f, err
	}
	if f.Name, err = cs.readString(); err != nil {
		return f, err
	}
	if f.Family, err = cs.readInt(); err != nil {
		return f, err
	}
	if f.Charset, err = cs.readInt(); err != nil {
		return f, err
	}
	if f.Color, err = cs.readString(); err != nil {
		return f, err
	}
	if f.Bold, err = cs.readBool(); err != nil {
		return f, err
	}
	if f.Italic, err = cs.readBool(); err != nil {
		return f, err
	}
	if f.Underline, err = cs.readBool(); err != nil {
		return f, err
	}
	return f, nil
}

//
func (cs *DiskVCellStore) writeAlignment(a Alignment) error {
	var err error
	if err = cs.writeString(a.Horizontal); err != nil {
		return err
	}
	if err = cs.writeInt(a.Indent); err != nil {
		return err
	}
	if err = cs.writeBool(a.ShrinkToFit); err != nil {
		return err
	}
	if err = cs.writeInt(a.TextRotation); err != nil {
		return err
	}
	if err = cs.writeString(a.Vertical); err != nil {
		return err
	}
	if err = cs.writeBool(a.WrapText); err != nil {
		return err
	}
	return nil
}

func (cs *DiskVCellStore) readAlignment() (Alignment, error) {
	var err error
	a := Alignment{}
	if a.Horizontal, err = cs.readString(); err != nil {
		return a, err
	}
	if a.Indent, err = cs.readInt(); err != nil {
		return a, err
	}
	if a.ShrinkToFit, err = cs.readBool(); err != nil {
		return a, err
	}
	if a.TextRotation, err = cs.readInt(); err != nil {
		return a, err
	}
	if a.Vertical, err = cs.readString(); err != nil {
		return a, err
	}
	if a.WrapText, err = cs.readBool(); err != nil {
		return a, err
	}
	return a, nil
}

func (cs *DiskVCellStore) writeStyle(s *Style) error {
	var err error
	if err = cs.writeBorder(s.Border); err != nil {
		return err
	}
	if err = cs.writeFill(s.Fill); err != nil {
		return err
	}
	if err = cs.writeFont(s.Font); err != nil {
		return err
	}
	if err = cs.writeAlignment(s.Alignment); err != nil {
		return err
	}
	if err = cs.writeBool(s.ApplyBorder); err != nil {
		return err
	}
	if err = cs.writeBool(s.ApplyFill); err != nil {
		return err
	}
	if err = cs.writeBool(s.ApplyFont); err != nil {
		return err
	}
	if err = cs.writeBool(s.ApplyAlignment); err != nil {
		return err
	}
	if err = cs.writeEndOfRecord(); err != nil {
		return err
	}
	return nil
}

func (cs *DiskVCellStore) readStyle() (*Style, error) {
	var err error
	s := &Style{}
	if s.Border, err = cs.readBorder(); err != nil {
		return s, err
	}
	if s.Fill, err = cs.readFill(); err != nil {
		return s, err
	}
	if s.Font, err = cs.readFont(); err != nil {
		return s, err
	}
	if s.Alignment, err = cs.readAlignment(); err != nil {
		return s, err
	}
	if s.ApplyBorder, err = cs.readBool(); err != nil {
		return s, err
	}
	if s.ApplyFill, err = cs.readBool(); err != nil {
		return s, err
	}
	if s.ApplyFont, err = cs.readBool(); err != nil {
		return s, err
	}
	if s.ApplyAlignment, err = cs.readBool(); err != nil {
		return s, err
	}
	if err = cs.readEndOfRecord(); err != nil {
		return s, err
	}
	return s, nil
}

func (cs *DiskVCellStore) writeDataValidation(dv *xlsxDataValidation) error {
	var err error
	if err = cs.writeBool(dv.AllowBlank); err != nil {
		return err
	}
	if err = cs.writeBool(dv.ShowInputMessage); err != nil {
		return err
	}
	if err = cs.writeBool(dv.ShowErrorMessage); err != nil {
		return err
	}
	if err = cs.writeStringPointer(dv.ErrorStyle); err != nil {
		return err
	}
	if err = cs.writeStringPointer(dv.ErrorTitle); err != nil {
		return err
	}
	if err = cs.writeString(dv.Operator); err != nil {
		return err
	}
	if err = cs.writeStringPointer(dv.Error); err != nil {
		return err
	}
	if err = cs.writeStringPointer(dv.PromptTitle); err != nil {
		return err
	}
	if err = cs.writeStringPointer(dv.Prompt); err != nil {
		return err
	}
	if err = cs.writeString(dv.Type); err != nil {
		return err
	}
	if err = cs.writeString(dv.Sqref); err != nil {
		return err
	}
	if err = cs.writeString(dv.Formula1); err != nil {
		return err
	}
	if err = cs.writeString(dv.Formula2); err != nil {
		return err
	}
	if err = cs.writeEndOfRecord(); err != nil {
		return err
	}
	return nil
}

func (cs *DiskVCellStore) readDataValidation() (*xlsxDataValidation, error) {
	var err error
	dv := &xlsxDataValidation{}
	if dv.AllowBlank, err = cs.readBool(); err != nil {
		return dv, err
	}
	if dv.ShowInputMessage, err = cs.readBool(); err != nil {
		return dv, err
	}
	if dv.ShowErrorMessage, err = cs.readBool(); err != nil {
		return dv, err
	}
	if dv.ErrorStyle, err = cs.readStringPointer(); err != nil {
		return dv, err
	}
	if dv.ErrorTitle, err = cs.readStringPointer(); err != nil {
		return dv, err
	}
	if dv.Operator, err = cs.readString(); err != nil {
		return dv, err
	}
	if dv.Error, err = cs.readStringPointer(); err != nil {
		return dv, err
	}
	if dv.PromptTitle, err = cs.readStringPointer(); err != nil {
		return dv, err
	}
	if dv.Prompt, err = cs.readStringPointer(); err != nil {
		return dv, err
	}
	if dv.Type, err = cs.readString(); err != nil {
		return dv, err
	}
	if dv.Sqref, err = cs.readString(); err != nil {
		return dv, err
	}
	if dv.Formula1, err = cs.readString(); err != nil {
		return dv, err
	}
	if dv.Formula2, err = cs.readString(); err != nil {
		return dv, err
	}
	if err = cs.readEndOfRecord(); err != nil {
		return dv, err
	}
	return dv, nil
}

func (cs *DiskVCellStore) writeRow(r *Row) error {
	var err error
	if err = cs.writeBool(r.Hidden); err != nil {
		return err
	}
	// We don't write the Sheet reference, it's always restorable from context.
	if err = cs.writeFloat(r.GetHeight()); err != nil {
		return err
	}
	if err = cs.writeInt(int(r.GetOutlineLevel())); err != nil {
		return err
	}
	if err = cs.writeBool(r.isCustom); err != nil {
		return err
	}
	if err = cs.writeInt(r.num); err != nil {
		return err
	}
	if err = cs.writeInt(r.cellCount); err != nil {
		return err
	}
	if err = cs.writeEndOfRecord(); err != nil {
		return err
	}
	for _, cell := range r.cells {
		err = cs.writeCell(cell)
		if err != nil {
			return err
		}
	}
	return cs.writeGroupSeparator()
}

func (cs *DiskVCellStore) writeCell(c *Cell) error {
	var err error
	if c == nil {
		if err := cs.writeBool(true); err != nil {

			return err
		}
		return cs.writeEndOfRecord()
	}
	if err := cs.writeBool(false); err != nil {
		return err
	}
	if err = cs.writeString(c.Value); err != nil {
		return err
	}
	if err = cs.writeString(c.formula); err != nil {
		return err
	}
	if err = cs.writeBool(c.style != nil); err != nil {
		return err
	}
	if err = cs.writeString(c.NumFmt); err != nil {
		return err
	}
	if err = cs.writeBool(c.date1904); err != nil {
		return err
	}
	if err = cs.writeBool(c.Hidden); err != nil {
		return err
	}
	if err = cs.writeInt(c.HMerge); err != nil {
		return err
	}
	if err = cs.writeInt(c.VMerge); err != nil {
		return err
	}
	if err = cs.writeInt(int(c.cellType)); err != nil {
		return err
	}
	if err = cs.writeBool(c.DataValidation != nil); err != nil {
		return err
	}
	if err = cs.writeString(c.Hyperlink.DisplayString); err != nil {
		return err
	}
	if err = cs.writeString(c.Hyperlink.Link); err != nil {
		return err
	}
	if err = cs.writeString(c.Hyperlink.Tooltip); err != nil {
		return err
	}
	if err = cs.writeInt(c.num); err != nil {
		return err
	}
	if err = cs.writeRichText(c.RichText); err != nil {
		return err
	}
	if err = cs.writeEndOfRecord(); err != nil {
		return err
	}
	if c.style != nil {
		if err = cs.writeStyle(c.style); err != nil {
			return err
		}
	}
	if c.DataValidation != nil {
		if err = cs.writeDataValidation(c.DataValidation); err != nil {
			return err
		}
	}
	return nil
}

func (cs *DiskVCellStore) readRow() (*Row, error) {
	var err error
	r := &Row{}

	r.Hidden, err = cs.readBool()
	if err != nil {
		return nil, err
	}
	height, err := cs.readFloat()
	if err != nil {
		return nil, err
	}
	r.SetHeight(height)
	outlineLevel, err := cs.readInt()
	if err != nil {
		return nil, err
	}
	r.SetOutlineLevel(uint8(outlineLevel))

	r.isCustom, err = cs.readBool()
	if err != nil {
		return nil, err
	}
	r.num, err = cs.readInt()
	if err != nil {
		return nil, err
	}
	r.cellCount, err = cs.readInt()
	if err != nil {
		return nil, err
	}
	err = cs.readEndOfRecord()
	if err != nil {
		return r, err
	}
	for {
		if err := cs.readGroupSeparator(); err == nil || err.Error() == "EOF" {
			break
		}
		cs.reader.UnreadByte()
		cell, err := cs.readCell()
		if err != nil {
			return r, err
		}
		r.cells = append(r.cells, cell)
	}
	return r, nil
}

func (cs *DiskVCellStore) readCell() (*Cell, error) {
	var err error
	var cellType int
	var hasStyle, hasDataValidation bool
	var cellIsNil bool
	if cellIsNil, err = cs.readBool(); err != nil {
		return nil, err
	}
	if cellIsNil {
		if err = cs.readEndOfRecord(); err != nil {
			return nil, err
		}
		return nil, nil
	}
	c := &Cell{}
	if c.Value, err = cs.readString(); err != nil {
		return c, err
	}
	if c.formula, err = cs.readString(); err != nil {
		return c, err
	}
	if hasStyle, err = cs.readBool(); err != nil {
		return c, err
	}
	if c.NumFmt, err = cs.readString(); err != nil {
		return c, err
	}
	if c.date1904, err = cs.readBool(); err != nil {
		return c, err
	}
	if c.Hidden, err = cs.readBool(); err != nil {
		return c, err
	}
	if c.HMerge, err = cs.readInt(); err != nil {
		return c, err
	}
	if c.VMerge, err = cs.readInt(); err != nil {
		return c, err
	}
	if cellType, err = cs.readInt(); err != nil {
		return c, err
	}
	c.cellType = CellType(cellType)
	if hasDataValidation, err = cs.readBool(); err != nil {
		return c, err
	}
	if c.Hyperlink.DisplayString, err = cs.readString(); err != nil {
		return c, err
	}
	if c.Hyperlink.Link, err = cs.readString(); err != nil {
		return c, err
	}
	if c.Hyperlink.Tooltip, err = cs.readString(); err != nil {
		return c, err
	}
	if c.num, err = cs.readInt(); err != nil {
		return c, err
	}
	if c.RichText, err = cs.readRichText(); err != nil {
		return c, err
	}
	if err = cs.readEndOfRecord(); err != nil {
		return c, err
	}
	if hasStyle {
		if c.style, err = cs.readStyle(); err != nil {
			return c, err
		}
	}
	if hasDataValidation {
		if c.DataValidation, err = cs.readDataValidation(); err != nil {
			return c, err
		}
	}
	return c, nil
}

// WriteRow writes a Row to persistant storage.
func (cs *DiskVCellStore) WriteRow(r *Row) error {
	cs.buf.Reset()
	err := cs.writeRow(r)
	if err != nil {
		return err
	}
	key := r.key()
	return cs.store.WriteStream(key, cs.buf, true)
}

func cellTransform(s string) []string {
	return strings.Split(s, ":")
}

func (cs *DiskVCellStore) writeRichTextColor(c *RichTextColor) error {
	var err error
	var hasIndexed bool
	var hasTheme bool

	hasIndexed = c.coreColor.Indexed != nil
	hasTheme = c.coreColor.Theme != nil

	if err = cs.writeString(c.coreColor.RGB); err != nil {
		return err
	}
	if err = cs.writeBool(hasTheme); err != nil {
		return err
	}
	if err = cs.writeFloat(c.coreColor.Tint); err != nil {
		return err
	}
	if err = cs.writeBool(hasIndexed); err != nil {
		return err
	}
	if err = cs.writeEndOfRecord(); err != nil {
		return err
	}

	if hasTheme {
		if err = cs.writeInt(*c.coreColor.Theme); err != nil {
			return err
		}
		if err = cs.writeEndOfRecord(); err != nil {
			return err
		}
	}

	if hasIndexed {
		if err = cs.writeInt(*c.coreColor.Indexed); err != nil {
			return err
		}
		if err = cs.writeEndOfRecord(); err != nil {
			return err
		}
	}

	return nil
}

func (cs *DiskVCellStore) readRichTextColor() (*RichTextColor, error) {
	var err error
	var hasIndexed bool
	var hasTheme bool

	c := &RichTextColor{}

	if c.coreColor.RGB, err = cs.readString(); err != nil {
		return nil, err
	}
	if hasTheme, err = cs.readBool(); err != nil {
		return nil, err
	}
	if c.coreColor.Tint, err = cs.readFloat(); err != nil {
		return nil, err
	}
	if hasIndexed, err = cs.readBool(); err != nil {
		return nil, err
	}
	if err = cs.readEndOfRecord(); err != nil {
		return nil, err
	}

	if hasTheme {
		var theme int
		if theme, err = cs.readInt(); err != nil {
			return nil, err
		}
		if err = cs.readEndOfRecord(); err != nil {
			return nil, err
		}
		c.coreColor.Theme = &theme
	}

	if hasIndexed {
		var indexed int
		if indexed, err = cs.readInt(); err != nil {
			return nil, err
		}
		if err = cs.readEndOfRecord(); err != nil {
			return nil, err
		}
		c.coreColor.Indexed = &indexed
	}

	return c, nil
}

func (cs *DiskVCellStore) writeRichTextFont(f *RichTextFont) error {
	var err error
	var hasColor bool

	hasColor = f.Color != nil

	if err = cs.writeString(f.Name); err != nil {
		return err
	}
	if err = cs.writeFloat(f.Size); err != nil {
		return err
	}
	if err = cs.writeInt(int(f.Family)); err != nil {
		return err
	}
	if err = cs.writeInt(int(f.Charset)); err != nil {
		return err
	}
	if err = cs.writeBool(hasColor); err != nil {
		return err
	}
	if err = cs.writeBool(f.Bold); err != nil {
		return err
	}
	if err = cs.writeBool(f.Italic); err != nil {
		return err
	}
	if err = cs.writeBool(f.Strike); err != nil {
		return err
	}
	if err = cs.writeString(string(f.VertAlign)); err != nil {
		return err
	}
	if err = cs.writeString(string(f.Underline)); err != nil {
		return err
	}
	if err = cs.writeEndOfRecord(); err != nil {
		return err
	}

	if hasColor {
		if err = cs.writeRichTextColor(f.Color); err != nil {
			return err
		}
	}

	return nil
}

func (cs *DiskVCellStore) readRichTextFont() (*RichTextFont, error) {
	var err error
	var hasColor bool
	var family int
	var charset int
	var verAlign string
	var underline string

	f := &RichTextFont{}

	if f.Name, err = cs.readString(); err != nil {
		return nil, err
	}
	if f.Size, err = cs.readFloat(); err != nil {
		return nil, err
	}
	if family, err = cs.readInt(); err != nil {
		return nil, err
	}
	f.Family = RichTextFontFamily(family)
	if charset, err = cs.readInt(); err != nil {
		return nil, err
	}
	f.Charset = RichTextCharset(charset)
	if hasColor, err = cs.readBool(); err != nil {
		return nil, err
	}
	if f.Bold, err = cs.readBool(); err != nil {
		return nil, err
	}
	if f.Italic, err = cs.readBool(); err != nil {
		return nil, err
	}
	if f.Strike, err = cs.readBool(); err != nil {
		return nil, err
	}
	if verAlign, err = cs.readString(); err != nil {
		return nil, err
	}
	f.VertAlign = RichTextVertAlign(verAlign)
	if underline, err = cs.readString(); err != nil {
		return nil, err
	}
	f.Underline = RichTextUnderline(underline)
	if err = cs.readEndOfRecord(); err != nil {
		return nil, err
	}

	if hasColor {
		if f.Color, err = cs.readRichTextColor(); err != nil {
			return nil, err
		}
	}

	return f, nil
}

func (cs *DiskVCellStore) writeRichTextRun(r *RichTextRun) error {
	var err error
	var hasFont bool

	hasFont = r.Font != nil

	if err = cs.writeBool(hasFont); err != nil {
		return err
	}
	if err = cs.writeString(r.Text); err != nil {
		return err
	}
	if err = cs.writeEndOfRecord(); err != nil {
		return err
	}

	if hasFont {
		if err = cs.writeRichTextFont(r.Font); err != nil {
			return err
		}
	}

	return nil
}

func (cs *DiskVCellStore) readRichTextRun() (*RichTextRun, error) {
	var err error
	var hasFont bool

	r := &RichTextRun{}

	if hasFont, err = cs.readBool(); err != nil {
		return nil, err
	}
	if r.Text, err = cs.readString(); err != nil {
		return nil, err
	}
	if err = cs.readEndOfRecord(); err != nil {
		return nil, err
	}

	if hasFont {
		if r.Font, err = cs.readRichTextFont(); err != nil {
			return nil, err
		}
	}

	return r, nil
}

func (cs *DiskVCellStore) writeRichText(rt []RichTextRun) error {
	var err error
	var length int

	length = len(rt)

	if err = cs.writeInt(length); err != nil {
		return err
	}

	for _, r := range rt {
		if err = cs.writeRichTextRun(&r); err != nil {
			return err
		}
	}

	return nil
}

func (cs *DiskVCellStore) readRichText() ([]RichTextRun, error) {
	var err error
	var length int

	if length, err = cs.readInt(); err != nil {
		return nil, err
	}

	var rt []RichTextRun

	var i int
	for i = 0; i < length; i++ {
		var r *RichTextRun
		if r, err = cs.readRichTextRun(); err != nil {
			return nil, err
		}
		rt = append(rt, *r)
	}

	return rt, nil
}
