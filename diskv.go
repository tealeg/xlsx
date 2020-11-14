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

type DiskVRow struct {
	row         *Row
	maxCol      int
	store       *diskv.Diskv
	buf         bytes.Buffer
	currentCell *Cell
}

func makeDiskVRow(sheet *Sheet, store *diskv.Diskv) *DiskVRow {
	dvr := &DiskVRow{
		row:    new(Row),
		maxCol: -1,
		store:  store,
	}
	dvr.row.Sheet = sheet
	dvr.row.cellStoreRow = dvr
	sheet.setCurrentRow(dvr.row)
	return dvr
}

func (dvr *DiskVRow) CellUpdatable(c *Cell) {
	if c != dvr.currentCell {
		panic("Attempt to update Cell that isn't the current cell whilst using the DiskVCellStore.  You must use the Cell returned by the most recent operation.")

	}
}
func (dvr *DiskVRow) Updatable() {
	if dvr.row != dvr.row.Sheet.currentRow {
		panic("Attempt to update Row that isn't the current row whilst using the DiskVCellStore.  You must use the row returned by the most recent operation.")
	}
}

func (dvr *DiskVRow) AddCell() *Cell {
	cell := newCell(dvr.row, dvr.maxCol+1)
	dvr.setCurrentCell(cell)
	return cell
}

func (dvr *DiskVRow) readCell(key string) (*Cell, error) {
	var err error
	var cellType int
	var hasStyle, hasDataValidation bool
	var cellIsNil bool

	b, err := dvr.store.Read(key)
	if err != nil {
		return nil, err
	}

	buf := bytes.NewReader(b)
	if cellIsNil, err = readBool(buf); err != nil {
		return nil, err
	}
	if cellIsNil {
		if err = readEndOfRecord(buf); err != nil {
			return nil, err
		}
		return nil, nil
	}
	c := &Cell{}
	if c.Value, err = readString(buf); err != nil {
		return c, err
	}
	if c.formula, err = readString(buf); err != nil {
		return c, err
	}
	if hasStyle, err = readBool(buf); err != nil {
		return c, err
	}
	if c.NumFmt, err = readString(buf); err != nil {
		return c, err
	}
	if c.date1904, err = readBool(buf); err != nil {
		return c, err
	}
	if c.Hidden, err = readBool(buf); err != nil {
		return c, err
	}
	if c.HMerge, err = readInt(buf); err != nil {
		return c, err
	}
	if c.VMerge, err = readInt(buf); err != nil {
		return c, err
	}
	if cellType, err = readInt(buf); err != nil {
		return c, err
	}
	c.cellType = CellType(cellType)
	if hasDataValidation, err = readBool(buf); err != nil {
		return c, err
	}
	if c.Hyperlink.DisplayString, err = readString(buf); err != nil {
		return c, err
	}
	if c.Hyperlink.Link, err = readString(buf); err != nil {
		return c, err
	}
	if c.Hyperlink.Tooltip, err = readString(buf); err != nil {
		return c, err
	}
	if c.num, err = readInt(buf); err != nil {
		return c, err
	}
	if c.RichText, err = readRichText(buf); err != nil {
		return c, err
	}
	if err = readEndOfRecord(buf); err != nil {
		return c, err
	}
	if hasStyle {
		if c.style, err = readStyle(buf); err != nil {
			return c, err
		}
	}
	if hasDataValidation {
		if c.DataValidation, err = readDataValidation(buf); err != nil {
			return c, err
		}
	}
	return c, nil
}

func (dvr *DiskVRow) writeCell(c *Cell) error {
	var err error
	dvr.buf.Reset()
	if c == nil {
		if err := writeBool(&dvr.buf, true); err != nil {

			return err
		}
		return writeEndOfRecord(&dvr.buf)
	}
	if err := writeBool(&dvr.buf, false); err != nil {
		return err
	}
	if err = writeString(&dvr.buf, c.Value); err != nil {
		return err
	}
	if err = writeString(&dvr.buf, c.formula); err != nil {
		return err
	}
	if err = writeBool(&dvr.buf, c.style != nil); err != nil {
		return err
	}
	if err = writeString(&dvr.buf, c.NumFmt); err != nil {
		return err
	}
	if err = writeBool(&dvr.buf, c.date1904); err != nil {
		return err
	}
	if err = writeBool(&dvr.buf, c.Hidden); err != nil {
		return err
	}
	if err = writeInt(&dvr.buf, c.HMerge); err != nil {
		return err
	}
	if err = writeInt(&dvr.buf, c.VMerge); err != nil {
		return err
	}
	if err = writeInt(&dvr.buf, int(c.cellType)); err != nil {
		return err
	}
	if err = writeBool(&dvr.buf, c.DataValidation != nil); err != nil {
		return err
	}
	if err = writeString(&dvr.buf, c.Hyperlink.DisplayString); err != nil {
		return err
	}
	if err = writeString(&dvr.buf, c.Hyperlink.Link); err != nil {
		return err
	}
	if err = writeString(&dvr.buf, c.Hyperlink.Tooltip); err != nil {
		return err
	}
	if err = writeInt(&dvr.buf, c.num); err != nil {
		return err
	}
	if err = writeRichText(&dvr.buf, c.RichText); err != nil {
		return err
	}
	if err = writeEndOfRecord(&dvr.buf); err != nil {
		return err
	}
	if c.style != nil {
		if err = writeStyle(&dvr.buf, c.style); err != nil {
			return err
		}
	}
	if c.DataValidation != nil {
		if err = writeDataValidation(&dvr.buf, c.DataValidation); err != nil {
			return err
		}
	}
	key := dvr.row.makeCellKey(c.num)
	return dvr.store.Write(key, dvr.buf.Bytes())

}

func (dvr *DiskVRow) setCurrentCell(cell *Cell) {
	if dvr.currentCell.Modified() {
		err := dvr.writeCell(dvr.currentCell)
		if err != nil {
			panic(err.Error())
		}
	}
	if cell.num > dvr.maxCol {
		dvr.maxCol = cell.num
	}
	dvr.currentCell = cell

}

func (dvr *DiskVRow) PushCell(c *Cell) {
	c.modified = true
	dvr.setCurrentCell(c)
}

func (dvr *DiskVRow) GetCell(colIdx int) *Cell {
	if dvr.currentCell != nil {
		if dvr.currentCell.num == colIdx {
			return dvr.currentCell
		}
	}
	key := dvr.row.makeCellKey(colIdx)
	cell, err := dvr.readCell(key)
	if err == nil {
		dvr.setCurrentCell(cell)
		return cell
	}
	cell = newCell(dvr.row, colIdx)
	dvr.PushCell(cell)
	return cell
}

func (dvr *DiskVRow) ForEachCell(cvf CellVisitorFunc, option ...CellVisitorOption) error {
	flags := &cellVisitorFlags{}
	for _, opt := range option {
		opt(flags)
	}
	fn := func(ci int, c *Cell) error {
		if c == nil {
			if flags.skipEmptyCells {
				return nil
			}
			c = dvr.GetCell(ci)
		}
		if !c.Modified() && flags.skipEmptyCells {
			return nil
		}
		c.Row = dvr.row
		dvr.setCurrentCell(c)
		return cvf(c)
	}

	for ci := 0; ci <= dvr.maxCol; ci++ {
		var cell *Cell
		key := dvr.row.makeCellKey(ci)
		b, err := dvr.store.Read(key)
		if err != nil {
			// If the file doesn't exist that's fine, it was just an empty cell.
			if !os.IsNotExist(err) {
				return err
			}

		} else {
			cell, err = readCell(bytes.NewReader(b))
			if err != nil {
				return err
			}
		}
		
		err = fn(ci, cell)
		if err != nil {
			return err
		}
	}

	if !flags.skipEmptyCells {
		for ci := dvr.maxCol + 1; ci < dvr.row.Sheet.MaxCol; ci++ {
			c := dvr.GetCell(ci)
			err := cvf(c)
			if err != nil {
				return err
			}

		}
	}

	return nil
}

// MaxCol returns the index of the rightmost cell in the row's column.
func (dvr *DiskVRow) MaxCol() int {
	return dvr.maxCol
}

// CellCount returns the total number of cells in the row.
func (dvr *DiskVRow) CellCount() int {
	return dvr.maxCol + 1
}

// DiskVCellStore is an implementation of the CellStore interface, backed by DiskV
type DiskVCellStore struct {
	baseDir string
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
	return cs, nil
}

// ReadRow reads a row from the persistant store, identified by key,
// into memory and returns it, with the provided Sheet set as the Row's Sheet.
func (cs *DiskVCellStore) ReadRow(key string, s *Sheet) (*Row, error) {
	b, err := cs.store.Read(key)
	if err != nil {
		if _, ok := err.(*os.PathError); ok {
			return nil, NewRowNotFoundError(key, err.Error())
		}
		return nil, err
	}
	r, err := readRow(bytes.NewReader(b), cs.store, s)
	if err != nil {
		return nil, err
	}
	return r, nil
}

// MoveRow moves a Row from one position in a Sheet (index) to another
// within the persistant store.
func (cs *DiskVCellStore) MoveRow(r *Row, index int) error {

	cell := r.cellStoreRow.(*DiskVRow).currentCell
	if cell != nil {
		cs.buf.Reset()
		if err := writeCell(cs.buf, cell); err != nil {
			return err
		}
		key := r.makeCellKey(cell.num)
		if err := cs.store.WriteStream(key, cs.buf, true); err != nil {
			return err
		}
	}
	oldKey := r.key()
	r.num = index
	newKey := r.key()
	if cs.store.Has(newKey) {
		return fmt.Errorf("Target index for row (%d) would overwrite a row already exists", index)
	}
	err := cs.store.Erase(oldKey)
	if err != nil {
		return err
	}
	cs.buf.Reset()
	err = writeRow(cs.buf, r)
	if err != nil {
		return err
	}
	var cBuf bytes.Buffer
	keys := cs.store.KeysPrefix(oldKey, nil)
	for key := range keys {
		if key != oldKey {
			b, err := cs.store.Read(key)
			if err != nil {
				return err
			}
			c, err := readCell(bytes.NewReader(b))
			if err != nil {
				return err
			}
			c.Row = r
			err = writeCell(&cBuf, c)
			if err != nil {
				return err
			}
			newCKey := r.makeCellKey(c.num)
			if err := cs.store.Write(newCKey, cBuf.Bytes()); err != nil {
				return err
			}
			cs.store.Erase(key)

		}
	}

	err = r.ForEachCell(func(c *Cell) error {
		c.key()
		c.Row = r
		if err := writeCell(&cBuf, c); err != nil {
			return err
		}
		key := r.makeCellKey(c.num)
		cs.store.WriteStream(key, &cBuf, true)
		cBuf.Reset()
		return nil
	}, SkipEmptyCells)
	if err != nil {
		return err
	}
	return cs.store.WriteStream(newKey, cs.buf, true)
}

// RemoveRow removes a Row from the Sheet's representation in the
// persistant store.
func (cs *DiskVCellStore) RemoveRow(key string) error {
	keys := cs.store.KeysPrefix(key, nil)
	for key := range keys {
		err := cs.store.Erase(key)
		if err != nil {
			return err
		}
	}
	return nil
}

// MakeRow returns an empty Row
func (cs *DiskVCellStore) MakeRow(sheet *Sheet) *Row {
	return makeDiskVRow(sheet, cs.store).row
}

// MakeRowWithLen returns an empty Row, with a preconfigured starting length.
func (cs *DiskVCellStore) MakeRowWithLen(sheet *Sheet, len int) *Row {
	mr := makeDiskVRow(sheet, cs.store)
	mr.maxCol = len - 1
	return mr.row
}

// Close will remove the persisant storage for a given Sheet completely.
func (cs *DiskVCellStore) Close() error {
	return os.RemoveAll(cs.baseDir)

}

func writeBool(buf *bytes.Buffer, b bool) error {
	if b {
		err := buf.WriteByte(TRUE)
		if err != nil {
			return err
		}
	} else {
		err := buf.WriteByte(FALSE)
		if err != nil {
			return err
		}
	}
	return writeUnitSeparator(buf)
}

func readUnitSeparator(reader *bytes.Reader) error {
	us, err := reader.ReadByte()
	if err != nil {
		return err
	}
	if us != US {
		return errors.New("Invalid format in cellstore, no unit separator found")
	}
	return nil
}

//
func writeUnitSeparator(buf *bytes.Buffer) error {
	return buf.WriteByte(US)
}

func writeGroupSeparator(buf *bytes.Buffer) error {
	return buf.WriteByte(GS)
}

func readGroupSeparator(reader *bytes.Reader) error {
	gs, err := reader.ReadByte()
	if err != nil {
		return err
	}
	if gs != GS {
		return errors.New("Invalid format in cellstore, no group separator found")
	}
	return nil
}

func readBool(reader *bytes.Reader) (bool, error) {
	b, err := reader.ReadByte()
	if err != nil {
		return false, err
	}
	err = readUnitSeparator(reader)
	if err != nil {
		return false, err
	}
	if b == TRUE {
		return true, nil
	}
	return false, nil
}

func writeString(buf *bytes.Buffer, s string) error {
	_, err := buf.WriteString(s)
	if err != nil {
		return err
	}
	return writeUnitSeparator(buf)
}

func readString(reader *bytes.Reader) (string, error) {
	var s strings.Builder
	for {
		b, err := reader.ReadByte()
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

func writeInt(buf *bytes.Buffer, i int) error {
	ibuf := make([]byte, binary.MaxVarintLen64)

	n := binary.PutVarint(ibuf, int64(i))
	_, err := buf.Write(ibuf[:n])
	if err != nil {
		return err
	}
	return writeUnitSeparator(buf)
}

func readFloat(reader *bytes.Reader) (float64, error) {
	i, err := binary.ReadUvarint(reader)
	if err != nil {
		return -1, err
	}
	err = readUnitSeparator(reader)
	if err != nil {
		return -2, err
	}
	return math.Float64frombits(i), nil

}

func writeFloat(buf *bytes.Buffer, f float64) error {
	ibuf := make([]byte, binary.MaxVarintLen64)
	bits := math.Float64bits(f)
	n := binary.PutUvarint(ibuf, bits)
	_, err := buf.Write(ibuf[:n])
	if err != nil {
		return err
	}
	return writeUnitSeparator(buf)
}

func readInt(reader *bytes.Reader) (int, error) {
	i, err := binary.ReadVarint(reader)
	if err != nil {
		return -1, err
	}
	err = readUnitSeparator(reader)
	if err != nil {
		return -1, err
	}
	return int(i), nil
}

func writeStringPointer(buf *bytes.Buffer, sp *string) error {
	err := writeBool(buf, sp == nil)
	if err != nil {
		return err
	}
	if sp != nil {
		_, err = buf.WriteString(*sp)
		if err != nil {
			return err
		}
	}
	return writeUnitSeparator(buf)
}

func readStringPointer(reader *bytes.Reader) (*string, error) {
	isNil, err := readBool(reader)
	if err != nil {
		return nil, err
	}
	if isNil {
		err := readUnitSeparator(reader)
		return nil, err
	}
	s, err := readString(reader)
	return &s, err
}

func writeEndOfRecord(buf *bytes.Buffer) error {
	return buf.WriteByte(RS)
}

func readEndOfRecord(reader *bytes.Reader) error {
	b, err := reader.ReadByte()
	if err != nil {
		return err
	}
	if b != RS {
		return errors.New("Expected end of record, but not found")
	}
	return nil
}

func writeBorder(buf *bytes.Buffer, b Border) error {
	if err := writeString(buf, b.Left); err != nil {
		return err
	}
	if err := writeString(buf, b.LeftColor); err != nil {
		return err
	}
	if err := writeString(buf, b.Right); err != nil {
		return err
	}
	if err := writeString(buf, b.RightColor); err != nil {
		return err
	}
	if err := writeString(buf, b.Top); err != nil {
		return err
	}
	if err := writeString(buf, b.TopColor); err != nil {
		return err
	}
	if err := writeString(buf, b.Bottom); err != nil {
		return err
	}
	if err := writeString(buf, b.BottomColor); err != nil {
		return err
	}
	return nil
}

func readBorder(reader *bytes.Reader) (Border, error) {
	var err error
	b := Border{}
	if b.Left, err = readString(reader); err != nil {
		return b, err
	}
	if b.LeftColor, err = readString(reader); err != nil {
		return b, err
	}
	if b.Right, err = readString(reader); err != nil {
		return b, err
	}
	if b.RightColor, err = readString(reader); err != nil {
		return b, err
	}
	if b.Top, err = readString(reader); err != nil {
		return b, err
	}
	if b.TopColor, err = readString(reader); err != nil {
		return b, err
	}
	if b.Bottom, err = readString(reader); err != nil {
		return b, err
	}
	if b.BottomColor, err = readString(reader); err != nil {
		return b, err
	}
	return b, nil
}

func writeFill(buf *bytes.Buffer, f Fill) error {
	if err := writeString(buf, f.PatternType); err != nil {
		return err
	}
	if err := writeString(buf, f.BgColor); err != nil {
		return err
	}
	if err := writeString(buf, f.FgColor); err != nil {
		return err
	}
	return nil
}

func readFill(reader *bytes.Reader) (Fill, error) {
	var err error
	f := Fill{}
	if f.PatternType, err = readString(reader); err != nil {
		return f, err
	}
	if f.BgColor, err = readString(reader); err != nil {
		return f, err
	}
	if f.FgColor, err = readString(reader); err != nil {
		return f, err
	}
	return f, nil
}

func writeFont(buf *bytes.Buffer, f Font) error {
	if err := writeFloat(buf, f.Size); err != nil {
		return err
	}
	if err := writeString(buf, f.Name); err != nil {
		return err
	}
	if err := writeInt(buf, f.Family); err != nil {
		return err
	}
	if err := writeInt(buf, f.Charset); err != nil {
		return err
	}
	if err := writeString(buf, f.Color); err != nil {
		return err
	}
	if err := writeBool(buf, f.Bold); err != nil {
		return err
	}
	if err := writeBool(buf, f.Italic); err != nil {
		return err
	}
	if err := writeBool(buf, f.Underline); err != nil {
		return err
	}
	return nil
}

func readFont(reader *bytes.Reader) (Font, error) {
	var err error
	f := Font{}
	if f.Size, err = readFloat(reader); err != nil {
		return f, err
	}
	if f.Name, err = readString(reader); err != nil {
		return f, err
	}
	if f.Family, err = readInt(reader); err != nil {
		return f, err
	}
	if f.Charset, err = readInt(reader); err != nil {
		return f, err
	}
	if f.Color, err = readString(reader); err != nil {
		return f, err
	}
	if f.Bold, err = readBool(reader); err != nil {
		return f, err
	}
	if f.Italic, err = readBool(reader); err != nil {
		return f, err
	}
	if f.Underline, err = readBool(reader); err != nil {
		return f, err
	}
	return f, nil
}

//
func writeAlignment(buf *bytes.Buffer, a Alignment) error {
	var err error
	if err = writeString(buf, a.Horizontal); err != nil {
		return err
	}
	if err = writeInt(buf, a.Indent); err != nil {
		return err
	}
	if err = writeBool(buf, a.ShrinkToFit); err != nil {
		return err
	}
	if err = writeInt(buf, a.TextRotation); err != nil {
		return err
	}
	if err = writeString(buf, a.Vertical); err != nil {
		return err
	}
	if err = writeBool(buf, a.WrapText); err != nil {
		return err
	}
	return nil
}

func readAlignment(reader *bytes.Reader) (Alignment, error) {
	var err error
	a := Alignment{}
	if a.Horizontal, err = readString(reader); err != nil {
		return a, err
	}
	if a.Indent, err = readInt(reader); err != nil {
		return a, err
	}
	if a.ShrinkToFit, err = readBool(reader); err != nil {
		return a, err
	}
	if a.TextRotation, err = readInt(reader); err != nil {
		return a, err
	}
	if a.Vertical, err = readString(reader); err != nil {
		return a, err
	}
	if a.WrapText, err = readBool(reader); err != nil {
		return a, err
	}
	return a, nil
}

func writeStyle(buf *bytes.Buffer, s *Style) error {
	var err error
	if err = writeBorder(buf, s.Border); err != nil {
		return err
	}
	if err = writeFill(buf, s.Fill); err != nil {
		return err
	}
	if err = writeFont(buf, s.Font); err != nil {
		return err
	}
	if err = writeAlignment(buf, s.Alignment); err != nil {
		return err
	}
	if err = writeBool(buf, s.ApplyBorder); err != nil {
		return err
	}
	if err = writeBool(buf, s.ApplyFill); err != nil {
		return err
	}
	if err = writeBool(buf, s.ApplyFont); err != nil {
		return err
	}
	if err = writeBool(buf, s.ApplyAlignment); err != nil {
		return err
	}
	if err = writeEndOfRecord(buf); err != nil {
		return err
	}
	return nil
}

func readStyle(reader *bytes.Reader) (*Style, error) {
	var err error
	s := &Style{}
	if s.Border, err = readBorder(reader); err != nil {
		return s, err
	}
	if s.Fill, err = readFill(reader); err != nil {
		return s, err
	}
	if s.Font, err = readFont(reader); err != nil {
		return s, err
	}
	if s.Alignment, err = readAlignment(reader); err != nil {
		return s, err
	}
	if s.ApplyBorder, err = readBool(reader); err != nil {
		return s, err
	}
	if s.ApplyFill, err = readBool(reader); err != nil {
		return s, err
	}
	if s.ApplyFont, err = readBool(reader); err != nil {
		return s, err
	}
	if s.ApplyAlignment, err = readBool(reader); err != nil {
		return s, err
	}
	if err = readEndOfRecord(reader); err != nil {
		return s, err
	}
	return s, nil
}

func writeDataValidation(buf *bytes.Buffer, dv *xlsxDataValidation) error {
	var err error
	if err = writeBool(buf, dv.AllowBlank); err != nil {
		return err
	}
	if err = writeBool(buf, dv.ShowInputMessage); err != nil {
		return err
	}
	if err = writeBool(buf, dv.ShowErrorMessage); err != nil {
		return err
	}
	if err = writeStringPointer(buf, dv.ErrorStyle); err != nil {
		return err
	}
	if err = writeStringPointer(buf, dv.ErrorTitle); err != nil {
		return err
	}
	if err = writeString(buf, dv.Operator); err != nil {
		return err
	}
	if err = writeStringPointer(buf, dv.Error); err != nil {
		return err
	}
	if err = writeStringPointer(buf, dv.PromptTitle); err != nil {
		return err
	}
	if err = writeStringPointer(buf, dv.Prompt); err != nil {
		return err
	}
	if err = writeString(buf, dv.Type); err != nil {
		return err
	}
	if err = writeString(buf, dv.Sqref); err != nil {
		return err
	}
	if err = writeString(buf, dv.Formula1); err != nil {
		return err
	}
	if err = writeString(buf, dv.Formula2); err != nil {
		return err
	}
	if err = writeEndOfRecord(buf); err != nil {
		return err
	}
	return nil
}

func readDataValidation(reader *bytes.Reader) (*xlsxDataValidation, error) {
	var err error
	dv := &xlsxDataValidation{}
	if dv.AllowBlank, err = readBool(reader); err != nil {
		return dv, err
	}
	if dv.ShowInputMessage, err = readBool(reader); err != nil {
		return dv, err
	}
	if dv.ShowErrorMessage, err = readBool(reader); err != nil {
		return dv, err
	}
	if dv.ErrorStyle, err = readStringPointer(reader); err != nil {
		return dv, err
	}
	if dv.ErrorTitle, err = readStringPointer(reader); err != nil {
		return dv, err
	}
	if dv.Operator, err = readString(reader); err != nil {
		return dv, err
	}
	if dv.Error, err = readStringPointer(reader); err != nil {
		return dv, err
	}
	if dv.PromptTitle, err = readStringPointer(reader); err != nil {
		return dv, err
	}
	if dv.Prompt, err = readStringPointer(reader); err != nil {
		return dv, err
	}
	if dv.Type, err = readString(reader); err != nil {
		return dv, err
	}
	if dv.Sqref, err = readString(reader); err != nil {
		return dv, err
	}
	if dv.Formula1, err = readString(reader); err != nil {
		return dv, err
	}
	if dv.Formula2, err = readString(reader); err != nil {
		return dv, err
	}
	if err = readEndOfRecord(reader); err != nil {
		return dv, err
	}
	return dv, nil
}

func writeRow(buf *bytes.Buffer, r *Row) error {
	var err error
	if err = writeBool(buf, r.Hidden); err != nil {
		return err
	}
	// We don't write the Sheet reference, it's always restorable from context.
	if err = writeFloat(buf, r.GetHeight()); err != nil {
		return err
	}
	if err = writeInt(buf, int(r.GetOutlineLevel())); err != nil {
		return err
	}
	if err = writeBool(buf, r.isCustom); err != nil {
		return err
	}
	if err = writeInt(buf, r.num); err != nil {
		return err
	}
	if err = writeInt(buf, r.cellStoreRow.MaxCol()); err != nil {
		return err
	}
	if err = writeEndOfRecord(buf); err != nil {
		return err
	}
	return writeGroupSeparator(buf)
}

func writeCell(buf *bytes.Buffer, c *Cell) error {
	var err error
	if c == nil {
		if err := writeBool(buf, true); err != nil {

			return err
		}
		return writeEndOfRecord(buf)
	}
	if err := writeBool(buf, false); err != nil {
		return err
	}
	if err = writeString(buf, c.Value); err != nil {
		return err
	}
	if err = writeString(buf, c.formula); err != nil {
		return err
	}
	if err = writeBool(buf, c.style != nil); err != nil {
		return err
	}
	if err = writeString(buf, c.NumFmt); err != nil {
		return err
	}
	if err = writeBool(buf, c.date1904); err != nil {
		return err
	}
	if err = writeBool(buf, c.Hidden); err != nil {
		return err
	}
	if err = writeInt(buf, c.HMerge); err != nil {
		return err
	}
	if err = writeInt(buf, c.VMerge); err != nil {
		return err
	}
	if err = writeInt(buf, int(c.cellType)); err != nil {
		return err
	}
	if err = writeBool(buf, c.DataValidation != nil); err != nil {
		return err
	}
	if err = writeString(buf, c.Hyperlink.DisplayString); err != nil {
		return err
	}
	if err = writeString(buf, c.Hyperlink.Link); err != nil {
		return err
	}
	if err = writeString(buf, c.Hyperlink.Tooltip); err != nil {
		return err
	}
	if err = writeInt(buf, c.num); err != nil {
		return err
	}
	if err = writeRichText(buf, c.RichText); err != nil {
		return err
	}
	if err = writeEndOfRecord(buf); err != nil {
		return err
	}
	if c.style != nil {
		if err = writeStyle(buf, c.style); err != nil {
			return err
		}
	}
	if c.DataValidation != nil {
		if err = writeDataValidation(buf, c.DataValidation); err != nil {
			return err
		}
	}
	return nil
}

func readRow(reader *bytes.Reader, store *diskv.Diskv, sheet *Sheet) (*Row, error) {
	var err error

	r := &Row{
		Sheet: sheet,
	}
	dr := &DiskVRow{
		row:   r,
		store: store,
	}
	r.cellStoreRow = dr

	r.Hidden, err = readBool(reader)
	if err != nil {
		return nil, err
	}
	height, err := readFloat(reader)
	if err != nil {
		return nil, err
	}
	r.height = height
	outlineLevel, err := readInt(reader)
	if err != nil {
		return nil, err
	}
	r.outlineLevel = uint8(outlineLevel)
	r.isCustom, err = readBool(reader)
	if err != nil {
		return nil, err
	}
	r.num, err = readInt(reader)
	if err != nil {
		return nil, err
	}
	dr.maxCol, err = readInt(reader)
	if err != nil {
		return nil, err
	}
	err = readEndOfRecord(reader)
	if err != nil {
		return r, err
	}
	return r, nil
}

func readCell(reader *bytes.Reader) (*Cell, error) {
	var err error
	var cellType int
	var hasStyle, hasDataValidation bool
	var cellIsNil bool
	if cellIsNil, err = readBool(reader); err != nil {
		return nil, err
	}
	if cellIsNil {
		if err = readEndOfRecord(reader); err != nil {
			return nil, err
		}
		return nil, nil
	}
	c := &Cell{}
	if c.Value, err = readString(reader); err != nil {
		return c, err
	}
	if c.formula, err = readString(reader); err != nil {
		return c, err
	}
	if hasStyle, err = readBool(reader); err != nil {
		return c, err
	}
	if c.NumFmt, err = readString(reader); err != nil {
		return c, err
	}
	if c.date1904, err = readBool(reader); err != nil {
		return c, err
	}
	if c.Hidden, err = readBool(reader); err != nil {
		return c, err
	}
	if c.HMerge, err = readInt(reader); err != nil {
		return c, err
	}
	if c.VMerge, err = readInt(reader); err != nil {
		return c, err
	}
	if cellType, err = readInt(reader); err != nil {
		return c, err
	}
	c.cellType = CellType(cellType)
	if hasDataValidation, err = readBool(reader); err != nil {
		return c, err
	}
	if c.Hyperlink.DisplayString, err = readString(reader); err != nil {
		return c, err
	}
	if c.Hyperlink.Link, err = readString(reader); err != nil {
		return c, err
	}
	if c.Hyperlink.Tooltip, err = readString(reader); err != nil {
		return c, err
	}
	if c.num, err = readInt(reader); err != nil {
		return c, err
	}
	if c.RichText, err = readRichText(reader); err != nil {
		return c, err
	}
	if err = readEndOfRecord(reader); err != nil {
		return c, err
	}
	if hasStyle {
		if c.style, err = readStyle(reader); err != nil {
			return c, err
		}
	}
	if hasDataValidation {
		if c.DataValidation, err = readDataValidation(reader); err != nil {
			return c, err
		}
	}
	return c, nil
}

// WriteRow writes a Row to persistant storage.
func (cs *DiskVCellStore) WriteRow(r *Row) error {
	dvr, ok := r.cellStoreRow.(*DiskVRow)
	if !ok {
		return fmt.Errorf("cellStoreRow for a DiskVCellStore is not DiskVRow (%T)!", r.cellStoreRow)
	}
	if dvr.currentCell != nil {
		err := dvr.writeCell(dvr.currentCell)
		if err != nil {
			return err
		}
	}
	cs.buf.Reset()
	err := writeRow(cs.buf, r)
	if err != nil {
		return err
	}
	key := r.key()
	return cs.store.WriteStream(key, cs.buf, true)
}

func cellTransform(s string) []string {
	return strings.Split(s, ":")
}

func writeRichTextColor(buf *bytes.Buffer, c *RichTextColor) error {
	var err error
	var hasIndexed bool
	var hasTheme bool

	hasIndexed = c.coreColor.Indexed != nil
	hasTheme = c.coreColor.Theme != nil

	if err = writeString(buf, c.coreColor.RGB); err != nil {
		return err
	}
	if err = writeBool(buf, hasTheme); err != nil {
		return err
	}
	if err = writeFloat(buf, c.coreColor.Tint); err != nil {
		return err
	}
	if err = writeBool(buf, hasIndexed); err != nil {
		return err
	}
	if err = writeEndOfRecord(buf); err != nil {
		return err
	}

	if hasTheme {
		if err = writeInt(buf, *c.coreColor.Theme); err != nil {
			return err
		}
		if err = writeEndOfRecord(buf); err != nil {
			return err
		}
	}

	if hasIndexed {
		if err = writeInt(buf, *c.coreColor.Indexed); err != nil {
			return err
		}
		if err = writeEndOfRecord(buf); err != nil {
			return err
		}
	}

	return nil
}

func readRichTextColor(reader *bytes.Reader) (*RichTextColor, error) {
	var err error
	var hasIndexed bool
	var hasTheme bool

	c := &RichTextColor{}

	if c.coreColor.RGB, err = readString(reader); err != nil {
		return nil, err
	}
	if hasTheme, err = readBool(reader); err != nil {
		return nil, err
	}
	if c.coreColor.Tint, err = readFloat(reader); err != nil {
		return nil, err
	}
	if hasIndexed, err = readBool(reader); err != nil {
		return nil, err
	}
	if err = readEndOfRecord(reader); err != nil {
		return nil, err
	}

	if hasTheme {
		var theme int
		if theme, err = readInt(reader); err != nil {
			return nil, err
		}
		if err = readEndOfRecord(reader); err != nil {
			return nil, err
		}
		c.coreColor.Theme = &theme
	}

	if hasIndexed {
		var indexed int
		if indexed, err = readInt(reader); err != nil {
			return nil, err
		}
		if err = readEndOfRecord(reader); err != nil {
			return nil, err
		}
		c.coreColor.Indexed = &indexed
	}

	return c, nil
}

func writeRichTextFont(buf *bytes.Buffer, f *RichTextFont) error {
	var err error
	var hasColor bool

	hasColor = f.Color != nil

	if err = writeString(buf, f.Name); err != nil {
		return err
	}
	if err = writeFloat(buf, f.Size); err != nil {
		return err
	}
	if err = writeInt(buf, int(f.Family)); err != nil {
		return err
	}
	if err = writeInt(buf, int(f.Charset)); err != nil {
		return err
	}
	if err = writeBool(buf, hasColor); err != nil {
		return err
	}
	if err = writeBool(buf, f.Bold); err != nil {
		return err
	}
	if err = writeBool(buf, f.Italic); err != nil {
		return err
	}
	if err = writeBool(buf, f.Strike); err != nil {
		return err
	}
	if err = writeString(buf, string(f.VertAlign)); err != nil {
		return err
	}
	if err = writeString(buf, string(f.Underline)); err != nil {
		return err
	}
	if err = writeEndOfRecord(buf); err != nil {
		return err
	}

	if hasColor {
		if err = writeRichTextColor(buf, f.Color); err != nil {
			return err
		}
	}

	return nil
}

func readRichTextFont(reader *bytes.Reader) (*RichTextFont, error) {
	var err error
	var hasColor bool
	var family int
	var charset int
	var verAlign string
	var underline string

	f := &RichTextFont{}

	if f.Name, err = readString(reader); err != nil {
		return nil, err
	}
	if f.Size, err = readFloat(reader); err != nil {
		return nil, err
	}
	if family, err = readInt(reader); err != nil {
		return nil, err
	}
	f.Family = RichTextFontFamily(family)
	if charset, err = readInt(reader); err != nil {
		return nil, err
	}
	f.Charset = RichTextCharset(charset)
	if hasColor, err = readBool(reader); err != nil {
		return nil, err
	}
	if f.Bold, err = readBool(reader); err != nil {
		return nil, err
	}
	if f.Italic, err = readBool(reader); err != nil {
		return nil, err
	}
	if f.Strike, err = readBool(reader); err != nil {
		return nil, err
	}
	if verAlign, err = readString(reader); err != nil {
		return nil, err
	}
	f.VertAlign = RichTextVertAlign(verAlign)
	if underline, err = readString(reader); err != nil {
		return nil, err
	}
	f.Underline = RichTextUnderline(underline)
	if err = readEndOfRecord(reader); err != nil {
		return nil, err
	}

	if hasColor {
		if f.Color, err = readRichTextColor(reader); err != nil {
			return nil, err
		}
	}

	return f, nil
}

func writeRichTextRun(buf *bytes.Buffer, r *RichTextRun) error {
	var err error
	var hasFont bool

	hasFont = r.Font != nil

	if err = writeBool(buf, hasFont); err != nil {
		return err
	}
	if err = writeString(buf, r.Text); err != nil {
		return err
	}
	if err = writeEndOfRecord(buf); err != nil {
		return err
	}

	if hasFont {
		if err = writeRichTextFont(buf, r.Font); err != nil {
			return err
		}
	}

	return nil
}

func readRichTextRun(reader *bytes.Reader) (*RichTextRun, error) {
	var err error
	var hasFont bool

	r := &RichTextRun{}

	if hasFont, err = readBool(reader); err != nil {
		return nil, err
	}
	if r.Text, err = readString(reader); err != nil {
		return nil, err
	}
	if err = readEndOfRecord(reader); err != nil {
		return nil, err
	}

	if hasFont {
		if r.Font, err = readRichTextFont(reader); err != nil {
			return nil, err
		}
	}

	return r, nil
}

func writeRichText(buf *bytes.Buffer, rt []RichTextRun) error {
	var err error
	var length int

	length = len(rt)

	if err = writeInt(buf, length); err != nil {
		return err
	}

	for _, r := range rt {
		if err = writeRichTextRun(buf, &r); err != nil {
			return err
		}
	}

	return nil
}

func readRichText(reader *bytes.Reader) ([]RichTextRun, error) {
	var err error
	var length int

	if length, err = readInt(reader); err != nil {
		return nil, err
	}

	var rt []RichTextRun

	var i int
	for i = 0; i < length; i++ {
		var r *RichTextRun
		if r, err = readRichTextRun(reader); err != nil {
			return nil, err
		}
		rt = append(rt, *r)
	}

	return rt, nil
}
