package xlsx

import (
	. "gopkg.in/check.v1"
)

type SheetSuite struct{}

var _ = Suite(&SheetSuite{})

// Test we can add a Row to a Sheet
func (s *SheetSuite) TestAddRow(c *C) {
	var f *File
	f = NewFile()
	sheet := f.AddSheet("MySheet")
	row := sheet.AddRow()
	c.Assert(row, NotNil)
	c.Assert(len(sheet.Rows), Equals, 1)
}

func (s *SheetSuite) TestMakeXLSXSheetFromRows(c *C) {
	file := NewFile()
	sheet := file.AddSheet("Sheet1")
	row := sheet.AddRow()
	cell := row.AddCell()
	cell.Value = "A cell!"
	refTable := NewSharedStringRefTable()
	// refTable.AddStringsFromSheet(sheet)
	xSheet:= sheet.makeXLSXSheet(refTable)
	// xSST := refTable.makeXLSXSST()
	c.Assert(xSheet.Dimension.Ref, Equals, "A1:A1")
	c.Assert(xSheet.SheetData.Row, HasLen, 1)
	xRow := xSheet.SheetData.Row[0]
	c.Assert(xRow.R, Equals, 1)
	c.Assert(xRow.Spans, Equals, "")
	c.Assert(xRow.C, HasLen, 1)
	xC := xRow.C[0]
	c.Assert(xC.R, Equals, "A1")
	c.Assert(xC.S, Equals, 0)
	c.Assert(xC.T, Equals, "s") // Shared string type
	c.Assert(xC.V, Equals, "0") // reference to shared string
}


// 	expectedXLSXSheet := `<?xml version="1.0" encoding="UTF-8"?>
//   <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
//     <dimension ref="A1:A1"></dimension>
//     <sheetData>
//       <row r="1">
//         <c r="A1" t="s">
//           <v>0</v>
//         </c>
//       </row>
//     </sheetData>
//   </worksheet>`
// 	expectedXLSXSST := `<?xml version="1.0" encoding="UTF-8"?>
//   <sst uniqueCount="1">
//     <si>
//       <t>A cell!</t>
//     </si>
//   </sst>`
// 	c.Assert(string(xSheet), Equals, expectedXLSXSheet)
// 	c.Assert(string(xSST), Equals, expectedXLSXSST)
// }

