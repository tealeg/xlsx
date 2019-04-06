package xlsx

import (
	"bytes"
	"fmt"
	"io"
	"reflect"
	"strings"

	. "gopkg.in/check.v1"
)

const (
	TestsShouldMakeRealFiles = true
)

type StreamSuite struct{}

var _ = Suite(&StreamSuite{})

func (s *StreamSuite) TestTestsShouldMakeRealFilesShouldBeFalse(t *C) {
	if TestsShouldMakeRealFiles {
		t.Fatal("TestsShouldMakeRealFiles should only be true for local debugging. Don't forget to switch back before commiting.")
	}
}

func (s *StreamSuite) TestXlsxStreamWrite(t *C) {
	// When shouldMakeRealFiles is set to true this test will make actual XLSX files in the file system.
	// This is useful to ensure files open in Excel, Numbers, Google Docs, etc.
	// In case of issues you can use "Open XML SDK 2.5" to diagnose issues in generated XLSX files:
	// https://www.microsoft.com/en-us/download/details.aspx?id=30425
	testCases := []struct {
		testName      string
		sheetNames    []string
		workbookData  [][][]StreamCell
		expectedError error
	}{
		{
			testName: "Number Row",
			sheetNames: []string{
				"Sheet1",
			},
			workbookData: [][][]StreamCell{
				{
					{MakeStringStreamCell("1"), MakeStringStreamCell("25"),
						MakeStringStreamCell("A"), MakeStringStreamCell("B")},
					{MakeIntegerStreamCell(1234), MakeStyledIntegerStreamCell(98, BoldIntegers),
						MakeStyledIntegerStreamCell(34, ItalicIntegers), MakeStyledIntegerStreamCell(26, UnderlinedIntegers)},
				},
			},
		},
		{
			testName: "One Sheet",
			sheetNames: []string{
				"Sheet1",
			},
			workbookData: [][][]StreamCell{
				{
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU")},
					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123)},
				},
			},
		},
		{
			testName: "One Column",
			sheetNames: []string{
				"Sheet1",
			},
			workbookData: [][][]StreamCell{
				{
					{MakeStringStreamCell("Token")},
					{MakeIntegerStreamCell(123)},
				},
			},
		},
		{
			testName: "Several Sheets, with different numbers of columns and rows",
			sheetNames: []string{
				"Sheet 1", "Sheet 2", "Sheet3",
			},
			workbookData: [][][]StreamCell{
				{
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU")},

					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123)},
				},
				{
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU"),
						MakeStringStreamCell("Stock")},

					{MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(1)},

					{MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(3)},
				},
				{
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price")},

					{MakeIntegerStreamCell(9853), MakeStringStreamCell("Guacamole"),
						MakeIntegerStreamCell(500)},

					{MakeIntegerStreamCell(2357), MakeStringStreamCell("Margarita"),
						MakeIntegerStreamCell(700)},
				},
			},
		},
		{
			testName: "Two Sheets with same the name",
			sheetNames: []string{
				"Sheet 1", "Sheet 1",
			},
			workbookData: [][][]StreamCell{
				{
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU")},

					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123)},
				},
				{
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU"),
						MakeStringStreamCell("Stock")},

					{MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(1)},

					{MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(3)},
				},
			},
			expectedError: fmt.Errorf("duplicate sheet name '%s'.", "Sheet 1"),
		},
		{
			testName: "One Sheet Registered, tries to write to two",
			sheetNames: []string{
				"Sheet 1",
			},
			workbookData: [][][]StreamCell{
				{
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU")},

					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123)},
				},
				{
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU")},

					{MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346)},
				},
			},
			expectedError: AlreadyOnLastSheetError,
		},
		{
			testName: "One Sheet, too many columns in row 1",
			sheetNames: []string{
				"Sheet 1",
			},
			workbookData: [][][]StreamCell{
				{
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU")},

					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeStringStreamCell("asdf")},
				},
			},
			expectedError: WrongNumberOfRowsError,
		},
		{
			testName: "One Sheet, too few columns in row 1",
			sheetNames: []string{
				"Sheet 1",
			},
			workbookData: [][][]StreamCell{
				{
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU")},

					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300)},
				},
			},
			expectedError: WrongNumberOfRowsError,
		},
		{
			testName: "Lots of Sheets, only writes rows to one, only writes headers to one, should not error and should still create a valid file",
			sheetNames: []string{
				"Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4", "Sheet 5", "Sheet 6",
			},
			workbookData: [][][]StreamCell{
				{
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU")},

					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123)},
				},
				{{}},
				{{MakeStringStreamCell("Id"), MakeStringStreamCell("Unit Cost")}},
				{{}},
				{{}},
				{{}},
			},
		},
		{
			testName: "Two Sheets, only writes to one, should not error and should still create a valid file",
			sheetNames: []string{
				"Sheet 1", "Sheet 2",
			},
			workbookData: [][][]StreamCell{
				{
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU")},

					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123)},
				},
				{{}},
			},
		},
		{
			testName: "Larger Sheet",
			sheetNames: []string{
				"Sheet 1",
			},
			workbookData: [][][]StreamCell{
				{
					{MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU"),
						MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU"),
						MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU"),
						MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU"),
						MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU"),
						MakeStringStreamCell("Token"), MakeStringStreamCell("Name"),
						MakeStringStreamCell("Price"), MakeStringStreamCell("SKU")},
					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),},
					{MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346)},
					{MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754)},
					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),},
					{MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346)},
					{MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754)},
					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),},
					{MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346)},
					{MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754)},
					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),},
					{MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346)},
					{MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754)},
					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),},
					{MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346)},
					{MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754)},
					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),},
					{MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346)},
					{MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754)},
					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),},
					{MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346)},
					{MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754)},
					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),
						MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123),},
					{MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346),
						MakeIntegerStreamCell(456), MakeStringStreamCell("Salsa"),
						MakeIntegerStreamCell(200), MakeIntegerStreamCell(346)},
					{MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754),
						MakeIntegerStreamCell(789), MakeStringStreamCell("Burritos"),
						MakeIntegerStreamCell(400), MakeIntegerStreamCell(754)},
				},
			},
		},
		{
			testName: "UTF-8 Characters. This XLSX File loads correctly with Excel, Numbers, and Google Docs. It also passes Microsoft's Office File Format Validator.",
			sheetNames: []string{
				"Sheet1",
			},
			workbookData: [][][]StreamCell{
				{
					// String courtesy of https://github.com/minimaxir/big-list-of-naughty-strings/
					// Header row contains the tags that I am filtering on
					{MakeStringStreamCell("Token"), MakeStringStreamCell(endSheetDataTag),
						MakeStringStreamCell("Price"), MakeStringStreamCell(fmt.Sprintf(dimensionTag, "A1:D1"))},
					// Japanese and emojis
					{MakeIntegerStreamCell(123), MakeStringStreamCell("パーティーへ行かないか"),
						MakeIntegerStreamCell(300), MakeStringStreamCell("🍕🐵 🙈 🙉 🙊")},
					// XML encoder/parser test strings
					{MakeIntegerStreamCell(123), MakeStringStreamCell(`<?xml version="1.0" encoding="ISO-8859-1"?>`),
						MakeIntegerStreamCell(300), MakeStringStreamCell(`<?xml version="1.0" encoding="ISO-8859-1"?><!DOCTYPE foo [ <!ELEMENT foo ANY ><!ENTITY xxe SYSTEM "file:///etc/passwd" >]><foo>&xxe;</foo>`)},
					// Upside down text and Right to Left Arabic text
					{MakeIntegerStreamCell(123), MakeStringStreamCell(`˙ɐnbᴉlɐ ɐuƃɐɯ ǝɹolop ʇǝ ǝɹoqɐl ʇn ʇunpᴉpᴉɔuᴉ ɹodɯǝʇ poɯsnᴉǝ op pǝs 'ʇᴉlǝ ƃuᴉɔsᴉdᴉpɐ ɹnʇǝʇɔǝsuoɔ 'ʇǝɯɐ ʇᴉs ɹolop ɯnsdᴉ ɯǝɹo˥
					00˙Ɩ$-`), MakeIntegerStreamCell(300), MakeStringStreamCell(`ﷺ`)} ,
					{MakeIntegerStreamCell(123), MakeStringStreamCell("Taco"),
						MakeIntegerStreamCell(300), MakeIntegerStreamCell(123)},
				},
			},
		},
	}

	for i, testCase := range testCases {
		var filePath string
		var buffer bytes.Buffer
		if TestsShouldMakeRealFiles {
			filePath = fmt.Sprintf("Workbook%d.xlsx", i)
		}

		err := writeStreamFile(filePath, &buffer, testCase.sheetNames, testCase.workbookData, TestsShouldMakeRealFiles)
		if err != testCase.expectedError && err.Error() != testCase.expectedError.Error() {
			t.Fatalf("Error differs from expected error. Error: %v, Expected Error: %v ", err, testCase.expectedError)
		}
		if testCase.expectedError != nil {
			return
		}
		// read the file back with the xlsx package
		var bufReader *bytes.Reader
		var size int64
		if !TestsShouldMakeRealFiles {
			bufReader = bytes.NewReader(buffer.Bytes())
			size = bufReader.Size()
		}
		actualSheetNames, actualWorkbookData := readXLSXFile(t, filePath, bufReader, size, TestsShouldMakeRealFiles)
		// check if data was able to be read correctly
		if !reflect.DeepEqual(actualSheetNames, testCase.sheetNames) {
			t.Fatal("Expected sheet names to be equal")
		}

		expectedWorkbookDataStrings := [][][]string{}
		for j,_ := range testCase.workbookData {
			expectedWorkbookDataStrings = append(expectedWorkbookDataStrings, [][]string{})
			for k,_ := range testCase.workbookData[j]{
				expectedWorkbookDataStrings[j] = append(expectedWorkbookDataStrings[j], []string{})
				for _, cell := range testCase.workbookData[j][k] {
					expectedWorkbookDataStrings[j][k] = append(expectedWorkbookDataStrings[j][k], cell.cellData)
				}
			}

		}
		if !reflect.DeepEqual(actualWorkbookData, expectedWorkbookDataStrings) {
			t.Fatal("Expected workbook data to be equal")
		}
	}
}

// The purpose of TestXlsxStyleBehavior is to ensure that initMaxStyleId has the correct starting value
// and that the logic in AddSheet() that predicts Style IDs is correct.
func (s *StreamSuite) TestXlsxStyleBehavior(t *C) {
	file := NewFile()
	sheet, err := file.AddSheet("Sheet 1")
	if err != nil {
		t.Fatal(err)
	}
	row := sheet.AddRow()
	rowData := []string{"testing", "1", "2", "3"}
	if count := row.WriteSlice(&rowData, -1); count != len(rowData) {
		t.Fatal("not enough cells written")
	}
	parts, err := file.MarshallParts()
	styleSheet, ok := parts["xl/styles.xml"]
	if !ok {
		t.Fatal("no style sheet")
	}
	// Created an XLSX file with only the default style.
	// We expect that the number of styles is one more than our max index constant.
	// This means the library adds two styles by default.
	if !strings.Contains(styleSheet, fmt.Sprintf(`<cellXfs count="%d">`, initMaxStyleId+1)) {
		t.Fatal("Expected sheet to have two styles")
	}

	file = NewFile()
	sheet, err = file.AddSheet("Sheet 1")
	if err != nil {
		t.Fatal(err)
	}
	row = sheet.AddRow()
	rowData = []string{"testing", "1", "2", "3", "4"}
	if count := row.WriteSlice(&rowData, -1); count != len(rowData) {
		t.Fatal("not enough cells written")
	}
	sheet.Cols[0].SetType(CellTypeString)
	sheet.Cols[1].SetType(CellTypeString)
	sheet.Cols[3].SetType(CellTypeNumeric)
	sheet.Cols[4].SetType(CellTypeString)
	parts, err = file.MarshallParts()
	styleSheet, ok = parts["xl/styles.xml"]
	if !ok {
		t.Fatal("no style sheet")
	}
	// Created an XLSX file with two distinct cell types, which should create two new styles.
	// The same cell type was added three times, this should be coalesced into the same style rather than
	// recreating the style. This XLSX stream library depends on this behavior when predicting the next style id.
	if !strings.Contains(styleSheet, fmt.Sprintf(`<cellXfs count="%d">`, initMaxStyleId+1+2)) {
		t.Fatal("Expected sheet to have four styles")
	}
}

// writeStreamFile will write the file using this stream package
func writeStreamFile(filePath string, fileBuffer io.Writer, sheetNames []string, workbookData [][][]StreamCell, shouldMakeRealFiles bool) error {
	var file *StreamFileBuilder
	var err error
	if shouldMakeRealFiles {
		file, err = NewStreamFileBuilderForPath(filePath)
		if err != nil {
			return err
		}
	} else {
		file = NewStreamFileBuilder(fileBuffer)
	}

	err = file.AddStreamStyle(Strings)
	err = file.AddStreamStyle(BoldStrings)
	err = file.AddStreamStyle(ItalicStrings)
	err = file.AddStreamStyle(UnderlinedStrings)
	err = file.AddStreamStyle(Integers)
	err = file.AddStreamStyle(BoldIntegers)
	err = file.AddStreamStyle(ItalicIntegers)
	err = file.AddStreamStyle(UnderlinedIntegers)
	if err != nil { // TODO handle all errors not just one
		return err
	}

	for i, sheetName := range sheetNames {
		header := workbookData[i][0]
		err := file.AddSheet(sheetName, header)
		if err != nil {
			return err
		}
	}
	streamFile, err := file.Build()
	if err != nil {
		return err
	}
	for i, sheetData := range workbookData {

		if i != 0 {
			err = streamFile.NextSheet()
			if err != nil {
				return err
			}
		}
		for i, row := range sheetData {
			if i == 0 {
				continue
			}
			err = streamFile.Write(row)
			if err != nil {
				return err
			}
		}
	}
	err = streamFile.Close()
	if err != nil {
		return err
	}
	return nil
}

// readXLSXFile will read the file using the xlsx package.
func readXLSXFile(t *C, filePath string, fileBuffer io.ReaderAt, size int64, shouldMakeRealFiles bool) ([]string, [][][]string) {
	var readFile *File
	var err error
	if shouldMakeRealFiles {
		readFile, err = OpenFile(filePath)
		if err != nil {
			t.Fatal(err)
		}
	} else {
		readFile, err = OpenReaderAt(fileBuffer, size)
		if err != nil {
			t.Fatal(err)
		}
	}
	var actualWorkbookData [][][]string
	var sheetNames []string
	for _, sheet := range readFile.Sheets {
		sheetData := [][]string{}
		for _, row := range sheet.Rows {
			data := []string{}
			for _, cell := range row.Cells {
				str, err := cell.FormattedValue()
				if err != nil {
					t.Fatal(err)
				}
				data = append(data, str)
			}
			sheetData = append(sheetData, data)
		}
		sheetNames = append(sheetNames, sheet.Name)
		actualWorkbookData = append(actualWorkbookData, sheetData)
	}
	return sheetNames, actualWorkbookData
}

func (s *StreamSuite) TestAddSheetErrorsAfterBuild(t *C) {
	file := NewStreamFileBuilder(bytes.NewBuffer(nil))

	err := file.AddSheet("Sheet1", []StreamCell{MakeStringStreamCell("Header")})
	if err != nil {
		t.Fatal(err)
	}
	err = file.AddSheet("Sheet2", []StreamCell{MakeStringStreamCell("Header2")})
	if err != nil {
		t.Fatal(err)
	}

	_, err = file.Build()
	if err != nil {
		t.Fatal(err)
	}
	err = file.AddSheet("Sheet3", []StreamCell{MakeStringStreamCell("Header3")})
	if err != BuiltStreamFileBuilderError {
		t.Fatal(err)
	}
}

func (s *StreamSuite) TestBuildErrorsAfterBuild(t *C) {
	file := NewStreamFileBuilder(bytes.NewBuffer(nil))

	err := file.AddSheet("Sheet1", []StreamCell{MakeStringStreamCell("Header")})
	if err != nil {
		t.Fatal(err)
	}
	err = file.AddSheet("Sheet2", []StreamCell{MakeStringStreamCell("Header")})
	if err != nil {
		t.Fatal(err)
	}

	_, err = file.Build()
	if err != nil {
		t.Fatal(err)
	}
	_, err = file.Build()
	if err != BuiltStreamFileBuilderError {
		t.Fatal(err)
	}
}

func (s *StreamSuite) TestCloseWithNothingWrittenToSheets(t *C) {
	buffer := bytes.NewBuffer(nil)
	file := NewStreamFileBuilder(buffer)

	sheetNames := []string{"Sheet1", "Sheet2"}
	workbookData := [][][]StreamCell{
		{{MakeStringStreamCell("Header1"), MakeStringStreamCell("Header2")}},
		{{MakeStringStreamCell("Header3"), MakeStringStreamCell("Header4")}},
	}
	err := file.AddSheet(sheetNames[0], workbookData[0][0])
	if err != nil {
		t.Fatal(err)
	}
	err = file.AddSheet(sheetNames[1], workbookData[1][0])
	if err != nil {
		t.Fatal(err)
	}

	stream, err := file.Build()
	if err != nil {
		t.Fatal(err)
	}
	err = stream.Close()
	if err != nil {
		t.Fatal(err)
	}
	bufReader := bytes.NewReader(buffer.Bytes())
	size := bufReader.Size()

	actualSheetNames, actualWorkbookData := readXLSXFile(t, "", bufReader, size, false)
	// check if data was able to be read correctly
	if !reflect.DeepEqual(actualSheetNames, sheetNames) {
		t.Fatal("Expected sheet names to be equal")
	}
	expectedWorkbookDataStrings := [][][]string{}
	for j,_ := range workbookData {
		expectedWorkbookDataStrings = append(expectedWorkbookDataStrings, [][]string{})
		for k,_ := range workbookData[j]{
			expectedWorkbookDataStrings[j] = append(expectedWorkbookDataStrings[j], []string{})
			for _, cell := range workbookData[j][k] {
				expectedWorkbookDataStrings[j][k] = append(expectedWorkbookDataStrings[j][k], cell.cellData)
			}
		}

	}
	if !reflect.DeepEqual(actualWorkbookData, expectedWorkbookDataStrings) {
		t.Fatal("Expected workbook data to be equal")
	}
}

func (s *StreamSuite) TestMakeNewStyleAndUseIt(t *C){
	var filePath string
	var buffer bytes.Buffer
	if TestsShouldMakeRealFiles {
		filePath = fmt.Sprintf("Workbook_newStyle.xlsx")
	}

	timesNewRoman12 := NewFont(12,TimesNewRoman)
	timesNewRoman12.Color = RGB_Dard_Green
	courier20:= NewFont(12, Courier)
	courier20.Color = RGB_Dark_Red

	greenFill := NewFill(Solid_Cell_Fill, RGB_Light_Green, RGB_White)
	redFill   := NewFill(Solid_Cell_Fill, RGB_Light_Red,   RGB_White)

	greenStyle := MakeStyle(0, timesNewRoman12, greenFill, DefaultAlignment(), DefaultBorder())
	redStyle :=   MakeStyle(0, courier20, redFill, DefaultAlignment(), DefaultBorder())

	sheetNames := []string{"Sheet1"}
	workbookData := [][][]StreamCell{
		{
			{MakeStringStreamCell("Header1"), MakeStringStreamCell("Header2")},
			{MakeStyledStringStreamCell("Good", greenStyle), MakeStyledStringStreamCell("Bad", redStyle)},
		},
	}

	err := writeStreamFile(filePath, &buffer, sheetNames, workbookData, TestsShouldMakeRealFiles)

	if err != nil {
		t.Fatal("Error during writing")
	}

	// read the file back with the xlsx package
	var bufReader *bytes.Reader
	var size int64
	if !TestsShouldMakeRealFiles {
		bufReader = bytes.NewReader(buffer.Bytes())
		size = bufReader.Size()
	}
	actualSheetNames, actualWorkbookData := readXLSXFile(t, filePath, bufReader, size, TestsShouldMakeRealFiles)
	// check if data was able to be read correctly
	if !reflect.DeepEqual(actualSheetNames, sheetNames) {
		t.Fatal("Expected sheet names to be equal")
	}

	expectedWorkbookDataStrings := [][][]string{}
	for j,_ := range workbookData {
		expectedWorkbookDataStrings = append(expectedWorkbookDataStrings, [][]string{})
		for k,_ := range workbookData[j]{
			expectedWorkbookDataStrings[j] = append(expectedWorkbookDataStrings[j], []string{})
			for _, cell := range workbookData[j][k] {
				expectedWorkbookDataStrings[j][k] = append(expectedWorkbookDataStrings[j][k], cell.cellData)
			}
		}

	}

	if !reflect.DeepEqual(actualWorkbookData, expectedWorkbookDataStrings) {
		t.Fatal("Expected workbook data to be equal")
	}
}













