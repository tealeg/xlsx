include $(GOROOT)/src/Make.inc

TARG=github.com/tealeg/xslx
GOFILES=\
	doc.go\
	lib.go\
	sharedstrings.go\
	workbook.go\
	worksheet.go\

include $(GOROOT)/src/Make.pkg
