XSLX
====
xlsx is intended to be a library to simplify reading the XML format
used by recent version of Microsoft Excel in Go programs.

I've recently updated the code (23rd September 2013) to support a
little more style information and made some of the code work
concurrently, which will hopefully improve performance a little.

Updates are sporadic as I no longer rely on this code myself, but I'm
always happy to look at patches and pull requests and I will try to
respond to specific issues you might be having.

There are no current plans to support writing documents.

Usage
-----

Presented here is a minimal example usage that will dump all cell data in a given XLSX file.  A more complete example of this kind of functionality is contained int the XLSX2CSV program <https://github.com/tealeg/xlsx2csv>.:

```go

import (
    "fmt"
    "github.com/tealeg/xlsx"
)

func main() {
    excelFileName := "/home/tealeg/foo.xlsx"
    xlFile, error := xlsx.OpenFile(excelFileName)
    if error != nil {
        ...
    }
    for _, sheet := range xlFile.Sheets {
        for _, row := range sheet.Rows {
            for _, cell := range row.Cells {
                fmt.Printf("%s\n", cell.String())
            }
        }
    }
}

```

Some additional information is available from the cell (for example, style information).  For more details see the godoc output for this package.

License
-------
This code is under a BSD style license:


Copyright 2011-2013 Geoffrey Teale. All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
THIS SOFTWARE IS PROVIDED BY Geoffrey Teale ``AS IS'' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE FREEBSD PROJECT OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.


Eat a peach - Geoff
