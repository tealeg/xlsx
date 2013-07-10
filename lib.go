// OpenFile take the name of an XLSX file and returns a populated xlsx.File struct for it.
func OpenFile(filename string) (*File, error) {
  var f *zip.ReadCloser
  f, err := zip.OpenReader(filename)
  if err != nil {
    return nil, err
  }
  return ReadZip(f)
}

// ReadZip takes a zip file of an XLSX file and returns a populated xlsx.File struct for it.
func ReadZip(f *zip.ReadCloser) (*File, error) {
  var error error
  var file *File
  var v *zip.File
  var workbook *zip.File
  var styles *zip.File
  var sharedStrings *zip.File
  var reftable []string
  var worksheets map[string]*zip.File
  var sheetMap map[string]*Sheet

  file = new(File)
  worksheets = make(map[string]*zip.File, len(f.File))
  for _, v = range f.File {
    switch v.Name {
    case "xl/sharedStrings.xml":
      sharedStrings = v
    case "xl/workbook.xml":
      workbook = v
    case "xl/styles.xml":
      styles = v
    default:
      if len(v.Name) > 12 {
        if v.Name[0:13] == "xl/worksheets" {
          worksheets[v.Name[14:len(v.Name)-4]] = v
        }
      }
    }
  }
  file.worksheets = worksheets
  reftable, error = readSharedStringsFromZipFile(sharedStrings)
  if error != nil {
    return nil, error
  }
  if reftable == nil {
    error := new(XLSXReaderError)
    error.Err = "No valid sharedStrings.xml found in XLSX file"
    return nil, error
  }
  file.referenceTable = reftable
  style, error := readStylesFromZipFile(styles)
  if error != nil {
    return nil, error
  }
  file.styles = style
  sheets, names, error := readSheetsFromZipFile(workbook, file)
  if error != nil {
    return nil, error
  }
  if sheets == nil {
    error := new(XLSXReaderError)
    error.Err = "No sheets found in XLSX File"
    return nil, error
  }
  file.Sheets = sheets
  sheetMap = make(map[string]*Sheet, len(names))
  for i := 0; i < len(names); i++ {
    sheetMap[names[i]] = sheets[i]
  }
  file.Sheet = sheetMap
  f.Close()
  return file, nil
}
