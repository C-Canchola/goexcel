package writing

import (
	"errors"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"os"
	"strconv"
)

// tableColumnRatio is used to space column widths
// when writing to an excel table format.
const tableColumnRatio = 1.229

// excelOffset is used to work with zero based code but translated for excel dimensions
const excelOffset = 1

// headerOffset is used to signify that a header row exists and data should start below
const headerOffset = 1

// nextSheetOffset is meant to signify that information is currently being written for a sheet being added.
// current sheet count is not incremented until all the information has been written successfully which
// is the why this is needed.
const nextSheetOffset = 1

// FileWriter is meant to provide a simple
// set of methods to write data to an excel file.
type FileWriter struct {
	file            *excelize.File
	hasWrittenSheet bool

}


// MakeNewFileWriter creates a new FileWriter
func MakeNewFileWriter() *FileWriter {
	return &FileWriter{
		file:            excelize.NewFile(),
		hasWrittenSheet: false,
	}
}


// MakeFileWriterFromExisting reads an existing Excel file
// so that its contents can be appended or changed.
func MakeFileWriterFromExisting(path string) (*FileWriter, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	return &FileWriter{
		file:            f,
		hasWrittenSheet: true,
	}, nil
}

func pathDoesNotExist(path string) bool {
	if _, err := os.Stat(path); os.IsNotExist(err) {
		return true
	}
	return false
}

// ErrNonOverwriteOnExistingPath is returned when attempting to write on
// existing file path when explicitly setting overwrite to false.
var ErrNonOverwriteOnExistingPath = errors.New("writing: attempting to write on an existing path with overwrite set to false")

// SaveFile attempts to save at the given path.
//		overwrite set to false will return ErrNonOverwriteOnExistingPath
//		when attempting to write on an existing path.
func (w *FileWriter) SaveFile(path string, overwrite bool) error {
	if !pathDoesNotExist(path) && !overwrite {
		return ErrNonOverwriteOnExistingPath
	}
	return w.file.SaveAs(path)
}

func (w *FileWriter) sheetExists(sheet string) bool {
	_, ok := w.file.Sheet[sheet]
	return ok
}

func (w *FileWriter) populateEmptySheet(sheet string) {
	if !w.hasWrittenSheet {
		defaultName := w.file.GetSheetName(0)
		w.file.SetSheetName(defaultName, sheet)
		w.hasWrittenSheet = true
		return
	}
	_ = w.addAndNameSheet(sheet)
}

// ErrSheetExists is returned when an operation that is not meant to be performed on
// an existing sheet is.
var ErrSheetExists = errors.New("writing: sheet already exists")

func (w *FileWriter) addAndNameSheet(sheet string) error {
	if w.sheetExists(sheet) {
		return ErrSheetExists
	}
	w.file.NewSheet(sheet)
	return nil
}

// WriteDataToSheet writes the given data fields to the given sheet.
func (w *FileWriter) WriteDataToSheet(header []string, data [][]interface{}, sheet string) error {
	w.populateEmptySheet(sheet)

	for colIdx, h := range header {
		coords, _ := excelize.CoordinatesToCellName(colIdx+1, 1)
		if err := w.file.SetCellValue(sheet, coords, h); err != nil {
			return err
		}
	}

	for rowIdx, row := range data {
		for colIdx, c := range row {
			coords, _ := excelize.CoordinatesToCellName(colIdx+1, rowIdx+2)
			if err := w.file.SetCellValue(sheet, coords, c); err != nil {
				return err
			}
		}
	}
	return nil
}
// WriteStringDataToSheet attempts to write purely string data to a sheet in a way where
// the format is maintained.
func (w *FileWriter)WriteStringDataToSheet(header []string, data[][]string, sheet string)error{
	w.populateEmptySheet(sheet)
	sw, err := w.file.NewStreamWriter(sheet)
	if err != nil{
		return err
	}

	coords, _ := excelize.CoordinatesToCellName(1, 1)

	if err := sw.SetRow(coords, convertStringArrToInterfaceArr(header)); err != nil{
		return err
	}

	for rowIdx, row := range data{
		coords, _ = excelize.CoordinatesToCellName(1, rowIdx + 2)
		if err := sw.SetRow(coords, convertStringArrToInterfaceArr(row)); err != nil{
			return err
		}
	}
	if err := sw.Flush(); err != nil{
		return err
	}
	return nil
}

// FreezeTopRow freezes the top row of a sheet and sets the active cell to the first cell (A1)
func (w *FileWriter)FreezeTopRow(sheet string)error{
	return w.file.SetPanes(sheet, `{"freeze":true,"split":false,"x_split":0,"y_split":1,"top_left_cell":"A2","active_pane":"bottomLeft","panes":[{"sqref":"A1","active_cell":"A1","pane":"bottomLeft"}]}`)
}

func getTableFormatString(sheet string)string{
	return fmt.Sprintf(`{
    "table_name": "%s",
    "table_style": "TableStyleMedium2",
    "show_first_column": false,
    "show_last_column": false,
    "show_row_stripes": true,
    "show_column_stripes": false
}`, "TBL_" + sheet)
}

func (w *FileWriter)WriteDataTableToSheet(header[]string, data[][]interface{}, sheet string)error{
	if err := w.WriteDataToSheet(header, data, sheet); err != nil{
		return err
	}
	endR, endC := len(header), len(data) + 1
	// Not sure why this appears to be backwards?
	coords, _ := excelize.CoordinatesToCellName(endR, endC)
	return w.file.AddTable(sheet, "A1", coords, getTableFormatString(sheet))
}

func convertStringArrToInterface(sa []string) []interface{} {
	ia := make([]interface{}, len(sa))
	for i := range sa {
		ia[i] = sa[i]
	}
	return ia
}

// IndexedWriter is a special type of FileWriter in which tabs are
// added with numeric ids, hyperlinks are added for navigation,
// and descriptor values are possible.
//	Meant to be specifically used for writing large amounts of data
//	which is separated by tabs.
type IndexedWriter struct {
	fw *FileWriter
	indexedSheetCount int
	hyperlinkStyle int
}

// indexSheetName is the constant name of the index tab
const indexSheetName = "INDEX"

// MakeNewIndexedWriter creates a new indexed writer.
//	An index tab is created with the headers "TAB_NAME" and any
//	additional columns provided.
//	Be sure to ensure that any additional details match the ordering given here.
func MakeNewIndexedWriter(additionalCols ...string)*IndexedWriter{
	indexedWriter :=  &IndexedWriter{
		fw:        MakeNewFileWriter(),
		indexedSheetCount: 0,
	}
	style, _ := indexedWriter.fw.file.NewStyle(`{"font":{"color":"#1265BE","underline":"single"}}`)
	indexedWriter.hyperlinkStyle = style
	header := append([]string{"TAB_NAME"}, additionalCols...)
	_ = indexedWriter.fw.WriteStringDataToSheet(header, make([][]string, 0), indexSheetName)
	_ = indexedWriter.fw.FreezeTopRow(indexSheetName)
	return indexedWriter
}

// getNextSheetName returns the numeric string used for naming a data file
func (iw *IndexedWriter) getNextSheetName()string{
	return strconv.Itoa(iw.indexedSheetCount + 1)
}

// getCurrentIndexCellAddress returns the cell of the current index row
// which is based on the number of sheets which have been written.
// 	Zero based, 0 -> column 1 e.g. 0 called with sheet count of 0 = B1
func (iw *IndexedWriter) getCurrentIndexCellAddress(col int)string{
	addr, _ := excelize.CoordinatesToCellName(col + excelOffset, iw.indexedSheetCount + headerOffset + nextSheetOffset)
	return addr
}

// writeIndexedValue writes the provided writeVal to the current row with the given column.
func (iw *IndexedWriter)writeIndexedValue(col int, writeVal interface{})error{
	return iw.fw.file.SetCellValue(indexSheetName, iw.getCurrentIndexCellAddress(col), writeVal)
}

// writeAdditionalDetails writes the sheet name and any additional details to the index tab
func (iw *IndexedWriter)writeAdditionalDetails(shtName string, additionalDetails ...interface{})error{
	if err := iw.writeIndexedValue(0, shtName); err != nil{
		return err
	}
	for colIdx, detail := range additionalDetails{
		if err := iw.writeIndexedValue(colIdx + 1, detail); err != nil{
			return err
		}
	}
	return nil
}

func (iw *IndexedWriter)nextSheetStartCellAddress()string{
	return fmt.Sprintf("%s!A1", iw.getNextSheetName())
}
func (iw *IndexedWriter)nextIndexHyperlinkAddress()string{
	return fmt.Sprintf("%s!%s", indexSheetName, iw.getCurrentIndexCellAddress(0))
}

func (iw *IndexedWriter)writeHyperlinks()error{
	currentName := iw.getNextSheetName()
	if err := iw.fw.file.SetCellHyperLink(indexSheetName, iw.getCurrentIndexCellAddress(0), iw.nextSheetStartCellAddress(), "Location"); err != nil{
		return err
	}
	if err := iw.fw.file.SetCellStyle(indexSheetName, iw.getCurrentIndexCellAddress(0), iw.getCurrentIndexCellAddress(0), iw.hyperlinkStyle); err != nil{
		return err
	}
	if err := iw.fw.file.SetCellHyperLink(currentName, "A1", iw.nextIndexHyperlinkAddress(), "Location"); err != nil{
		return err
	}
	if err := iw.fw.file.SetCellStyle(currentName, "A1", "A1", iw.hyperlinkStyle); err != nil{
		return err
	}
	return nil
}

// WriteStringDataToSheet writes a tab with the given string data and adds the details to the index tab as well
// as navigation links.
func (iw *IndexedWriter) WriteStringDataToSheet(header []string, data [][]string, additionalDetails ...interface{})error{
	shtName := iw.getNextSheetName()
	if err := iw.fw.WriteStringDataToSheet(header, data, shtName); err != nil{
		return err
	}
	if err := iw.writeAdditionalDetails(shtName, additionalDetails...); err != nil{
		return err
	}
	if err := iw.writeHyperlinks(); err != nil{
		return err
	}
	if err := iw.fw.FreezeTopRow(shtName); err != nil{
		return err
	}
	iw.indexedSheetCount++
	return nil
}
// WriteInterfaceDataToSheet writes empty interface data to an indexed tab.
func (iw *IndexedWriter)WriteInterfaceDataToSheet(header []string, data [][]interface{}, additionalDetails ...interface{})error{
	shtName := iw.getNextSheetName()
	if err := iw.fw.WriteDataToSheet(header, data, shtName); err != nil{
		return err
	}
	if err := iw.writeAdditionalDetails(shtName, additionalDetails...); err != nil{
		return err
	}
	if err := iw.writeHyperlinks(); err != nil{
		return err
	}
	if err := iw.fw.FreezeTopRow(shtName); err != nil{
		return err
	}
	iw.indexedSheetCount++
	return nil
}

// SaveFile saves the indexed file at the given path.
// 	overwrite flag will replace an existing file.
func (iw *IndexedWriter)SaveFile(path string, overwrite bool)error{
	return iw.fw.SaveFile(path, overwrite)
}

