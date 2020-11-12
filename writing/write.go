package writing

import (
	"errors"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"os"
)

// tableColumnRatio is used to space column widths
// when writing to an excel table format.
const tableColumnRatio = 1.229

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
func getTableFormatString(sheet string)string{
	return fmt.Sprintf(`{
    "table_name": "%s",
    "table_style": "TableStyleMedium2",
    "show_first_column": false,
    "show_last_column": false,
    "show_row_stripes": true,
    "show_column_stripes": false
}`, sheet+"tbl")
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
