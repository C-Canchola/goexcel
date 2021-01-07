package parse

import (
	"errors"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"path/filepath"
	"strconv"
	"time"
)

// ErrInvalidData is returned when a parsed sheet is attempted to be created from a sheet with
// no rows or a sheet without a header row.
var ErrInvalidData = errors.New("sheetParse: invalid data either none or no header row")

// ErrInvalidIndices is returned when attempting to access invalid pair of indices on a parsed sheet.
var ErrInvalidIndices = errors.New("sheetParse: invalid pair of indices")

// SheetParse returns all the data on a sheet both formatted
// as it is originally and as a decimal. This will be used to
// create go types.
type ParsedSheet struct {
	Original      [][]string
	DecimalFormat [][]string
	// name of the sheet parsed
	Name string
	// Path of the file containing the sheet. Empty if parsed directly from excelize file.
	Path string
	// Name of file containing sheet. Empty if parsed directly from excelize file.
	FileName string
}

func (ps *ParsedSheet) indexErr(r, c int) error {
	if r > len(ps.Original) || c > len(ps.Original[0]) {
		return ErrInvalidIndices
	}
	return nil
}

// ParseString attempts to return the originally formatted cell value as a string.
func (ps *ParsedSheet) ParsedString(r, c int) (string, error) {
	if err := ps.indexErr(r, c); err != nil {
		return "", err
	}
	return ps.Original[r][c], nil
}

// ParseFloat attempts to parse the cell value using the decimal number formatted string
// as a float64.
func (ps *ParsedSheet) ParsedFloat(r, c int) (float64, error) {
	if err := ps.indexErr(r, c); err != nil {
		return 0, err
	}
	return strconv.ParseFloat(ps.DecimalFormat[r][c], 64)
}

// ParseInt attempts to parse the cell value using the decimal number
//formatted string as an int.
func (ps *ParsedSheet) ParsedInt(r, c int) (int, error) {
	f, err := ps.ParsedFloat(r, c)
	if err != nil {
		return 0, err
	}
	return int(f), nil
}

// ParseInt attempts to parse the cell value using the decimal number
//formatted string as an time.Time.
func (ps *ParsedSheet) ParsedTime(r, c int) (time.Time, error) {
	f, err := ps.ParsedFloat(r, c)
	if err != nil {
		return time.Time{}, err
	}
	return excelize.ExcelDateToTime(f, false)
}

// RemoveColumnFromRowPred removes columns of data where the given predicate function
// returns true on the given row index
func (ps *ParsedSheet)RemoveColumnFromRowPred(rIdx int, pred func(s string)bool){
	filterNeeded := false
	keepMap := make(map[int]bool)
	for colIdx, rowS := range ps.Original[rIdx]{
		if pred(rowS){
			filterNeeded = true
			keepMap[colIdx] = true
		}
	}
	if !filterNeeded{
		return
	}
	original, decimal := make([][]string, 0, len(ps.Original)), make([][]string, 0, len(ps.DecimalFormat))
	for rowIdx := range ps.Original{
		originalRow, decimalRow := make([]string, 0, len(ps.Original[rowIdx])), make([]string, 0, len(ps.DecimalFormat[rowIdx]))
		for colIdx := range ps.Original[rowIdx]{
			if keepMap[colIdx]{
				continue
			}
			originalRow = append(originalRow, ps.Original[rowIdx][colIdx])
			decimalRow = append(decimalRow, ps.DecimalFormat[rowIdx][colIdx])
		}
		original = append(original, originalRow)
		decimal = append(decimal, decimalRow)
	}
	ps.Original = original
	ps.DecimalFormat = decimal
}
// RemoveDuplicateColumnsFromRow removes all columns which have a duplicate value in the row with the given index.
func (ps *ParsedSheet)RemoveDuplicateColumnsFromRow(rowIdx int){
	valCountMap := make(map[string]int)
	for _, v := range ps.Original[rowIdx]{
		valCountMap[v] = valCountMap[v] + 1
	}
	pred := func(s string)bool{
		return valCountMap[s] >= 2
	}
	ps.RemoveColumnFromRowPred(rowIdx, pred)
}

// RemoveRightDuplicateColumnsFromRow keeps the first found column of a row with a given value
// and removes any subsequent columns with that same value.
func (ps *ParsedSheet)RemoveRightDuplicateColumnsFromRow(rowIdx int){
	valCountMap := make(map[string]int)
	pred := func(s string)bool{
		valCountMap[s] = valCountMap[s] + 1
		return valCountMap[s] >= 2
	}
	ps.RemoveColumnFromRowPred(rowIdx, pred)
}
// MakeParsedSheet returns a ParsedSheet to provide quick access to both the originally formatted
// cell values as well as the decimal formatted cell values.
func MakeParsedSheet(f *excelize.File, sheet string) (*ParsedSheet, error) {
	cells, err := f.GetRows(sheet)
	if err != nil {
		return nil, err
	}
	if len(cells) == 0 || len(cells[0]) == 0 {
		return nil, ErrInvalidData
	}
	shapedCells := shapeCells(cells)

	startAdd, _ := excelize.CoordinatesToCellName(1, 1)
	endAddr, _ := excelize.CoordinatesToCellName(len(shapedCells[0]), len(shapedCells))
	numberStyle, _ := f.NewStyle(`{"decimal_places":15}`)
	err = f.SetCellStyle(sheet, startAdd, endAddr, numberStyle)
	if err != nil {
		return nil, err
	}

	decCells, _ := f.GetRows(sheet)
	shapedDecCells := shapeCells(decCells)

	return &ParsedSheet{
		Original:      shapedCells,
		DecimalFormat: shapedDecCells,
		Name: sheet,
	}, nil
}

// MakeParsedSheetFromPath attempts to parse a sheet from a given file path
func MakeParsedSheetFromPath(path string, sheet string)(*ParsedSheet, error){
	f, err := excelize.OpenFile(path)
	if err != nil{
		return nil, err
	}
	sht, err := MakeParsedSheet(f, sheet)
	if err != nil{
		return nil, err
	}
	sht.Name = filepath.Base(path)
	sht.Path = path
	return sht, nil
}

// MakeParsedSheetFromPathAndSheetIndex provides a way to parse a sheet by expected sheet
// position.
//	Note: sheetIdx is zero based
func MakeParsedSheetFromPathAndSheetIndex(path string, sheetIdx int)(*ParsedSheet, error){
	f, err := excelize.OpenFile(path)
	if err != nil{
		return nil, err
	}
	sheetName := f.GetSheetName(sheetIdx)
	if sheetName == ""{
		return nil, errors.New(fmt.Sprintf("sheet idx %d does not exist in file %s", sheetIdx, path))
	}
	sht, err := MakeParsedSheet(f, sheetName)
	if err != nil{
		return nil, err
	}
	sht.FileName = filepath.Base(path)
	sht.Path = path
	return sht, nil
}

// ParsedFile is the result of attempting to
// parse all the tabs of a given excel file.
type ParsedFile struct {
	ParsedSheets map[string]*ParsedSheet
	FailedSheets []string
	name string
	path string
}

func (pf *ParsedFile)Name()string{
	return pf.name
}
func (pf *ParsedFile)Path()string{
	return pf.path
}

// ListParsedSheets returns a slice of ParsedSheets with no specific order.
func (pf *ParsedFile)ListParsedSheets()[]ParsedSheet{
	sheets := make([]ParsedSheet,0, len(pf.ParsedSheets))
	for _, sheet := range pf.ParsedSheets{
		sheets = append(sheets, *sheet)
	}
	return sheets
}

func makeParsedFileSync(path string)(*ParsedFile, error){
	f, err := excelize.OpenFile(path)
	if err != nil{
		return nil, err
	}
	parsedSheetMap := make(map[string]*ParsedSheet)
	failedSheets := make([]string, 0)

	for _, nm := range f.GetSheetList(){
		parsedSheet, err := MakeParsedSheet(f, nm)
		switch err {
		case nil:
			parsedSheetMap[nm] = parsedSheet
		default:
			failedSheets = append(failedSheets, nm)
		}
	}
	pf := &ParsedFile{
		ParsedSheets: parsedSheetMap,
		FailedSheets: failedSheets,
		name:         filepath.Base(path),
		path:         path,
	}
	for _, sht := range pf.ParsedSheets{
		sht.FileName = filepath.Base(path)
		sht.Path = path
	}
	return pf, nil
}

// MakeParsedFile attempts to parse every sheet of a given file path.
// TODO parse the sheets concurrently to improve performance
func MakeParsedFile(path string)(*ParsedFile, error){
	return makeParsedFileSync(path)
}

// shapeCells calculates the number of columns
// for a given tabular data structure and
// re-dimensions each row to have that number of columns
func shapeCells2(cells [][]string) [][]string {
	colCount := getColumnCount(cells)
	for i := range cells {
		cells[i] = shapeRow(cells[i], colCount)
		if !rowEmpty(cells[i], colCount) {
			continue
		}
		return cells[:i]
	}
	return cells
}

// shapeCells calculates the number of columns
// for a given tabular data structure and
// re-dimensions each row to have that number of columns
func shapeCells(cells [][]string) [][]string {
	colCount := getColumnCount(cells)
	for i := range cells {
		cells[i] = shapeRow(cells[i], colCount)
	}
	return removeEmptyTrailingRows(cells)
}

// shapeRow will re-dimension the array of strings
// to have colCount values.
//	If colCount < len(r), values will be removed
// 	If colCount > len(r), empty strings will be added
func shapeRow(r []string, colCount int) []string {
	if colCount > len(r) {
		originalLen := len(r)
		for i := 0; i < colCount-originalLen; i++ {
			r = append(r, "")
		}
	}
	return r[:colCount]
}
func getColumnCount(cells [][]string) int {
	if len(cells) == 0 {
		return 0
	}
	firstRow := cells[0]
	if len(firstRow) == 0 {
		return 0
	}
	cnt := 0
	for firstRow[cnt] != "" {
		cnt++

		if cnt == len(firstRow) {
			return cnt
		}
	}
	return cnt
}

func rowEmpty(row []string, colCount int) bool {
	if len(row) == 0 {
		return true
	}

	for i := 0; i < colCount; i++ {
		if row[i] != "" {
			return false
		}
	}
	return true
}

func removeEmptyTrailingRows(cells [][]string)[][]string{
	emptyCount := 0
	for{
		if emptyCount == len(cells){
			break
		}
		row := cells[len(cells) - 1 - emptyCount]
		if !rowEmpty(row, len(row)){
			break
		}
		emptyCount++
	}
	return cells[:len(cells) - emptyCount]
}
