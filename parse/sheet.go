package parse

import (
	"errors"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
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
	Original [][]string
	DecimalFormat [][]string
}
func (ps *ParsedSheet)indexErr(r, c int)error{
	if r > len(ps.Original) || c > len(ps.Original[0]){
		return ErrInvalidIndices
	}
	return nil
}
// ParseString attempts to return the originally formatted cell value as a string.
func (ps *ParsedSheet)ParsedString(r, c int)(string, error){
	if err := ps.indexErr(r, c); err != nil{
		return "", err
	}
	return ps.Original[r][c], nil
}
// ParseFloat attempts to parse the cell value using the decimal number formatted string
// as a float64.
func (ps *ParsedSheet)ParsedFloat(r, c int)(float64, error){
	if err := ps.indexErr(r, c); err != nil{
		return 0, err
	}
	return strconv.ParseFloat(ps.DecimalFormat[r][c],64)
}
// ParseInt attempts to parse the cell value using the decimal number
//formatted string as an int.
func (ps *ParsedSheet)ParsedInt(r, c int)(int, error){
	f, err := ps.ParsedFloat(r, c)
	if err != nil{
		return 0, err
	}
	return int(f), nil
}
// ParseInt attempts to parse the cell value using the decimal number
//formatted string as an time.Time.
func (ps *ParsedSheet)ParsedTime(r, c int)(time.Time, error){
	f, err := ps.ParsedFloat(r, c)
	if err != nil{
		return time.Time{}, err
	}
	return excelize.ExcelDateToTime(f, false)
}


// MakeParsedSheet returns a ParsedSheet to provide quick access to both the originally formatted
// cell values as well as the decimal formatted cell values.
func MakeParsedSheet(f *excelize.File, sheet string)(*ParsedSheet, error){
	cells, err := f.GetRows(sheet)
	if err != nil{
		return nil, err
	}
	if len(cells) == 0 || len(cells[0]) == 0{
		return nil, ErrInvalidData
	}
	shapedCells := shapeCells(cells)

	startAdd, _ := excelize.CoordinatesToCellName(1, 1)
	endAddr, _ := excelize.CoordinatesToCellName(len(shapedCells[0]), len(shapedCells))
	numberStyle, _ := f.NewStyle(`{"decimal_places":15}`)
	err = f.SetCellStyle(sheet, startAdd, endAddr, numberStyle)
	if err != nil{
		return nil, err
	}

	decCells, _ := f.GetRows(sheet)
	shapedDecCells := shapeCells(decCells)

	return &ParsedSheet{
		Original:      shapedCells,
		DecimalFormat: shapedDecCells,
	}, nil
}

func shapeCells(cells [][]string)[][]string{
	colCount := getColumnCount(cells)
	for i := range cells{
		cells[i] = shapeRow(cells[i], colCount)
		if !rowEmpty(cells[i], colCount){
			continue
		}
		return cells[:i]
	}
	return cells
}

func shapeRow(r []string, colCount int)[]string{
	if colCount > len(r){
		originalLen := len(r)
		for i := 0; i < colCount - originalLen; i++{
			r = append(r, "")
		}
	}
	return r[:colCount]
}
func getColumnCount(cells [][]string)int{
	if len(cells) == 0{
		return 0
	}
	firstRow := cells[0]
	if len(firstRow) == 0{
		return 0
	}
	cnt := 0
	for firstRow[cnt] != ""{
		cnt++

		if cnt == len(firstRow){
			return cnt
		}
	}
	return cnt
}

func rowEmpty(row []string, colCount int)bool{
	if len(row) == 0{
		return true
	}

	for i := 0; i < colCount;i++{
		if row[i] != ""{
			return false
		}
	}
	return true
}
