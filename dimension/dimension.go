package dimension

import (
	"errors"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"time"
)

// Package to work with the dimensioning of an excel tab
// assumed to be in a tabular format

// TableDimension provides information about a tab's
// tabular dimensions.
// These being properties such as ColumnCount and RowCount.
// 	ColumnCount is the number of values in the header row.
//
// 	RowCount is the number of continuous rows which are not
// 	empty under the ColumnCount subspace.
type TableDimension struct {
	RowCount, ColumnCount int
}
// ErrNoHeaderRow occurs when a dimension is made on a tab
// with now values in the first row.
var ErrNoHeaderRow = errors.New("dimension: no header row on tab")

// MakeTableDimension make a TableDimension struct from
// an excelize file and sheet name.
func MakeTableDimension(f *excelize.File, sheet string)(TableDimension, error){
	getRStart := time.Now()
	cells, err := f.GetRows(sheet)
	fmt.Println("took seconds to get all rows", time.Now().Sub(getRStart))
	if err != nil{
		return TableDimension{},err
	}
	colCount := getColumnCount(cells)
	if colCount == 0{
		return TableDimension{}, ErrNoHeaderRow
	}
	return TableDimension{
		RowCount:    getNonEmptyRowCount(cells, colCount),
		ColumnCount: colCount,
	}, nil
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

// getNonEmptyRowCount returns the number of consecutive non empty rows
// not including the header row.
// Includes column count in order to not include rows
// of data that may have empty values under the column
// subset but non empty values outside.
func getNonEmptyRowCount(cells [][]string, colCount int)int{
	if colCount == 0{
		return 0
	}
	nonEmptyRowCount := 0
	for i := 1; i < len(cells); i++{
		if rowEmpty(cells[i], colCount){
			return nonEmptyRowCount
		}
		nonEmptyRowCount++
	}
	return nonEmptyRowCount
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
