package dimension

import (
	"errors"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

// Package to work with the dimensioning of an excel tab
// assumed to be in a tabular format

// TableDimension provides information about a tab's
// tabular dimensions.
// These being properties such as row count and column count.
type TableDimension struct {
	RowCount, ColumnCount int
}

var ErrNoHeaderRow = errors.New("dimension: no header row on tab")

//TODO Function which takes an excelize file and sheet name
// and returns a TableDimension
func MakeTableDimension(f *excelize.File, sheet string)(TableDimension, error){

}

func getColumnCount(cells [][]string)int{
	if len(cells) == 0{
		return 0
	}
	firstRow := cells[0]
	cnt := 0
	for firstRow[cnt] != ""{
		cnt++
	}
	return cnt
}

