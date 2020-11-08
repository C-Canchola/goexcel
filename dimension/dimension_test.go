package dimension

import (
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"path/filepath"
	"testing"
)

var dataFilePath = filepath.Join("data", "data.xlsx")
const twoRowFourColumnName = "TWO_ROW_FOUR_COLUMN"
const zeroRowFourColumnName = "ZERO_ROW_FOUR_COLUMN"
const blankName = "BLANK"
const nonBlankNoHeaderName = "NON_BLANK_NO_HEADER"
const discontinuousThreeColFourRowName = "DISCON_3_COL_4_ROW"

func getDataFile()(*excelize.File, error){
	return excelize.OpenFile(dataFilePath)
}

func getTabCells(tabName string)([][]string, error){
	f, err := getDataFile()
	if err != nil{
		return nil, err
	}
	return f.GetRows(tabName)
}

func TestDimensionGetters(t *testing.T){
	cells, err := getTabCells(twoRowFourColumnName)
	if err != nil{
		t.Fatal(err)
	}

	colCount := getColumnCount(cells)
	if colCount != 4{
		t.Error("colCount should equal 4 but equals", colCount)
	}

	rowCount := getNonEmptyRowCount(cells, colCount)
	if rowCount != 2{
		t.Error("rowCount should be equal to 2 but equals", rowCount)
	}

}

func TestMakeTableDimension(t *testing.T) {
	f, err := getDataFile()
	if err != nil{
		t.Fatal(err)
	}
	dimension, err := MakeTableDimension(f, twoRowFourColumnName)
	if err != nil{
		t.Fatal(err)
	}

	if dimension.ColumnCount != 4{
		t.Error("dimension column count should equal 4 but equals", dimension.ColumnCount)
	}
	if dimension.RowCount != 2 {
		t.Error("dimension row count should equal 2 but equals", dimension.RowCount)
	}

	zeroDataRowDim, err := MakeTableDimension(f, zeroRowFourColumnName)
	if err != nil{
		t.Fatal(err)
	}
	if zeroDataRowDim.RowCount != 0{
		t.Error("zero row dimension should have 0 rows but has", zeroDataRowDim.RowCount)
	}
	if zeroDataRowDim.ColumnCount != 4{
		t.Error("zero row dimension should have 4 columns but has", zeroDataRowDim.ColumnCount)
	}

	_, blankErr := MakeTableDimension(f, blankName)
	if blankErr != ErrNoHeaderRow{
		t.Error("ErrNoHeaderRow should be returned when reading a blank tab")
	}

	_, nonBlankNoHeader := MakeTableDimension(f, nonBlankNoHeaderName)
	if nonBlankNoHeader != ErrNoHeaderRow{
		t.Error("ErrNoHeaderRow should be returned when reading a non blank tab without a header row")
	}

	discontinuousDim, err := MakeTableDimension(f, discontinuousThreeColFourRowName)
	if err != nil{
		t.Fatal(err)
	}

	if discontinuousDim.ColumnCount != 3{
		t.Error("discontinuous dimension should have 3 columns")
	}
	if discontinuousDim.RowCount != 4 {
		t.Error("discontinuous dimension should have 4 rows")
	}

}
