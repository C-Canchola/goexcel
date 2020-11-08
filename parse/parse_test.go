package parse

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"path/filepath"
	"testing"
	"time"
)

var dataFilePath = filepath.Join("data", "data.xlsx")

const oneCellName = "ONE_FIRST_CELL"
const oneDateCellName = "ONE_DATE"

func TestStyler(t *testing.T){
	f, err := excelize.OpenFile(dataFilePath)
	if err != nil{
		t.Fatal(err)
	}
	styler, err := makeStyler(f)
	if err != nil{
		t.Fatal(err)
	}
	val, err := styler.getNumericStyledCellValue(oneCellName, 1, 1)
	if err != nil{
		t.Error(err)
	}
	fmt.Println("numeric styled cell value is", val)

	originalDateValue, err := styler.getCurrentStyledCellValue(oneDateCellName, 1, 1)
	if err != nil{
		t.Error(err)
	}
	dateVal, err := styler.getNumericStyledCellValue(oneDateCellName, 1, 1)
	if err != nil{
		t.Error(err)
	}
	afterStyledDateValue, err := styler.getCurrentStyledCellValue(oneDateCellName, 1, 1)
	if err != nil{
		t.Error(err)
	}
	fmt.Println("before styling date cell value is", originalDateValue)
	fmt.Println("date formatted as number value is", dateVal)
	fmt.Println("date formatted originally after parse styling is", afterStyledDateValue)
	if afterStyledDateValue != originalDateValue{
		t.Error("after styled date value and original styled date value should be the same")
	}
}

func TestParser(t *testing.T){
	f, err := excelize.OpenFile(dataFilePath)
	if err != nil{
		t.Fatal(err)
	}
	parser, err := MakeParser(f)
	if err != nil{
		t.Fatal(err)
	}

	parsedNumericCellValue, err := parser.ParseFloat(oneCellName, 1, 1)
	if err != nil{
		t.Error("error parsing float")
	}
	if parsedNumericCellValue != 1.1{
		t.Error("parsed float should equal 1.1 but equals", parsedNumericCellValue)
	}

	expectedTime := time.Date(2020, time.Month(11), 8, 0, 0, 0, 0, time.UTC)
	parsedTime, err := parser.ParseTime(oneDateCellName, 1, 1)
	if err != nil{
		t.Error("error parsing time")
	}
	if expectedTime != parsedTime{
		t.Errorf("expected and parsed times do not match: expected %v parsed %v", expectedTime, parsedTime)
	}
}
