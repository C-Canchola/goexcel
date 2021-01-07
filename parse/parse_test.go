package parse

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"path/filepath"
	"testing"
)

var dataFilePath = filepath.Join("data", "data.xlsx")
var largeDataFilePath = filepath.Join("data", "largeWrite.xlsx")
var aggDataPath = filepath.Join("data", "aggTest.xlsx")

func TestMakeParsedSheet(t *testing.T) {
	f, err := excelize.OpenFile(dataFilePath)
	if err != nil {
		t.Fatal(err)
	}
	ps, err := MakeParsedSheet(f, "PARSE")
	if err != nil {
		t.Fatal(err)
	}

	fmt.Println(ps.ParsedString(1, 0))
	fmt.Println(ps.ParsedInt(1, 1))
	fmt.Println(ps.ParsedFloat(1, 2))
	fmt.Println(ps.ParsedTime(1, 3))
}

func TestMakeParsedSheetFromPath(t *testing.T) {
	ps, err := MakeParsedSheetFromPath(dataFilePath, "PARSE")
	if err != nil{
		t.Fatal(err)
	}
	if ps.Name != "PARSE"{
		t.Error("parsed sheet name should be PARSE, is:", ps.Name)
	}
}

func TestMakeParsedFile(t *testing.T) {
	pf, err := MakeParsedFile(dataFilePath)
	if err != nil{
		t.Fatal(err)
	}
	if len(pf.ParsedSheets) != 3{
		t.Error("number of parsed sheets should be 3, is", len(pf.ParsedSheets))
	}
	if pf.Path() != dataFilePath {
		t.Error("path should be",dataFilePath, "is", pf.path)
	}
	if pf.Name() != "data.xlsx"{
		t.Error("name should be data.xlsx is", pf.Name())
	}
}
func TestMakeParsedLargeFile(t *testing.T){
	pf, err := MakeParsedFile(largeDataFilePath)
	if err != nil{
		t.Fatal(err)
	}
	if len(pf.ParsedSheets) != 100{
		t.Error("number of parsed sheets should be 100, is", len(pf.ParsedSheets))
	}
	if pf.Path() != largeDataFilePath {
		t.Error("path should be",largeDataFilePath, "is", pf.path)
	}
	if pf.Name() != "largeWrite.xlsx"{
		t.Error("name should be data.xlsx is", pf.Name())
	}
}

func TestAggregateAllSheetsDefaultInfo(t *testing.T) {
	pf, err := MakeParsedFile(largeDataFilePath)
	if err != nil{
		t.Fatal(err)
	}
	rowsPerTab := 900
	tabs := 100
	agg, err := AggregateAllSheetsDefaultInfo(pf.ListParsedSheets()...)
	if err != nil{
		t.Error(err)
	}
	if rowsPerTab * tabs != len(agg.Items){
		t.Error("total aggregated rows should be", rowsPerTab * tabs, "is", len(agg.Items))
	}

	pf2, err := MakeParsedFile(aggDataPath)
	if err != nil{
		t.Fatal(err)
	}
	agg2, err := AggregateAllSheetsDefaultInfo(pf2.ListParsedSheets()...)
	if err != nil{
		t.Error(err)
	}
	fmt.Println(len(agg2.Header))
	ps1, err := MakeParsedSheetFromPathAndSheetIndex(aggDataPath, 0)
	if err != nil{
		t.Fatal(err)
	}
	if ps1.Name != "AGG_1"{
		t.Error("parsed sheet name should equal AGG_1, is:",ps1.Name)
	}
}
