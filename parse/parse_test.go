package parse

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"path/filepath"
	"testing"
)

var dataFilePath = filepath.Join("data", "data.xlsx")

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
