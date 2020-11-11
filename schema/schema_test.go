package schema

import (
	"fmt"
	"path/filepath"
	"testing"
)

type IdData struct {
	Id   StringField `gxl:"ID"`
	Date TimeField   `gxl:"DATE"`
}

func TestSchema_ApplySchema(t *testing.T) {

	var idArr = make([]IdData, 0)
	s, err := MakeSchema(filepath.Join("data", "data.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	err = s.ApplySchema("STRING_ID", &idArr)
	if err != nil {
		t.Error(err)
	}
	fmt.Println(len(idArr))
	for _, v := range idArr {
		fmt.Println(v)
	}
}

type LargeStringOnly struct {
	ReferenceId StringField `gxl:"Reference ID"`
}
type LargeIntOnly struct {
	Month IntField `gxl:"Month Reported"`
}

type LargeTwoInts struct {
	Month IntField `gxl:"Month Reported"`
	Year  IntField `gxl:"Year Reported"`
}

type LargeTwoString struct {
	ReferenceId StringField `gxl:"Reference ID"`
	Month       StringField `gxl:"Month Reported"`
}

func TestSchema_LargeRead(t *testing.T) {
	var itemArr = make([]LargeTwoInts, 0)
	s, err := MakeSchema(filepath.Join("data", "large_data.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	err = s.ApplySchema("Savings Report", &itemArr)
	if err != nil {
		t.Error(err)
	}
	fmt.Println(len(itemArr))
	fmt.Println(itemArr[0])

}
