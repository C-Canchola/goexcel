package schema

import (
	"fmt"
	"path/filepath"
	"testing"
)
type IdData struct {
	Id StringField `gxl:"ID"`
	Date TimeField `gxl:"DATE"`
}
func TestSchema_ApplySchema(t *testing.T) {

	var idArr = make([]IdData, 0)
	s, err := MakeSchema(filepath.Join("data", "data.xlsx"))
	if err != nil{
		t.Fatal(err)
	}
	err = s.ApplySchema("STRING_ID", &idArr)
	if err != nil{
		t.Error(err)
	}
	fmt.Println(len(idArr))
	for _, v := range idArr {
		fmt.Println(v)
	}
}
