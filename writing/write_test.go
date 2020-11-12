package writing

import (
	"github.com/C-Canchola/goexcel/schema"
	"path/filepath"
	"testing"
)

type IdData struct {
	Id   schema.StringField `gxl:"ID"`
	Date schema.TimeField   `gxl:"DATE"`
}

func getDataToWrite() ([]IdData, error) {

	var idArr = make([]IdData, 0)
	s, err := schema.MakeSchema(filepath.Join("data", "data.xlsx"))
	if err != nil {
		return nil, err
	}

	err = s.ApplySchema("STRING_ID", &idArr)
	if err != nil {
		return nil, err
	}
	return idArr, nil
}

func TestFileWriter_SaveFile(t *testing.T) {
	data, err := getDataToWrite()
	if err != nil {
		t.Fatal(err)
	}
	writeData := make([][]interface{}, len(data))
	for i := range data {
		writeData[i] = make([]interface{}, 2)
		writeData[i][0] = data[i].Date.ParsedValue
		writeData[i][1] = data[i].Id.ParsedValue
	}

	header := []string{"DATE", "ID"}

	writer := MakeNewFileWriter()
	err = writer.WriteDataToSheet(header, writeData, "STRING_ID")
	if err != nil {
		t.Fatal(err)
	}

	err = writer.SaveFile(filepath.Join("data", "firstWrite.xlsx"), true)
	if err != nil {
		t.Error(err)
	}
}
func TestWriterMultipeWrite(t *testing.T){
	data, err := getDataToWrite()
	if err != nil{
		t.Fatal(err)
	}
	writeData := make([][]interface{}, len(data))
	for i := range data {
		writeData[i] = make([]interface{}, 2)
		writeData[i][0] = data[i].Date.ParsedValue
		writeData[i][1] = data[i].Id.ParsedValue
	}

	header := []string{"DATE", "ID"}

	writer := MakeNewFileWriter()
	err = writer.WriteDataToSheet(header, writeData, "STRING_ID")
	if err != nil {
		t.Fatal(err)
	}
	if err := writer.WriteDataToSheet(header, writeData, "SECOND_SHEET"); err != nil{
		t.Fatal(err)
	}
	err = writer.SaveFile(filepath.Join("data", "multiWrite.xlsx"), true)
	if err != nil {
		t.Error(err)
	}
}
