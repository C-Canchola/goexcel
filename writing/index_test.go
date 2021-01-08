package writing

import (
	"path/filepath"
	"testing"
)

func getTestData()[][]string{
	return [][]string{
		{"a1", "b1", "c1", "d1"},
		{"a2", "b2", "c2", "d2"},
	}
}

func TestIndexedWriter_WriteStringDataToSheet(t *testing.T) {
	iw := MakeNewIndexedWriter("Company", "File Type")
	if err := iw.WriteStringDataToSheet([]string{"a", "b", "c", "d"}, getTestData(), "TEP", "Existing Homes"); err != nil{
		t.Error(err)
	}
	if err := iw.SaveFile(filepath.Join("data", "testIndex.xlsx"), true); err != nil{
		t.Error(err)
	}
}
