package schema

import (
	"reflect"
	"testing"
)

type testStruct struct {
	A StringField `gxl:"AY"`
	B TimeField
}

func TestIsStructSlice(t *testing.T) {
	var structSlice []testStruct
	var intSlice []int

	if !typeIsStructSlice(reflect.ValueOf(structSlice)){
		t.Error("structSlice should return true for typeIsStructSlice")
	}
	if typeIsStructSlice(reflect.ValueOf(intSlice)){
		t.Error("intSlice should return false for typeIsStructSlice")
	}
}

func TestPreProcessing(t *testing.T){
	var tStruct testStruct
	m, err := taggedFieldMap(reflect.ValueOf(tStruct))
	if err != nil{
		t.Error("taggedFieldMap should not return an error for tStruct")
	}
	if m["AY"] != 0{
		t.Error("AY should correspond to the first field of tStruct")
	}

	fieldTypeMap, err := taggedFieldFieldTypeMap(reflect.ValueOf(tStruct), m)
	if err != nil{
		t.Error("creating fieldTypeMap should not result in an error")
	}
	if fieldTypeMap[0] != reflect.TypeOf(StringField{}){
		t.Error("the first index should have a type of StringField")
	}

	preProcessor, err := makePreprocessor(reflect.ValueOf(tStruct))
	if err != nil{
		t.Fatal("there should not be an error in creating the preprocessor for tStruct")
	}
	if preProcessor.headerFieldMap["AY"] != 0{
		t.Error("preprocessors headerFieldMap with key AY should correspond to field index 0")
	}
	if preProcessor.taggedFieldTypeMap[0] != reflect.TypeOf(StringField{}){
		t.Error("preprocessors taggedFieldTypeMap with key of zero should have a type of StringField")
	}
}
