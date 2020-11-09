package schema

import (
	"errors"
	"reflect"
)

var ErrNotStructSlice = errors.New("schema: type is not a slice of structs")

// typeIsStructSlice returns if a value is slice of structs
func typeIsStructSlice(v reflect.Value)bool{
	vType := v.Type()
	if vType.Kind() != reflect.Slice{
		return false
	}
	return vType.Elem().Kind() == reflect.Struct
}

func typeIsValidTaggedType(t reflect.Type)bool{
	switch t {
	case reflect.TypeOf(TimeField{}), reflect.TypeOf(IntField{}), reflect.TypeOf(FloatField{}), reflect.TypeOf(StringField{}):
		return true
	default:
		return false
	}
}
func preProcessorHasAllValidTaggedTypes(p preProcessor)bool{
	for _, t := range p.taggedFieldTypeMap{
		if !typeIsValidTaggedType(t){
			return false
		}
	}
	return true
}

var ErrTaggedHeaderDNEInData = errors.New("schema: not all tagged headers exist in sheet")
var ErrTaggedHeaderNotUnique = errors.New("schema: tagged header appears more than once in data")
// preProcessorIsValidWithHeaderRow returns if the preprocessor
// will result in a valid schema parse.
// The following must be true:
//		Every tagged fields value should exist in the header row exactly once.
func preProcessorIsValidWithHeaderRow(p preProcessor, d sheetDetails)error{
	sheetHeaderColIndices := d.headerExcelColumnIndices()
	for taggedHeader := range p.headerFieldMap{
		indices, ok := sheetHeaderColIndices[taggedHeader]
		if !ok{
			return ErrTaggedHeaderDNEInData
		}
		if len(indices) > 1{
			return ErrTaggedHeaderNotUnique
		}
	}
	return nil
}



