package schema

import (
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/C-Canchola/goexcel/parse"
	"reflect"
	"time"
)

// schema package uses reflection to populate slices of structs
// to reduce amount of code to read in tabular excel sheets
// as slices of go types.

// SchemaTagKey is used to identify which fields of a
// struct should the parsing be applied to.
const TagKey = "gxl"

// ExcelOffset is used to convert between zero and one based indices for parsing purposes.
const ExcelOffset = 1

// Schema is used to provide parsing to a single excel file reference
// in order to populate struct slices.
type Schema struct {
	f *excelize.File
}

// MakeSchema creates a Schema for a given excel file.
// One schema should exist per file to keep workbook
// level variables in sync.
func MakeSchema(filePath string) (Schema, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return Schema{}, err
	}
	return Schema{
		f: f,
	}, nil
}

type sheetSchema struct {
	sheetName string

	schema      Schema
	parsedSheet *parse.ParsedSheet
}

// TimeField is a valid type for Schema parsing.
// Its parsed value is a time.Time value.
type TimeField struct {
	ParsedValue time.Time
	Successful  bool
	StringValue string
	HeaderValue string
}

// FloatField is a valid type for Schema parsing.
// Its parsed value is float64 value.
type FloatField struct {
	ParsedValue float64
	Successful  bool
	StringValue string
	HeaderValue string
}

// IntField is a valid type for Schema parsing.
// Its parsed value is an int value.
type IntField struct {
	ParsedValue int
	Successful  bool
	StringValue string
	HeaderValue string
}

// StringField is a valid type for Schema parsing.
// Its parsed value is a string value.
type StringField struct {
	ParsedValue string
	Successful  bool
	HeaderValue string
}

func (sc Schema) makeSheetSchema(sheetName string) (sheetSchema, error) {
	parsedSheet, err := parse.MakeParsedSheet(sc.f, sheetName)
	if err != nil {
		return sheetSchema{}, err
	}
	return sheetSchema{
		sheetName:   sheetName,
		schema:      sc,
		parsedSheet: parsedSheet,
	}, nil
}

func (shtSc sheetSchema) makeTimeField(rowIdx int, fieldIdx int, colIdx int, pp preProcessor) TimeField {
	s, _ := shtSc.parsedSheet.ParsedString(rowIdx+ExcelOffset, colIdx)
	t, err := shtSc.parsedSheet.ParsedTime(rowIdx+ExcelOffset, colIdx)
	success := err == nil
	return TimeField{
		ParsedValue: t,
		Successful:  success,
		StringValue: s,
		HeaderValue: pp.headerIdxMap[fieldIdx],
	}
}

func (shtSc sheetSchema) makeFloatField(rowIdx int, fieldIdx int, colIdx int, pp preProcessor) FloatField {
	s, _ := shtSc.parsedSheet.ParsedString(rowIdx+ExcelOffset, colIdx)
	f, err := shtSc.parsedSheet.ParsedFloat(rowIdx+ExcelOffset, colIdx)
	success := err == nil
	return FloatField{
		ParsedValue: f,
		Successful:  success,
		StringValue: s,
		HeaderValue: pp.headerIdxMap[fieldIdx],
	}
}

func (shtSc sheetSchema) makeIntField(rowIdx int, fieldIdx int, colIdx int, pp preProcessor) IntField {
	s, _ := shtSc.parsedSheet.ParsedString(rowIdx+ExcelOffset, colIdx)
	i, err := shtSc.parsedSheet.ParsedInt(rowIdx+ExcelOffset, colIdx)
	success := err == nil
	return IntField{
		ParsedValue: i,
		Successful:  success,
		StringValue: s,
		HeaderValue: pp.headerIdxMap[fieldIdx],
	}
}

func (shtSc sheetSchema) makeStringField(rowIdx int, fieldIdx int, colIdx int, pp preProcessor) StringField {
	s, _ := shtSc.parsedSheet.ParsedString(rowIdx+ExcelOffset, colIdx)

	return StringField{
		ParsedValue: s,
		Successful:  true,
		HeaderValue: pp.headerIdxMap[fieldIdx],
	}
}

// MakeAndApplySchema creates a schema based on the given file path
// and attempts the application on the given sheet and value (pointer to slice of whatever
// type which contains the tagged struct fields to be read from the excel file)
func MakeAndApplySchema(filePath string, sheet string, v interface{})error{
	sch, err := MakeSchema(filePath)
	if err != nil{
		return err
	}
	return sch.ApplySchema(sheet, v)
}

// ApplySchema attempts to apply the schema to a worksheet
// and struct slice based upon the tags of the slice's elements
func (sc Schema) ApplySchema(sheet string, v interface{}) error {
	sheetSchema, err := sc.makeSheetSchema(sheet)
	if err != nil {
		return err
	}

	vSlicePtr := reflect.ValueOf(v)
	vSlice := vSlicePtr.Elem()

	if !typeIsStructSlice(vSlice) {
		return ErrNotStructSlice
	}
	sliceEl := vSlice.Type().Elem()
	tempElVal := reflect.New(sliceEl).Elem()

	preProcessor, err := makePreprocessor(tempElVal)
	if err != nil {
		return err
	}

	sheetDetails, err := sheetSchema.makeSheetDetails()
	if err != nil {
		return err
	}

	taggedFieldMap, err := preProcessor.getTaggedFieldColumnIndexMap(sheetDetails)
	if err != nil {
		return err
	}

	for i := 0; i < sheetDetails.tblDimension.RowCount; i++ {
		newSliceEl := sheetSchema.makeNewSliceEl(sliceEl, preProcessor, taggedFieldMap, i)
		vSlice.Set(reflect.Append(vSlice, newSliceEl))
	}
	return nil
}

// makeNewSliceEl iterates each tagged field and applies the correct
// parsing functions to each field.
// NOTE: EXCEL ROW OFFSET IS APPLIED IN PARSING FUNCTIONS
func (shtSc sheetSchema) makeNewSliceEl(el reflect.Type, pp preProcessor, taggedFieldMap map[int]int, rowIdx int) reflect.Value {
	newElValPtr := reflect.New(el)
	newElVal := newElValPtr.Elem()

	for fieldIdx, colIdx := range taggedFieldMap {
		fieldPtr := newElVal.Field(fieldIdx)

		switch pp.taggedFieldTypeMap[fieldIdx] {

		case reflect.TypeOf(TimeField{}):
			timeField := shtSc.makeTimeField(rowIdx, fieldIdx, colIdx, pp)
			fieldPtr.Set(reflect.ValueOf(timeField))

		case reflect.TypeOf(FloatField{}):
			floatField := shtSc.makeFloatField(rowIdx, fieldIdx, colIdx, pp)
			fieldPtr.Set(reflect.ValueOf(floatField))

		case reflect.TypeOf(IntField{}):
			intField := shtSc.makeIntField(rowIdx, fieldIdx, colIdx, pp)
			fieldPtr.Set(reflect.ValueOf(intField))

		case reflect.TypeOf(StringField{}):
			stringField := shtSc.makeStringField(rowIdx, fieldIdx, colIdx, pp)
			fieldPtr.Set(reflect.ValueOf(stringField))

		}
	}
	return newElVal
}
