package schema

import (
	"errors"
	"reflect"
)

// taggedFieldMap returns a map from a given reflect.Type
// which maps the values associated with the constant tag
// value (column header) to their field position.

var ErrNotStructType = errors.New("schema: type is not a struct")
var ErrTagsWithSameKey = errors.New("schema: struct has multiple fields with same tag value")

// taggedFieldMap returns a map where the keys are the column headers
// to be searched for upon parsing a tabular excel sheet.
// These column headers are added to the map if they have the constant tag key.
func taggedFieldMap(v reflect.Value) (map[string]int, error) {
	t := v.Type()

	if t.Kind() != reflect.Struct {
		return nil, ErrNotStructType
	}
	m := make(map[string]int)
	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)

		value, ok := field.Tag.Lookup(TagKey)
		if !ok {
			continue
		}
		_, exists := m[value]
		if exists {
			return nil, ErrTagsWithSameKey
		}
		m[value] = i
	}

	return m, nil
}

// taggedFieldFieldTypeMap returns a map of the indices schema tagged fields
// with their field Types.
func taggedFieldFieldTypeMap(v reflect.Value, taggedFieldMap map[string]int) (map[int]reflect.Type, error) {
	t := v.Type()
	if t.Kind() != reflect.Struct {
		return nil, ErrNotStructType
	}
	m := make(map[int]reflect.Type)
	for _, i := range taggedFieldMap {
		field := t.Field(i)
		m[i] = field.Type
	}
	return m, nil
}

// preProcessor is used to hold the type and tag information
// of a type which will be parsed from a tabular excel sheet.
type preProcessor struct {
	headerFieldMap     map[string]int
	headerIdxMap       map[int]string
	taggedFieldTypeMap map[int]reflect.Type
}

var ErrPreprocessorHasInvalidTaggedFields = errors.New("schema: preprocessor has tagged fields which are not valid")

// makePreprocessor creates the preprocessor from a given value of a type which will be parsed.
// It
func makePreprocessor(v reflect.Value) (preProcessor, error) {
	headerFieldMap, err := taggedFieldMap(v)
	if err != nil {
		return preProcessor{}, err
	}
	headerIdxMap := make(map[int]string)
	for s, i := range headerFieldMap {
		headerIdxMap[i] = s
	}
	taggedFieldFieldTypeMap, err := taggedFieldFieldTypeMap(v, headerFieldMap)
	if err != nil {
		return preProcessor{}, err
	}
	madePreProcessor := preProcessor{
		headerFieldMap:     headerFieldMap,
		headerIdxMap:       headerIdxMap,
		taggedFieldTypeMap: taggedFieldFieldTypeMap,
	}
	if !preProcessorHasAllValidTaggedTypes(madePreProcessor) {
		return preProcessor{}, ErrPreprocessorHasInvalidTaggedFields
	}

	return madePreProcessor, nil
}

//getTaggedFieldColumnIndexMap returns a map of key: taggedFieldIndex value:columnIndex where
// the unique header appears in the data.
func (pp preProcessor) getTaggedFieldColumnIndexMap(d sheetDetails) (map[int]int, error) {
	err := preProcessorIsValidWithHeaderRow(pp, d)
	if err != nil {
		return nil, err
	}
	colIndices := d.headerExcelColumnIndices()
	idxMap := make(map[int]int)

	// Can assume valid without checks as validation occurs above.
	for k, v := range pp.headerFieldMap {
		idxMap[v] = colIndices[k][0]
	}
	return idxMap, nil
}
