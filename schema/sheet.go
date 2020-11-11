package schema

// Functions to help store information about the sheet to be parsed by
// schema.

// TableDimension provides information about a tab's
// tabular dimensions.
// These being properties such as ColumnCount and RowCount.
// 	ColumnCount is the number of values in the header row.
//
// 	RowCount is the number of continuous rows which are not
// 	empty under the ColumnCount subspace.
type TableDimension struct {
	RowCount, ColumnCount int
}

type sheetDetails struct {
	headerRow []string

	tblDimension TableDimension
}

func (shtSc sheetSchema) makeSheetDetails() (sheetDetails, error) {
	d := TableDimension{
		RowCount:    len(shtSc.parsedSheet.Original) - 1,
		ColumnCount: len(shtSc.parsedSheet.Original[0]),
	}
	return sheetDetails{
		headerRow:    shtSc.parsedSheet.Original[0],
		tblDimension: d,
	}, nil
}

func (d sheetDetails) headerExcelColumnIndices() map[string][]int {
	m := make(map[string][]int)

	for i, header := range d.headerRow {
		_, ok := m[header]
		if ok {
			m[header] = append(m[header], i)
		} else {
			m[header] = append(make([]int, 0), i)
		}
	}
	return m
}
