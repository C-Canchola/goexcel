package schema

import (
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/C-Canchola/goexcel/dimension"
)

// Functions to help store information about the sheet to be parsed by
// schema.

type sheetDetails struct {
	headerRow []string

	tblDimension dimension.TableDimension
}

func makeSheetDetails(f *excelize.File, sheet string)(sheetDetails, error) {
	tblDimension, err := dimension.MakeTableDimension(f, sheet)
	if err != nil {
		return sheetDetails{}, err
	}
	headerRow := make([]string, tblDimension.ColumnCount)
	for i := range headerRow {
		addr, _ := excelize.CoordinatesToCellName(i+1, 1)
		header, _ := f.GetCellValue(sheet, addr)
		headerRow[i] = header
	}
	return sheetDetails{
		headerRow:    headerRow,
		tblDimension: tblDimension,
	}, nil
}

func (d sheetDetails)headerExcelColumnIndices()map[string][]int{
	m := make(map[string][]int)

	for i, header := range d.headerRow{
		_, ok := m[header]
		if ok {
			m[header] = append(m[header], i)
		}else{
			m[header] = append(make([]int, 0), i)
		}
	}
	return m
}
