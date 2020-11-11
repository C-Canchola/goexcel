package schema

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/C-Canchola/goexcel/dimension"
	"time"
)

// Functions to help store information about the sheet to be parsed by
// schema.

type sheetDetails struct {
	headerRow []string

	tblDimension dimension.TableDimension
}

func (shtSc sheetSchema)makeSheetDetails()(sheetDetails, error){
	d := dimension.TableDimension{
		RowCount:    len(shtSc.parsedSheet.Original) - 1,
		ColumnCount: len(shtSc.parsedSheet.Original[0]),
	}
	return sheetDetails{
		headerRow:    shtSc.parsedSheet.Original[0],
		tblDimension: d,
	},nil
}

func makeSheetDetails(f *excelize.File, sheet string)(sheetDetails, error) {
	tdS := time.Now()
	tblDimension, err := dimension.MakeTableDimension(f, sheet)
	fmt.Println("took tblDimension seconds:", time.Now().Sub(tdS).Seconds())
	if err != nil {
		return sheetDetails{}, err
	}
	headerRow := make([]string, tblDimension.ColumnCount)
	hStart := time.Now()
	for i := range headerRow {
		addr, _ := excelize.CoordinatesToCellName(i+1, 1)
		header, _ := f.GetCellValue(sheet, addr)
		headerRow[i] = header
	}
	fmt.Println("took tblDimension seconds:", time.Now().Sub(hStart).Seconds())
	dStart := time.Now()
	d := sheetDetails{
		headerRow:    headerRow,
		tblDimension: tblDimension,
	}
	fmt.Println("took sheet detail struct seconds:", time.Now().Sub(dStart).Seconds())
	return d, nil
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
