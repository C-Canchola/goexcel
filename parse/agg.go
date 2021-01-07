package parse

import (
	"errors"
	"fmt"
)

const ExcelRowOffset = 1

// AggItem represents information for the aggregated items
type AggItem struct {
	SheetName string
	FileName string
	FilePath string
	RowIdx int
	OriginalFormat []string
	DecimalFormat []string
}

// AggregatedParse represents an aggregation similar
// to the appending of pandas dataframes.
type AggregatedParse struct {
	Header []string
	Items []AggItem
}

// AggregateInfo provides a way to change the start
// cell of a ParsedSheet in the case that
// the first cell is not the beginning of the data.
type AggregateInfo struct {
	Sheet ParsedSheet
	StartRow int
	StartCol int
}

// Header returns the row which represents the header columns
// using the information provided.
func (ai AggregateInfo)Header()[]string{
	return ai.Sheet.Original[ai.StartRow][ai.StartCol:]
}
func dataFromInfo(data [][]string, startRow int, startCol int)[][]string{
	return data[startRow + 1:][startCol:]
}
// DecimalFormattedData uses the aggregate info to return a sheets
// decimal formatted data.
func (ai AggregateInfo)DecimalFormattedData()[][]string{
	return dataFromInfo(ai.Sheet.DecimalFormat, ai.StartRow, ai.StartCol)
}
// OriginalFormattedData uses the aggregate info to return a sheets
// originally formatted data.
func (ai AggregateInfo)OriginalFormattedData()[][]string{
	return dataFromInfo(ai.Sheet.Original, ai.StartRow, ai.StartCol)
}

// checkHeaderDuplicates checks for existence of duplicate column headers
// and returns the duplicated value if true.
func checkHeaderDuplicates(ai AggregateInfo)(bool, string){
	dupMap := make(map[string]int)
	for idx, header := range ai.Header(){
		if _, ok := dupMap[header]; ok {
			return true, header
		}
		dupMap[header] = idx
	}
	return false, ""
}

// allHeaderIndices creates a map which will be used to
// determine the position of a value tied to a specific header
// in an aggregation.
func allHeaderIndices(ais ...AggregateInfo)(map[string]int, error){
	posMap := make(map[string]int)
	cnt := 0
	for _, ai := range ais{
		if hasDup, dupVal := checkHeaderDuplicates(ai); hasDup{
			return nil, errors.New(fmt.Sprintf("duplicate column header %s in %s", dupVal, ai.Sheet.Name))
		}
		for _, headerVal := range ai.Header(){
			if _, ok := posMap[headerVal]; !ok{
				posMap[headerVal] = cnt
				cnt++
			}
		}
	}
	return posMap, nil
}

// createAggregateRowMapper returns a function which creates an aggregate
// row from a sheet's row after all the column headers of every sheet
// to be aggregated are considered.
func createAggregateRowMapper(ai AggregateInfo, aggPosMap map[string]int)func([]string)[]string{
	sheetPosMap := make(map[string]int)
	for idx, header := range ai.Header(){
		if _, ok := sheetPosMap[header];!ok{
			sheetPosMap[header] = idx
		}
	}
	return func(r []string)[]string{
		aggRow := make([]string, len(aggPosMap))
		for header, aggIdx := range aggPosMap{
			if _, ok := sheetPosMap[header]; !ok{
				continue
			}
			sheetIdx := sheetPosMap[header]
			aggRow[aggIdx] = r[sheetIdx]
		}
		return aggRow
	}
}
// AggregateAllSheets returns an AggregatedParse from all the given AggregateInfos
func AggregateAllSheets(ais ...AggregateInfo)(AggregatedParse, error){
	aggPosMap, err := allHeaderIndices(ais...)
	if err != nil{
		return AggregatedParse{}, err
	}
	aggItems := make([]AggItem, 0)
	for _, ai := range ais{
		mapper := createAggregateRowMapper(ai, aggPosMap)
		for i := range ai.OriginalFormattedData(){
			aggItem := AggItem{
				SheetName:      ai.Sheet.Name,
				FileName: ai.Sheet.FileName,
				FilePath: ai.Sheet.Path,
				RowIdx:         i + ExcelRowOffset,
				OriginalFormat: mapper(ai.OriginalFormattedData()[i]),
				DecimalFormat:  mapper(ai.DecimalFormattedData()[i]),
			}
			aggItems = append(aggItems, aggItem)
		}
	}
	headerRow := make([]string, len(aggPosMap))
	for header, pos := range aggPosMap{
		headerRow[pos] = header
	}
	return AggregatedParse{
		Header:        headerRow,
		Items: aggItems,
	}, nil
}
// AggregatedAllSheetsDefaultInfo returns an AggregatedParse from
// the given ParsedSheets and assumes that each ParsedSheet's data begins on the first
// cell of the tab
func AggregateAllSheetsDefaultInfo(sheets ...ParsedSheet)(AggregatedParse, error){
	infos := make([]AggregateInfo, len(sheets))
	for i := range infos{
		infos[i] = AggregateInfo{
			Sheet:    sheets[i],
			StartRow: 0,
			StartCol: 0,
		}
	}
	return AggregateAllSheets(infos...)
}

