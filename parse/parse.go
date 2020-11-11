package parse

import (
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"strconv"
	"time"
)

// The goal of this package is to create
// a way to parse a table to a list of
// go types.
type Parser struct {
	f *excelize.File

	styler *styler
}

// MakeParser creates a Parser for a given excel file reference.
func MakeParser(f *excelize.File)(Parser, error){
	styler, err := makeStyler(f)
	if err != nil{
		return Parser{}, err
	}
	return Parser{
		f:      f,
		styler: styler,
	},nil
}
// ParseFloat attempts to parse the given cell address as float64 by applying
// a numeric styling before accessing the value.
func (p *Parser)ParseFloat(sheet string, row int, col int)(float64, error){
	s, err := p.styler.getNumericStyledCellValue(sheet, row, col)
	if err != nil{
		return -1, err
	}
	return strconv.ParseFloat(s, 64)
}
// ParseInt attempts to parse the given cell address as an int by applying
// a numeric styling before accessing the value.
func (p *Parser)ParseInt(sheet string, row int, col int)(int, error){
	v, err := p.ParseFloat(sheet, row, col)
	if err != nil{
		return -1, err
	}
	return int(v), nil
}
// ParseTime attempts to parse the given cell address as a time.Time by
// applying a numeric styling before accessing the value.
func (p *Parser)ParseTime(sheet string, row int, col int)(time.Time, error){
	v, err := p.ParseFloat(sheet, row, col)
	if err != nil{
		return time.Time{}, err
	}
	return excelize.ExcelDateToTime(v, false)
}
// ParseString attempts to parse the given cell address as a string
// by applying NO styling before accessing the value.
func (p *Parser)ParseString(sheet string, row int, col int)(string, error){
	return p.styler.getCurrentStyledCellValue(sheet, row, col)
}
type styler struct {
	f *excelize.File

	numberStyle int
}

// makeStyler makes a styler for a given excel file reference
// one styler should be made for one reference to keep consistency with applied
// style ids
func makeStyler(f *excelize.File)(*styler, error){
	numberStyle, err := f.NewStyle(`{"decimal_places":15}`)
	if err != nil{
		return nil, err
	}

	return &styler{
		f:           f,
		numberStyle: numberStyle,
	}, nil
}

// getCellStyle returns the current style of the sheet cell with given
// row and column index
func (s *styler)getCellStyle(sheet string, row int, col int)(int, error){
	addr, err := excelize.CoordinatesToCellName(col, row)
	if err != nil{
		return -1, err
	}
	return s.f.GetCellStyle(sheet, addr)
}

// setCellNumericStyle sets the style of a cell to a decimal
// with maximum precision in terms of excel's significant number limit
func (s *styler)setCellNumericStyle(sheet string, row int, col int)error{
	addr, err := excelize.CoordinatesToCellName(col, row)
	if err != nil{
		return err
	}
	return s.f.SetCellStyle(sheet, addr, addr, s.numberStyle)
}

// getNumericStyledCellValue attempts to convert a cell to it's maximum significant
// decimal digit string representation and return said value.
// The original style is also re-applied after accessing the numeric styled value.
func (s *styler)getNumericStyledCellValue(sheet string, row int, col int)(string, error){
	addr, err := excelize.CoordinatesToCellName(col, row)
	if err != nil{
		return "", nil
	}
	originalStyle, err := s.getCellStyle(sheet, row, col)
	defer s.f.SetCellStyle(sheet, addr, addr, originalStyle)

	if err != nil{
		return "", err
	}
	err = s.setCellNumericStyle(sheet, row, col)
	if err != nil{
		return "", err
	}

	return s.f.GetCellValue(sheet, addr)
}

// getCurrentStyledCellValue returns the currently styled cell value
// at the given coordinates
func (s *styler)getCurrentStyledCellValue(sheet string, row int, col int)(string, error){
	addr, err := excelize.CoordinatesToCellName(col, row)
	if err != nil{
		return "", err
	}
	return s.f.GetCellValue(sheet, addr)
}

