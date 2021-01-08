package writing

import (
	"errors"
	"strconv"
	"strings"
	"time"
)

// dateParseHack attempts to convert a string to a time using a variety of date formats.
func dateParseHack(strVal string)(time.Time, error){
	strSplit := strings.Split(strVal, "-")
	if len(strSplit) != 3{
		return time.Now(), errors.New("")
	}
	monthVal, err := strconv.ParseInt(strSplit[0], 10, 64)
	if err != nil{
		return time.Now(), errors.New("err time parse hack")
	}
	if monthVal > 12 || monthVal < 1{
		return time.Now(), errors.New("err time parse hack")
	}
	dayVal, err := strconv.ParseInt(strSplit[1], 10, 64)
	if err != nil{
		return time.Now(), errors.New("err time parse hack")
	}
	if dayVal < 1 || dayVal > 31 {
		return time.Now(), errors.New("err time parse hack")
	}
	yearVal, err := strconv.ParseInt(strSplit[2], 10, 64)
	if err != nil {
		return time.Now(), errors.New("err time parse hack")
	}
	if yearVal < 1 || yearVal > 99{
		return time.Now(), errors.New("err time parse hack")
	}
	timeVal := time.Date(int(yearVal + 2000), time.Month(monthVal), int(dayVal), 0, 0, 0, 0, time.UTC)
	return timeVal, nil
}
// convertStringValToInterfaceVal attempts to convert string values in order of
// time -> int -> float -> string (default)
func convertStringValToInterfaceVal(strVal string) interface{} {
	if timeVal, err := dateParseHack(strVal); err == nil {
		return timeVal
	}
	if timeVal, err := time.Parse("12/31/2020", strVal); err == nil {
		return timeVal
	}
	if timeVal, err := time.Parse("12/31/2020 1:59", strVal); err == nil{
		return timeVal
	}
	if timeVal, err := time.Parse("01-09-20", strVal); err == nil{
		return timeVal
	}

	if timeVal, err := time.Parse("12-31-2020", strVal); err == nil{
		return timeVal
	}
	if timeVal, err := time.Parse("12-31-2020 1:59", strVal); err == nil{
		return timeVal
	}
	if timeVal, err := time.Parse("12-31-99 1:59", strVal); err == nil{
		return timeVal
	}

	if intVal, err := strconv.ParseInt(strVal, 10, 64); err == nil {
		return intVal
	}

	if floatVal, err := strconv.ParseFloat(strVal, 64); err == nil {
		return floatVal
	}
	return strVal
}

// convertStringArrToInterfaceArr attempts to convert an array of strings
// to potential parsed types.
func convertStringArrToInterfaceArr(a []string) []interface{} {
	retArr := make([]interface{}, len(a))
	for idx, val := range a {
		retArr[idx] = convertStringValToInterfaceVal(val)
	}
	return retArr
}

func convertTwoDimStringArrToInterface(d [][]string)[][]interface{}{
	interfaceD := make([][]interface{}, 0, len(d))
	for _, r := range d{
		interfaceD = append(interfaceD, convertStringArrToInterfaceArr(r))
	}
	return interfaceD
}
