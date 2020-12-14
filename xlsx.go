package xlsx

import (
	"fmt"
	"reflect"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

const (
	timeFormat       = "2006-01-02 15:04:05"
	dateFormat       = "2006-01-02"
	defaultXlsxName  = "result.xlsx"
	defaultSheetName = "Sheet1"
)

type Xlsx struct {
	XlsxName  string
	SheetName string
	Headers   []interface{}
	Rows      [][]interface{}
}

func GetHeaders(s []interface{}) (headers []interface{}) {
	if len(s) == 0 {
		return headers
	}
	xt := reflect.TypeOf(s[0])

	for i := 0; i < xt.NumField(); i++ {
		head, ok := xt.Field(i).Tag.Lookup("xlsx")
		if ok {
			headers = append(headers, head)
		}
	}
	return headers
}

func GetRows(s []interface{}) (rows [][]interface{}) {
	if len(s) == 0 {
		return rows
	}
	xt := reflect.TypeOf(s[0])

	for _, e := range s {
		cells := []interface{}{}
		xv := reflect.ValueOf(e)
		for i := 0; i < xv.NumField(); i++ {
			_, ok := xt.Field(i).Tag.Lookup("xlsx")
			if ok {
				cells = append(cells, xv.Field(i).Interface())
			}
		}
		rows = append(rows, cells)
	}
	return rows
}

func (x Xlsx) ToXlsx() {
	xlsxName := x.XlsxName
	sheetName := x.SheetName
	headers := x.Headers
	rows := x.Rows

	if xlsxName == "" {
		xlsxName = defaultXlsxName
	}

	if sheetName == "" {
		xlsxName = defaultSheetName
	}

	file := excelize.NewFile()

	streamWriter, err := file.NewStreamWriter(defaultSheetName)
	if err != nil {
		fmt.Println(err)
	}
	// set sheetName
	file.SetSheetName(defaultSheetName, sheetName)

	if err := streamWriter.SetRow("A1", headers); err != nil {
		fmt.Println(err)
	}

	for i := 0; i <= len(rows)-1; i++ {
		row := rows[i]
		rowID := i + 2
		cell, _ := excelize.CoordinatesToCellName(1, rowID)
		if err := streamWriter.SetRow(cell, row); err != nil {
			fmt.Println(err)
		}
	}
	if err := streamWriter.Flush(); err != nil {
		fmt.Println(err)
	}
	// set xlsxNmae
	if err := file.SaveAs(xlsxName); err != nil {
		fmt.Println(err)
	}
}
