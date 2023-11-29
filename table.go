package autoxlsx

import (
	"reflect"

	"github.com/tealeg/xlsx"
)

// AddTableHeaders creates headers row
func (g *Generator) AddTableHeaders(row *xlsx.Row, sheetNo int, t reflect.Type, count int) (int, error) {
	if row == nil {
		sheet, err := g.GetSheet(sheetNo)
		if err != nil {
			return 0, err
		}

		row = sheet.AddRow()
	}

	var currentCount int
	for i := 0; i < t.NumField(); i++ {
		f := t.Field(i)
		added, err := g.addTableHeader(row, sheetNo, f, count)
		if err != nil {
			return 0, err
		}
		count += added
		currentCount += added
	}

	return currentCount, nil
}

func (g *Generator) addTableHeader(row *xlsx.Row, sheetNo int, f reflect.StructField, currentCount int) (int, error) {
	fv := f.Type
	kind := fv.Kind()
	if kind == reflect.Pointer {
		fv = fv.Elem()
		kind = fv.Kind()
	}

	if kind == reflect.Struct {
		return g.AddTableHeaders(row, sheetNo, fv, currentCount)
	}

	tagValue, ok := f.Tag.Lookup("xlsx")
	if !ok {
		tagValue = ""
	}

	fieldOptions, err := g.parseTagValue(sheetNo, tagValue)
	if err != nil {
		return 0, err
	}

	if fieldOptions.Skip {
		return 0, nil
	}
	cell := row.AddCell()

	fieldOptions.ApplyToHeaderCell(cell)

	sheet, err := g.GetSheet(sheetNo)
	if err != nil {
		return 0, err
	}

	fieldOptions.ApplyToCol(sheet.Cols[currentCount])

	return 1, nil
}

// AddTableDataCells creates new data cells
func (g *Generator) AddTableDataCells(row *xlsx.Row, sheetNo int, t reflect.Type, data reflect.Value, count int) (int, error) {
	if row == nil {
		sheet, err := g.GetSheet(sheetNo)
		if err != nil {
			return 0, err
		}

		row = sheet.AddRow()
	}

	var currentCount int
	for i := 0; i < t.NumField(); i++ {
		added, err := g.addTableDataCell(row, sheetNo, data, t.Field(i), count)
		if err != nil {
			return 0, err
		}

		count += added
		currentCount += added
	}

	return currentCount, nil
}

func (g *Generator) addTableDataCell(row *xlsx.Row, sheetNo int, data reflect.Value, f reflect.StructField, currentCount int) (int, error) {
	fv := data.FieldByName(f.Name)
	kind := fv.Kind()
	if kind == reflect.Pointer {
		fv = fv.Elem()
		kind = fv.Kind()
	}

	if kind == reflect.Struct {
		return g.AddTableDataCells(row, sheetNo, fv.Type(), fv, currentCount)
	}

	fieldOptions := g.customOptions[sheetNo][currentCount]
	if fieldOptions.Skip {
		return 1, nil
	}

	cell := row.AddCell()
	addValueToCell(fv, cell)

	fieldOptions.ApplyToCell(cell)
	return 1, nil
}

func addValueToCell(data reflect.Value, cell *xlsx.Cell) {
	switch data.Kind() {
	case reflect.Pointer:
		if data.IsNil() {
			cell.SetValue(nil)
			return
		}

		addValueToCell(data.Elem(), cell)
		return
	case reflect.Invalid:
		cell.SetValue(nil)
		return
	case reflect.Int:
		v := data.Int()
		cell.SetValue(v)
		return
	}
	v := data.Interface()

	cell.SetValue(v)
}
