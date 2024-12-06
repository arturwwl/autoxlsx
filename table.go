package autoxlsx

import (
	"reflect"
	"sort"

	"github.com/tealeg/xlsx/v3"

	"github.com/arturwwl/autoxlsx/pkg/helpers"
)

// AddTableDataCells creates new data cells
func (g *Generator) AddTableDataCells(row *xlsx.Row, sheetNo int, t reflect.Type, data reflect.Value, count int) (int, error) {
	sheet, err := g.GetSheet(sheetNo)
	if err != nil {
		return 0, err
	}
	if row == nil {

		row = sheet.AddRow()
	}

	var currentCount int
	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)
		added, err := g.addTableDataCell(row, sheetNo, data, &field, count)
		if err != nil {
			return 0, err
		}

		count += added
		currentCount += added
	}

	return currentCount, nil
}

func (g *Generator) addTableDataCell(row *xlsx.Row, sheetNo int, data reflect.Value, field *reflect.StructField, currentCount int) (int, error) {
	var fv reflect.Value
	if field != nil {
		fv = data.FieldByName(field.Name)
	} else {
		fv = data
	}

	kind := fv.Kind()
	if kind == reflect.Pointer {
		fv = fv.Elem()
		kind = fv.Kind()
	}

	if kind == reflect.Map {
		return g.addMapTableCells(row, sheetNo, data, *field, currentCount)
	}

	if kind == reflect.Struct {
		if !helpers.IsCommonGoStruct(fv.Type()) {
			return g.AddTableDataCells(row, sheetNo, fv.Type(), fv, currentCount)
		}
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

func (g *Generator) addMapTableCells(row *xlsx.Row, sheetNo int, data reflect.Value, field reflect.StructField, currentCount int) (int, error) {
	fv := data.FieldByName(field.Name)
	var added int

	keys := fv.MapKeys()
	// Sort the keys by name
	sort.Slice(keys, func(i, j int) bool {
		return keys[i].String() < keys[j].String()
	})

	for _, key := range keys {
		value := fv.MapIndex(key)
		nAdded, err := g.addTableDataCell(row, sheetNo, value, nil, currentCount+added)
		if err != nil {
			return 0, err
		}
		added += nAdded
	}

	return added, nil
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
	}

	cell.SetValue(data.Interface())
}
