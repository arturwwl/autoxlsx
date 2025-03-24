package autoxlsx

import (
	"reflect"

	"github.com/tealeg/xlsx/v3"

	"github.com/arturwwl/autoxlsx/pkg/helpers"
)

// AddTableHeaders creates headers row
func (g *Generator) AddTableHeaders(row *xlsx.Row, sheetNo int, t reflect.Type, value reflect.Value, count int) (int, bool, error) {
	if row == nil {
		sheet, err := g.GetSheet(sheetNo)
		if err != nil {
			return 0, false, err
		}

		row = sheet.AddRow()
	}

	var currentCount int
	var hasMapField bool
	for i := 0; i < t.NumField(); i++ {
		f := t.Field(i)
		added, withMapField, err := g.addTableHeader(row, sheetNo, f, value, count, nil)
		if err != nil {
			return 0, false, err
		}

		if !hasMapField {
			hasMapField = withMapField
		}

		count += added
		currentCount += added
	}

	return currentCount, hasMapField, nil
}

func (g *Generator) addTableHeader(row *xlsx.Row, sheetNo int, field reflect.StructField, data reflect.Value, currentCount int, fieldOptions *CustomOptions) (int, bool, error) {
	fv := field.Type
	kind := fv.Kind()

	if reflect.Pointer == kind {
		fv = fv.Elem()
		kind = fv.Kind()
	}

	if reflect.Map == kind {
		return g.addMapTableHeader(row, sheetNo, field, data, currentCount)
	}

	if kind == reflect.Struct {
		if !helpers.IsCommonGoStruct(fv) {
			return g.AddTableHeaders(row, sheetNo, fv, data, currentCount)
		}
	}

	fieldOptions, err := g.parseTagValue(sheetNo, field)
	if err != nil {
		return 0, false, err
	}

	if fieldOptions.Skip {
		return 0, false, nil
	}

	err = g.addTableHeaderCell(row, sheetNo, currentCount, fieldOptions, "")
	if err != nil {
		return 0, false, err
	}

	return 1, false, nil
}

func (g *Generator) addTableHeaderCell(row *xlsx.Row, sheetNo int, currentCount int, fieldOptions *CustomOptions, customValue string) error {
	cell := row.AddCell()

	err := fieldOptions.ApplyToHeaderCell(cell, currentCount, customValue)
	if err != nil {
		return err
	}

	if customValue != "" {
		cell.SetValue(customValue)
	}

	sheet, err := g.GetSheet(sheetNo)
	if err != nil {
		return err
	}

	col := xlsx.NewColForRange(currentCount+1, currentCount+1)
	sheet.Cols.Add(col)

	fieldOptions.ApplyToCol(col)

	return nil
}

func (g *Generator) addMapTableHeader(row *xlsx.Row, sheetNo int, field reflect.StructField, data reflect.Value, currentCount int) (int, bool, error) {
	var added int

	keys, err := helpers.GetMapKeys(data.FieldByName(field.Name))
	if err != nil {
		return 0, false, err
	}

	for _, key := range keys {
		fieldOptions, err := g.parseTagValue(sheetNo, field)
		if err != nil {
			return 0, false, err
		}

		if fieldOptions.Skip {
			return 0, false, nil
		}

		err = g.addTableHeaderCell(row, sheetNo, currentCount+added, fieldOptions, key)
		if err != nil {
			return 0, false, err
		}

		added++
	}

	return added, true, nil
}
