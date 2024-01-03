package autoxlsx

import (
	"fmt"
	"io"
	"reflect"
	"sync"

	"github.com/arturwwl/gointtoletters"
	"github.com/tealeg/xlsx"

	"github.com/arturwwl/autoxlsx/pkg/helpers"
)

// Generator holds data needed for generating
type Generator struct {
	sync.Mutex
	sheets        []*xlsx.Sheet
	customOptions [][]*CustomOptions
	wb            *xlsx.File
}

// NewGenerator creates new generator instance
func NewGenerator() *Generator {
	return &Generator{
		Mutex:         sync.Mutex{},
		sheets:        nil,
		customOptions: nil,
		wb:            xlsx.NewFile(),
	}
}

// GenerateXLSX generate xlsx for provided slice
func (g *Generator) GenerateXLSX(list map[string]interface{}) error {
	for sheetName, data := range list {
		sheetNo, err := g.AddSheet(sheetName)
		if err != nil {
			return err
		}

		err = g.AddData(sheetNo, data)
		if err != nil {
			return err
		}
	}

	return nil
}

// AddSheet creates new sheet
func (g *Generator) AddSheet(sheetName string) (int, error) {
	sheet, err := g.wb.AddSheet(sheetName)
	if err != nil {
		return -1, err
	}

	g.Mutex.Lock()
	defer g.Mutex.Unlock()
	g.sheets = append(g.sheets, sheet)
	g.customOptions = append(g.customOptions, []*CustomOptions{})

	return len(g.sheets) - 1, nil
}

// GetSheet if not found it returns an error
func (g *Generator) GetSheet(sheetNo int) (*xlsx.Sheet, error) {
	if len(g.sheets) <= sheetNo {
		return nil, &ErrSheetNotFound{}
	}

	return g.sheets[sheetNo], nil
}

// validateAndLength validates the input data, returns the slice length, and an error if validation fails
func validateAndLength(data interface{}) (int, error) {
	sliceValue := reflect.ValueOf(data)
	if sliceValue.Kind() != reflect.Slice {
		return 0, &ErrExpectedSlice{}
	}

	sliceLen := sliceValue.Len()
	if sliceLen == 0 {
		return 0, &ErrEmptySlice{}
	}

	return sliceLen, nil
}

// processHeaders processes headers for the given item and updates mapFields if needed
func (g *Generator) processHeaders(sheetNo int, itemType reflect.Type, itemValue reflect.Value, mapFields *[]string) (int, error) {
	count, withMap, err := g.AddTableHeaders(nil, sheetNo, itemType, itemValue, 0)
	if err != nil {
		return 0, err
	}

	// Identify fields with map type
	if withMap {
		for j := 0; j < itemType.NumField(); j++ {
			field := itemType.Field(j)
			if field.Type.Kind() == reflect.Map {
				*mapFields = append(*mapFields, field.Name)
			}
		}
	}

	return count, nil
}

// processItem processes an individual item, updating mapValues and processing data cells
func (g *Generator) processItem(sheetNo int, itemType reflect.Type, itemValue reflect.Value, mapFields []string, mapValues map[string][]reflect.Value) error {
	// Collect map values for comparison
	for _, fieldName := range mapFields {
		field, found := itemType.FieldByName(fieldName)
		if !found {
			return fmt.Errorf("field %s not found in type %s", fieldName, itemType.Name())
		}
		index := field.Index[0] // Assuming map fields have only one index
		mapValues[fieldName] = append(mapValues[fieldName], itemValue.Field(index))
	}

	// Process data cells
	_, err := g.AddTableDataCells(nil, sheetNo, itemType, itemValue, 0)
	return err
}

// setSheetProperties sets up sheet properties such as AutoFilter and SheetViews
func (g *Generator) setSheetProperties(sheetNo, rowLength, sliceLen int) error {
	sheet, err := g.GetSheet(sheetNo)
	if err != nil {
		return err
	}

	sheet.AutoFilter = &xlsx.AutoFilter{
		TopLeftCell:     "A1",
		BottomRightCell: fmt.Sprintf("%s%d", gointtoletters.IntToLetters(rowLength), sliceLen),
	}

	sheet.SheetViews = append(sheet.SheetViews, xlsx.SheetView{
		Pane: &xlsx.Pane{
			XSplit:      0,
			YSplit:      1,
			TopLeftCell: "A2",
			ActivePane:  "bottomLeft",
			State:       "frozen",
		},
	})

	return nil
}

func (g *Generator) AddData(sheetNo int, data interface{}) error {
	sliceLen, err := validateAndLength(data)
	if err != nil {
		return err
	}

	rowLength, mapValues, err := g.processData(sheetNo, data, sliceLen)
	if err != nil {
		return err
	}

	if err := g.checkConsistentMapKeys(mapValues); err != nil {
		return err
	}

	if err := g.setSheetProperties(sheetNo, rowLength, sliceLen); err != nil {
		return err
	}

	return nil
}

func (g *Generator) processData(sheetNo int, data interface{}, sliceLen int) (int, map[string][]reflect.Value, error) {
	var rowLength int
	var mapFields []string
	mapValues := make(map[string][]reflect.Value)

	for i := 0; i < sliceLen; i++ {
		itemValue := reflect.ValueOf(data).Index(i)
		itemType := itemValue.Type()

		// Handle pointers
		if itemType.Kind() == reflect.Ptr {
			if itemValue.IsNil() {
				continue
			}

			itemValue = itemValue.Elem()
			itemType = itemValue.Type()
		}

		// Process headers for the first item
		if i == 0 {
			var err error
			rowLength, err = g.processHeaders(sheetNo, itemType, itemValue, &mapFields)
			if err != nil {
				return 0, nil, err
			}
		}

		// Process the item
		if err := g.processItem(sheetNo, itemType, itemValue, mapFields, mapValues); err != nil {
			return 0, nil, err
		}
	}

	return rowLength, mapValues, nil
}

func (g *Generator) checkConsistentMapKeys(mapValues map[string][]reflect.Value) error {
	for _, maps := range mapValues {
		if sameKeys, err := helpers.AreAllMapKeysSame(maps); err != nil || !sameKeys {
			return &ErrInconsistentMapKeys{}
		}
	}
	return nil
}

func (g *Generator) parseTagValue(sheetNo int, f reflect.StructField) (*CustomOptions, error) {
	tagValue, ok := f.Tag.Lookup("xlsx")
	if !ok {
		tagValue = ""
	}

	_, err := g.GetSheet(sheetNo)
	if err != nil {
		return nil, err
	}

	options, err := NewCustomOptions(tagValue)
	if err != nil {
		return nil, err
	}

	g.customOptions[sheetNo] = append(g.customOptions[sheetNo], options)

	return options, nil
}

// SaveTo writes generated xlsx to io.Writer
func (g *Generator) SaveTo(out io.Writer) error {
	return g.wb.Write(out)
}
