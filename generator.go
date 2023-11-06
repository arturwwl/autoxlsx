package autoxlsx

import (
	"fmt"
	"io"
	"reflect"
	"sync"

	"github.com/arturwwl/gointtoletters"
	"github.com/tealeg/xlsx"
)

// custom errors
var (
	ErrExpectedSlice = fmt.Errorf("invalid data type provided, expected slice")
	ErrEmptySlice    = fmt.Errorf("empty slice provided")
	ErrSheetNotFound = fmt.Errorf("sheet not found")
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
		return nil, ErrSheetNotFound
	}

	return g.sheets[sheetNo], nil
}

// AddData creates headers and rows
func (g *Generator) AddData(sheetNo int, data interface{}) error {
	sData := reflect.ValueOf(data)
	if sData.Kind() != reflect.Slice {
		return ErrExpectedSlice
	}

	sLen := sData.Len()
	if sLen == 0 {
		return ErrEmptySlice
	}

	var i int
	for i = 0; i < sLen; i++ {
		s := sData.Index(i)
		t := s.Type()

		if i == 0 {
			err := g.AddHeaders(sheetNo, t)
			if err != nil {
				return err
			}
		}

		err := g.AddRow(sheetNo, t, s)
		if err != nil {
			return err
		}
	}

	sheet, err := g.GetSheet(sheetNo)
	if err != nil {
		return err
	}

	sheet.AutoFilter = &xlsx.AutoFilter{TopLeftCell: "A1", BottomRightCell: fmt.Sprintf("%s%d", gointtoletters.IntToLetters(i), 1)}
	return nil
}

// AddHeaders creates headers row
func (g *Generator) AddHeaders(sheetNo int, t reflect.Type) error {
	sheet, err := g.GetSheet(sheetNo)
	if err != nil {
		return err
	}

	row := sheet.AddRow()
	for i := 0; i < t.NumField(); i++ {
		f := t.Field(i)
		tagValue, ok := f.Tag.Lookup("xlsx")
		if !ok || tagValue == "-" {
			continue
		}

		fieldOptions, err := g.parseTagValue(sheetNo, tagValue)
		if err != nil {
			return err
		}

		cell := row.AddCell()

		fieldOptions.ApplyToHeaderCell(cell)
		fieldOptions.ApplyToCol(sheet.Cols[i])
	}

	return nil
}

func (g *Generator) parseTagValue(sheetNo int, tagValue string) (*CustomOptions, error) {
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

// AddRow creates new data row
func (g *Generator) AddRow(sheetNo int, t reflect.Type, s reflect.Value) error {
	sheet, err := g.GetSheet(sheetNo)
	if err != nil {
		return err
	}

	row := sheet.AddRow()
	for i := 0; i < t.NumField(); i++ {
		cell := row.AddCell()
		cell.SetValue(s.Field(i).Interface())
		g.customOptions[sheetNo][i].ApplyToCell(cell)
	}

	return nil
}

// SaveTo writes generated xlsx to io.Writer
func (g *Generator) SaveTo(out io.Writer) error {
	return g.wb.Write(out)
}
