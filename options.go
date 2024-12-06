package autoxlsx

import (
	"strconv"
	"strings"

	"github.com/tealeg/xlsx/v3"
)

// GeneratorOption holds generator option
type GeneratorOption interface{}

// GeneratorOptionAutoFilter holds option for auto filter
type GeneratorOptionAutoFilter struct{}

// GeneratorOptionFreezeFirstColumn holds option for freeze first column
type GeneratorOptionFreezeFirstColumn struct{}

// GeneratorOptionFreezeFirstRow holds option for freeze first row
type GeneratorOptionFreezeFirstRow struct{}

// generatorOptionCustomDropdown holds option for custom dropdown
type generatorOptionCustomDropdown struct {
	values map[string][]string
}

// GeneratorOptionCustomDropdown creates custom dropdown option
func GeneratorOptionCustomDropdown(values map[string][]string) GeneratorOption {
	return generatorOptionCustomDropdown{values: values}
}

// CustomOptions holds options for cells and cols
type CustomOptions struct {
	Format         string
	Width          float64
	ColumnName     string
	Skip           bool
	CustomDropdown CustomDropdown
}

type CustomDropdown struct {
	Rows   int
	Sheet  string
	Values []string
}

// NewCustomOptions creates CustomOptions from tag value
func (g *Generator) NewCustomOptions(tagValue string) (*CustomOptions, error) {
	if tagValue == "" || tagValue == "-" {
		return &CustomOptions{
			Skip: true,
		}, nil
	}

	values := strings.Split(tagValue, ",")
	options := &CustomOptions{
		ColumnName: values[0],
	}

	for k, v := range values {
		if k > 0 {
			var err error
			if strings.Contains(v, "format:") {
				options.Format = strings.TrimPrefix(v, "format:")
			}

			if strings.Contains(v, "width:") {
				options.Width, err = strconv.ParseFloat(strings.TrimPrefix(v, "width:"), 64)
				if err != nil {
					return options, err
				}
			}

			if strings.Contains(v, "dropdown:") {
				options.CustomDropdown.Rows, err = strconv.Atoi(strings.TrimPrefix(v, "dropdown:"))
				if err != nil {
					return options, err
				}

				values, ok := g.customDropdown[options.ColumnName]
				if ok {
					options.CustomDropdown.Values = values
				}
			}

			if strings.Contains(v, "dropdown-sheet:") {
				options.CustomDropdown.Sheet = strings.TrimPrefix(v, "dropdown-sheet:")
			}
		}
	}
	return options, nil
}

var defaultWidth = 12.0

// ApplyToCol applies options to column
func (co *CustomOptions) ApplyToCol(col *xlsx.Col) {
	if co.Width > 0 {
		col.Width = &co.Width
	} else {
		col.Width = &defaultWidth
	}
}

// ApplyToHeaderCell applies options to header's cell
func (co *CustomOptions) ApplyToHeaderCell(cell *xlsx.Cell, colIndex int, customName string) error {
	cell.SetValue(co.ColumnName)
	if co.CustomDropdown.Rows > 0 {
		sheet := cell.Row.Sheet
		dv := xlsx.NewDataValidation(1, colIndex, co.CustomDropdown.Rows+1, colIndex, true)

		if len(co.CustomDropdown.Values) > 0 {
			err := dv.SetDropList(co.CustomDropdown.Values)
			if err != nil {
				return err
			}
		}
		if co.CustomDropdown.Sheet != "" {
			sheetName := co.CustomDropdown.Sheet
			if co.CustomDropdown.Sheet == "auto" {
				co.CustomDropdown.Sheet = co.ColumnName
				if customName != "" {
					sheetName = customName
				}
			}

			err := dv.SetInFileList(sheetName, 1, 1, 1, -1)
			if err != nil {
				return err
			}
		}

		sheet.AddDataValidation(dv)

		return nil
	}
	return nil
}

// ApplyToCell applies options to cell
func (co *CustomOptions) ApplyToCell(cell *xlsx.Cell) {
	if co.Format != "" {
		cell.SetFormat(co.Format)
	}
}
