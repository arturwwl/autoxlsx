package autoxlsx

import (
	"strconv"
	"strings"

	"github.com/tealeg/xlsx"
)

// CustomOptions holds options for cells and cols
type CustomOptions struct {
	Format     string
	Width      float64
	ColumnName string
}

// NewCustomOptions creates CustomOptions from tag value
func NewCustomOptions(tagValue string) (*CustomOptions, error) {
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
		}
	}
	return options, nil
}

// ApplyToCol applies options to column
func (co *CustomOptions) ApplyToCol(col *xlsx.Col) {
	if co.Width > 0 {
		col.Width = co.Width
	}
}

// ApplyToHeaderCell applies options to header's cell
func (co *CustomOptions) ApplyToHeaderCell(cell *xlsx.Cell) {
	cell.SetValue(co.ColumnName)
}

// ApplyToCell applies options to cell
func (co *CustomOptions) ApplyToCell(cell *xlsx.Cell) {
	if co.Format != "" {
		cell.SetFormat(co.Format)
	}
}
