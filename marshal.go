package autoxlsx

import (
	"io"

	"github.com/arturwwl/autoxlsx/sheetList"
)

// Marshal expects map, which key is sheet name and value is slice of objects
func Marshal(in *sheetList.List, out io.Writer, options ...GeneratorOption) error {
	g := NewGenerator(options...)
	err := g.GenerateXLSX(in)
	if err != nil {
		return err
	}

	return g.SaveTo(out)
}
