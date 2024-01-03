package helpers

import (
	"reflect"
	"slices"
	"time"
)

var commonGoStructs = []reflect.Type{
	reflect.TypeOf(time.Time{}),
}

func IsCommonGoStruct(t reflect.Type) bool {
	return slices.Contains(commonGoStructs, t)
}
