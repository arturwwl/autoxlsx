package helpers_test

import (
	"reflect"
	"testing"
	"time"

	"github.com/google/go-cmp/cmp"

	"github.com/arturwwl/autoxlsx/pkg/helpers"
)

func TestIsCommonGoStruct(t *testing.T) {
	testCases := []struct {
		name     string
		input    reflect.Type
		expected bool
	}{
		{
			name:     "TimeType",
			input:    reflect.TypeOf(time.Time{}),
			expected: true,
		},
		{
			name:     "NonCommonStruct",
			input:    reflect.TypeOf(struct{ Field int }{}),
			expected: false,
		},
	}

	for _, testCase := range testCases {
		t.Run(testCase.name, func(t *testing.T) {
			result := helpers.IsCommonGoStruct(testCase.input)

			if diff := cmp.Diff(testCase.expected, result); diff != "" {
				t.Errorf("AddData cols value differs from expected (-want +got)\n%s", diff)
			}
		})
	}
}
