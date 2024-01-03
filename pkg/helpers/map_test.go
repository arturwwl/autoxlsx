package helpers_test

import (
	"reflect"
	"testing"

	"github.com/google/go-cmp/cmp"

	"github.com/arturwwl/autoxlsx/pkg/helpers"
)

func TestGetMapKeys(t *testing.T) {
	testCases := []struct {
		name     string
		input    interface{}
		expected []string
	}{
		{
			name:     "Empty Map",
			input:    map[string]interface{}{},
			expected: nil,
		},
		{
			name:     "Non-Empty Map",
			input:    map[string]interface{}{"key1": 42, "key2": "value"},
			expected: []string{"key1", "key2"},
		},
	}

	for _, testCase := range testCases {
		t.Run(testCase.name, func(t *testing.T) {
			inputValue := reflect.ValueOf(testCase.input)
			result, err := helpers.GetMapKeys(inputValue)

			if err != nil {
				t.Errorf("Unexpected error: %v", err)
			}

			if diff := cmp.Diff(testCase.expected, result); diff != "" {
				t.Errorf("AddData cols value differs from expected (-want +got)\n%s", diff)
			}
		})
	}
}

func TestAreAllMapKeysSame(t *testing.T) {
	testCases := []struct {
		name     string
		fields   []interface{}
		expected bool
	}{
		{
			name:     "Empty Fields",
			fields:   []interface{}{},
			expected: true,
		},
		{
			name:     "Fields with Same Keys",
			fields:   []interface{}{map[string]interface{}{"key1": 42, "key2": "value"}, map[string]interface{}{"key1": 23, "key2": "value"}},
			expected: true,
		},
		{
			name:     "Fields with Different Keys",
			fields:   []interface{}{map[string]interface{}{"key1": 42, "key2": "value"}, map[string]interface{}{"key1": 23, "key3": "value"}},
			expected: false,
		},
	}

	for _, testCase := range testCases {
		t.Run(testCase.name, func(t *testing.T) {
			var fields []reflect.Value
			for _, field := range testCase.fields {
				fields = append(fields, reflect.ValueOf(field))
			}

			result, err := helpers.AreAllMapKeysSame(fields)
			if err != nil {
				t.Errorf("Unexpected error: %v", err)
			}

			if diff := cmp.Diff(testCase.expected, result); diff != "" {
				t.Errorf("AddData cols value differs from expected (-want +got)\n%s", diff)
			}
		})
	}
}
