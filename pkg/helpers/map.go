package helpers

import (
	"reflect"
	"slices"
)

// GetMapKeys retrieves the keys of a map from a reflect.StructField.
func GetMapKeys(data reflect.Value) ([]string, error) {
	// Convert keys to string slice
	var keyStrings []string
	for _, key := range data.MapKeys() {
		keyStrings = append(keyStrings, key.String())
	}

	return keyStrings, nil
}

// AreAllMapKeysSame checks if all records (reflect.Value) have the same keys.
func AreAllMapKeysSame(fields []reflect.Value) (bool, error) {
	if len(fields) == 0 {
		return true, nil
	}

	// Get the keys of the first field
	firstFieldKeys, err := GetMapKeys(fields[0])
	if err != nil {
		return false, err
	}
	slices.Sort(firstFieldKeys)

	// Iterate over the remaining fields and compare keys
	for _, field := range fields[1:] {
		currentKeys, err := GetMapKeys(field)
		if err != nil {
			return false, err
		}
		slices.Sort(currentKeys)

		// Check if the keys are the same
		if !reflect.DeepEqual(firstFieldKeys, currentKeys) {
			return false, nil
		}
	}

	return true, nil
}
