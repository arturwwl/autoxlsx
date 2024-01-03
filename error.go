package autoxlsx

// ErrExpectedSlice is returned when a function expects a slice but receives a different data type.
type ErrExpectedSlice struct{}

func (e *ErrExpectedSlice) Error() string {
	return "expected a slice, got a different data type"
}

// ErrEmptySlice is returned when a provided slice is empty.
type ErrEmptySlice struct{}

func (e *ErrEmptySlice) Error() string {
	return "provided slice is empty"
}

// ErrSheetNotFound is returned when a specified sheet is not found.
type ErrSheetNotFound struct{}

func (e *ErrSheetNotFound) Error() string {
	return "specified sheet not found"
}

// ErrInconsistentMapKeys is returned when all entities must have consistent keys for map fields.
type ErrInconsistentMapKeys struct{}

func (e *ErrInconsistentMapKeys) Error() string {
	return "all entities must have consistent keys for map fields"
}
