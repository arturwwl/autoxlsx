// Package sheetList provides a simple data structure List
// for managing a list of sheets with associated data.
package sheetList

import (
	"slices"
	"sort"
)

// Option is a functional option type for configuring List.
type Option func(*List)

// List represents a list of sheets with associated data.
type List struct {
	list map[string]interface{} // The underlying map to store sheet data.
	keys []string
}

// New creates a new List instance with the provided map of sheet data and applies options.
func New(m map[string]interface{}, options ...Option) *List {
	l := &List{
		list: m,
	}

	keys := make([]string, 0, len(l.list))
	for key := range l.list {
		keys = append(keys, key)
	}
	l.keys = keys

	for _, opt := range options {
		opt(l)
	}

	return l
}

// WithSort is an option to set the SortAsc or SortDesc flag.
// If sortAsc is true, the list is sorted in ascending order.
// If sortAsc is false, the list is sorted in descending order.
func WithSort(sortAsc bool) Option {
	return func(l *List) {
		l.sortList(sortAsc)
	}
}

// sortList sorts the list based on the provided sortAsc flag.
// If sortAsc is true, the list is sorted in ascending order.
// If sortAsc is false, the list is sorted in descending order.
func (l *List) sortList(sortAsc bool) {
	sort.Strings(l.keys)
	if !sortAsc {
		slices.Reverse(l.keys)
	}
}

// Get retrieves the sorted list.
func (l *List) Get() (map[string]interface{}, []string) {
	return l.list, l.keys
}
