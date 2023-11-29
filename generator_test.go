package autoxlsx

import (
	"fmt"
	"testing"
	"time"

	"github.com/google/go-cmp/cmp"
	"github.com/google/go-cmp/cmp/cmpopts"
	"github.com/tealeg/xlsx"
)

type WithNilStruct struct {
	ID             int     `xlsx:"id"`
	NillableString *string `xlsx:"string"`
}
type WithOmittedFieldStruct struct {
	ID             int     `xlsx:"id"`
	OmittedCheck   bool    `xlsx:"-"`
	NillableString *string `xlsx:"string"`
}

type WithNestedStruct struct {
	WithOmittedFieldStruct
	WithNilStruct
}

type WithNestedStruct2 struct {
	*WithOmittedFieldStruct
	WithNilStruct
}

type WithTypeStruct struct {
	ID       int       `xlsx:"id"`
	SomeTime time.Time `xlsx:"time,format:yy-mm-dd hh:mm"`
}

var exampleString = "example"

func TestGenerator_AddSheet(t *testing.T) {
	tests := []struct {
		name          string
		arg           string
		currentSheets int
		want          int
		wantErr       bool
	}{
		{
			name:          "base case",
			arg:           "Some Name",
			currentSheets: 0,
			want:          0,
			wantErr:       false,
		},
		{
			name:          "invalid name",
			arg:           "Some very long name what is not allowed",
			currentSheets: 0,
			want:          0,
			wantErr:       true,
		},
		{
			name:          "got some sheets case",
			arg:           "Some Name",
			currentSheets: 3,
			want:          3,
			wantErr:       false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			generator := &Generator{
				wb: xlsx.NewFile(),
			}
			for i := 0; i < tt.currentSheets; i++ {
				generator.sheets = append(generator.sheets, nil)
			}

			got, err := generator.AddSheet(tt.arg)

			if (err == nil) == tt.wantErr {
				t.Errorf("AddSheet got err= %v, want %v", err, tt.wantErr)
			}

			if err == nil {
				if diff := cmp.Diff(tt.want, got); diff != "" {
					t.Errorf("AddSheet return value differs from expected (-want +got)\n%s", diff)
				}
			}
		})
	}
}

func TestGenerator_GetSheet(t *testing.T) {
	tests := []struct {
		name          string
		arg           int
		currentSheets int
		want          *xlsx.Sheet
		wantErr       bool
	}{
		{
			name:          "empty sheets",
			arg:           0,
			currentSheets: 0,
			want:          nil,
			wantErr:       true,
		},
		{
			name:          "sheet no. out of range",
			arg:           5,
			currentSheets: 5,
			want:          nil,
			wantErr:       true,
		},
		{
			name:          "valid sheet no.",
			arg:           4,
			currentSheets: 5,
			want: &xlsx.Sheet{
				Name:        "some-sheet",
				File:        nil,
				Rows:        nil,
				Cols:        nil,
				MaxRow:      0,
				MaxCol:      0,
				Hidden:      false,
				Selected:    false,
				SheetViews:  nil,
				SheetFormat: xlsx.SheetFormat{},
				AutoFilter:  nil,
			},
			wantErr: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			generator := &Generator{
				wb: xlsx.NewFile(),
			}
			for i := 0; i < tt.currentSheets; i++ {
				generator.sheets = append(generator.sheets, &xlsx.Sheet{
					Name:        "some-sheet",
					File:        nil,
					Rows:        nil,
					Cols:        nil,
					MaxRow:      0,
					MaxCol:      0,
					Hidden:      false,
					Selected:    false,
					SheetViews:  nil,
					SheetFormat: xlsx.SheetFormat{},
					AutoFilter:  nil,
				})
			}

			got, err := generator.GetSheet(tt.arg)

			if (err == nil) == tt.wantErr {
				t.Errorf("GetSheet got err= %v, want %v", err, tt.wantErr)
			}

			if err == nil {
				if diff := cmp.Diff(tt.want, got); diff != "" {
					t.Errorf("GetSheet return value differs from expected (-want +got)\n%s", diff)
				}
			}
		})
	}
}

func TestGenerator_AddData(t *testing.T) {
	type args struct {
		sheetNo int
		data    interface{}
	}
	tests := []struct {
		name          string
		args          args
		currentSheets int
		want          []*xlsx.Row
		wantErr       bool
	}{
		{
			name: "empty sheets",
			args: args{
				sheetNo: 0,
				data: []SomeStruct{
					{
						ID:    1,
						Value: 2.2,
					},
				},
			},
			currentSheets: 0,
			want:          nil,
			wantErr:       true,
		},
		{
			name: "invalid data",
			args: args{
				sheetNo: 0,
				data: SomeStruct{
					ID:    1,
					Value: 2.2,
				},
			},
			currentSheets: 1,
			want:          nil,
			wantErr:       true,
		},
		{
			name: "empty data",
			args: args{
				sheetNo: 0,
				data:    []SomeStruct{},
			},
			currentSheets: 1,
			want:          nil,
			wantErr:       true,
		},
		{
			name: "add basic data",
			args: args{
				sheetNo: 0,
				data: []SomeStruct{
					{
						ID:    1,
						Value: 2.2,
					},
				},
			},
			currentSheets: 1,
			want: []*xlsx.Row{
				{
					Cells: []*xlsx.Cell{
						{
							Value: "id",
						},
						{
							Value: "value",
						},
					},
				},
				{
					Cells: []*xlsx.Cell{
						{
							Value:  "1",
							NumFmt: "general",
						},
						{
							Value:  "2.2",
							NumFmt: "0.000000000000",
						},
					},
				},
			},
			wantErr: false,
		},
		{
			name: "add nillable data",
			args: args{
				sheetNo: 0,
				data: []WithNilStruct{
					{
						ID:             1,
						NillableString: &exampleString,
					},
				},
			},
			currentSheets: 1,
			want: []*xlsx.Row{
				{
					Cells: []*xlsx.Cell{
						{
							Value: "id",
						},
						{
							Value: "string",
						},
					},
				},
				{
					Cells: []*xlsx.Cell{
						{
							Value:  "1",
							NumFmt: "general",
						},
						{
							Value:  "example",
							NumFmt: "",
						},
					},
				},
			},
			wantErr: false,
		},
		{
			name: "add nil data",
			args: args{
				sheetNo: 0,
				data: []WithNilStruct{
					{
						ID:             1,
						NillableString: nil,
					},
				},
			},
			currentSheets: 1,
			want: []*xlsx.Row{
				{
					Cells: []*xlsx.Cell{
						{
							Value: "id",
						},
						{
							Value: "string",
						},
					},
				},
				{
					Cells: []*xlsx.Cell{
						{
							Value:  "1",
							NumFmt: "general",
						},
						{
							Value:  "",
							NumFmt: "",
						},
					},
				},
			},
			wantErr: false,
		},
		{
			name: "add pointer to data",
			args: args{
				sheetNo: 0,
				data: []*WithNilStruct{
					{
						ID:             1,
						NillableString: nil,
					},
				},
			},
			currentSheets: 1,
			want: []*xlsx.Row{
				{
					Cells: []*xlsx.Cell{
						{
							Value: "id",
						},
						{
							Value: "string",
						},
					},
				},
				{
					Cells: []*xlsx.Cell{
						{
							Value:  "1",
							NumFmt: "general",
						},
						{
							Value:  "",
							NumFmt: "",
						},
					},
				},
			},
			wantErr: false,
		},
		{
			name: "check skip field",
			args: args{
				sheetNo: 0,
				data: []WithOmittedFieldStruct{
					{
						ID:             1,
						OmittedCheck:   true,
						NillableString: &exampleString,
					},
				},
			},
			currentSheets: 1,
			want: []*xlsx.Row{
				{
					Cells: []*xlsx.Cell{
						{
							Value: "id",
						},
						{
							Value: "string",
						},
					},
				},
				{
					Cells: []*xlsx.Cell{
						{
							Value:  "1",
							NumFmt: "general",
						},
						{
							Value:  "example",
							NumFmt: "",
						},
					},
				},
			},
			wantErr: false,
		},
		{
			name: "check nested struct",
			args: args{
				sheetNo: 0,
				data: []WithNestedStruct{
					{
						WithOmittedFieldStruct: WithOmittedFieldStruct{
							ID:             2,
							OmittedCheck:   false,
							NillableString: &exampleString,
						},
						WithNilStruct: WithNilStruct{
							ID:             1,
							NillableString: nil,
						},
					},
				},
			},
			currentSheets: 1,
			want: []*xlsx.Row{
				{
					Cells: []*xlsx.Cell{
						{
							Value: "id",
						},
						{
							Value: "string",
						},
						{
							Value: "id",
						},
						{
							Value: "string",
						},
					},
				},
				{
					Cells: []*xlsx.Cell{
						{
							Value:  "2",
							NumFmt: "general",
						},
						{
							Value:  "example",
							NumFmt: "",
						},
						{
							Value:  "1",
							NumFmt: "general",
						},
						{
							Value:  "",
							NumFmt: "",
						},
					},
				},
			},
			wantErr: false,
		},
		{
			name: "check nested struct (with pointer)",
			args: args{
				sheetNo: 0,
				data: []WithNestedStruct2{
					{
						WithOmittedFieldStruct: &WithOmittedFieldStruct{
							ID:             2,
							OmittedCheck:   false,
							NillableString: &exampleString,
						},
						WithNilStruct: WithNilStruct{
							ID:             1,
							NillableString: nil,
						},
					},
				},
			},
			currentSheets: 1,
			want: []*xlsx.Row{
				{
					Cells: []*xlsx.Cell{
						{
							Value: "id",
						},
						{
							Value: "string",
						},
						{
							Value: "id",
						},
						{
							Value: "string",
						},
					},
				},
				{
					Cells: []*xlsx.Cell{
						{
							Value:  "2",
							NumFmt: "general",
						},
						{
							Value:  "example",
							NumFmt: "",
						},
						{
							Value:  "1",
							NumFmt: "general",
						},
						{
							Value:  "",
							NumFmt: "",
						},
					},
				},
			},
			wantErr: false,
		},
		{
			name: "check with time",
			args: args{
				sheetNo: 0,
				data: []WithTypeStruct{
					{
						ID:       2,
						SomeTime: time.Date(2020, 1, 1, 0, 0, 0, 0, time.UTC),
					},
				},
			},
			currentSheets: 1,
			want: []*xlsx.Row{
				{
					Cells: []*xlsx.Cell{
						{
							Value: "id",
						},
						{
							Value: "time",
						},
					},
				},
				{
					Cells: []*xlsx.Cell{
						{
							Value:  "2",
							NumFmt: "general",
						},
						{
							Value:  "43831",
							NumFmt: "yy-mm-dd hh:mm",
						},
					},
				},
			},
			wantErr: false,
		},
	}

	ignoreOpts := cmp.Options{
		cmpopts.IgnoreUnexported(xlsx.Row{}, xlsx.Cell{}),
		cmpopts.IgnoreFields(xlsx.Row{}, "Sheet"),
		cmpopts.IgnoreFields(xlsx.Cell{}, "Row"),
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			generator := &Generator{
				wb: xlsx.NewFile(),
			}
			for i := 0; i < tt.currentSheets; i++ {
				_, err := generator.AddSheet(fmt.Sprintf("test-%d", i))
				if err != nil {
					t.Errorf("unable to prepare sheet, err= %v", err)
				}
			}

			err := generator.AddData(tt.args.sheetNo, tt.args.data)

			if (err == nil) == tt.wantErr {
				t.Errorf("AddData got err= %v, want %v", err, tt.wantErr)
			}

			if !tt.wantErr {
				got := generator.wb.Sheets[tt.args.sheetNo].Rows
				if diff := cmp.Diff(tt.want, got, ignoreOpts); diff != "" {
					t.Errorf("AddData cols value differs from expected (-want +got)\n%s", diff)
				}
			}
		})
	}
}
