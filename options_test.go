package autoxlsx

import (
	"testing"

	"github.com/google/go-cmp/cmp"
)

func TestNewCustomOptions(t *testing.T) {
	tests := []struct {
		name    string
		arg     string
		want    *CustomOptions
		wantErr bool
	}{
		{
			name: "only name",
			arg:  "Some Name",
			want: &CustomOptions{
				Format:     "",
				Width:      0,
				ColumnName: "Some Name",
			},
			wantErr: false,
		},
		{
			name: "name and width",
			arg:  "Some Name,width:123.11",
			want: &CustomOptions{
				Format:     "",
				Width:      123.11,
				ColumnName: "Some Name",
			},
			wantErr: false,
		},
		{
			name:    "name and invalid width",
			arg:     "Some Name,width:12o3.11",
			want:    nil,
			wantErr: true,
		},
		{
			name: "name, format and width",
			arg:  "Some Name,format:asd,width:123.11",
			want: &CustomOptions{
				Format:     "asd",
				Width:      123.11,
				ColumnName: "Some Name",
			},
			wantErr: false,
		},
		{
			name: "name, width and format",
			arg:  "Some Name,width:123.11,format:asd",
			want: &CustomOptions{
				Format:     "asd",
				Width:      123.11,
				ColumnName: "Some Name",
			},
			wantErr: false,
		},
		{
			name: "name and format",
			arg:  "Some Name,format:123",
			want: &CustomOptions{
				Format:     "123",
				Width:      0,
				ColumnName: "Some Name",
			},
			wantErr: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			g := &Generator{}
			got, err := g.NewCustomOptions(tt.arg)

			if (err == nil) == tt.wantErr {
				t.Errorf("NewCustomOptions got err= %v, want %v", err, tt.wantErr)
			}

			if err == nil {
				if diff := cmp.Diff(tt.want, got); diff != "" {
					t.Errorf("NewCustomOptions return value differs from expected (-want +got)\n%s", diff)
				}
			}
		})
	}
}
