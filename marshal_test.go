package autoxlsx

import (
	"bytes"
	"testing"
)

type SomeStruct struct {
	ID    int     `xlsx:"id"`
	Value float64 `xlsx:"value,format:0.000000000000,width:25"`
}

func TestMarshal(t *testing.T) {
	tests := []struct {
		name    string
		arg     map[string]interface{}
		wantErr bool
	}{
		{
			name:    "nil arg",
			arg:     nil,
			wantErr: true,
		},
		{
			name: "success",
			arg: map[string]interface{}{
				"sheet1": []SomeStruct{
					{
						ID:    1,
						Value: 2.2,
					},
				},
			},
			wantErr: false,
		},
		{
			name: "invalid sheet name",
			arg: map[string]interface{}{
				"invalid sheet name too long to match": []SomeStruct{
					{
						ID:    1,
						Value: 2.2,
					},
				},
			},
			wantErr: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			buff := new(bytes.Buffer)
			err := Marshal(tt.arg, buff)

			if (err == nil) == tt.wantErr {
				t.Errorf("Marshal got err= %v, want %v", err, tt.wantErr)
			}

			if !tt.wantErr && buff.Len() == 0 {
				t.Errorf("Marshal no error expected, but got empty buffer")
			}
		})
	}
}
