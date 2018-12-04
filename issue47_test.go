package xls

import (
	"testing"
)

func TestIssue47(t *testing.T) {
	e := compareXlsXlsx("issue47")

	if e != "" {
		t.Fatalf("XLS an XLSX are not equal: %s", e)
	}

}
