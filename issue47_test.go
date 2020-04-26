package xls

import (
	"io/ioutil"
	"path"
	"path/filepath"
	"strings"
	"testing"
)

func TestIssue47(t *testing.T) {
	testdatapath := "testdata"
	files, err := ioutil.ReadDir(testdatapath)
	if err != nil {
		t.Fatalf("Cant read testdata directory contents: %s", err)
	}
	for _, f := range files {
		if filepath.Ext(f.Name()) == ".xls" {
			xlsfilename := f.Name()
			xlsxfilename := strings.TrimSuffix(xlsfilename, filepath.Ext(xlsfilename)) + ".xlsx"
			err := CompareXlsXlsx(path.Join(testdatapath, xlsfilename),
				path.Join(testdatapath, xlsxfilename))
			if err != "" {
				t.Fatalf("XLS file %s an XLSX file are not equal: %s", xlsfilename, err)
			}

		}
	}

}
