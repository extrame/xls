package xls

import (
	"bytes"
	"fmt"
	"os"
	"testing"
)

func TestOpen(t *testing.T) {
	xlFile, _ := Open("Table.xls", "")
	sheet1 := xlFile.GetSheet(0)
	fmt.Println(sheet1.Name)
	fmt.Print(sheet1.Rows)
	for k, row1 := range sheet1.Rows {
		// row1 := sheet1.Rows[1]
		fmt.Printf("\n[%d]", k)
		for _, col1 := range row1.Cols {
			// col1 := row1.Cols[0]
			fmt.Print(col1.LastCol())
			fmt.Print(" ")
		}
	}
}

func TestBof(t *testing.T) {
	b := new(bof)
	b.Id = 0x41E
	b.Size = 55
	buf := bytes.NewReader([]byte{0x07, 0x00, 0x19, 0x00, 0x01, 0x22, 0x00, 0xE5, 0xFF, 0x22, 0x00, 0x23, 0x00, 0x2C, 0x00, 0x23, 0x00, 0x23, 0x00, 0x30, 0x00, 0x2E, 0x00, 0x30, 0x00, 0x30, 0x00, 0x3B, 0x00, 0x22, 0x00, 0xE5, 0xFF, 0x22, 0x00, 0x5C, 0x00, 0x2D, 0x00, 0x23, 0x00, 0x2C, 0x20, 0x00})
	new(WorkBook).parseBof(buf, b, b, 0)
}

func TestMaxRow(t *testing.T) {
	xlFile, err := Open("Table.xls", "utf-8")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failure: %v\n", err)
		t.Error(err)
	}

	if sheet1 := xlFile.GetSheet(0); sheet1 != nil {
		if sheet1.MaxRow != 11 {
			t.Errorf("max row is error,is %d instead of 11", sheet1.MaxRow)
		}
	}
}
