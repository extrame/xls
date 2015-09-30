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

//read the content of first two cols in each row
func ExampleReadXls(t *testing.T) {
	xlFile, err := Open("Table.xls", "utf-8")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failure: %v\n", err)
		t.Error(err)
	}

	if sheet1 := xlFile.GetSheet(0); sheet1 != nil {
		fmt.Print("Total Lines ", sheet1.MaxRow, sheet1.Name)
		col1 := sheet1.Rows[0].Cols[0]
		col2 := sheet1.Rows[0].Cols[0]
		for i := 0; i <= (int(sheet1.MaxRow)); i++ {
			row1 := sheet1.Rows[uint16(i)]
			col1 = row1.Cols[0]
			col2 = row1.Cols[1]
			fmt.Print("\n", col1.String(xlFile), ",", col2.String(xlFile))
		}
	}
}
