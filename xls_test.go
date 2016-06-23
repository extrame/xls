package xls

import (
	"fmt"
	"testing"
	"unicode/utf16"
)

func TestOpen(t *testing.T) {
	if xlFile, err := Open("Book1.xls", "utf-8"); err == nil {
		if sheet1 := xlFile.GetSheet(0); sheet1 != nil {
			fmt.Println("Total Lines ", sheet1.MaxRow, sheet1.Name)
			for i := 0; i < int(sheet1.MaxRow); i++ {
				fmt.Printf("row %v point %v \n", i, sheet1.Rows[uint16(i)])
				if sheet1.Rows[uint16(i)] == nil {
					continue
				}
				row := sheet1.Rows[uint16(i)]
				for n, col := range row.Cols {
					fmt.Println(n, "==>", col.String(xlFile), " ")
				}
				// col1 := .Cols[0]
				// col2 := sheet1.Rows[uint16(i)].Cols[1]
				// fmt.Printf("\ncol1 %v \nCol2 %v \n", col1.String(xlFile), col2.String(xlFile))
			}
		}
	}
}

func TestEuropeString(t *testing.T) {
	bts := []byte{66, 233, 114, 232}
	var bts1 = make([]uint16, 4)
	for k, v := range bts {
		bts1[k] = uint16(v)
	}
	runes := utf16.Decode(bts1)
	fmt.Println(string(runes))
}

// func TestOpen1(t *testing.T) {
// 	xlFile, _ := Open("000.xls", "")
// 	for i := 0; i < xlFile.NumSheets(); i++ {
// 		fmt.Println(xlFile.GetSheet(i).Name)
// 		sheet := xlFile.GetSheet(i)
// 		row := sheet.Rows[1]
// 		for i, col := range row.Cols {
// 			fmt.Println(i, col.String(xlFile))
// 		}
// 	}
// 	// sheet1 := xlFile.GetSheet(0)
// 	// fmt.Println(sheet1.Name)
// 	// fmt.Print(sheet1.Rows)
// 	// for k, row1 := range sheet1.Rows {
// 	// 	// row1 := sheet1.Rows[1]
// 	// 	fmt.Printf("\n[%d]", k)
// 	// 	for _, col1 := range row1.Cols {
// 	// 		// col1 := row1.Cols[0]
// 	// 		fmt.Print(col1.LastCol())
// 	// 		fmt.Print(" ")
// 	// 	}
// 	// }
// }

// func TestBof(t *testing.T) {
// 	b := new(bof)
// 	b.Id = 0x41E
// 	b.Size = 55
// 	buf := bytes.NewReader([]byte{0x07, 0x00, 0x19, 0x00, 0x01, 0x22, 0x00, 0xE5, 0xFF, 0x22, 0x00, 0x23, 0x00, 0x2C, 0x00, 0x23, 0x00, 0x23, 0x00, 0x30, 0x00, 0x2E, 0x00, 0x30, 0x00, 0x30, 0x00, 0x3B, 0x00, 0x22, 0x00, 0xE5, 0xFF, 0x22, 0x00, 0x5C, 0x00, 0x2D, 0x00, 0x23, 0x00, 0x2C, 0x20, 0x00})
// 	wb := new(WorkBook)
// 	wb.Formats = make(map[uint16]*Format)
// 	wb.parseBof(buf, b, b, 0)
// }

// func TestMaxRow(t *testing.T) {
// 	xlFile, err := Open("Table.xls", "utf-8")
// 	if err != nil {
// 		fmt.Fprintf(os.Stderr, "Failure: %v\n", err)
// 		t.Error(err)
// 	}

// 	if sheet1 := xlFile.GetSheet(0); sheet1 != nil {
// 		if sheet1.MaxRow != 11 {
// 			t.Errorf("max row is error,is %d instead of 11", sheet1.MaxRow)
// 		}
// 	}
// }
