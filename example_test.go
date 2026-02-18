package xls

import (
	"fmt"
)

func ExampleOpen() {
	if xlFile, err := Open("Table.xls", "utf-8"); err == nil {
		fmt.Println(xlFile.Author)
	}
}

func ExampleWorkBook_NumSheets() {
	if xlFile, closer, err := OpenWithCloser("Table.xls", "utf-8"); err == nil {
		defer closer.Close()
		for i := 0; i < xlFile.NumSheets(); i++ {
			sheet := xlFile.GetSheet(i)
			fmt.Println(sheet.Name)
		}
	}
}

//Output: read the content of first two cols in each row
func ExampleWorkBook_GetSheet() {
	if xlFile, closer, err := OpenWithCloser("Table.xls", "utf-8"); err == nil {
		defer closer.Close()
		if sheet1 := xlFile.GetSheet(0); sheet1 != nil {
			fmt.Print("Total Lines ", sheet1.MaxRow, sheet1.Name)
			col1 := sheet1.Row(0).Col(0)
			col2 := sheet1.Row(0).Col(0)
			for i := 0; i <= (int(sheet1.MaxRow)); i++ {
				row1 := sheet1.Row(i)
				if row1 == nil {
					continue
				}
				col1 = row1.Col(0)
				col2 = row1.Col(1)
				fmt.Print("\n", col1, ",", col2)
			}
		}
	}
}
