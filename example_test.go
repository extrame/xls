package xls

import (
	"fmt"
)

func ExampleOpen() {
	if xlFile, err := Open("Table.xls", "utf-8"); err == nil {
		fmt.Println(xlFile.Author)
	}
}

func ExampleWorkBook_NumberSheets() {
	if xlFile, err := Open("Table.xls", "utf-8"); err == nil {
		for i := 0; i < xlFile.NumSheets(); i++ {
			sheet := xlFile.GetSheet(i)
			fmt.Println(sheet.Name)
		}
	}
}

//Output: read the content of first two cols in each row
func ExampleWorkBook_GetSheet() {
	if xlFile, err := Open("Table.xls", "utf-8"); err == nil {
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
}
