package xls

import (
	"fmt"
	"testing"
	"time"
)

func TestBigTable(t *testing.T) {
	xlFile, err := Open("BigTable.xls", "utf-8")
	if err != nil {
		t.Fatalf("Cant open xls file: %s", err)
	}

	sheet := xlFile.GetSheet(0)
	if sheet == nil {
		t.Fatal("Cant get sheet")
	}

	cnt1 := 1
	cnt2 := 10000
	cnt3 := 20000
	date1, _ := time.Parse("2006-01-02", "2015-01-01")
	date2, _ := time.Parse("2006-01-02", "2016-01-01")
	date3, _ := time.Parse("2006-01-02", "2017-01-01")

	for i := 1; i <= 4999; i++ {
		row := sheet.Row(i)
		if row == nil {
			continue
		}

		col2sample := fmt.Sprintf("%d от %s", cnt1, date1.Format("02.01.2006"))
		col5sample := fmt.Sprintf("%d от %s", cnt2, date2.Format("02.01.2006"))
		col8sample := fmt.Sprintf("%d от %s", cnt3, date3.Format("02.01.2006"))

		col2 := row.Col(2)
		col5 := row.Col(5)
		col8 := row.Col(8)

		if col2 != col2sample {
			t.Fatalf("Row %d: col 2 val not eq base value: %s != %s", i, col2, col2sample)
		}
		if col5 != col5sample {
			t.Fatalf("Row %d: col 5 val not eq base value: %s != %s", i, col5, col5sample)
		}
		if col8 != col8sample {
			t.Fatalf("Row %d: col 8 val not eq base value: %s != %s", i, col8, col8sample)
		}

		cnt1++
		cnt2++
		cnt3++
		date1 = date1.AddDate(0, 0, 1)
		date2 = date2.AddDate(0, 0, 1)
		date3 = date3.AddDate(0, 0, 1)

	}
}
