package xls

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"math"
	"strconv"
)

//Compares xls and xlsx files
func compareXlsXlsx(xlsfilepathname string, xlsxfilepathname string) string {
	xlsFile, err := Open(xlsfilepathname, "utf-8")
	if err != nil {
		return fmt.Sprintf("Cant open xls file: %s", err)
	}

	xlsxFile, err := xlsx.OpenFile(xlsxfilepathname)
	if err != nil {
		return fmt.Sprintf("Cant open xlsx file: %s", err)
	}

	for sheet, xlsxSheet := range xlsxFile.Sheets {
		xlsSheet := xlsFile.GetSheet(sheet)
		if xlsSheet == nil {
			return fmt.Sprintf("Cant get xls sheet")
		}
		for row, xlsxRow := range xlsxSheet.Rows {
			xlsRow := xlsSheet.Row(row)
			for cell, xlsxCell := range xlsxRow.Cells {
				xlsText := xlsRow.Col(cell)
				xlsxText := xlsxCell.String()
				if xlsText != xlsxText {
					//try to convert to numbers
					xlsFloat, xlsErr := strconv.ParseFloat(xlsText, 64)
					xlsxFloat, xlsxErr := strconv.ParseFloat(xlsxText, 64)
					//check if numbers have no significant difference
					if xlsErr == nil && xlsxErr == nil {
						diff := math.Abs(xlsFloat - xlsxFloat)
						if diff > 0.0000001 {
							return fmt.Sprintf("sheet:%d, row/col: %d/%d, xlsx: (%s)[%d], xls: (%s)[%d], numbers difference: %f.",
								sheet, row, cell, xlsxText, len(xlsxText),
								xlsText, len(xlsText), diff)
						}
					} else {
						return fmt.Sprintf("sheet:%d, row/col: %d/%d, xlsx: (%s)[%d], xls: (%s)[%d].",
							sheet, row, cell, xlsxText, len(xlsxText),
							xlsText, len(xlsText))
					}
				}
			}
		}
	}

	return ""
}
