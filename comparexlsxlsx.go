package xls

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"path"
)

//Compares xls and xlsx files
func compareXlsXlsx(filepathname string) string {
	xlsFile, err := Open(path.Join("testdata", filepathname)+".xls", "utf-8")
	if err != nil {
		return fmt.Sprintf("Cant open xls file: %s", err)
	}

	xlsxFile, err := xlsx.OpenFile(path.Join("testdata", filepathname) + ".xlsx")
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
					return fmt.Sprintf("Sheet: %d, row: %d, col: %d, xlsx: (%s)[%d], xls: (%s)[%d].",
						sheet, row, cell, xlsxText, len(xlsxText), xlsText, len(xlsText))
				}
			}
		}
	}

	return ""
}
