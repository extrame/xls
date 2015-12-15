# xls

[![GoDoc](https://godoc.org/github.com/extrame/xls?status.svg)](https://godoc.org/github.com/extrame/xls)

Pure Golang xls library writen by [MinkTech(chinese)](http://www.mink-tech.com). 

Thanks for contributions from Tamás Gulácsi. 

**English User please mailto** [Liu Ming](mailto:liuming@mink-tech.com)

This is a xls library writen in pure Golang. Almostly it is translated from the libxls library in c.

It has just the reading function without the format.

# Basic Usage

* Use **Open** function for open file
* Use **OpenReader** function for open xls from a reader

These methods will open a workbook object for reading, like

	func (w *WorkBook) ReadAllCells() (res [][]string) {
		for _, sheet := range w.Sheets {
			w.PrepareSheet(sheet)
			if sheet.MaxRow != 0 {
				temp := make([][]string, sheet.MaxRow+1)
				for k, row := range sheet.Rows {
					data := make([]string, 0)
					if len(row.Cols) > 0 {
						for _, col := range row.Cols {
							if uint16(len(data)) <= col.LastCol() {
								data = append(data, make([]string, col.LastCol()-uint16(len(data))+1)...)
							}
							str := col.String(w)
							for i := uint16(0); i < col.LastCol()-col.FirstCol()+1; i++ {
								data[col.FirstCol()+i] = str[i]
							}
						}
						temp[k] = data
					}
				}
				res = append(res, temp...)
			}
		}
		return
	}

