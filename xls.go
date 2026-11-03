package xls

import (
	"errors"
	"io"
	"os"

	"github.com/extrame/ole2"
)

//Open one xls file. The caller must Close the returned workbook.
func Open(file string, charset string) (*WorkBook, error) {
	fi, err := os.Open(file)
	if err != nil {
		return nil, err
	}
	wb, err := OpenReader(fi, charset)
	if err != nil {
		fi.Close()
		return nil, err
	}
	wb.closer = fi
	return wb, nil
}

//Open one xls file and return the closer
func OpenWithCloser(file string, charset string) (*WorkBook, io.Closer, error) {
	if fi, err := os.Open(file); err == nil {
		wb, err := OpenReader(fi, charset)
		return wb, fi, err
	} else {
		return nil, nil, err
	}
}

//Open xls file from reader
func OpenReader(reader io.ReadSeeker, charset string) (wb *WorkBook, err error) {
	var ole *ole2.Ole
	if ole, err = ole2.Open(reader, charset); err == nil {
		var dir []*ole2.File
		if dir, err = ole.ListDir(); err == nil {
			var book *ole2.File
			var root *ole2.File
			for _, file := range dir {
				name := file.Name()
				if name == "Workbook" {
					if book == nil {
						book = file
					}
					//book = file
					// break
				}
				if name == "Book" {
					book = file
					// break
				}
				if name == "Root Entry" {
					root = file
				}
			}
			if book != nil {
				wb = newWorkBookFromOle2(ole.OpenFile(book, root))
				return
			}
			err = errors.New("xls: Workbook stream not found")
		}
	}
	return
}
