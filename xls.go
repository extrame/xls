package xls

import (
	"github.com/extrame/ole2"
	"io"
	"io/ioutil"
)

//Open one xls file
func Open(file string, charset string) (*WorkBook, error) {
	if bts, err := ioutil.ReadFile(file); err == nil {
		return parse(bts, charset)
	} else {
		return nil, err
	}

}

//Open xls file from reader
func OpenReader(reader io.ReadCloser, charset string) (*WorkBook, error) {
	if bts, err := ioutil.ReadAll(reader); err == nil {
		return parse(bts, charset)
	} else {
		return nil, err
	}
}

func parse(bts []byte, charset string) (wb *WorkBook, err error) {
	var ole *ole2.Ole
	if ole, err = ole2.Open(bts, charset); err == nil {
		var dir []*ole2.File
		if dir, err = ole.ListDir(); err == nil {
			var book *ole2.File
			for _, file := range dir {
				name := file.Name()
				if name == "Workbook" {
					book = file
					// break
				}
				if name == "Book" {
					book = file
					// break
				}
			}
			if book != nil {
				wb = newWorkBookFromOle2(ole.OpenFile(book))
				return
			}
		}
	}
	return
}
