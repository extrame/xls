package xls

import (
	"github.com/extrame/ole2"
	"io"
	"io/ioutil"
)

func Open(file string, charset string) (*WorkBook, error) {

	if bts, err := ioutil.ReadFile(file); err == nil {
		return parse(bts, charset)
	} else {
		return nil, err
	}

}

func OpenReader(reader io.ReadCloser, charset string) (*WorkBook, error) {
	bts, _ := ioutil.ReadAll(reader)
	return parse(bts, charset)
}

func parse(bts []byte, charset string) (*WorkBook, error) {
	ole, _ := ole2.Open(bts, charset)
	dir, err := ole.ListDir()
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
		wb := newWookBookFromOle2(ole.OpenFile(book))
		return wb, nil
	}
	return nil, err
}
