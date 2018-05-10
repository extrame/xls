package xls

import (
	formatter "github.com/extrame/xls/format"
)

type format struct {
	Head struct {
		Index uint16
		Size  uint16
	}
	str string
}

func (f *format) Format(val float64, date1904 bool) string {
	_, tokens := formatter.Lexer(f.str)
	ds := formatter.Parse(tokens)
	return ds.Format(val, date1904)
}
