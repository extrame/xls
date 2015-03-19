package xls

type FormatB struct {
	Index uint16
	Size  uint16
}

type Format struct {
	b   *FormatB
	str string
}
