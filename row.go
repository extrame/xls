package xls

type RowInfo struct {
	Index    uint16
	Fcell    uint16
	Lcell    uint16
	Height   uint16
	Notused  uint16
	Notused2 uint16
	Flags    uint32
}

type Row struct {
	info *RowInfo
	Cols map[uint16]contentHandler
}
