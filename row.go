package xls

type rowInfo struct {
	Index    uint16
	Fcell    uint16
	Lcell    uint16
	Height   uint16
	Notused  uint16
	Notused2 uint16
	Flags    uint32
}

//Row the data of one row
type Row struct {
	wb   *WorkBook
	info *rowInfo
	cols map[uint16]contentHandler
}

//Col Get the Nth Col from the Row, if has not, return nil.
//Suggest use Has function to test it.
func (r *Row) Col(i int) string {
	serial := uint16(i)
	if ch, ok := r.cols[serial]; ok {
		// Some records (e.g. an empty MULBLANK run whose last column is before
		// its first) hold no cells at all; that is a blank cell, not a crash.
		if strs := ch.String(r.wb); len(strs) > 0 {
			return strs[0]
		}
		return ""
	}
	for _, v := range r.cols {
		if v.FirstCol() <= serial && v.LastCol() >= serial {
			strs := v.String(r.wb)
			if index := int(serial) - int(v.FirstCol()); index < len(strs) {
				return strs[index]
			}
			return ""
		}
	}
	return ""
}

//ColExact Get the Nth Col from the Row, if has not, return nil.
//For merged cells value is returned for first cell only
func (r *Row) ColExact(i int) string {
	serial := uint16(i)
	if ch, ok := r.cols[serial]; ok {
		strs := ch.String(r.wb)
		return strs[0]
	}
	return ""
}

//LastCol Get the number of Last Col of the Row.
func (r *Row) LastCol() int {
	return int(r.info.Lcell)
}

//FirstCol Get the number of First Col of the Row.
func (r *Row) FirstCol() int {
	return int(r.info.Fcell)
}
