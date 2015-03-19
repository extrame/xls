package xls

import ()

type BOF struct {
	Id   uint16
	Size uint16
}

type BIFFHeader struct {
	Ver     uint16
	Type    uint16
	Id_make uint16
	Year    uint16
	Flags   uint32
	Min_ver uint32
}

// func parseBofsForWb(bts []byte, wb *WorkBook) {
// 	bof := new(BOF)
// 	var bof_pre *BOF
// 	buf := bytes.NewReader(bts)
// 	for {
// 		if err := binary.Read(buf, binary.LittleEndian, bof); err == nil {
// 			bof_pre = bof.ActForWb(buf, wb, bof_pre)
// 		} else {
// 			break
// 		}
// 	}
// }
