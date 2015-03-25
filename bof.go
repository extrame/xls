package xls

import (
	"bytes"
	"encoding/binary"
	"io"
	"unicode/utf16"
)

type BOF struct {
	Id   uint16
	Size uint16
}

func (b *BOF) Reader(buf io.ReadSeeker) io.ReadSeeker {
	rts := make([]byte, b.Size)
	buf.Read(rts)
	return bytes.NewReader(rts)
}

func (b *BOF) Utf16String(buf io.ReadSeeker, count uint32) string {
	var bts = make([]uint16, count)
	binary.Read(buf, binary.LittleEndian, &bts)
	runes := utf16.Decode(bts[:len(bts)-1])
	return string(runes)
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
