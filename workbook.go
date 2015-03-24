package xls

import (
	"bytes"
	"encoding/binary"
	"io"
	"unicode/utf16"
)

type WorkBook struct {
	Is5ver   bool
	Type     uint16
	Codepage uint16
	Xfs      []st_xf_data
	Fonts    []Font
	Formats  map[uint16]*Format
	Sheets   []*WorkSheet
	Author   string
	rs       io.ReadSeeker
	sst      []string
}

func newWookBookFromOle2(rs io.ReadSeeker) *WorkBook {
	wb := new(WorkBook)
	wb.Formats = make(map[uint16]*Format)
	// wb.bts = bts
	wb.rs = rs
	wb.Sheets = make([]*WorkSheet, 0)
	wb.Parse(rs)
	return wb
}

func (w *WorkBook) Parse(buf io.ReadSeeker) {
	bof := new(BOF)
	var bof_pre *BOF
	// buf := bytes.NewReader(bts)
	offset := 0
	for {
		if err := binary.Read(buf, binary.LittleEndian, bof); err == nil {
			bof_pre, offset = w.parseBof(buf, bof, bof_pre, offset)
		} else {
			break
		}
	}
}

func (w *WorkBook) addXf(xf st_xf_data) {
	w.Xfs = append(w.Xfs, xf)
}

func (w *WorkBook) addFont(font *FontInfo, buf io.ReadSeeker) {
	name := w.get_string(buf, uint16(font.NameB))
	w.Fonts = append(w.Fonts, Font{Info: font, Name: name})
}

func (w *WorkBook) addFormat(format *FormatB, buf io.ReadSeeker) {
	w.Formats[format.Index] = &Format{b: format, str: w.get_string(buf, uint16(format.Size))}
}

func (wb *WorkBook) parseBof(buf io.ReadSeeker, b *BOF, pre *BOF, offset_pre int) (after *BOF, offset int) {
	after = b
	var bts = make([]byte, b.Size)
	binary.Read(buf, binary.LittleEndian, bts)
	buf_item := bytes.NewReader(bts)
	switch b.Id {
	case 0x809:
		bif := new(BIFFHeader)
		binary.Read(buf_item, binary.LittleEndian, bif)
		if bif.Ver != 0x600 {
			wb.Is5ver = true
		}
		wb.Type = bif.Type
	case 0x042: // CODEPAGE
		binary.Read(buf_item, binary.LittleEndian, &wb.Codepage)
	case 0x3c: // CONTINUE
		var bts = make([]byte, b.Size)
		binary.Read(buf_item, binary.LittleEndian, bts)
		buf_item := bytes.NewReader(bts)
		if pre.Id == 0xfc {
			var size uint16
			if err := binary.Read(buf_item, binary.LittleEndian, &size); err == nil {
				wb.sst[offset_pre] = wb.get_string(buf_item, size)
				offset_pre++
			}
		}
		offset = offset_pre
		after = pre
	case 0xfc: // SST
		info := new(SstInfo)
		binary.Read(buf_item, binary.LittleEndian, info)
		wb.sst = make([]string, info.Count)
		for i := 0; i < int(info.Count); i++ {
			var size uint16
			if err := binary.Read(buf_item, binary.LittleEndian, &size); err == nil || err == io.EOF {
				wb.sst[i] = wb.sst[i] + wb.get_string(buf_item, size)
				if err == io.EOF {
					offset = i
					break
				}
			}
		}
	case 0x85: // BOUNDSHEET
		var bs = new(Boundsheet)
		binary.Read(buf_item, binary.LittleEndian, bs)
		// different for BIFF5 and BIFF8
		wb.addSheet(bs, buf_item)
	case 0x0e0: // XF
		if wb.Is5ver {
			xf := new(Xf5)
			binary.Read(buf_item, binary.LittleEndian, xf)
			wb.addXf(xf)
		} else {
			xf := new(Xf8)
			binary.Read(buf_item, binary.LittleEndian, xf)
			wb.addXf(xf)
		}
	case 0x031: // FONT
		f := new(FontInfo)
		binary.Read(buf_item, binary.LittleEndian, f)
		wb.addFont(f, buf_item)
	case 0x41E: //FORMAT
		// var bts = make([]byte, b.Size)
		// binary.Read(buf, binary.LittleEndian, bts)
		// buf_item := bytes.NewReader(bts)
		f := new(FormatB)
		binary.Read(buf_item, binary.LittleEndian, f)
		wb.addFormat(f, buf_item)
		// case 0x5c:
		// var bts = make([]byte, b.Size)
		// binary.Read(buf_item, binary.LittleEndian, bts)
		// if wb.Is5ver {
		// 	wb.Author = wb.get_string_from_bytes(bts[1:], uint16(bts[1]))
		// } else {
		// 	size := binary.LittleEndian.Uint16(bts)
		// 	wb.Author = wb.get_string_from_bytes(bts[2:], size)
		// }
	}
	return
}

func (w *WorkBook) get_string(buf io.ReadSeeker, size uint16) string {
	var res string
	if w.Is5ver {
		var bts = make([]byte, size)
		buf.Read(bts)
		return string(bts)
	} else {
		var richtext_num uint16
		var phonetic_size uint32
		var flag byte
		binary.Read(buf, binary.LittleEndian, &flag)
		if flag&0x8 != 0 {
			binary.Read(buf, binary.LittleEndian, &richtext_num)
		}
		if flag&0x4 != 0 {
			binary.Read(buf, binary.LittleEndian, &phonetic_size)
		}
		if flag&0x1 != 0 {
			var bts = make([]uint16, size)
			binary.Read(buf, binary.LittleEndian, &bts)
			runes := utf16.Decode(bts)
			res = string(runes)
		} else {
			var bts = make([]byte, size)
			binary.Read(buf, binary.LittleEndian, &bts)
			res = string(bts)
		}
		if flag&0x8 != 0 {
			var bts []byte
			if w.Is5ver {
				bts = make([]byte, 2*richtext_num)
			} else {
				bts = make([]byte, 4*richtext_num)
			}
			binary.Read(buf, binary.LittleEndian, bts)
		}
		if flag&0x4 != 0 {
			var bts []byte
			bts = make([]byte, phonetic_size)
			binary.Read(buf, binary.LittleEndian, bts)
		}
	}
	return res
}

func (w *WorkBook) get_string_from_bytes(bts []byte, size uint16) string {
	buf := bytes.NewReader(bts)
	return w.get_string(buf, size)
}

func (w *WorkBook) addSheet(sheet *Boundsheet, buf io.ReadSeeker) {
	name := w.get_string(buf, uint16(sheet.Name))
	w.Sheets = append(w.Sheets, &WorkSheet{bs: sheet, Name: name, wb: w})
}

func (w *WorkBook) PrepareSheet(sheet *WorkSheet) {
	w.rs.Seek(int64(sheet.bs.Filepos), 0)
	sheet.Parse(w.rs)
}

func (w *WorkBook) ReadAllCells(max int) (res [][]string) {
	res = make([][]string, 0)
	for _, sheet := range w.Sheets {
		if len(res) < max {
			max = max - len(res)
			w.PrepareSheet(sheet)
			if sheet.MaxRow != 0 {
				leng := int(sheet.MaxRow) + 1
				if max < leng {
					leng = max
				}
				temp := make([][]string, leng)
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
						if leng > int(k) {
							temp[k] = data
						}
					}
				}
				res = append(res, temp...)
			}
		}
	}
	return
}
