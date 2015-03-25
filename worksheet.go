package xls

import (
	"encoding/binary"
	"fmt"
	"io"
	"log"
	"unicode/utf16"
)

type Boundsheet struct {
	Filepos uint32
	Type    byte
	Visible byte
	Name    byte
}

type WorkSheet struct {
	bs     *Boundsheet
	wb     *WorkBook
	Name   string
	Rows   map[uint16]*Row
	MaxRow uint16
}

func (w *WorkSheet) Parse(buf io.ReadSeeker) {
	w.Rows = make(map[uint16]*Row)
	bof := new(BOF)
	var bof_pre *BOF
	for {
		if err := binary.Read(buf, binary.LittleEndian, bof); err == nil {
			bof_pre = w.parseBof(buf, bof, bof_pre)
		} else {
			fmt.Println(err)
			break
		}
	}
}

func (w *WorkSheet) parseBof(buf io.ReadSeeker, bof *BOF, pre *BOF) *BOF {
	var col interface{}
	switch bof.Id {
	// case 0x0E5: //MERGEDCELLS
	// ws.mergedCells(buf)
	case 0x208: //ROW
		r := new(RowInfo)
		binary.Read(buf, binary.LittleEndian, r)
		w.addRow(r)
	case 0x0BD: //MULRK
		mc := new(MulrkCol)
		size := (bof.Size - 6) / 6
		binary.Read(buf, binary.LittleEndian, &mc.Col)
		mc.Xfrks = make([]XfRk, size)
		for i := uint16(0); i < size; i++ {
			binary.Read(buf, binary.LittleEndian, &mc.Xfrks[i])
		}
		binary.Read(buf, binary.LittleEndian, &mc.LastColB)
		col = mc
	case 0x0BE: //MULBLANK
		mc := new(MulBlankCol)
		size := (bof.Size - 6) / 2
		binary.Read(buf, binary.LittleEndian, &mc.Col)
		mc.Xfs = make([]uint16, size)
		for i := uint16(0); i < size; i++ {
			binary.Read(buf, binary.LittleEndian, &mc.Xfs[i])
		}
		binary.Read(buf, binary.LittleEndian, &mc.LastColB)
		col = mc
	case 0x203: //NUMBER
		col = new(NumberCol)
		binary.Read(buf, binary.LittleEndian, col)
	case 0x06: //FORMULA
		c := new(FormulaCol)
		binary.Read(buf, binary.LittleEndian, &c.Header)
		c.Bts = make([]byte, bof.Size-20)
		binary.Read(buf, binary.LittleEndian, &c.Bts)
		col = c
	case 0x27e: //RK
		col = new(RkCol)
		binary.Read(buf, binary.LittleEndian, col)
	case 0xFD: //LABELSST
		col = new(LabelsstCol)
		binary.Read(buf, binary.LittleEndian, col)
	case 0x201: //BLANK
		col = new(BlankCol)
		binary.Read(buf, binary.LittleEndian, col)
	case 0x1b8: //HYPERLINK
		var hy HyperLink
		binary.Read(buf, binary.LittleEndian, &hy.CellRange)
		buf.Seek(20, 1)
		var flag uint32
		binary.Read(buf, binary.LittleEndian, &flag)
		var count uint32

		if flag&0x14 != 0 {
			binary.Read(buf, binary.LittleEndian, &count)
			hy.Description = bof.Utf16String(buf, count)
		}
		if flag&0x80 != 0 {
			binary.Read(buf, binary.LittleEndian, &count)
			hy.TargetFrame = bof.Utf16String(buf, count)
		}
		if flag&0x1 != 0 {
			var guid [2]uint64
			binary.Read(buf, binary.BigEndian, &guid)
			if guid[0] == 0xE0C9EA79F9BACE11 && guid[1] == 0x8C8200AA004BA90B { //URL
				hy.IsUrl = true
				binary.Read(buf, binary.LittleEndian, &count)
				hy.Url = bof.Utf16String(buf, count/2)
			} else if guid[0] == 0x303000000000000 && guid[1] == 0xC000000000000046 { //URL{
				var upCount uint16
				binary.Read(buf, binary.LittleEndian, &upCount)
				binary.Read(buf, binary.LittleEndian, &count)
				bts := make([]byte, count)
				binary.Read(buf, binary.LittleEndian, &bts)
				hy.ShortedFilePath = string(bts)
				buf.Seek(24, 1)
				binary.Read(buf, binary.LittleEndian, &count)
				if count > 0 {
					binary.Read(buf, binary.LittleEndian, &count)
					buf.Seek(2, 1)
					hy.ExtendedFilePath = bof.Utf16String(buf, count/2+1)
				}
				log.Println(hy)
			}
		}
		if flag&0x8 != 0 {
			binary.Read(buf, binary.LittleEndian, &count)
			var bts = make([]uint16, count)
			binary.Read(buf, binary.LittleEndian, &bts)
			runes := utf16.Decode(bts[:len(bts)-1])
			hy.TextMark = string(runes)
		}

		w.addRange(&hy.CellRange, &hy)
	default:
		fmt.Printf("Unknow %X,%d\n", bof.Id, bof.Size)
		buf.Seek(int64(bof.Size), 1)
	}
	if col != nil {
		w.add(col)
	}
	return bof
}

func (w *WorkSheet) add(content interface{}) {
	if ch, ok := content.(ContentHandler); ok {
		if col, ok := content.(Coler); ok {
			w.addCell(col, ch)
		}
	}

}

func (w *WorkSheet) addCell(col Coler, ch ContentHandler) {
	var row *Row
	var ok bool
	if row, ok = w.Rows[col.Row()]; !ok {
		info := new(RowInfo)
		info.Index = col.Row()
		row = w.addRow(info)
	}
	row.Cols[ch.FirstCol()] = ch
}

func (w *WorkSheet) addRange(rang Ranger, ch ContentHandler) {
	var row *Row
	var ok bool
	for i := rang.FirstRow(); i <= rang.LastRow(); i++ {
		if row, ok = w.Rows[i]; !ok {
			info := new(RowInfo)
			info.Index = i
			row = w.addRow(info)
		}
		row.Cols[ch.FirstCol()] = ch
	}
}

func (w *WorkSheet) addRow(info *RowInfo) (row *Row) {
	if info.Index > w.MaxRow {
		w.MaxRow = info.Index
	}
	var ok bool
	if row, ok = w.Rows[info.Index]; ok {
		row.info = info
	} else {
		row = &Row{info: info, Cols: make(map[uint16]ContentHandler)}
		w.Rows[info.Index] = row
	}
	return
}
