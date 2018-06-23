package xls

import (
	"encoding/binary"
	"io"
	"unicode/utf16"
)

type boundsheet struct {
	Filepos uint32
	Type    byte
	Visible byte
	Name    byte
}

type extSheetRef struct {
	Num  uint16
	Info []ExtSheetInfo
}

// ExtSheetInfo external sheet references provided for named cells
type ExtSheetInfo struct {
	ExternalBookIndex uint16
	FirstSheetIndex   uint16
	LastSheetIndex    uint16
}

//WorkSheet in one WorkBook
type WorkSheet struct {
	bs   *boundsheet
	wb   *WorkBook
	Name string
	rows map[uint16]*Row
	//NOTICE: this is the max row number of the sheet, so it should be count -1
	MaxRow uint16
	id     int
	parsed bool
}

func (w *WorkSheet) Row(i int) *Row {
	row := w.rows[uint16(i)]
	if row != nil {
		row.wb = w.wb
	}
	return row
}

func (w *WorkSheet) parse(buf io.ReadSeeker) {
	w.rows = make(map[uint16]*Row)
	b := new(bof)
	var bp *bof
	for {
		if err := binary.Read(buf, binary.LittleEndian, b); err == nil {
			bp = w.parseBof(buf, b, bp)
			if b.Id == 0xa {
				break
			}
		} else {
			break
		}
	}
	w.parsed = true
}

func (w *WorkSheet) parseBof(buf io.ReadSeeker, b *bof, pre *bof) *bof {
	var col interface{}
	switch b.Id {
	// case 0x0E5: //MERGEDCELLS
	// ws.mergedCells(buf)
	case 0x208: //ROW
		r := new(rowInfo)
		binary.Read(buf, binary.LittleEndian, r)
		w.addRow(r)
	case 0x0BD: //MULRK
		mc := new(MulrkCol)
		size := (b.Size - 6) / 6
		mc.Xfrks = make([]XfRk, size)
		binary.Read(buf, binary.LittleEndian, &mc.Col)
		binary.Read(buf, binary.LittleEndian, &mc.Xfrks)
		binary.Read(buf, binary.LittleEndian, &mc.LastColB)
		col = mc
	case 0x0BE: //MULBLANK
		mc := new(MulBlankCol)
		size := (b.Size - 6) / 2
		mc.Xfs = make([]uint16, size)
		binary.Read(buf, binary.LittleEndian, &mc.Col)
		binary.Read(buf, binary.LittleEndian, &mc.Xfs)
		binary.Read(buf, binary.LittleEndian, &mc.LastColB)
		col = mc
	case 0x203: //NUMBER
		col = new(NumberCol)
		binary.Read(buf, binary.LittleEndian, col)
	case 0x06: //FORMULA
		c := new(FormulaCol)
		c.ws = w.id
		c.Header = new(FormulaColHeader)
		c.Bts = make([]byte, b.Size-20)
		binary.Read(buf, binary.LittleEndian, c.Header)
		binary.Read(buf, binary.LittleEndian, &c.Bts)
		col = c
		c.parse(w.wb, false)

		if TYPE_STRING == c.vType {
			binary.Read(buf, binary.LittleEndian, &c.Code)
			binary.Read(buf, binary.LittleEndian, &c.Btl)
			binary.Read(buf, binary.LittleEndian, &c.Btc)

			var fms, fme = w.wb.parseString(buf, c.Btc)
			if nil == fme {
				c.value = fms
			}

			buf.Seek(-int64(c.Btl+4), 1)
		}
	case 0x27e: //RK
		col = new(RkCol)
		binary.Read(buf, binary.LittleEndian, col)
	case 0xFD: //LABELSST
		col = new(LabelsstCol)
		binary.Read(buf, binary.LittleEndian, col)
	case 0x204:
		c := new(labelCol)
		binary.Read(buf, binary.LittleEndian, &c.BlankCol)
		var count uint16
		binary.Read(buf, binary.LittleEndian, &count)
		c.Str, _ = w.wb.parseString(buf, count)
		col = c
	case 0x201: //BLANK
		col = new(BlankCol)
		binary.Read(buf, binary.LittleEndian, col)
	case 0x1b8: //HYPERLINK
		var flag uint32
		var count uint32
		var hy HyperLink
		binary.Read(buf, binary.LittleEndian, &hy.CellRange)
		buf.Seek(20, 1)
		binary.Read(buf, binary.LittleEndian, &flag)

		if flag&0x14 != 0 {
			binary.Read(buf, binary.LittleEndian, &count)
			hy.Description = b.utf16String(buf, count)
		}
		if flag&0x80 != 0 {
			binary.Read(buf, binary.LittleEndian, &count)
			hy.TargetFrame = b.utf16String(buf, count)
		}
		if flag&0x1 != 0 {
			var guid [2]uint64
			binary.Read(buf, binary.BigEndian, &guid)
			if guid[0] == 0xE0C9EA79F9BACE11 && guid[1] == 0x8C8200AA004BA90B { //URL
				hy.IsUrl = true
				binary.Read(buf, binary.LittleEndian, &count)
				hy.Url = b.utf16String(buf, count/2)
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
					hy.ExtendedFilePath = b.utf16String(buf, count/2+1)
				}
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
	case 0x809:
		buf.Seek(int64(b.Size), 1)
	case 0xa:
	default:
		// log.Printf("Unknow %X,%d\n", b.Id, b.Size)
		buf.Seek(int64(b.Size), 1)
	}
	if col != nil {
		w.add(col)
	}
	return b
}

func (w *WorkSheet) add(content interface{}) {
	if ch, ok := content.(contentHandler); ok {
		if col, ok := content.(Coler); ok {
			w.addCell(col, ch)
		}
	}
}

func (w *WorkSheet) addCell(col Coler, ch contentHandler) {
	w.addContent(col.Row(), ch)
}

func (w *WorkSheet) addRange(rang Ranger, ch contentHandler) {
	for i := rang.FirstRow(); i <= rang.LastRow(); i++ {
		w.addContent(i, ch)
	}
}

func (w *WorkSheet) addContent(num uint16, ch contentHandler) {
	var row *Row
	var ok bool
	if row, ok = w.rows[num]; !ok {
		info := new(rowInfo)
		info.Index = num
		row = w.addRow(info)
	}
	row.cols[ch.FirstCol()] = ch
}

func (w *WorkSheet) addRow(info *rowInfo) *Row {
	var ok bool
	var row *Row

	if info.Index > w.MaxRow {
		w.MaxRow = info.Index
	}

	if row, ok = w.rows[info.Index]; ok {
		row.info = info
	} else {
		row = &Row{info: info, cols: make(map[uint16]contentHandler, int(info.Last-info.First))}
		w.rows[info.Index] = row
	}

	return row
}
