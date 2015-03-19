package xls

import (
	"encoding/binary"
	"fmt"
	"io"
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
	var col Coler
	switch bof.Id {
	case 0x0E5: //MERGEDCELLS
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
		col = new(FormulaCol)
		binary.Read(buf, binary.LittleEndian, col)
	case 0x27e: //RK
		col = new(RkCol)
		binary.Read(buf, binary.LittleEndian, col)
	case 0xFD: //LABELSST
		col = new(LabelsstCol)
		binary.Read(buf, binary.LittleEndian, col)
	case 0x201: //BLANK
		col = new(BlankCol)
		binary.Read(buf, binary.LittleEndian, col)
	default:
		buf.Seek(int64(bof.Size), 1)
	}
	if col != nil {
		w.addCell(col)
	}
	return bof
}

func (w *WorkSheet) addCell(col Coler) {
	var row *Row
	var ok bool
	if row, ok = w.Rows[col.Row()]; !ok {
		info := new(RowInfo)
		info.Index = col.Row()
		row = w.addRow(info)
	}
	row.Cols[col.FirstCol()] = col
}

func (w *WorkSheet) addRow(info *RowInfo) (row *Row) {
	if info.Index > w.MaxRow {
		w.MaxRow = info.Index
	}
	var ok bool
	if row, ok = w.Rows[info.Index]; ok {
		row.info = info
	} else {
		row = &Row{info: info, Cols: make(map[uint16]Coler)}
		w.Rows[info.Index] = row
	}
	return
}
