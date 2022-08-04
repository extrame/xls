package xls

import (
	"encoding/binary"
	"io"
)

func ReadBytes(r io.Reader, size int) ([]byte, error) {
	buf := make([]byte, size)
	if _, err := r.Read(buf); err != nil {
		return buf, err
	}
	return buf, nil
}

func MustReadBytes(r io.Reader, size int) []byte {
	buf, _ := ReadBytes(r, size)
	return buf
}

func ReadByte(r io.Reader) (byte, error) {
	buf, err := ReadBytes(r, 1)
	if err != nil {
		return 0, err
	}
	return buf[0], nil
}

func ReadUint16(r io.Reader) (uint16, error) {
	buf, err := ReadBytes(r, 2)
	if err != nil {
		return 0, err
	}
	return binary.LittleEndian.Uint16(buf), nil
}

func ReadUint32(r io.Reader) (uint32, error) {
	buf, err := ReadBytes(r, 4)
	if err != nil {
		return 0, err
	}
	return binary.LittleEndian.Uint32(buf), nil
}

func ReadBoundSheet(r io.Reader) *boundsheet {
	var bs = new(boundsheet)
	buf, _ := ReadBytes(r, 7)
	bs.Filepos = binary.LittleEndian.Uint32(buf[0:4])
	bs.Visible = buf[4]
	bs.Type = buf[5]
	bs.Name = buf[6]
	return bs
}

func ReadRowInfo(r io.Reader) *rowInfo {
	row := new(rowInfo)
	buf, _ := ReadBytes(r, 16)
	row.Index = binary.LittleEndian.Uint16(buf[0:2])
	row.Fcell = binary.LittleEndian.Uint16(buf[2:4])
	row.Lcell = binary.LittleEndian.Uint16(buf[4:6])
	row.Height = binary.LittleEndian.Uint16(buf[6:8])
	row.Notused = binary.LittleEndian.Uint16(buf[8:10])
	row.Notused2 = binary.LittleEndian.Uint16(buf[10:12])
	row.Flags = binary.LittleEndian.Uint32(buf[12:16])
	return row
}

func ReadLabelsstCol(r io.Reader) *LabelsstCol {
	col := new(LabelsstCol)
	buf, _ := ReadBytes(r, 10)
	col.RowB = binary.LittleEndian.Uint16(buf[0:2])
	col.FirstColB = binary.LittleEndian.Uint16(buf[2:4])
	col.Xf = binary.LittleEndian.Uint16(buf[4:6])
	col.Sst = binary.LittleEndian.Uint32(buf[6:10])
	return col
}

func ReadBof(r io.Reader, row *bof) error {
	buf, err := ReadBytes(r, 4)
	if err != nil {
		return err
	}
	row.Id = binary.LittleEndian.Uint16(buf[0:2])
	row.Size = binary.LittleEndian.Uint16(buf[2:4])
	return err
}
