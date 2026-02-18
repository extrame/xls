package xls

import (
	"bytes"
	"encoding/binary"
	"testing"
)

func writeBIFFRecord(buf *bytes.Buffer, id uint16, data []byte) {
	binary.Write(buf, binary.LittleEndian, id)
	binary.Write(buf, binary.LittleEndian, uint16(len(data)))
	buf.Write(data)
}

func makeBOFData() []byte {
	var b bytes.Buffer
	binary.Write(&b, binary.LittleEndian, uint16(0x0600)) // Ver: BIFF8
	binary.Write(&b, binary.LittleEndian, uint16(0x0005)) // Type: workbook globals
	b.Write(make([]byte, 12))                              // Id_make, Year, Flags, Min_ver
	return b.Bytes()
}

// TestSSTContinueRichtext tests that richtext formatting runs spanning
// a record boundary don't corrupt subsequent SST entries.
// String 0: "ABC" with 2 formatting runs — chars fit in SST, runs overflow to CONTINUE.
// String 1: "DEF" in CONTINUE after the formatting runs.
func TestSSTContinueRichtext(t *testing.T) {
	// SST record: SstInfo + string 0 header + "ABC" (no room for formatting runs)
	var sstData bytes.Buffer
	binary.Write(&sstData, binary.LittleEndian, uint32(2)) // Total
	binary.Write(&sstData, binary.LittleEndian, uint32(2)) // Count
	binary.Write(&sstData, binary.LittleEndian, uint16(3)) // string 0 char count
	sstData.WriteByte(0x08)                                 // flag: has richtext
	binary.Write(&sstData, binary.LittleEndian, uint16(2)) // richtext_num = 2 runs
	sstData.Write([]byte("ABC"))                            // char data (Latin1)
	// 0 bytes left for 8-byte formatting runs → overflow to CONTINUE

	// CONTINUE record: formatting runs + string 1
	var contData bytes.Buffer
	contData.Write(make([]byte, 8))                          // 2 runs × 4 bytes (dummy)
	binary.Write(&contData, binary.LittleEndian, uint16(3)) // string 1 char count
	contData.WriteByte(0x00)                                 // flag: plain Latin1
	contData.Write([]byte("DEF"))

	var stream bytes.Buffer
	writeBIFFRecord(&stream, 0x0809, makeBOFData())
	writeBIFFRecord(&stream, 0x00FC, sstData.Bytes())
	writeBIFFRecord(&stream, 0x003C, contData.Bytes())

	wb := &WorkBook{Formats: make(map[uint16]*Format)}
	wb.Parse(bytes.NewReader(stream.Bytes()))

	if len(wb.sst) != 2 {
		t.Fatalf("expected 2 SST entries, got %d", len(wb.sst))
	}
	if wb.sst[0] != "ABC" {
		t.Errorf("sst[0]: expected %q, got %q", "ABC", wb.sst[0])
	}
	if wb.sst[1] != "DEF" {
		t.Errorf("sst[1]: expected %q, got %q", "DEF", wb.sst[1])
	}
}

// TestSSTContinueRichtextPartial tests the case where richtext formatting
// runs are partially read in the SST record (boundary falls mid-formatting).
func TestSSTContinueRichtextPartial(t *testing.T) {
	// SST record: SstInfo + string 0 header + "ABC" + 4 of 8 formatting bytes
	var sstData bytes.Buffer
	binary.Write(&sstData, binary.LittleEndian, uint32(2)) // Total
	binary.Write(&sstData, binary.LittleEndian, uint32(2)) // Count
	binary.Write(&sstData, binary.LittleEndian, uint16(3)) // string 0 char count
	sstData.WriteByte(0x08)                                 // flag: has richtext
	binary.Write(&sstData, binary.LittleEndian, uint16(2)) // richtext_num = 2
	sstData.Write([]byte("ABC"))                            // chars
	sstData.Write(make([]byte, 4))                          // partial: 4 of 8 formatting bytes

	// CONTINUE: remaining 4 formatting bytes + string 1
	var contData bytes.Buffer
	contData.Write(make([]byte, 4))                          // remaining formatting bytes
	binary.Write(&contData, binary.LittleEndian, uint16(3)) // string 1 char count
	contData.WriteByte(0x00)                                 // flag
	contData.Write([]byte("DEF"))

	var stream bytes.Buffer
	writeBIFFRecord(&stream, 0x0809, makeBOFData())
	writeBIFFRecord(&stream, 0x00FC, sstData.Bytes())
	writeBIFFRecord(&stream, 0x003C, contData.Bytes())

	wb := &WorkBook{Formats: make(map[uint16]*Format)}
	wb.Parse(bytes.NewReader(stream.Bytes()))

	if len(wb.sst) != 2 {
		t.Fatalf("expected 2 SST entries, got %d", len(wb.sst))
	}
	if wb.sst[0] != "ABC" {
		t.Errorf("sst[0]: expected %q, got %q", "ABC", wb.sst[0])
	}
	if wb.sst[1] != "DEF" {
		t.Errorf("sst[1]: expected %q, got %q", "DEF", wb.sst[1])
	}
}

// TestSSTContinuePhonetic tests that phonetic data spanning a record
// boundary is properly skipped in the CONTINUE handler.
func TestSSTContinuePhonetic(t *testing.T) {
	// SST record: SstInfo + string 0 with phonetic flag, chars fit but phonetic overflows
	var sstData bytes.Buffer
	binary.Write(&sstData, binary.LittleEndian, uint32(2))  // Total
	binary.Write(&sstData, binary.LittleEndian, uint32(2))  // Count
	binary.Write(&sstData, binary.LittleEndian, uint16(3))  // string 0 char count
	sstData.WriteByte(0x04)                                  // flag: has phonetic
	binary.Write(&sstData, binary.LittleEndian, uint32(12)) // phonetic_size = 12 bytes
	sstData.Write([]byte("ABC"))                             // chars
	// 0 bytes for 12-byte phonetic data → overflow

	// CONTINUE: phonetic data + string 1
	var contData bytes.Buffer
	contData.Write(make([]byte, 12))                         // phonetic data (dummy)
	binary.Write(&contData, binary.LittleEndian, uint16(3)) // string 1 char count
	contData.WriteByte(0x00)                                 // flag
	contData.Write([]byte("DEF"))

	var stream bytes.Buffer
	writeBIFFRecord(&stream, 0x0809, makeBOFData())
	writeBIFFRecord(&stream, 0x00FC, sstData.Bytes())
	writeBIFFRecord(&stream, 0x003C, contData.Bytes())

	wb := &WorkBook{Formats: make(map[uint16]*Format)}
	wb.Parse(bytes.NewReader(stream.Bytes()))

	if len(wb.sst) != 2 {
		t.Fatalf("expected 2 SST entries, got %d", len(wb.sst))
	}
	if wb.sst[0] != "ABC" {
		t.Errorf("sst[0]: expected %q, got %q", "ABC", wb.sst[0])
	}
	if wb.sst[1] != "DEF" {
		t.Errorf("sst[1]: expected %q, got %q", "DEF", wb.sst[1])
	}
}
