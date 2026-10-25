package xls

import (
	"bytes"
	"encoding/binary"
	"testing"
)

// TestHyperlinkZeroCountPanic demonstrates a slice-bounds panic in the
// worksheet record parser. In worksheet.go case 0x1b8 (HYPERLINK) and in the
// shared helper bof.utf16String (bof.go:19), a UTF-16 string length `count` is
// read straight from the file and then:
//
//	var bts = make([]uint16, count)
//	binary.Read(buf, binary.LittleEndian, &bts)
//	runes := utf16.Decode(bts[:len(bts)-1])   // bts[:-1] when count == 0
//
// A HYPERLINK record whose flags request a string but whose length field is 0
// makes `bts` empty, so `bts[:len(bts)-1]` is `bts[:-1]` -> panic.
//
// parse() / parseBof() have no recover(), and the record bytes are the raw,
// attacker-controlled worksheet stream reached from the public API via
// OpenReader(...).GetSheet(n) / ReadAllCells(...). This test drives parse()
// directly with that stream. It FAILS on current code (panics).
func TestHyperlinkZeroCountPanic(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Fatalf("parse panicked on a malformed HYPERLINK record (count==0): %v", r)
		}
	}()

	var rec bytes.Buffer
	// record body: CellRange(8) + 20 skipped + flag(4) + count(4) = 36 bytes
	body := make([]byte, 0, 36)
	body = append(body, make([]byte, 8)...)  // CellRange (zeros)
	body = append(body, make([]byte, 20)...) // 20 bytes skipped by buf.Seek(20,1)
	flag := make([]byte, 4)
	binary.LittleEndian.PutUint32(flag, 0x14) // request Description string
	body = append(body, flag...)
	body = append(body, make([]byte, 4)...) // count = 0  -> make([]uint16, 0)

	// BIFF record header: Id=0x1b8 (HYPERLINK), Size=len(body)
	var hdr [4]byte
	binary.LittleEndian.PutUint16(hdr[0:2], 0x1b8)
	binary.LittleEndian.PutUint16(hdr[2:4], uint16(len(body)))
	rec.Write(hdr[:])
	rec.Write(body)

	w := &WorkSheet{}
	w.parse(bytes.NewReader(rec.Bytes()))

	// If we get here the input was handled without panicking.
	t.Logf("parse handled the malformed record without panic (%d rows)", len(w.rows))
}
