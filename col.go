package xls

import (
	"errors"
	"fmt"
	"math"
	"strconv"
	"time"
)

var ErrIsInt = errors.New("is int")

/* Data types */
const TYPE_STRING2 = 1
const TYPE_STRING = 2
const TYPE_FORMULA = 3
const TYPE_NUMERIC = 4
const TYPE_BOOL = 5
const TYPE_NULL = 6
const TYPE_INLINE = 7
const TYPE_ERROR = 8
const TYPE_DATETIME = 9
const TYPE_PERCENTAGE = 10
const TYPE_CURRENCY = 11

//content type
type contentHandler interface {
	Debug(wb *WorkBook)
	String(*WorkBook) []string
	FirstCol() uint16
	LastCol() uint16
}

type Col struct {
	RowB      uint16
	FirstColB uint16
}

type Coler interface {
	Row() uint16
}

func (c *Col) Debug(wb *WorkBook) {
	fmt.Printf("col dump:%#+v\n", c)
}

func (c *Col) Row() uint16 {
	return c.RowB
}

func (c *Col) FirstCol() uint16 {
	return c.FirstColB
}

func (c *Col) LastCol() uint16 {
	return c.FirstColB
}

func (c *Col) String(wb *WorkBook) []string {
	return []string{"default"}
}

type RK uint32

func (rk RK) Debug(wb *WorkBook) {
	fmt.Printf("rk dump:%#+v\n", rk)
}

func (rk RK) number() (intNum int64, floatNum float64, isFloat bool) {
	multiplied := rk & 1
	isInt := rk & 2
	val := rk >> 2
	if isInt == 0 {
		isFloat = true
		floatNum = math.Float64frombits(uint64(val) << 34)
		if multiplied != 0 {
			floatNum = floatNum / 100
		}
		return
	}
	//+++ add lines from here
	if multiplied != 0 {
		isFloat = true
		floatNum = float64(val) / 100
		return
	}
	//+++end
	return int64(val), 0, false
}

func (rk RK) float() float64 {
	var i, f, isFloat = rk.number()
	if !isFloat {
		f = float64(i)
	}

	return f
}

func (rk RK) String(wb *WorkBook) string {
	i, f, isFloat := rk.number()
	if isFloat {
		return strconv.FormatFloat(f, 'f', -1, 64)
	}

	return strconv.FormatInt(i, 10)
}

type XfRk struct {
	Index uint16
	Rk    RK
}

func (xf *XfRk) Debug(wb *WorkBook) {
	fmt.Printf("xfrk dump:%#+v\n", wb.Xfs[xf.Index])
	xf.Rk.Debug(wb)
}

func (xf *XfRk) String(wb *WorkBook) string {
	if val, ok := wb.Format(xf.Index, xf.Rk.float()); ok {
		return val
	}

	return xf.Rk.String(wb)
}

type MulrkCol struct {
	Col
	Xfrks    []XfRk
	LastColB uint16
}

func (c *MulrkCol) Debug(wb *WorkBook) {
	fmt.Printf("mulrk dump:%#+v\n", c)

	for _, v := range c.Xfrks {
		v.Debug(wb)
	}
}

func (c *MulrkCol) LastCol() uint16 {
	return c.LastColB
}

func (c *MulrkCol) String(wb *WorkBook) []string {
	var res = make([]string, len(c.Xfrks))
	for i, v := range c.Xfrks {
		res[i] = v.String(wb)
	}

	return res
}

type MulBlankCol struct {
	Col
	Xfs      []uint16
	LastColB uint16
}

func (c *MulBlankCol) Debug(wb *WorkBook) {
	fmt.Printf("mul blank dump:%#+v\n", c)
}

func (c *MulBlankCol) LastCol() uint16 {
	return c.LastColB
}

func (c *MulBlankCol) String(wb *WorkBook) []string {
	return make([]string, len(c.Xfs))
}

type NumberCol struct {
	Col
	Index uint16
	Float float64
}

func (c *NumberCol) Debug(wb *WorkBook) {
	fmt.Printf("number col dump:%#+v\n", c)
}

func (c *NumberCol) String(wb *WorkBook) []string {
	if v, ok := wb.Format(c.Index, c.Float); ok {
		return []string{v}
	}

	return []string{strconv.FormatFloat(c.Float, 'f', -1, 64)}
}

type FormulaColHeader struct {
	Col
	IndexXf uint16
	Result  [8]byte
	Flags   uint16
	_       uint32
}

// Value formula header value
func (f *FormulaColHeader) Value() float64 {
	var rknumhigh = ByteToUint32(f.Result[4:8])
	var rknumlow = ByteToUint32(f.Result[0:4])
	var sign = (rknumhigh & 0x80000000) >> 31
	var exp = ((rknumhigh & 0x7ff00000) >> 20) - 1023
	var mantissa = (0x100000 | (rknumhigh & 0x000fffff))
	var mantissalow1 = (rknumlow & 0x80000000) >> 31
	var mantissalow2 = (rknumlow & 0x7fffffff)
	var value = float64(mantissa) / math.Pow(2, float64(20-exp))

	if mantissalow1 != 0 {
		value += 1 / math.Pow(2, float64(21-exp))
	}

	value += float64(mantissalow2) / math.Pow(2, float64(52-exp))
	if 0 != sign {
		value *= -1
	}

	return value
}

// IsPart part of shared formula check
// WARNING:
// We can apparently not rely on $isPartOfSharedFormula. Even when $isPartOfSharedFormula = true
// the formula data may be ordinary formula data, therefore we need to check
// explicitly for the tExp token (0x01)
func (f *FormulaColHeader) IsPart() bool {
	return 0 != (0x0008 & ByteToUint16(f.Result[6:8]))
}

type FormulaCol struct {
	parsed bool
	Code   uint16
	Btl    uint16
	Btc    uint16
	Bts    []byte
	Header *FormulaColHeader
	ws     int
	vType  int
	value  string
}

func (c *FormulaCol) Debug(wb *WorkBook) {
	fmt.Printf("formula col dump:%#+v\n", c)
}

func (c *FormulaCol) Row() uint16 {
	return c.Header.Col.RowB
}

func (c *FormulaCol) FirstCol() uint16 {
	return c.Header.Col.FirstColB
}

func (c *FormulaCol) LastCol() uint16 {
	return c.Header.Col.FirstColB
}

func (c *FormulaCol) String(wb *WorkBook) []string {
	if !c.parsed {
		c.parse(wb, true)
	}

	return []string{c.value}
}

func (c *FormulaCol) parse(wb *WorkBook, ref bool) {
	c.parsed = true

	if 0 == c.Header.Result[0] && 255 == c.Header.Result[6] && 255 == c.Header.Result[7] {
		// String formula. Result follows in appended STRING record
		c.vType = TYPE_STRING
	} else if 1 == c.Header.Result[0] && 255 == c.Header.Result[6] && 255 == c.Header.Result[7] {
		// Boolean formula. Result is in +2; 0=false, 1=true
		c.vType = TYPE_BOOL
		if 0 == c.Header.Result[3] {
			c.value = "false"
		} else {
			c.value = "true"
		}
	} else if 2 == c.Header.Result[0] && 255 == c.Header.Result[6] && 255 == c.Header.Result[7] {
		// Error formula. Error code is in +2
		c.vType = TYPE_ERROR
		switch c.Header.Result[3] {
		case 0x00:
			c.value = "#NULL!"
		case 0x07:
			c.value = "#DIV/0"
		case 0x0F:
			c.value = "#VALUE!"
		case 0x17:
			c.value = "#REF!"
		case 0x1D:
			c.value = "#NAME?"
		case 0x24:
			c.value = "#NUM!"
		case 0x2A:
			c.value = "#N/A"
		}
	} else if 3 == c.Header.Result[0] && 255 == c.Header.Result[6] && 255 == c.Header.Result[7] {
		// Formula result is a null string
		c.vType = TYPE_NULL
		c.value = ""
	} else {
		// formula result is a number, first 14 bytes like _NUMBER record
		c.vType = TYPE_NUMERIC

		var flag bool
		if c.isGetCurTime() {
			// if date time format is not support, use time.RFC3339
			if c.value, flag = wb.Format(c.Header.IndexXf, 0); !flag {
				c.value = parseTime(0, time.RFC3339)
			}
		} else if c.isRef() {
			if ref {
				var ws = -1
				var find bool
				var rIdx uint16
				var cIdx uint16

				if 0x07 == c.Bts[0] {
					var exi = ByteToUint16(c.Bts[3:5])
					rIdx = ByteToUint16(c.Bts[5:7])
					cIdx = 0x00FF & ByteToUint16(c.Bts[7:9])
					if exi <= wb.ref.Num {
						ws = int(wb.ref.Info[int(exi)].FirstSheetIndex)
					}
				} else {
					ws = c.ws
					rIdx = ByteToUint16(c.Bts[3:5])
					cIdx = 0x00FF & ByteToUint16(c.Bts[5:7])
				}

				if ws < len(wb.sheets) {
					if row := wb.GetSheet(ws).Row(int(rIdx)); nil != row {
						find = true
						c.value = row.Col(int(cIdx))
					}
				}
				if !find {
					c.value = "#REF!"
				}
			} else {
				c.parsed = false
			}
		} else {
			c.value, flag = wb.Format(c.Header.IndexXf, c.Header.Value())
			if !flag {
				c.value = strconv.FormatFloat(c.Header.Value(), 'f', -1, 64)
			}
		}
	}
}

// isRef return cell is reference to other cell
func (c *FormulaCol) isRef() bool {
	if 0x05 == c.Bts[0] && (0x24 == c.Bts[2] || 0x44 == c.Bts[2] || 0x64 == c.Bts[2]) {
		return true
	} else if 0x07 == c.Bts[0] && (0x3A == c.Bts[2] || 0x5A == c.Bts[2] || 0x7A == c.Bts[2]) {
		return true
	}

	return false
}

// isGetCurTime return cell value is get current date or datetime flag
func (c *FormulaCol) isGetCurTime() bool {
	var ret bool
	var next byte

	if 0x19 == c.Bts[2] && (0x21 == c.Bts[6] || 0x41 == c.Bts[6] || 0x61 == c.Bts[6]) {
		next = c.Bts[7]
	} else if 0x21 == c.Bts[2] || 0x41 == c.Bts[2] || 0x61 == c.Bts[2] {
		next = c.Bts[3]
	}

	if 0x4A == next || 0xDD == next {
		ret = true
	}

	return ret
}

type RkCol struct {
	Col
	Xfrk XfRk
}

func (c *RkCol) Debug(wb *WorkBook) {
	fmt.Printf("rk col dump:%#+v\n", c)
}

func (c *RkCol) String(wb *WorkBook) []string {
	return []string{c.Xfrk.String(wb)}
}

type LabelsstCol struct {
	Col
	Xf  uint16
	Sst uint32
}

func (c *LabelsstCol) Debug(wb *WorkBook) {
	fmt.Printf("label sst col dump:%#+v\n", c)
}

func (c *LabelsstCol) String(wb *WorkBook) []string {
	return []string{wb.sst[int(c.Sst)]}
}

type labelCol struct {
	BlankCol
	Str string
}

func (c *labelCol) Debug(wb *WorkBook) {
	fmt.Printf("label col dump:%#+v\n", c)
}

func (c *labelCol) String(wb *WorkBook) []string {
	return []string{c.Str}
}

type BlankCol struct {
	Col
	Xf uint16
}

func (c *BlankCol) Debug(wb *WorkBook) {
	fmt.Printf("blank col dump:%#+v\n", c)
}

func (c *BlankCol) String(wb *WorkBook) []string {
	return []string{""}
}
