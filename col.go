package xls

import (
	"fmt"
	"math"
)

//content type
type contentHandler interface {
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

type XfRk struct {
	Index uint16
	Rk    RK
}

func (xf *XfRk) String(wb *WorkBook) string {
	if len(wb.Xfs) > 21 {
		switch wb.Xfs[21].formatNo() {
		case 27:
			if f, e := xf.Rk.Float(); e == nil {
				t := timeFromExcelTime(f, true)
				return t.Format("2006.01") //TODO it should be international
			}
		}
	}
	return fmt.Sprintf("%s", xf.Rk.String())
}

type RK uint32

func (rk RK) String() string {
	multiplied := rk & 1
	isInt := rk & 2
	val := rk >> 2
	if isInt == 0 {
		f := math.Float64frombits(uint64(val) << 34)
		if multiplied != 0 {
			f = f / 100
		}
		return fmt.Sprintf("%.1f", f)
	} else {
		return fmt.Sprint(val)
	}
}

var ErrIsInt = fmt.Errorf("is int")

func (rk RK) Float() (float64, error) {
	multiplied := rk & 1
	isInt := rk & 2
	val := rk >> 2
	if isInt == 0 {
		f := math.Float64frombits(uint64(val) << 34)
		if multiplied != 0 {
			f = f / 100
		}
		return f, nil
	} else {
		return 0.0, ErrIsInt
	}
}

type MulrkCol struct {
	Col
	Xfrks    []XfRk
	LastColB uint16
}

func (c *MulrkCol) LastCol() uint16 {
	return c.LastColB
}

func (c *MulrkCol) String(wb *WorkBook) []string {
	var res = make([]string, len(c.Xfrks))
	for i := 0; i < len(c.Xfrks); i++ {
		xfrk := c.Xfrks[i]
		res[i] = xfrk.String(wb)
	}
	return res
}

type MulBlankCol struct {
	Col
	Xfs      []uint16
	LastColB uint16
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

func (c *NumberCol) String(wb *WorkBook) []string {
	return []string{fmt.Sprintf("%f", c.Float)}
}

type FormulaCol struct {
	Header struct {
		Col
		IndexXf uint16
		Result  [8]byte
		Flags   uint16
		_       uint32
	}
	Bts []byte
}

func (c *FormulaCol) String(wb *WorkBook) []string {
	return []string{"FormulaCol"}
}

type RkCol struct {
	Col
	Xfrk XfRk
}

func (c *RkCol) String(wb *WorkBook) []string {
	return []string{c.Xfrk.Rk.String()}
}

type LabelsstCol struct {
	Col
	Xf  uint16
	Sst uint32
}

func (c *LabelsstCol) String(wb *WorkBook) []string {
	return []string{wb.sst[int(c.Sst)]}
}

type BlankCol struct {
	Col
	Xf uint16
}

func (c *BlankCol) String(wb *WorkBook) []string {
	return []string{""}
}
