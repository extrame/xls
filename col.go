package xls

import (
	"fmt"
	"math"
)

type Col struct {
	RowB      uint16
	FirstColB uint16
}

type Coler interface {
	String(*WorkBook) []string
	Row() uint16
	FirstCol() uint16
	LastCol() uint16
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
	return []string{""}
}

type XfRk struct {
	Index uint16
	Rk    RK
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
		res[i] = xfrk.Rk.String()
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
	Col
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
