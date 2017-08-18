package format

import (
	"errors"
	"fmt"
)

var InvalidTimeFormat = errors.New("invalid time format")

type Formatter struct {
	typ   int
	Items []ItemFormatter
}

func (f *Formatter) Format(val float64, date1904 bool) string {
	var gf string
	for _, i := range f.Items {
		itemFormatString, _ := i.translateToGolangFormat()
		gf += itemFormatString
	}
	if f.typ == DATEFORMAT {
		t := TimeFromExcelTime(val, date1904)
		return t.Format(gf)
	} else {
		return fmt.Sprintf("%f", val)
	}
}

type ItemFormatter interface {
	translateToGolangFormat() (string, error)
	setOriginal(string)
}

type basicFormatter struct {
	origin string
}

func (self *basicFormatter) translateToGolangFormat() (string, error) {
	return self.origin, nil
}

func (self *basicFormatter) setOriginal(o string) {
	self.origin = o
}

func (self *basicFormatter) String() string {
	return fmt.Sprintf("basic formatter as (%s)", self.origin)
}

type commaFormatter struct {
	basicFormatter
}

func (self *commaFormatter) translateToGolangFormat() (string, error) {
	return self.origin, nil
}

func (self *commaFormatter) setOriginal(o string) {
	self.origin = o
}

func (self *commaFormatter) String() string {
	return fmt.Sprintf("basic formatter as (%s)", self.origin)
}

type YearFormatter struct {
	basicFormatter
}

func (y *YearFormatter) translateToGolangFormat() (string, error) {
	switch len(y.origin) {
	case 4:
		return "2006", nil
	case 2:
		return "06", nil
	default:
		return "", InvalidTimeFormat
	}
}

func (y *YearFormatter) String() string {
	return fmt.Sprintf("year formatter as (%s)", y.origin)
}

type MonthFormatter struct {
	basicFormatter
}

func (self *MonthFormatter) translateToGolangFormat() (string, error) {
	switch len(self.origin) {
	case 2:
		return "01", nil
	case 1:
		return "1", nil
	default:
		return "", InvalidTimeFormat
	}
}

func (self *MonthFormatter) String() string {
	return fmt.Sprintf("month formatter as (%s)", self.origin)
}

type DayFormatter struct {
	basicFormatter
}

func (self *DayFormatter) translateToGolangFormat() (string, error) {
	switch len(self.origin) {
	case 2:
		return "02", nil
	case 1:
		return "2", nil
	default:
		return "", InvalidTimeFormat
	}
}

func (self *DayFormatter) String() string {
	return fmt.Sprintf("day formatter as (%s)", self.origin)
}
