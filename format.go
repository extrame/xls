package xls

import (
	"regexp"
	"strconv"
	"strings"
	"time"
)

// Excel styles can reference number formats that are built-in, all of which
// have an id less than 164. This is a possibly incomplete list comprised of as
// many of them as I could find.
var builtInNumFmt = map[uint16]string{
	0:  "general",
	1:  "0",
	2:  "0.00",
	3:  "#,##0",
	4:  "#,##0.00",
	9:  "0%",
	10: "0.00%",
	11: "0.00e+00",
	12: "# ?/?",
	13: "# ??/??",
	14: "mm-dd-yy",
	15: "d-mmm-yy",
	16: "d-mmm",
	17: "mmm-yy",
	18: "h:mm am/pm",
	19: "h:mm:ss am/pm",
	20: "h:mm",
	21: "h:mm:ss",
	22: "m/d/yy h:mm",
	37: "#,##0 ;(#,##0)",
	38: "#,##0 ;[red](#,##0)",
	39: "#,##0.00;(#,##0.00)",
	40: "#,##0.00;[red](#,##0.00)",
	41: `_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)`,
	42: `_("$"* #,##0_);_("$* \(#,##0\);_("$"* "-"_);_(@_)`,
	43: `_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)`,
	44: `_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)`,
	45: "mm:ss",
	46: "[h]:mm:ss",
	47: "mmss.0",
	48: "##0.0e+0",
	49: "@",
	58: time.RFC3339,
}

// Excel date time mapper to go system
var dateTimeMapper = []struct{ xls, golang string }{
	{"yyyy", "2006"},
	{"yy", "06"},
	{"mmmm", "%%%%"},
	{"dddd", "&&&&"},
	{"dd", "02"},
	{"d", "2"},
	{"mmm", "Jan"},
	{"mmss", "0405"},
	{"ss", "05"},
	{"mm:", "04:"},
	{":mm", ":04"},
	{"mm", "01"},
	{"am/pm", "pm"},
	{"m/", "1/"},
	{"%%%%", "January"},
	{"&&&&", "Monday"},
}

// Format value interface
type Format struct {
	Head struct {
		Index uint16
		Size  uint16
	}
	Raw   []string
	bts   int
	vType int
}

// Prepare format meta data
func (f *Format) Prepare() {
	var regexColor = regexp.MustCompile("^\\[[a-zA-Z]+\\]")
	var regexFraction = regexp.MustCompile("#\\,?#*")

	for k, v := range f.Raw {
		// In Excel formats, "_" is used to add spacing, which we can't do in HTML
		v = strings.Replace(v, "_", "", -1)

		// Some non-number characters are escaped with \, which we don't need
		v = strings.Replace(v, "\\", "", -1)

		// Some non-number strings are quoted, so we'll get rid of the quotes, likewise any positional * symbols
		v = strings.Replace(v, "*", "", -1)
		v = strings.Replace(v, "\"", "", -1)

		// strip ()
		v = strings.Replace(v, "(", "", -1)
		v = strings.Replace(v, ")", "", -1)

		// strip color information
		v = regexColor.ReplaceAllString(v, "")

		// Strip #
		v = regexFraction.ReplaceAllString(v, "")

		if 0 == f.vType {
			if regexp.MustCompile("^(\\[\\$[A-Z]*-[0-9A-F]*\\])*[hmsdy]").MatchString(v) {
				f.vType = TYPE_DATETIME
			} else if strings.HasSuffix(v, "%") {
				f.vType = TYPE_PERCENTAGE
			} else if strings.HasPrefix(v, "$") || strings.HasPrefix(v, "￥") {
				f.vType = TYPE_CURRENCY
			}
		}

		f.Raw[k] = strings.Trim(v, "\r\n\t ")
	}

	if 0 == f.vType {
		f.vType = TYPE_NUMERIC
	}

	if TYPE_NUMERIC == f.vType || TYPE_CURRENCY == f.vType || TYPE_PERCENTAGE == f.vType {
		var t []string
		if t = strings.SplitN(f.Raw[0], ".", 2); 2 == len(t) {
			f.bts = strings.Count(t[1], "")

			if f.bts > 0 {
				f.bts = f.bts - 1
			}
		}
	}
}

// String format content to spec string
// see http://www.openoffice.org/sc/excelfileformat.pdf Page #174
func (f *Format) String(v float64) string {
	var ret string

	switch f.vType {
	case TYPE_NUMERIC:
		if 0 == f.bts {
			ret = strconv.FormatInt(int64(v), 10)
		} else {
			ret = strconv.FormatFloat(v, 'f', f.bts, 64)
		}
	case TYPE_CURRENCY:
		if 0 == f.bts {
			ret = strconv.FormatInt(int64(v), 10)
		} else {
			ret = strconv.FormatFloat(v, 'f', f.bts, 64)
		}
	case TYPE_PERCENTAGE:
		if 0 == f.bts {
			ret = strconv.FormatInt(int64(v)*100, 10) + "%"
		} else {
			ret = strconv.FormatFloat(v*100, 'f', f.bts, 64) + "%"
		}
	case TYPE_DATETIME:
		ret = parseTime(v, f.Raw[0])
	default:
		ret = strconv.FormatFloat(v, 'f', -1, 64)
	}

	return ret
}

// ByteToUint32 Read 32-bit unsigned integer
func ByteToUint32(b []byte) uint32 {
	return uint32(b[0]) | uint32(b[1])<<8 | uint32(b[2])<<16 | uint32(b[3])<<24
}

// ByteToUint16 Read 16-bit unsigned integer
func ByteToUint16(b []byte) uint16 {
	return (uint16(b[0]) | (uint16(b[1]) << 8))
}

// parseTime provides function to returns a string parsed using time.Time.
// Replace Excel placeholders with Go time placeholders. For example, replace
// yyyy with 2006. These are in a specific order, due to the fact that m is used
// in month, minute, and am/pm. It would be easier to fix that with regular
// expressions, but if it's possible to keep this simple it would be easier to
// maintain. Full-length month and days (e.g. March, Tuesday) have letters in
// them that would be replaced by other characters below (such as the 'h' in
// March, or the 'd' in Tuesday) below. First we convert them to arbitrary
// characters unused in Excel Date formats, and then at the end, turn them to
// what they should actually be.
// Based off: http://www.ozgrid.com/Excel/CustomFormats.htm
func parseTime(v float64, f string) string {
	var val time.Time
	if 0 == v {
		val = time.Now()
	} else {
		val = timeFromExcelTime(v, false)
	}

	// It is the presence of the "am/pm" indicator that determines if this is
	// a 12 hour or 24 hours time format, not the number of 'h' characters.
	if is12HourTime(f) {
		f = strings.Replace(f, "hh", "03", 1)
		f = strings.Replace(f, "h", "3", 1)
	} else {
		f = strings.Replace(f, "hh", "15", 1)
		f = strings.Replace(f, "h", "15", 1)
	}
	for _, repl := range dateTimeMapper {
		f = strings.Replace(f, repl.xls, repl.golang, 1)
	}

	// If the hour is optional, strip it out, along with the possible dangling
	// colon that would remain.
	if val.Hour() < 1 {
		f = strings.Replace(f, "]:", "]", 1)
		f = strings.Replace(f, "[03]", "", 1)
		f = strings.Replace(f, "[3]", "", 1)
		f = strings.Replace(f, "[15]", "", 1)
	} else {
		f = strings.Replace(f, "[3]", "3", 1)
		f = strings.Replace(f, "[15]", "15", 1)
	}

	return val.Format(f)
}

// is12HourTime checks whether an Excel time format string is a 12 hours form.
func is12HourTime(format string) bool {
	return strings.Contains(format, "am/pm") || strings.Contains(format, "AM/PM") || strings.Contains(format, "a/p") || strings.Contains(format, "A/P")
}
