package format

import (
	"fmt"
	"testing"
)

func TestParseYY(t *testing.T) {
	_, tokens := Lexer(`yy-`)
	ds := Parse(tokens)
	fmt.Println(ds.Format(8100, true))
}

func TestParseMM(t *testing.T) {
	_, tokens := Lexer(`yyyymm`)
	ds := Parse(tokens)
	fmt.Println(ds.Format(8100, true))
}

func TestParseMM2(t *testing.T) {
	_, tokens := Lexer(`yyyy-mm"fasd65af"----`)
	ds := Parse(tokens)
	fmt.Println(ds.Format(8800, false))
}

func TestParseDD(t *testing.T) {
	_, tokens := Lexer(`yyyy-mm-dd`)
	ds := Parse(tokens)
	fmt.Println(ds.Format(8100, true))
}

func TestParseNum(t *testing.T) {
	_, tokens := Lexer(`"$"#,##0.00`)
	ds := Parse(tokens)
	fmt.Println(ds.Format(8800, false))
}
