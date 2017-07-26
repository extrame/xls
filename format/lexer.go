package format

import (
	"strconv"
	"strings"
	"unicode/utf8"
)

const (
	DATEFORMAT    = iota
	DECIMALFORMAT = iota
)

// LexToken holds is a (type, value) array.
type LexToken [3]string

// EOF character
var EOF string = "+++EOF+++"

// lexerState represents the state of the scanner
// as a function that returns the next state.
type lexerState func(*lexer) lexerState

// run lexes the input by executing state functions until
// the state is nil.
func (l *lexer) Run() {
	for state := l.initialState; state != nil; {
		state = state(l)
	}
}

// Lexer creates a new scanner for the input string.
func Lexer(input string) (*lexer, []LexToken) {
	l := &lexer{
		input:  input,
		tokens: make([]LexToken, 0),
		lineno: 1,
	}
	l.initialState = initLexerState
	l.Run()
	return l, l.tokens
}

// lexer holds the state of the scanner.
type lexer struct {
	input        string     // the string being scanned.
	start        int        // start position of this item.
	pos          int        // current position in the input.
	width        int        // width of last rune read from input.
	tokens       []LexToken // scanned items.
	initialState lexerState
	typ          int
	lineno       int
}

// next returns the next rune in the input.
func (l *lexer) next() string {
	var r rune
	if l.pos >= len(l.input) {
		l.width = 0
		return EOF
	}
	r, l.width = utf8.DecodeRuneInString(l.input[l.pos:])
	l.pos += l.width
	return string(r)
}

// ignore skips over the pending input before this point.
func (l *lexer) ignore() {
	l.start = l.pos
}

// backup steps back one rune.
// Can be called only once per call of next.
func (l *lexer) backup() {
	l.pos -= l.width
}

// acceptRun consumes a run of runes from the valid set.
func (l *lexer) acceptRun(valid string) {
	for strings.Index(valid, l.next()) >= 0 {
	}
	l.backup()
}

// acceptRun consumes a run of runes from the valid set.
func (l *lexer) acceptUntil(marker string) {
	for r := l.next(); r != EOF && strings.Index(marker, r) < 0; r = l.next() {
	}
}

// emit passes an item back to the client.
func (l *lexer) emit(t string) {
	l.tokens = append(l.tokens, LexToken{t, l.input[l.start:l.pos], strconv.Itoa(l.lineno)})
	l.start = l.pos
}

// emit passes an item back to the client.
func (l *lexer) emitRaw() {
	if l.pos-l.start > 1 {
		l.tokens = append(l.tokens, LexToken{T_RAW_MARK, l.input[l.start : l.pos-1], strconv.Itoa(l.lineno)})
		l.start = l.pos - 1
	}
}

// emit passes an item back to the client.
func (l *lexer) emitWithoutEnd(t string) {
	if l.pos-l.start > 1 {
		l.tokens = append(l.tokens, LexToken{t, l.input[l.start : l.pos-1], strconv.Itoa(l.lineno)})
		l.start = l.pos
	}
}

// initialState is the starting point for the
// scanner. It scans through each character and decides
// which state to create for the lexer. lexerState == nil
// is exit scanner.
func initLexerState(l *lexer) lexerState {
	for r := l.next(); r != EOF; r = l.next() {
		if r == "y" {
			l.emitRaw()
			l.acceptRun("y")
			l.emit(T_YEAR_MARK)
		} else if r == "m" {
			l.emitRaw()
			l.acceptRun("m")
			l.emit(T_MONTH_MARK)
		} else if r == "d" {
			l.emitRaw()
			l.acceptRun("d")
			l.emit(T_DAY_MARK)
		} else if r == "\"" {
			l.emitRaw()
			l.ignore()
			l.acceptUntil("\"")
			l.emitWithoutEnd(T_STRING_MARK)
		} else if r == "#" {
			l.emitRaw()
			l.acceptRun("#,")
			l.emit(T_COMMA_MARK)
		} else if r == "0" {
			l.emitRaw()
			l.acceptRun("0123456789.")
			l.emit(T_DECIMAL_MARK)
		}
	}

	l.emit(T_EOF)
	return nil
}
