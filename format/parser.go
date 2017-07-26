package format

// Parse creates a new parser with the recommended
// parameters.
func Parse(tokens []LexToken) Formatter {
	p := &parser{
		tokens: tokens,
		pos:    -1,
	}
	p.initState = initialParserState
	return p.run()
}

// run starts the statemachine
func (p *parser) run() Formatter {
	var f Formatter
	for state := p.initState; state != nil; {
		state = state(p, &f)
	}
	return f
}

// parserState represents the state of the scanner
// as a function that returns the next state.
type parserState func(*parser, *Formatter) parserState

// nest returns what the next token AND
// advances p.pos.
func (p *parser) next() *LexToken {
	if p.pos >= len(p.tokens)-1 {
		return nil
	}
	p.pos += 1
	return &p.tokens[p.pos]
}

// the parser type
type parser struct {
	tokens []LexToken
	pos    int
	serial int

	initState parserState
}

// the starting state for parsing
func initialParserState(p *parser, f *Formatter) parserState {
	var t *LexToken
	for t = p.next(); t[0] != T_EOF; t = p.next() {
		var item ItemFormatter
		switch t[0] {
		case T_YEAR_MARK:
			f.typ = DATEFORMAT
			item = new(YearFormatter)
		case T_MONTH_MARK:
			f.typ = DATEFORMAT
			item = new(MonthFormatter)
		case T_DAY_MARK:
			f.typ = DATEFORMAT
			item = new(DayFormatter)
		case T_RAW_MARK:
			item = new(basicFormatter)
		case T_STRING_MARK:
			item = new(basicFormatter)
		case T_COMMA_MARK, T_DECIMAL_MARK:
			f.typ = DECIMALFORMAT
			item = new(basicFormatter)
		}
		item.setOriginal(t[1])
		f.Items = append(f.Items, item)
	}
	if len(t[1]) > 0 {
		r := new(basicFormatter)
		r.origin = t[1]
		f.Items = append(f.Items, r)
	}
	return nil
}
