# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Pure Go library for parsing Excel 97-2004 (.xls) files in BIFF5/BIFF8 format. Fork of `github.com/extrame/xls`. Does **not** handle .xlsx files.

## Build & Test

No `go.mod` — pre-modules project. Requires `GO111MODULE=on`:

```bash
GO111MODULE=on go test -v -race ./...   # run all tests
GO111MODULE=on go test -v -run TestBig  # run single test
```

No Makefile, no linter config.

## Architecture

```
WorkBook → WorkSheet → Row → Col (cell)
```

- **Entry points**: `Open()`, `OpenWithCloser()`, `OpenReader()` in `xls.go`
- **WorkBook** (`workbook.go`): parses BIFF record stream, manages sheets/fonts/formats/SST
- **WorkSheet** (`worksheet.go`): lazy-loaded on `GetSheet(i)`, contains rows by index
- **Row** (`row.go`): `Col(i)` (merged-cell aware) vs `ColExact(i)` (exact match)
- **Cell types** (`col.go`): `contentHandler` interface — implementations: `NumberCol`, `LabelsstCol`, `labelCol`, `BlankCol`, `RkCol`, `MulrkCol`, `FormulaCol`, `HyperLink`, etc.
- **Binary records** (`bof.go`): all data as `bof` structs (ID uint16 + Size uint16)
- **Date handling** (`date.go`): Excel serial dates → `time.Time`, supports 1900 & 1904 systems
- **Formatting** (`xf.go`, `font.go`, `format.go`): XF records map cells to fonts/formats

Key dependency: `github.com/extrame/ole2` for OLE2 container parsing.

## Parsing Flow

1. `Open()` → OLE2 parse → create WorkBook
2. `Parse()` → read BIFF record stream → extract sheets, SST, XF, fonts, formats
3. `GetSheet(n)` → lazy-parse worksheet records into rows/cells
4. Cell access → format values via XF index lookup
