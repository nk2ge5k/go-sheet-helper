package spreadsheet

import (
	"fmt"
	"math"
	"net/url"
	"regexp"
	"strconv"
	"strings"
)

const base int = 26

var (
	// RegexpSpeadsheetId is regexp for extracting spreadsheet id from url
	RegexpSpeadsheetId *regexp.Regexp = regexp.MustCompile("spreadsheets/d/([a-zA-Z0-9-_]+)")
	// ErrNotFound error represents error that returns when spreadsheet id not found
	ErrNotFound error = fmt.Errorf("spreadsheet id not found")

	emptyCellAddr CellAddr
	emptyRange    Range
)

// NewCellAddr returns new CellAddr from string address representation (e.g A1)
func NewCellAddr(addr string) (CellAddr, error) {
	if len(addr) < 2 {
		return emptyCellAddr, fmt.Errorf("invalid cell address '%s'", addr)
	}

	var (
		i    int
		char rune
	)

	for i, char = range addr {
		if !isLetter(char) {
			break
		}
	}

	if i < 1 || i == len(addr) {
		return emptyCellAddr, fmt.Errorf("invalid cell address '%s'", addr)
	}

	c, r := strings.ToUpper(addr[:i]), addr[i:]

	cell := CellAddr{}

	res, err := strconv.ParseUint(r, 10, 16)
	if err != nil {
		return emptyCellAddr, err
	}
	cell.Row = uint16(res - 1)

	num, err := colNum(c)
	if err != nil {
		return emptyCellAddr, err
	}
	cell.Col = num - 1

	return cell, nil
}

// CellAddr represents addres of sheet cell (e.g A1)
type CellAddr struct {
	Col, Row uint16
}

// String implements fmt.Stringer interface
func (c CellAddr) String() string {
	col, row := int(c.Col), int(c.Row)
	return string(colRunes(col+1)) + (strconv.Itoa(row + 1))
}

// Equal compares addres with another and returns true if they are eqal
func (c CellAddr) Equal(b CellAddr) bool {
	return c.Col == b.Col && c.Row == b.Row
}

// GreaterThan compares addres with another and returns true
//	addres greater than another
func (c CellAddr) GreaterThan(b CellAddr) bool {
	if c.Row > b.Row || c.Col > b.Row {
		return true
	}
	return false
}

// Move moves cell
// TODO: test
func (c CellAddr) Move(ver, hor int) CellAddr {
	// ???
	row, col := int(c.Row)+ver, int(c.Col)+hor
	return CellAddr{uint16(col), uint16(row)}
}

// colRunes return runes describing excel column name
func colRunes(col int) []rune {
	i := digitsCount(col, base)

	r := make([]rune, i)

	for i--; col > 0; i-- {
		mod := (col - 1) % base
		r[i] = rune('A' + mod)

		col = (col - mod) / base
	}

	return r[i+1:]
}

// colNum decodes column name to integer (e.g B to 2)
func colNum(name string) (uint16, error) {

	num, fbase := 0, float64(base)

	for i, c := range name {
		if !isLetter(c) {
			return 0, fmt.Errorf("col num: invalid address '%s'", name)
		}

		r := 'A'
		if 'a' <= c && c <= 'z' {
			r = 'a'
		}

		d := int(c-r) + 1

		digit := float64(len(name) - i - 1)
		num += d * int(math.Pow(fbase, digit))
	}

	return uint16(num), nil
}

// digitsCount return count of digits in integer, in any base
func digitsCount(i, base int) (count int) {
	if i == 0 {
		return 1
	}

	for i != 0 {
		i /= base
		count++
	}

	return
}

// isLetter checks if rune is valid range letter (a-zA-Z)
func isLetter(r rune) bool {
	switch {
	case 'a' <= r && r <= 'z':
		return true
	case 'A' <= r && r <= 'Z':
		return true
	}
	return false
}

// NewRange is a Range constructor from string
func NewRange(str string) (Range, error) {
	s := strings.Split(str, ":")
	if len(s) != 2 {
		return emptyRange, fmt.Errorf("invalid range %s", str)
	}

	min, err := NewCellAddr(s[0])
	if err != nil {
		return emptyRange, fmt.Errorf("new range: %v", err)
	}

	max, err := NewCellAddr(s[1])
	if err != nil {
		return emptyRange, fmt.Errorf("new range: %v", err)
	}

	if min.GreaterThan(max) {
		min, max = max, min
	}

	return Range{min, max}, nil
}

// Range represents excel range (e.g A1:B223)
// TODO: add optional sheet name
type Range struct {
	Min, Max CellAddr
}

// String implements fmt.Stringer interface
func (r Range) String() string {
	min, max := r.Min, r.Max

	if min.GreaterThan(max) {
		min, max = max, min
	}

	return fmt.Sprintf("%v:%v", min, max)
}

// Square calculates square of range
func (r Range) Square() int {
	min, max := r.Min, r.Max

	if min.GreaterThan(max) {
		min, max = max, min
	}

	w := max.Col - min.Col + 1
	h := max.Row - min.Row + 1

	return int(w * h)
}

// Move moves entire range
// TODO: test
func (r Range) Move(ver, hor int) Range {
	return Range{
		r.Min.Move(ver, hor),
		r.Max.Move(ver, hor),
	}
}

// ID extracts spreadsheet id from given url
func ID(src string) (string, error) {
	if len(src) == 0 {
		return "", fmt.Errorf("spreadsheet id: link is empty")
	}

	link, err := url.Parse(src)
	if err != nil {
		return "", err
	}

	if host := link.Hostname(); "docs.google.com" != host {
		return "", fmt.Errorf(
			"spreadsheet id: '%s' not a google docs hostname", host,
		)
	}

	if res := RegexpSpeadsheetId.FindStringSubmatch(src); len(res) == 2 {
		return res[1], nil
	}

	return "", ErrNotFound
}
