package spreadsheet

import "testing"

func TestID(t *testing.T) {
	tt := map[string]string{
		"https://docs.google.com/spreadsheets/d/232jfks": "232jfks",
		"https://docs.yahoo.com/spreadsheets/d/23sksfjh": "",
		"fhejk": "",
		"":      "",
	}

	for u, w := range tt {

		id, err := ID(u)
		if w != id {
			t.Errorf("ID(%s) = (%s, %v), want %s", u, id, err, w)
		}
	}
}

func TestDigitsCount(t *testing.T) {
	tt := []struct{ i, base, want int }{
		{0, 10, 1},
		{4, 2, 3},
		{11, 16, 1},
		{256, 16, 3},
		{22, 16, 2},
		{7, 8, 1},
		{100, 10, 3},
		{12, 26, 1},
		{34, 26, 2},
	}

	for _, set := range tt {
		if result := digitsCount(set.i, set.base); result != set.want {
			t.Errorf(
				"digitsCount(%d, %d) = %d, want %d",
				set.i, set.base, result, set.want,
			)
		}
	}
}

func TestColRune(t *testing.T) {
	tt := map[int]string{
		1:     "A",
		2:     "B",
		5:     "E",
		26:    "Z",
		27:    "AA",
		28:    "AB",
		52:    "AZ",
		676:   "YZ",
		702:   "ZZ",
		705:   "AAC",
		16384: "XFD",
	}

	for n, w := range tt {
		if res := colRunes(n); string(res) != w {
			t.Errorf("colRunes(%d) = %c, want %c", n, res, []rune(w))
		}
	}
}

func TestColNum(t *testing.T) {

	tt := map[string]uint16{
		"ф1":  0,
		"11":  0,
		"ЁЁ":  0,
		"A":   1,
		"B":   2,
		"E":   5,
		"Z":   26,
		"aa":  27,
		"AB":  28,
		"AZ":  52,
		"YZ":  676,
		"ZZ":  702,
		"aac": 705,
		"XFD": 16384,
	}

	for n, w := range tt {
		if res, err := colNum(n); res != w {
			t.Errorf("colNum(%s) = %d, %v, want %d", n, res, err, w)
		}
	}

}

func TestCellAddrString(t *testing.T) {
	tt := map[string]CellAddr{
		"A1":   {0, 0},
		"A2":   {0, 1},
		"XFD3": {16383, 2},
	}
	for w, a := range tt {
		if a.String() != w {
			t.Errorf("CellAddr{%d, %d}.String() = %v, want %s", a.Col, a.Row, a, w)
		}
	}
}

func TestNewCellAddr(t *testing.T) {
	tt := map[string]struct {
		res CellAddr
		err bool
	}{
		"a1":    {CellAddr{0, 0}, false},
		"b5":    {CellAddr{1, 4}, false},
		"Z2303": {CellAddr{25, 2302}, false},
		"AA23":  {CellAddr{26, 22}, false},
		"ЁцЭ":   {emptyCellAddr, true},
		"":      {emptyCellAddr, true},
		"5A1":   {emptyCellAddr, true},
		"XFD3":  {CellAddr{16383, 2}, false},
	}

	for a, w := range tt {
		addr, err := NewCellAddr(a)

		isErr := (err != nil)

		if addr.String() != w.res.String() || isErr != w.err {
			t.Errorf(
				"NewCellAddr(%s) = (%v, %v), want %v", a, addr, err, w.res,
			)
		}
	}
}

func TestNewRange(t *testing.T) {
	tt := map[string]bool{
		"A1:A2":      false,
		"A2:A1":      false,
		"aa23:XFD27": false,
		"sd":         true,
		"5F:Ad":      true,
	}

	for r, e := range tt {
		if res, err := NewRange(r); (err != nil) != e {
			t.Errorf("NewRange(%s) = (%v, %v), want %s", r, res, err, r)
		}
	}

}

func TestRangeString(t *testing.T) {
	tt := map[Range]string{
		Range{CellAddr{0, 0}, CellAddr{16383, 2}}: "A1:XFD3",
		Range{CellAddr{1, 4}, CellAddr{25, 2302}}: "B5:Z2303",
	}

	for r, w := range tt {
		str := r.String()
		if str != w {
			t.Errorf("%#v.String() = %s, want %s", r, str, w)
		}
	}
}
