package app

import "testing"

func TestFormatColumn(t *testing.T) {
	data := []struct {
		in  uint64
		out string
	}{
		{in: 0, out: "A"},
		{in: 1, out: "B"},
		{in: 25, out: "Z"},
		{in: 26, out: "AA"},
		{in: 27, out: "AB"},
		{in: 51, out: "AZ"},
		{in: 52, out: "BA"},
		{in: 53, out: "BB"},
		{in: 675, out: "YZ"},
		{in: 676, out: "ZA"},
		{in: 701, out: "ZZ"},
		{in: 702, out: "AAA"},
		{in: 703, out: "AAB"},
		{in: 728, out: "ABA"},
		{in: 729, out: "ABB"},
	}
	for _, test := range data {
		out := formatColumn(test.in)
		if out != test.out {
			t.Errorf("formatColumn(%v) = %v want %v", test.in, out, test.out)
		}
	}
}

func TestParseColumn(t *testing.T) {
	data := []struct {
		in  string
		out uint64
	}{
		{in: "A", out: 0},
		{in: "B", out: 1},
		{in: "Z", out: 25},
		{in: "AA", out: 26},
		{in: "AB", out: 27},
		{in: "AZ", out: 51},
		{in: "BA", out: 52},
		{in: "BB", out: 53},
		{in: "YZ", out: 675},
		{in: "ZA", out: 676},
		{in: "ZZ", out: 701},
		{in: "AAA", out: 702},
		{in: "AAB", out: 703},
		{in: "ABA", out: 728},
		{in: "ABB", out: 729},
		{in: "0", out: 0},
	}
	for _, test := range data {
		out := parseColumn(test.in)
		if out != test.out {
			t.Errorf("parseColumn(%v) = %v want %v", test.in, out, test.out)
		}
	}
}
