package main

import (
	"bufio"
	"io"
	"unicode/utf8"

	"github.com/nyaosorg/go-windows-mbcs"

	"github.com/hymkor/pipe2excel/internal/filter"
)

const supportole = true

func mbcsReader(fd io.Reader) io.Reader {
	br := bufio.NewReader(fd)
	return &filter.Filter{
		In: func() ([]byte, error) {
			bin, err := br.ReadBytes('\n')
			if err != nil && err != io.EOF {
				return nil, err
			}
			if utf8.Valid(bin) {
				return bin, err
			}
			text, _err := mbcs.AtoU(bin, mbcs.ACP)
			if _err != nil {
				return nil, _err
			}
			return []byte(text), err
		},
	}
}
