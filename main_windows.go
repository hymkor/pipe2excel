package main

import (
	"bufio"
	"io"
	"unicode/utf8"

	"github.com/nyaosorg/go-windows-mbcs"
)

const supportole = true

type _MbcsReader struct {
	reader *bufio.Reader
	rest   []byte
	eof    bool
}

func (this *_MbcsReader) Read(buffer []byte) (int, error) {
	copiedBytes := 0
	for len(buffer) > 0 {
		if this.eof {
			if this.rest == nil || len(this.rest) <= 0 {
				return copiedBytes, io.EOF
			}
		} else {
			bytes, err := this.reader.ReadBytes('\n')
			if err != nil {
				if err != io.EOF {
					return copiedBytes, err
				}
				this.eof = true
			}
			if utf8.Valid(bytes) {
				this.rest = append(this.rest, bytes...)
			} else {
				text, _err := mbcs.AtoU(bytes, mbcs.ACP)
				if _err != nil {
					return copiedBytes, _err
				}
				this.rest = append(this.rest, []byte(text)...)
			}
		}
		n := copy(buffer, this.rest)
		buffer = buffer[n:]
		copiedBytes += n
		newlen := len(this.rest[n:])
		copy(this.rest, this.rest[n:])
		this.rest = this.rest[:newlen]
	}
	return copiedBytes, nil
}

func mbcsReader(fd io.Reader) io.Reader {
	return &_MbcsReader{reader: bufio.NewReader(fd)}
}
