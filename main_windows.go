package main

import (
	"github.com/zetamatta/go-texts/mbcs"
	"io"
)

const supportole = true

func mbcsReader(fd io.Reader) io.Reader {
	return mbcs.NewAutoDetectReader(fd, mbcs.ACP)
}
