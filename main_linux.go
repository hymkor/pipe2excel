package main

import (
	"io"
)

const supportole = false

func mbcsReader(fd io.Reader) io.Reader {
	return fd
}
