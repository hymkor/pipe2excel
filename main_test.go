package main

import (
	"bytes"
	"strings"

	"testing"
)

func TestMbcsReader(t *testing.T) {
	const testStr = `1234567890\nabcdefghijklmnopqrstuvwxyz\nABCDEFGHIJLMNOPQRSTUVWXYZ`
	r := mbcsReader(strings.NewReader(testStr))
	var buffer [4]byte
	var output []byte
	for {
		n, err := r.Read(buffer[:])
		output = append(output, buffer[:n]...)
		if err != nil {
			break
		}
	}
	if !bytes.Equal(output, []byte(testStr)) {
		t.Fail()
	}
}
