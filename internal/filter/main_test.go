package filter

import (
	"bytes"
	"strings"

	"testing"
)

func TestFilter(t *testing.T) {
	const testStr = `1234567890\nabcdefghijklmnopqrstuvwxyz\nABCDEFGHIJLMNOPQRSTUVWXYZ`
	source := strings.NewReader(testStr)
	r := Filter{
		In: func() ([]byte, error) {
			var tmp [8]byte
			n, err := source.Read(tmp[:])
			return tmp[:n], err
		},
	}
	var output []byte
	for {
		var buffer [7]byte
		n, err := r.Read(buffer[:])
		output = append(output, buffer[:n]...)
		if err != nil {
			break
		}
	}
	if !bytes.Equal(output, []byte(testStr)) {
		t.Fail()
	}
	// println(string(output))
	// println(testStr)
}
