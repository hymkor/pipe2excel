package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"os"
	"reflect"
	"strings"
)

// SendCsv is the interface to send csv somewhere
type SendCsv interface {
	Send(csv1 []string)
}

func isErrFieldCount(err error) bool {
	if err == csv.ErrFieldCount {
		return true
	}
	if e, ok := err.(*csv.ParseError); ok && e.Err == csv.ErrFieldCount {
		return true
	}
	return false
}

func parseCsvReader(r io.Reader, f SendCsv) error {
	reader := csv.NewReader(r)
	for {
		csv1, err := reader.Read()
		if err != nil && !isErrFieldCount(err) {
			if err == io.EOF {
				return nil
			}
			return err
		}
		f.Send(csv1)
	}
}

// SendCsvToStdout is the sample class implmentation for SendCsv interface
type SendCsvToStdout struct{}

// SendCsvToStdout output csv-line to STDOUT
func (*SendCsvToStdout) Send(csv []string) {
	fmt.Printf("<%s>\n", strings.Join(csv, "> <"))
}

func parseCsvFile(fname string, f SendCsv) error {
	if fname == "-" {
		return parseCsvReader(os.Stdin, f)
	}
	fd, err := os.Open(fname)
	if err != nil {
		return err
	}
	defer fd.Close()
	return parseCsvReader(fd, f)
}

func main1(args []string) error {
	if len(args) <= 0 {
		args = []string{"-"}
	}
	for _, arg1 := range args {
		if err := parseCsvFile(arg1, &SendCsvToStdout{}); err != nil {
			return err
		}
	}
	return nil
}

func main() {
	if err := main1(os.Args[1:]); err != nil {
		fmt.Fprintf(os.Stderr, "%s (%s)\n",
			err.Error(),
			reflect.TypeOf(err).String())
		os.Exit(1)
	}
	os.Exit(0)
}
