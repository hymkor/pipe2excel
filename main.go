package main

import (
	"bufio"
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"unicode/utf8"

	"github.com/mattn/go-isatty"
	"github.com/pkg/errors"
	"github.com/zetamatta/go-mbcs"
)

var optionUtf8 = flag.Bool("u", false, "Read CSV as UTF8")

// SendCsv is the interface to send csv somewhere
type SendCsv interface {
	Send(csv1 []string) error
	NewSheet(name string) error
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
		if err := f.Send(csv1); err != nil {
			return err
		}
	}
}

// SendCsvToStdout is the sample class implmentation for SendCsv interface
type SendCsvToStdout struct{}

// SendCsvToStdout output csv-line to STDOUT
func (*SendCsvToStdout) Send(csv []string) error {
	fmt.Printf("<%s>\n", strings.Join(csv, "> <"))
	return nil
}

func mbcsReader(fd io.Reader, onError func(error, io.Writer) bool) io.ReadCloser {
	reader, writer := io.Pipe()
	go func() {
		sc := bufio.NewScanner(fd)
		defer writer.Close()
		notUtf8 := false
		for sc.Scan() {
			line := sc.Bytes()
			if !notUtf8 && utf8.Valid(line) {
				fmt.Fprintln(writer, string(line))
			} else {
				text, err := mbcs.AtoU(line)
				if err != nil {
					if !onError(err, writer) {
						return
					}
				} else {
					notUtf8 = true
					fmt.Fprintln(writer, text)
				}
			}
		}
	}()
	return reader
}

func onError(err error, w io.Writer) bool {
	fmt.Fprintf(w, "<%s>\n", err.Error())
	return true
}

func parseCsvFile(fname string, f SendCsv) error {
	if fname == "-" {
		if *optionUtf8 || isatty.IsTerminal(os.Stdin.Fd()) {
			return parseCsvReader(os.Stdin, f)
		} else {
			return parseCsvReader(mbcsReader(os.Stdin, onError), f)
		}
	}
	if err := f.NewSheet(filepath.Base(fname)); err != nil {
		return errors.Wrap(err, "on parseCsvFile")
	}
	fd, err := os.Open(fname)
	if err != nil {
		return err
	}
	defer fd.Close()
	var reader io.ReadCloser
	if *optionUtf8 {
		reader = fd
	} else {
		reader = mbcsReader(fd, onError)
		defer reader.Close()
	}
	return parseCsvReader(reader, f)
}

func main1(args []string) error {
	if len(args) <= 0 {
		args = []string{"-"}
	}
	send2excel, err := NewSendCsvToExcel()
	if err != nil {
		return err
	}
	defer send2excel.Close()
	// send2excel := &SendCsvToStdout{}
	for _, arg1 := range args {
		if err := parseCsvFile(arg1, send2excel); err != nil {
			return err
		}
	}
	return nil
}

func main() {
	flag.Parse()
	if err := main1(flag.Args()); err != nil {
		fmt.Fprintln(os.Stderr, err.Error())
		os.Exit(1)
	}
	os.Exit(0)
}
