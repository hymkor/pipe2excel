package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"runtime"

	"github.com/pkg/errors"
	"github.com/zetamatta/go-mbcs"
	"github.com/mattn/go-isatty"
)

const version = "0.3"

var versionOption = flag.Bool("v", false, "Show version")
var quitOption = flag.Bool("q", false, "Quit immediately")
var saveAsOption = flag.String("o", "", "SaveAs")
var hideAsOption = flag.Bool("s", false, "Silent mode (hide Excel)")

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

func onError(err error, w io.Writer) bool {
	fmt.Fprintf(w, "<%s>\n", err.Error())
	return true
}

func parseCsvFile(fname string, f SendCsv) error {
	if fname == "-" {
		return parseCsvReader(mbcs.NewReader(os.Stdin), f)
	}
	if err := f.NewSheet(filepath.Base(fname)); err != nil {
		return errors.Wrap(err, "parseCsvFile")
	}
	fd, err := os.Open(fname)
	if err != nil {
		return err
	}
	defer fd.Close()
	reader := mbcs.NewReader(fd)
	return parseCsvReader(reader, f)
}

func main1(args []string) error {
	if len(args) <= 0 {
		if isatty.IsTerminal(os.Stdin.Fd()) {
			flag.PrintDefaults()
			return nil
		}
		args = []string{"-"}
	}

	if *hideAsOption {
		*quitOption = true
		if *saveAsOption == "" {
			return errors.New("-s option requires `-o FILENAME`")
		}
	}

	send2excel, err := NewSendCsvToExcel(!*hideAsOption)
	if err != nil {
		return err
	}
	if *quitOption {
		send2excel.DoQuit = true
	}
	if *saveAsOption != "" {
		if fullpath, err := filepath.Abs(*saveAsOption); err == nil {
			send2excel.SaveAs = fullpath
		}
	}
	defer send2excel.Close()
	for _, arg1 := range args {
		if err := parseCsvFile(arg1, send2excel); err != nil {
			return err
		}
	}
	return nil
}

func main() {
	flag.Parse()
	if *versionOption {
		fmt.Printf("%s-%s\n", version, runtime.GOARCH)
		os.Exit(0)
	}
	if err := main1(flag.Args()); err != nil {
		fmt.Fprintln(os.Stderr, err.Error())
		os.Exit(1)
	}
	os.Exit(0)
}
