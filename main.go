package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"runtime"

	"github.com/mattn/go-isatty"
	"github.com/pkg/errors"
)

const version = "0.5"

var versionOption = flag.Bool("v", false, "Show version")
var saveAsOption = flag.String("o", "", "Save to file and quit immediately without EXCEL.EXE")
var fieldSeparater = flag.String("f", ",", "Field Separater")

// SendCsv is the interface to send csv somewhere
type SendCsv interface {
	Send(csv1 []string) error
	NewSheet(name string) error
	Close()
	SetDoQuit(bool)
	SetSaveAs(string)
}

func parseCsvReader(r io.Reader, f SendCsv) error {
	reader := csv.NewReader(r)
	reader.FieldsPerRecord = -1 // do not error how many field exists per line
	if len(*fieldSeparater) != 1 {
		return fmt.Errorf("Invalid lendth of field seperator(%s)", *fieldSeparater)
	}
	for _, v := range *fieldSeparater {
		reader.Comma = v
	}
	for {
		csv1, err := reader.Read()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return err
		}
		if err := f.Send(csv1); err != nil {
			return err
		}
	}
}

func parseCsvFile(fname string, f SendCsv) error {
	if fname == "-" {
		return parseCsvReader(mbcsReader(os.Stdin), f)
	}
	if err := f.NewSheet(filepath.Base(fname)); err != nil {
		return errors.Wrap(err, "parseCsvFile")
	}
	fd, err := os.Open(fname)
	if err != nil {
		return err
	}
	defer fd.Close()
	reader := mbcsReader(fd)
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

	var send2excel SendCsv
	if *saveAsOption != "" {
		send2excel = NewSendCsvToXlsx()
		if fullpath, err := filepath.Abs(*saveAsOption); err == nil {
			send2excel.SetSaveAs(fullpath)
		}
	} else {
		var err error
		send2excel, err = NewSendCsvToExcel(true)
		if err != nil {
			return err
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
	if ! supportole && *saveAsOption == "" {
		fmt.Fprintf(os.Stderr,"%s: -o option requires on not Windows platform.\n",os.Args[0])
		os.Exit(1)
	}

	if err := main1(flag.Args()); err != nil {
		fmt.Fprintln(os.Stderr, err.Error())
		os.Exit(1)
	}
	os.Exit(0)
}
