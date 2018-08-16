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
	"github.com/zetamatta/go-texts/mbcs"
)

const version = "0.4"

var versionOption = flag.Bool("v", false, "Show version")
var quitOption = flag.Bool("q", false, "Quit immediately")
var saveAsOption = flag.String("o", "", "SaveAs")
var hideAsOption = flag.Bool("s", false, "Silent mode (hide Excel)")
var fieldSeperator = flag.String("f", ",", "Field Sperator")

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
	if len(*fieldSeperator) != 1 {
		return fmt.Errorf("Invalid lendth of field seperator(%s)", *fieldSeperator)
	}
	for _, v := range *fieldSeperator {
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
		return parseCsvReader(mbcs.NewAutoDetectReader(os.Stdin, mbcs.ACP), f)
	}
	if err := f.NewSheet(filepath.Base(fname)); err != nil {
		return errors.Wrap(err, "parseCsvFile")
	}
	fd, err := os.Open(fname)
	if err != nil {
		return err
	}
	defer fd.Close()
	reader := mbcs.NewAutoDetectReader(fd, mbcs.ACP)
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
		send2excel, err = NewSendCsvToExcel(!*hideAsOption)
		if err != nil {
			return err
		}
		if *quitOption {
			send2excel.SetDoQuit(true)
		}
	}
	defer send2excel.Close()

	if *hideAsOption {
		*quitOption = true
		if *saveAsOption == "" {
			return errors.New("-s option requires `-o FILENAME`")
		}
	}

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
