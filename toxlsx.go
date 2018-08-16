package main

import (
	"github.com/tealeg/xlsx"
)

type SendCsvToXlsx struct {
	file  *xlsx.File
	sheet *xlsx.Sheet
	fname string
}

func NewSendCsvToXlsx() SendCsv {
	xlsx.SetDefaultFont(11, "MS PGothic")
	return &SendCsvToXlsx{
		file: xlsx.NewFile(),
	}
}

func (this *SendCsvToXlsx) NewSheet(name string) error {
	sheet, err := this.file.AddSheet(name)
	if err != nil {
		return err
	}
	this.sheet = sheet
	return nil
}

func (this *SendCsvToXlsx) Send(csv1 []string) error {
	row := this.sheet.AddRow()
	for _, val := range csv1 {
		cell := row.AddCell()
		cell.SetString(val)
	}
	return nil
}

func (this *SendCsvToXlsx) Close() {
	if this.fname != "" {
		this.file.Save(this.fname)
	}
}

func (this *SendCsvToXlsx) SetDoQuit(bool) {
}

func (this *SendCsvToXlsx) SetSaveAs(name string) {
	this.fname = name
}
