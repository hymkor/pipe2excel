package main

import (
	"regexp"

	"github.com/pkg/errors"

	"github.com/hymkor/pipe2excel/excel"
)

type SendCsvToExcel struct {
	excel  *excel.Application
	sheet  *excel.Sheet
	book   *excel.Book
	row    int
	DoQuit bool
	SaveAs string
}

func (this *SendCsvToExcel) SetDoQuit(value bool) {
	this.DoQuit = value
}

func (this *SendCsvToExcel) SetSaveAs(name string) {
	this.SaveAs = name
}

func NewSendCsvToExcel(visible bool) (SendCsv, error) {
	excel1, err := excel.New(visible)
	if err != nil {
		return nil, errors.Wrap(err, "NewSendCsvToExcel")
	}
	book, err := excel1.NewBook()
	if err != nil {
		return nil, err
	}
	sheet, err := book.Item(1)
	if err != nil {
		book.Release()
		return nil, err
	}
	return &SendCsvToExcel{
		excel: excel1,
		sheet: sheet,
		book:  book,
		row:   1,
	}, nil
}

func (this *SendCsvToExcel) Close() {
	if this.sheet != nil {
		this.sheet.Release()
		this.sheet = nil
	}
	if this.book != nil {
		if this.SaveAs != "" {
			this.book.CallMethod("SaveAs", this.SaveAs)
		}
		this.book.Release()
		this.book = nil
	}
	if this.excel != nil {
		if this.DoQuit {
			this.excel.CallMethod("Quit")
		}
		this.excel.Close()
		this.excel = nil
	}
}

func (this *SendCsvToExcel) NewSheet(name string) error {
	this.sheet.Release()
	s, err := this.book.Add()
	if err != nil {
		return err
	}
	this.sheet = s
	s.SetName(name)
	this.row = 1
	return nil
}

var rxNumber = regexp.MustCompile(`^\-?[1-9]\d*(\.\d*[1-9])?$`)

func (this *SendCsvToExcel) Send(csv []string) error {
	for key, val := range csv {
		_cell, err := this.sheet.GetProperty("Cells", this.row, key+1)
		if err != nil {
			return errors.Wrap(err, "(*SendCsvToExcel)Send")
		}
		cell := _cell.ToIDispatch()
		if rxNumber.MatchString(val) {
			cell.PutProperty("NumberFormatLocal", "0_")
		} else {
			cell.PutProperty("NumberFormatLocal", "@")
		}
		cell.PutProperty("Value", val)
		cell.Release()
	}
	this.row++
	return nil
}
