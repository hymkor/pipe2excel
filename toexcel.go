package main

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/pkg/errors"
	"regexp"
)

type SendCsvToExcel struct {
	excel  *Excel
	sheet  *Sheet
	book   *Book
	row    int
	DoQuit bool
}

func NewSendCsvToExcel() (*SendCsvToExcel, error) {
	excel, err := NewExcel()
	if err != nil {
		return nil, errors.Wrap(err, "NewSendCsvToExcel")
	}
	book, err := excel.NewBook()
	if err != nil {
		return nil, err
	}
	sheet, err := book.Item(1)
	if err != nil {
		book.Release()
		return nil, err
	}
	return &SendCsvToExcel{
		excel: excel,
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
		this.book.Release()
		this.book = nil
	}
	if this.excel != nil {
		if this.DoQuit {
			this.excel.CallMethod("Quit")
		}
		this.excel.Release()
		this.excel = nil
	}
	ole.CoUninitialize()
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

var rxNumber = regexp.MustCompile(`^[1-9]\d*(\.\d*[1-9])?$`)

func (this *SendCsvToExcel) Send(csv []string) error {
	for key, val := range csv {
		_cell, err := oleutil.GetProperty(this.sheet.IDispatch, "Cells", this.row, key+1)
		if err != nil {
			return errors.Wrap(err, "(*SendCsvToExcel)Send")
		}
		cell := _cell.ToIDispatch()
		if rxNumber.MatchString(val) {
			oleutil.PutProperty(cell, "NumberFormatLocal", "0_")
		} else {
			oleutil.PutProperty(cell, "NumberFormatLocal", "@")
		}
		oleutil.PutProperty(cell, "Value", val)
		cell.Release()
	}
	this.row++
	return nil
}
