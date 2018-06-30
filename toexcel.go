package main

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/pkg/errors"
)

type SendCsvToExcel struct {
	sheet *Sheet
	book  *Book
	row   int
}

func NewSendCsvToExcel() (*SendCsvToExcel, error) {
	book, err := NewBook()
	if err != nil {
		return nil, err
	}
	sheet, err := book.Item(1)
	if err != nil {
		book.Release()
		return nil, err
	}
	return &SendCsvToExcel{
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

func (this *SendCsvToExcel) Send(csv []string) error {
	for key, val := range csv {
		_cell, err := oleutil.GetProperty(this.sheet.IDispatch, "Cells", this.row, key+1)
		if err != nil {
			return errors.Wrap(err, "on SendCsvToExcel.Send")
		}
		cell := _cell.ToIDispatch()
		oleutil.PutProperty(cell, "Value", val)

		text, err := oleutil.GetProperty(cell, "Text")
		if err == nil && text.ToString() != val {
			oleutil.PutProperty(cell, "Value", "'"+val)
		}
		cell.Release()
	}
	this.row++
	return nil
}
