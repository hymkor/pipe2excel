package main

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/pkg/errors"
)

type SendCsvToExcel struct {
	*WorkSheet
	row int
}

func NewSendCsvToExcel() (*SendCsvToExcel, error) {
	book, err := NewWorkbook()
	if err != nil {
		return nil, err
	}
	defer book.Release()
	sheet, err := book.Item(1)
	if err != nil {
		return nil, err
	}
	return &SendCsvToExcel{
		WorkSheet: sheet,
		row:       1,
	}, nil
}

func (this *SendCsvToExcel) Close() {
	if this.WorkSheet != nil {
		this.WorkSheet.Release()
		this.WorkSheet = nil
		ole.CoUninitialize()
	}
}

func (this *SendCsvToExcel) Send(csv []string) error {
	for key, val := range csv {
		_cell, err := oleutil.GetProperty(this.IDispatch, "Cells", this.row, key+1)
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
