package main

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/pkg/errors"
)

type SendCsvToExcel struct {
	worksheet *ole.IDispatch
	row       int
}

func NewSendCsvToExcel() (*SendCsvToExcel, error) {
	ole.CoInitializeEx(0, 0)

	_excel, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		return nil, errors.Wrap(err, "on NewSendCsvToExcel")
	}
	excel, err := _excel.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, errors.Wrap(err, "on NewSendCsvToExcel")
	}
	defer excel.Release()

	oleutil.PutProperty(excel, "Visible", true)

	_workbooks, err := oleutil.GetProperty(excel, "Workbooks")
	if err != nil {
		return nil, errors.Wrap(err, "on NewSendCsvToExcel")
	}
	workbooks := _workbooks.ToIDispatch()
	defer workbooks.Release()

	_workbook, err := oleutil.CallMethod(workbooks, "Add", nil)
	if err != nil {
		return nil, errors.Wrap(err, "on NewSendCsvToExcel")
	}
	workbook := _workbook.ToIDispatch()
	defer workbook.Release()

	_worksheet, err := oleutil.GetProperty(workbook, "Worksheets", 1)
	if err != nil {
		return nil, errors.Wrap(err, "on NewSendCsvToExcel")
	}

	return &SendCsvToExcel{
		worksheet: _worksheet.ToIDispatch(),
		row:       1,
	}, nil
}

func (this *SendCsvToExcel) Close() {
	if this.worksheet != nil {
		this.worksheet.Release()
		this.worksheet = nil
		ole.CoUninitialize()
	}
}

func (this *SendCsvToExcel) Send(csv []string) error {
	for key, val := range csv {
		_cell, err := oleutil.GetProperty(this.worksheet, "Cells", this.row, key+1)
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
