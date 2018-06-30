package main

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type SendCsvToExcel struct {
	worksheet *ole.IDispatch
	row       int
}

func NewSendCsvToExcel() (*SendCsvToExcel, error) {
	ole.CoInitialize(0)

	_excel, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		return nil, err
	}
	excel, err := _excel.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, err
	}
	defer excel.Release()

	oleutil.PutProperty(excel, "Visible", true)

	_workbooks, err := oleutil.GetProperty(excel, "Workbooks")
	if err != nil {
		return nil, err
	}
	workbooks := _workbooks.ToIDispatch()
	defer workbooks.Release()

	_workbook, err := oleutil.CallMethod(workbooks, "Add", nil)
	if err != nil {
		return nil, err
	}
	workbook := _workbook.ToIDispatch()
	defer workbook.Release()

	_worksheet, err := oleutil.GetProperty(workbook, "Worksheets", 1)
	if err != nil {
		return nil, err
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
			return err
		}
		cell := _cell.ToIDispatch()
		oleutil.PutProperty(cell, "Value", val)
		cell.Release()
	}
	this.row++
	return nil
}
