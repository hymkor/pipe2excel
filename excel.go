package main

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/pkg/errors"
)

type Workbook struct {
	*ole.IDispatch
}

func NewWorkbook() (*Workbook, error) {
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
	_workbook, err := oleutil.CallMethod(workbooks, "Add", nil)
	if err != nil {
		return nil, errors.Wrap(err, "on NewWorkbook.Add")
	}
	return &Workbook{_workbook.ToIDispatch()}, nil
}

type WorkSheet struct {
	*ole.IDispatch
}

func (b *Workbook) Item(name interface{}) (*WorkSheet, error) {
	_worksheet, err := oleutil.GetProperty(b.IDispatch, "Worksheets", name)
	if err != nil {
		return nil, errors.Wrap(err, "on Workbook.Item")
	}
	return &WorkSheet{_worksheet.ToIDispatch()}, nil
}
