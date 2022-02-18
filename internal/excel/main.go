package excel

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/pkg/errors"
)

type Application struct {
	*ole.IDispatch
}

func New(visible bool) (*Application, error) {
	ole.CoInitializeEx(0, 0)

	_excel, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		return nil, errors.Wrap(err, "excel.New: CreateObject")
	}
	excel, err := _excel.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, errors.Wrap(err, "excel.New: QueryInterface")
	}
	if visible {
		excel.PutProperty("Visible", true)
	}
	return &Application{excel}, nil
}

func (this *Application) Close() {
	this.Release()
	ole.CoUninitialize()
}

type Book struct {
	*ole.IDispatch
}

func (excel *Application) NewBook() (*Book, error) {
	_workbooks, err := excel.GetProperty("Workbooks")
	if err != nil {
		return nil, errors.Wrap(err, "(*excel.Application)NewBook: GetProperty(\"WorkBooks\")")
	}
	workbooks := _workbooks.ToIDispatch()
	_workbook, err := workbooks.CallMethod("Add", nil)
	if err != nil {
		return nil, errors.Wrap(err, "(*excel.Application)NewBook: CallMethod(\"Add\")")
	}
	workbooks.Release()
	return &Book{_workbook.ToIDispatch()}, nil
}

func (excel *Application) Open(fname string) (*Book, error) {
	_workbooks, err := excel.GetProperty("Workbooks")
	if err != nil {
		return nil, errors.Wrap(err, "(*excel.Application)Open: GetProperty(\"WorkBooks\")")
	}
	workbooks := _workbooks.ToIDispatch()
	_workbook, err := workbooks.CallMethod("Open", fname)
	if err != nil {
		return nil, errors.Wrap(err, "(*excel.Application)Open: CallMethod(\"Open\")")
	}
	workbooks.Release()
	return &Book{_workbook.ToIDispatch()}, nil
}

type Sheet struct {
	*ole.IDispatch
}

type Sheets struct {
	*ole.IDispatch
}

func (b *Book) Item(name interface{}) (*Sheet, error) {
	_worksheet, err := b.GetProperty("Worksheets", name)
	if err != nil {
		return nil, errors.Wrap(err, "(*excel.Book)Item")
	}
	return &Sheet{_worksheet.ToIDispatch()}, nil
}

func (b *Book) Sheets() (*Sheets, error) {
	_sheets, err := b.GetProperty("Sheets")
	if err != nil {
		return nil, errors.Wrap(err, "(*excel.Book)Sheets")
	}
	return &Sheets{_sheets.ToIDispatch()}, nil
}

func (st *Sheets) Count() (int, error) {
	value, err := st.GetProperty("Count")
	if err != nil {
		return -1, errors.Wrap(err, "(*excel.Sheets)Count")
	}
	return int(value.Val), nil
}

func (b *Book) Add() (*Sheet, error) {
	sheets, err := b.Sheets()
	if err != nil {
		return nil, errors.Wrap(err, "(*excel.Book)Add")
	}
	defer sheets.Release()
	_worksheet, err := sheets.CallMethod("Add", nil)
	if err != nil {
		return nil, errors.Wrap(err, "(*excel.Book)Add")
	}
	return &Sheet{_worksheet.ToIDispatch()}, nil
}

func (s *Sheet) SetName(name string) error {
	if _, err := s.PutProperty("Name", name); err != nil {
		return errors.Wrap(err, "(*excel.Sheet)SetName")
	}
	return nil
}
