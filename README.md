Pipe To Excel
=============
[![GoDev](https://pkg.go.dev/badge/github.com/hymkor/pipe2excel)](https://pkg.go.dev/github.com/hymkor/pipe2excel)

`pipe2excel` sends the contents of STDIN or specified files to Microsoft Excel as CSV data.

- Each CSV value is inserted as a string.
    - However, values matching `/^\-?[1-9]\d*(\.\d*[1-9])?$/` are treated as numbers.
- The CSV encoding is automatically detected, supporting both UTF-8 and the system's current code page.

Install
-------

You can install pipe2excel in the following ways:

### Downloading the Binary

Download the binary package from the [Releases] page and extract the executable.

[Releases]: https://github.com/hymkor/pipe2excel/releases

### Using `go install`

```
go install github.com/hymkor/pipe2excel@latest
```

### Using Scoop (Windows)

```
scoop install https://raw.githubusercontent.com/hymkor/pipe2excel/master/pipe2excel.json
```

or

```
scoop bucket add hymkor https://github.com/hymkor/scoop-bucket
scoop install hymkor/pipe2excel
```

How to Use
----------

### OLE Mode (Windows Only)

```
C:\> pipe2excel foo.csv bar.csv
```

```
C:\> type foo.csv | pipe2excel
```

This launches Microsoft Excel and opens the CSV files.

### Non-OLE Mode (Windows &amp; Linux)

```
C:\> pipe2excel -o foo.xlsx foo.csv
```

It does not start Microsoft Excel. Instead, it creates `foo.xlsx` directly.

### Options

* `-f string` → Field Sperator (default `,`)
* `-o string` → Save to a file and exit immediately without starting Excel
* `-v` → Show version information

Source CSV Data (Sample)
------------------------

![image](foo-csv.png)

Output Excel Data (Sample)
--------------------------

![image](foo-xls.png)

History
-------

### v0.5.2 (Feb 19 2022)

- Fixed package dependency issues.
- Moved or copied some packages into the `internal` directory.
- Changed package owner to `hymkor`.

### v0.5.1 (May 28, 2021)

- Added Linux support (requires the `-o` option).
- [#1] Fixed a panic when reading CSV from STDIN and using `-o`.

[#1]: https://github.com/hymkor/pipe2excel/issues/1

### v0.5.0 (Aug 08, 2018)

- When `-o FILENAME` is specified, use [tealeg/xlsx](https://github.com/tealeg/xlsx) instead of [go-ole/go-ole](https://github.com/go-ole/go-ole).
- Removed `-s` and `-q` options; their features are now enabled by `-o`.

### v0.4.0 (Aug 05, 2018)

- Added the `-f` option to set the field separator.
- No longer treats negative integers as strings.

### v0.3.0 (Jul 11, 2018)

- Fixed COM release leak.
- Display help if no arguments are provided and STDIN is not redirected.

### 0.2.0 (Jul 01, 2018)

- Only treats values matching `/^[1-9]\d*(\.\d*[1-9])?$/` as numbers.

### v0.1.0 (Jul 01, 2018)

- Inserted all cells as strings.
- Automatically detected encoding (UTF-8 or system code page) and removed the `-u` option.
- Added the `-v` option to display the version.

### v0.0.1 (Jul 01, 2018)

- Initial prototype.
