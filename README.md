Pipe To Excel
=============
[![GoDev](https://pkg.go.dev/badge/github.com/hymkor/pipe2excel)](https://pkg.go.dev/github.com/hymkor/pipe2excel)

`pipe2excel` sends the contents of STDIN 
or files of arguments to Excel as CSV to Microsoft Excel.

- The each value of the csv is inserted as a string.
    - Only matching `/^\-?[1-9]\d*(\.\d*[1-9])?$/`, as a number
- The encoding of the CSV is detected automatically whether it is written in UTF8 or the encoding of the current codepage.

Install
-------

Download the binary package from [Releases](https://github.com/hymkor/pipe2excel/releases) and extract the executable.

### for scoop-installer

```
scoop install https://raw.githubusercontent.com/hymkor/pipe2excel/master/pipe2excel.json
```

or

```
scoop bucket add hymkor https://github.com/hymkor/scoop-bucket
scoop install pipe2excel
```

How to use
----------

### OLE-MODE (Supported only Windows)

```
C:\> pipe2excel foo.csv bar.csv
```

```
C:\> type foo.csv | pipe2excel
```

It starts Microsoft Excel.

### Non-OLE (Both Windows and Linux Ok)

```
C:\> pipe2excel -o foo.xlsx foo.csv
```

It does not start Microsoft Excel. It creates foo.xlsx statically.

### Options

* `-f string` Field Sperator (default ",")
* `-o string` Save to file and quit immediately without EXCEL.EXE
* `-v` Show version

Source CSV data (sample)
------------------------

![image](foo-csv.png)

Destinate Excel data (sample)
----------------------------

![image](foo-xls.png)

history
-------
- v0.5.2 (Feb 19,2022)
    - Fix package dependency problems
    - Move or copy some packages into internal directory
    - Change package owner to hymkor
- v0.5.1 (May 28,2021)
    - Support Linux as platform (but required -o option always)
    - (#1) Fix the panic on reading CSV from STDIN and using -o
- v0.5.0 (Aug 08,2018)
    - When `-o FILENAME` is given, use "[tealeg/xlsx](https://github.com/tealeg/xlsx)" instead of "[go-ole/go-ole](https://github.com/go-ole/go-ole)"
    - Remove options -s and -q. Their features are enabled with -o.
- v0.4.0 (Aug 05,2018)
    - Add -f option to set field seperator
    - Do not treat as string when the value is a negative integer.
- v0.3.0 (Jul 11,2018)
    - Fix leak release COM
    - Print help if no arguments and stdin is not redirected
- v0.2.0 (Jul 1,2018)
    - Only matching `/^[1-9]\d*(\.\d*[1-9])?$/`, as a number
- v0.1.0 (Jul 1,2018)
    - Set cell as a string
    - Detect encoding utf8 or codepage automatically and remove -u option
    - Add -v option to show version
- v0.0.1 (Jul 1,2018)
    - prototype
