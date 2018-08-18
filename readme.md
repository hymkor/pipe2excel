Pipe To Excel
=============

`pipe2excel.exe` sends the contents of STDIN 
or files of arguments to Excel as CSV to Microsoft Excel.

- The each value of the csv is inserted as a string.
    - Only matching `/^\-?[1-9]\d*(\.\d*[1-9])?$/`, as a number
- The encoding of the CSV is detected automatically whether it is written in UTF8 or the encoding of the current codepage.


How to use
----------

```
C:\> pipe2excel foo.csv bar.csv
```

```
C:\> type foo.csv | pipe2excel
```

Source CSV data (sample)
------------------------

<img src="foo-csv.png" />

Destinate Excel data (sample)
----------------------------

<img src="foo-xls.png" />

history
-------
- 0.5 (Aug 08,2018)
    - When `-o FILENAME` is given, use "[tealeg/xlsx](https://github.com/tealeg/xlsx)" instead of "[go-ole/go-ole](https://github.com/go-ole/go-ole)"
    - Remove options -s and -q. Their features are enabled with -o.
- 0.4 (Aug 05,2018)
    - Add -f option to set field seperator
    - Do not treat as string when the value is a negative integer.
- 0.3 (Jul 11,2018)
    - Fix leak release COM
    - Print help if no arguments and stdin is not redirected
- 0.2 (Jul 1,2018)
    - Only matching `/^[1-9]\d*(\.\d*[1-9])?$/`, as a number
- 0.1 (Jul 1,2018)
    - Set cell as a string
    - Detect encoding utf8 or codepage automatically and remove -u option
    - Add -v option to show version
- 20180701
    - prototype
