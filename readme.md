Pipe To Excel
=============

`pipe2excel.exe` sends the contents of STDIN 
or files of arguments to Excel as CSV to Microsoft Excel.

- The each value of the csv is inserted as a string.
- The encoding of the CSV is detected automatically whether it is written in UTF8 or the encoding of the current codepage.

```
C:\> pipe2excel foo.csv bar.csv
```

```
C:\> type foo.csv | pipe2excel
```

