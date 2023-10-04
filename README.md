# mdtable.py

## Abstract
Program to convert .xlsx to grid_tables for markdown


## Usage

```sh
$ mdtable.py [-h] -i INPUT.xlsx [-s SHEET_NAME] [-r CELL_NAME CELL_NAME] [-c]
```

* `-h`, `--help`
	: show this help message and exit
* `-i INPUT.xlsx`
	: input .xlsx file
* `-s SHEET_NAME`
	: sheet name for converting table
* `-r CELL_NAME CELL_NAME`
	: Start and End position cells for target square area (Ex. `C3 E9`)
* `-c`
	: send clipboard



## Requirements
* Python 3
	* pyperclip
	* openpyxl


## License
The MIT License (MIT)

Copyright (c) 2023 Tatsuya Ohyama


## Authors
* Tatsuya Ohyama


## ChangeLog
### Ver. 2.3 (2023-10-04)
* Output cell values start with spaces at both side.
* Recognize all bold cells in row as header.

### Ver. 2.2 (2023-09-27)
* Fix bug that cannot convert cell with number format.

### Ver. 2.1 (2023-09-15)
* Support multiline in a cell.

### Ver. 2.0 (2023-09-15)
* Change drawing method of borders.
* Support number format.
* Fix bug that cannnot read workbook with specific sheet name.

### Ver. 1.0 (2023-09-14)
* Release.

