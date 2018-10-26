# excel-to-mptt-json

Tool to make MPTT-ready data list from Excel files

## Requirements:
* Python 3.6.x
* openpyxl==2.4.8

## Args:
```
python excel2json.py --help

usage: excel2json.py [-h] [-f FILE] [-l LVL] [-t TITLE] [-n NESTING] [-v]
                     [-O OUTPUT] [-C]

optional arguments:
  -h, --help            show this help message and exit
  -f FILE, --file FILE  Excel file path
  -l LVL, --lvl LVL     start col
  -t TITLE, --title TITLE
                        first useful line
  -n NESTING, --nesting NESTING
                        number nested levels
  -v                    show INFO log
  -O OUTPUT, --output OUTPUT
                        save to file
  -C, --capitalize      Capitalize titles
```

## Usage
```
    python excel2json.py -f Классификатор\ -\ Дрогери.xlsx -v -C -n 3 -O categories.py
```
Parse file "Классификатор Дрогери.xlsx" for 3 levels to categories.py with DEBUG output log
