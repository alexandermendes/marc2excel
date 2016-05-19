# marc2excel

Convert MARC files to Excel spreadsheets and vice-versa.


## Installation

From GitHub:

```Shell
git clone https://github.com/alexandermendes/marc2excel
cd marc2excel
python setup.py install
```


## Usage

### Running scripts:

Converting MARC to Excel:

```Shell
Usage: marc2excel_cli.py [OPTIONS] SOURCE_PATH SAVE_PATH

  Convert MARC (.mrc, .marc) to Excel (.xlsx)

Options:
  -d        Specify directories for SOURCE_PATH and SAVE_PATH.
  --utf8    Force records to be decoded as UTF-8.
  --silent  Don't display progress.
  --help    Show this message and exit.
```

Converting Excel to MARC:

```Shell
Usage: excel2marc_cli.py [OPTIONS] SOURCE_PATH SAVE_PATH

  Convert Excel (.xlsx) to MARC (.mrc)

Options:
  -s, --sheet INTEGER  Index of the sheet from which to extract data.
  -d                   Specify directories for SOURCE_PATH and SAVE_PATH.
  --utf8               Force records to be encoded as UTF-8.
  --silent             Don't display progress.
  --help               Show this message and exit.
```

The above scripts can be run from anywhere, once the package is installed.


### Running from Python:

Converting MARC to Excel:

```Python
import marc2excel
marc2excel.marc2excel('path/to/file.mrc', 'path/to/save/file.xls')

Converting Excel to MARC:

```Python
import marc2excel
marc2excel.excel2marc('path/to/file.xls', 'path/to/save/file.mrc')
```


## Spreadsheet guidelines

If a spreadsheet is created or edited manually the following guidelines for the
header row must be adhered to:

- The field tag is required for all fields.
- For any non-control fields the indicator and subfield codes are also required.
- Backslashes should be used to indicate blank spaces in indicators.
- Subfields should be prepended with a dollar symbol.
- Optionally, leaders can be added from a column with the heading LDR.
- Repeated fields can be created by sequential prepending headers with [number].

Example:

|    001    |    245 \\ $a   |   852 \1 $j   |   [2] 852 \1 $j   |
|:---------:|:--------------:|:-------------:|:-----------------:|
|    123    |    some_value  | another_value |   another_value   |
