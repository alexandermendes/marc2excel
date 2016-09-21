marc2excel
==========

|travis| |coveralls| |pypi|

Convert MARC files to Excel spreadsheets and vice-versa.


Installation
------------

From PyPI:

::

    pip install marc2excel

From GitHub:

::

    git clone https://github.com/alexandermendes/marc2excel
    cd marc2excel
    python setup.py install


Usage
-----

Running scripts:
~~~~~~~~~~~~~~~~

The following scripts can be run from anywhere, once the package is
installed.

Converting MARC to Excel:

::

    Usage: marc2excel_cli.py [OPTIONS] MARC_PATH

      Convert MARC_PATH to an Excel spreadsheet.

      The path to the saved spreadsheet will be the same as MARC_PATH but with a
      .xlsx extension.

      If a directory is specified for MARC_PATH all files found in that
      directory with .mrc, .marc or .lex extensions will be converted and the
      .xlsx versions saved in the same directory.

    Options:
      -d        Specify a directory for MARC_PATH.
      --silent  Don't display progress.
      --help    Show this message and exit.


Converting Excel to MARC:

::

    Usage: excel2marc_cli.py [OPTIONS] EXCEL_PATH

      Convert EXCEL_PATH to an MARC file.

      The path to the saved MARC file will be the same as EXCEL_PATH but with a
      .mrc extension.

      If a directory is specified for EXCEL_PATH all files found in that
      directory with .xlsx extensions will be converted and the .mrc versions
      saved in the same directory.

      Options:
        -d        Specify a directory for EXCEL_PATH.
        --silent  Don't display progress.
        --help    Show this message and exit.


Running from Python:
~~~~~~~~~~~~~~~~~~~~

Converting MARC to Excel:

.. code-block:: python

    import marc2excel
    marc2excel.marc2excel('path/to/file.mrc', 'path/to/save/file.xls')

Converting Excel to MARC:

.. code-block:: python

    import marc2excel
    marc2excel.excel2marc('path/to/file.xls', 'path/to/save/file.mrc')


Spreadsheet guidelines
----------------------

Spreadsheets require a header row that must adhere to the following
guidelines:

-  The field tag is required for all fields.
-  For non-control fields, the indicator and subfield tags are also
   required.
-  Backslashes should be used to indicate blank spaces in indicators.
-  Subfields should be prepended with a dollar symbol.
-  Leaders can be added from a column with the heading LDR (optional).
-  Repeated fields can be created by appending headers with [*number*].

**Example:**

+-------+---------------+------------------+------------------+
| 001   |  245 \\\\ $a  |     852 \\1 $j   |  852 \\1 $j [2]  |
+=======+===============+==================+==================+
| 123   |  some\_value  |  another\_value  |  another\_value  |
+-------+---------------+------------------+------------------+


Testing
-------

Tests can be run using:

::

    python setup.py test

.. |travis| image:: https://travis-ci.org/alexandermendes/marc2excel.svg?branch=master
    :target: https://travis-ci.org/alexandermendes/marc2excel
    :alt: Test success
.. |coveralls| image:: https://coveralls.io/repos/github/alexandermendes/marc2excel/badge.svg?branch=master
    :target: https://coveralls.io/github/alexandermendes/marc2excel?branch=master
    :alt: Test coverage
.. |pypi| image:: https://img.shields.io/pypi/v/marc2excel.svg?label=latest%20version
    :target: https://pypi.python.org/pypi/marc2excel
    :alt: Latest version released on PyPi