#!/usr/bin/env python

import os
import sys
import click
from marc2excel import marc2excel


@click.command()
@click.argument('marc_path', type=str)
@click.option('-d', flag_value=True, default="False",
              help='Specify a directory for MARC_PATH.')
@click.option('--silent', flag_value=True, default="False",
              help="Don't display progress.")
def main(marc_path, d=False, silent=False):
    """Convert MARC_PATH to an Excel spreadsheet.

    The path to the saved spreadsheet will be the same as MARC_PATH but with a
    .xlsx extension.

    If a directory is specified for MARC_PATH all files found in that directory
    with .mrc, .marc or .lex extensions will be converted and the .xlsx
    versions saved in the same directory.
    """
    if not silent:
        sys.stdout.write('Running MARC to Excel conversion\n')

    if not d:
        save_path = _get_save_path(marc_path)
        marc2excel(marc_path, save_path, silent=silent)
    else:
        paths = (os.path.join(marc_path, f) for f in os.listdir(marc_path)
                 if not f.startswith('.') and
                 f.lower().endswith(('.mrc', '.marc', '.lex')))
        for p in paths:
            save_path = _get_save_path(p)
            marc2excel(p, save_path, silent=silent)

    if not silent:
        sys.stdout.write('Finished\n')


def _get_save_path(path):
    """Return """
    fn = os.path.splitext(os.path.basename(path))[0]
    d = os.path.dirname(path)
    return os.path.join(d, fn + '.xlsx')


if __name__ == '__main__':
    main()
