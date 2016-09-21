#!/usr/bin/env python

import os
import sys
import click
from marc2excel import excel2marc


@click.command()
@click.argument('excel_path', type=str)
@click.option('-d', flag_value=True, default="False",
              help='Specify a directory for EXCEL_PATH.')
@click.option('--silent', flag_value=True, default="False",
              help="Don't display progress.")
@click.option('--utf8', flag_value=True, default="False",
              help="Force utf-8 encoding.")
def main(excel_path, d=False, silent=False, utf8=False):
    """Convert EXCEL_PATH to an MARC file.

    The path to the saved MARC file will be the same as EXCEL_PATH but with a
    .mrc extension.

    If a directory is specified for EXCEL_PATH all files found in that
    directory with .xlsx extensions will be converted and the .mrc versions
    saved in the same directory.

    The option to force utf-8 encoding is included to help deal with any poorly
    formed data found in the spreadsheet. For example, where rows contain data
    that is utf-8 encoded but without the leader set appropriately, or rows
    that contain mostly utf-8 data but with a few non-utf-8 characters (any
    non-utf-8 characters will be replaced with the official U+FFFD REPLACEMENT
    CHARACTER).
    """
    if not silent:
        sys.stdout.write('Running Excel to MARC conversion\n')

    if not d:
        save_path = _get_save_path(excel_path)
        excel2marc(excel_path, save_path, silent=silent, force_utf8=utf8)
    else:
        paths = (os.path.join(excel_path, f) for f in os.listdir(excel_path)
                 if not f.startswith('.') and f.lower().endswith('.xlsx'))
        for p in paths:
            save_path = _get_save_path(p)
            excel2marc(p, save_path, silent=silent, force_utf8=utf8)

    if not silent:
        sys.stdout.write('Finished\n')


def _get_save_path(path):
    """Return """
    fn = os.path.splitext(os.path.basename(path))[0]
    d = os.path.dirname(path)
    return os.path.join(d, fn + '.mrc')


if __name__ == '__main__':
    main()
