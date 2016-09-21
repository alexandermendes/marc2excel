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
@click.option('--utf8', flag_value=True, default="False",
              help="Force utf-8 encoding.")
def main(marc_path, d=False, silent=False, utf8=False):
    """Convert MARC_PATH to an Excel spreadsheet.

    The path to the saved spreadsheet will be the same as MARC_PATH but with a
    .xlsx extension.

    If a directory is specified for MARC_PATH all files found in that directory
    with .mrc, .marc or .lex extensions will be converted and the .xlsx
    versions saved in the same directory.

    The option to force utf-8 encoding is included to help deal with any poorly
    formed MARC records that you might encounter. For example, those that are
    utf-8 encoded but without the leader set appropriately, or those that
    contain mostly utf-8 data but with a few non-utf-8 characters (any
    non-utf-8 characters will be replaced with the official U+FFFD REPLACEMENT
    CHARACTER).
    """
    if not silent:
        sys.stdout.write('Running MARC to Excel conversion\n')

    if not d:
        save_path = _get_save_path(marc_path)
        marc2excel(marc_path, save_path, silent=silent, force_utf8=utf8)
    else:
        paths = (os.path.join(marc_path, f) for f in os.listdir(marc_path)
                 if not f.startswith('.') and
                 f.lower().endswith(('.mrc', '.marc', '.lex')))
        for p in paths:
            save_path = _get_save_path(p)
            marc2excel(p, save_path, silent=silent, force_utf8=utf8)

    if not silent:
        sys.stdout.write('Finished\n')


def _get_save_path(path):
    """Return """
    fn = os.path.splitext(os.path.basename(path))[0]
    d = os.path.dirname(path)
    return os.path.join(d, fn + '.xlsx')


if __name__ == '__main__':
    main()
