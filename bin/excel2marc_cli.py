#!/usr/bin/env python

import os
import sys
import click
from marc2excel import excel2marc


@click.command()
@click.argument('source_path', type=str)
@click.argument('save_path', type=str)
@click.option('-s', '--sheet', default=0,
              help="Index of the sheet from which to extract data.")
@click.option('-d', flag_value=True, default="False",
              help='Specify directories for SOURCE_PATH and SAVE_PATH.')
@click.option('--utf8', flag_value=True, default="False",
              help='Force records to be encoded as UTF-8.')
@click.option('--silent', flag_value=True, default="False",
              help="Don't display progress.")
def main(source_path, save_path, sheet=0, d=False, utf8=False, silent=False):
    """Convert Excel (.xlsx) to MARC (.mrc)"""
    if not silent:
        sys.stdout.write('Running Excel to MARC conversion\n')

    if not d:
        excel2marc(source_path, save_path, sheet=sheet, silent=silent,
                   force_utf8=utf8)
    else:
        paths = (os.path.join(source_path, f) for f in os.listdir(source_path)
                 if not f.startswith('.') and f.lower().endswith('.xlsx'))
        for p in paths:
            fn = '{0}.mrc'.format(os.path.splitext(os.path.basename(p))[0])
            out_path = os.path.join(save_path, p)
            excel2marc(source_path, save_path, sheet=sheet, silent=silent,
                       force_utf8=utf8)

    if not silent:
        sys.stdout.write('Finished\n')


if __name__ == '__main__':
    main()
