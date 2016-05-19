#!/usr/bin/env python

import os
import sys
import click
from marc2excel import marc2excel


@click.command()
@click.argument('source_path', type=str)
@click.argument('save_path', type=str)
@click.option('-d', flag_value=True, default="False",
              help='Specify directories for SOURCE_PATH and SAVE_PATH.')
@click.option('--utf8', flag_value=True, default="True",
              help='Force records to be decoded as UTF-8.')
@click.option('--silent', flag_value=True, default="False",
              help="Don't display progress.")
def main(source_path, save_path, d=False, utf8=False, silent=False):
    """Convert MARC (.mrc, .marc) to Excel (.xlsx)"""
    if not silent:
        sys.stdout.write('Running MARC to Excel conversion\n')

    if not d:
        marc2excel(source_path, save_path, force_utf8=utf8, silent=silent)
    else:
        paths = (os.path.join(source_path, f) for f in os.listdir(source_path)
                 if not f.startswith('.') and
                 f.lower().endswith(('.mrc', '.marc')))
        for p in paths:
            fn = '{0}.xlsx'.format(os.path.splitext(os.path.basename(p))[0])
            out_path = os.path.join(save_path, p)
            marc2excel(source_path, save_path, force_utf8=utf8, silent=silent)

    if not silent:
        sys.stdout.write('Finished\n')


if __name__ == '__main__':
    main()
