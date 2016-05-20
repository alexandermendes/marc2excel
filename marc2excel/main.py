# -*- coding: utf-8 -*-
"""Main module for marc2excel."""

import re
import sys
import pymarc
from tqdm import tqdm
from marc2excel.sheets import Sheets


class Converter(object):
    """A class for converting MARC to Excel and vice-versa

    :param silent: Don't display CLI progress.
    """

    def __init__(self, silent=True):
        self.silent = silent
        self.sheets = Sheets(silent)
        self.errors = []

    def marc2excel(self, path, out_path, force_utf8=False):
        """Convert a MARC file to an Excel spreadsheet.

        :param path: Path to the MARC file.
        :param out_path: Path to save the Excel spreadsheet.
        :param force_utf8: Force decoding of records as UTF8.
        """
        if not self.silent:
            sys.stdout.write('Processing: {0}\n'.format(path))
        headings = set()
        headings.add('LDR')
        records = []
        with open(path, 'rb') as marc_file:
            reader = pymarc.MARCReader(marc_file, force_utf8=force_utf8)
            for record in reader:
                rec_dict = {'LDR': record.leader}
                rheadings = []
                for field in record.fields:
                    heading = field.tag
                    rheadings.append(heading)

                    if field.is_control_field():
                        rec_dict[heading] = field.value()
                        headings.add(heading)

                    if hasattr(field, 'subfields'):
                        i1 = field.indicator1.replace(' ', '\\')
                        i2 = field.indicator1.replace(' ', '\\')
                        heading = '{0} {1}{2}'.format(heading, i1, i2)

                        # Split subfields
                        c = [code for code in field.subfields[::2]]
                        v = [value for value in field.subfields[1::2]]
                        subfields = [(c[i], v[i]) for i in range(len(c))]
                        for s in subfields:
                            subheading = '{0} ${1}'.format(heading, s[0])
                            n = rheadings.count(field.tag)
                            if n > 1:
                                subheading = '{0} [{1}]'.format(subheading, n)
                            rec_dict[subheading] = s[1]
                            headings.add(subheading)

                records.append(rec_dict)
            headings = sorted(list(headings))
            data = [headings] + [[r.get(h, '') for h in headings]
                                 for r in records]
            self.sheets.write(out_path, data)
        if not self.silent:
            sys.stdout.write('Saved: {0}\n'.format(out_path))

    def excel2marc(self, path, out_path, sheet=0, force_utf8=False):
        """Convert an Excel spreadsheet to a MARC file.

        :param path: Path to the Excel spreadsheet.
        :param out_path: Path to save the MARC file.
        :param sheet: The index of the sheet from which to extract data.
        :param force_utf8: Force encoding or records as UTF8.
        """
        if not self.silent:
            sys.stdout.write('Processing: {0}\n'.format(path))
        data = self.sheets.extract_data(path, sheet)
        with open(out_path, 'wb') as out_file:
            for item in tqdm(data, desc="Writing MARC data", unit="records",
                             disable=self.silent):
                record = pymarc.Record(force_utf8=force_utf8)
                for k in sorted(item):
                    value = item[k]
                    key = re.sub(r'\s+', '', k.lower())

                    # Repeated field indicator
                    pos = 0
                    ind = re.search(r'\[[1-9]+\]', key)
                    if ind:
                        pos = int(ind.group(0)[1])
                        key = re.sub(r'\[[1-9]+\]', '', key)

                    tag = key[:3]

                    # Leader
                    if tag == 'ldr':
                        record.leader = value
                        continue

                    # Control field
                    if len(key) == 3:
                        field = pymarc.Field(tag, data=value)
                        record.add_field(field)
                        continue

                    indicators = list(key[3:5])
                    subfield_code = key[6]

                    # Get or create field
                    fields = record.get_fields(tag)
                    if fields and pos <= len(fields):
                        field = fields[pos - 1]
                    else:
                        field = pymarc.Field(tag=tag, indicators=indicators)
                        record.add_field(field)
                    field.add_subfield(subfield_code, value)

                out_file.write(record.as_marc())
        if not self.silent:
            sys.stdout.write('Saved: {0}\n'.format(out_path))
