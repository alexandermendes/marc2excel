# -*- coding: utf8 -*-

import os
import pytest
import tempfile
from pymarc import MARCReader


@pytest.fixture()
def marc_dir(tmpdir):
    samples = os.path.join(os.path.dirname(__file__), 'samples')
    marc_dir = tmpdir.mkdir("marc")
    for fn in os.listdir(samples):
        path = os.path.join(samples, fn)
        with open(path, 'rb') as f:
            try:
                marc_dir.join(fn).write(f.read())
            except UnicodeDecodeError:
                continue


@pytest.fixture()
def marc_path():
    here = os.path.dirname(__file__)
    return os.path.join(here, 'samples', 'marc_file.mrc')


@pytest.fixture()
def bad_marc_path():
    here = os.path.dirname(__file__)
    return os.path.join(here, 'samples', 'bad_marc_file.mrc')


@pytest.fixture()
def marc_records():
    here = os.path.dirname(__file__)
    path = os.path.join(here, 'samples', 'marc_file.mrc')
    records = []
    with open(path, 'rb') as marc_file:
        reader = MARCReader(marc_file, to_unicode=True)
        for record in reader:
            records.append(record)
    return records


@pytest.fixture()
def tmp_marc_file():
    return tempfile.NamedTemporaryFile(suffix='.mrc')


@pytest.fixture()
def tmp_sheet():
    return tempfile.NamedTemporaryFile(suffix='.xlsx')
