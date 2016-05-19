#!/usr/bin/env python

import re
from setuptools import setup


version = re.search('^__version__\s*=\s*"(.*)"',
                    open('marc2excel/__init__.py').read(),
                    re.M).group(1)


with open("README.md", "rb") as f:
    long_description = f.read().decode("utf-8")


requirements = ["openpyxl>=2.3.3, <3.0",
                "pymarc>=3.1.2, <4.0",
                "click>=6.6, <7.0",
                "tqdm>=4.7.0, <5.0"
                ]

setup(
    name='marc2excel',
    version=version,
    author='Alexander Mendes',
    author_email='alexanderhmendes@gmail.com',
    url='https://github.com/alexandermendes/marc2excel',
    description='Convert MARC files to Excel spreadsheets and vice-versa.',
    long_description=long_description,
    install_requires=requirements,
    packages=['marc2excel'],
    scripts=["bin/marc2excel_cli.py", "bin/excel2marc_cli.py"],
)
