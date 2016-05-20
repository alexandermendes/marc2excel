import re
from setuptools import setup


version = re.search('^__version__\s*=\s*"(.*)"',
                    open('marc2excel/__init__.py').read(),
                    re.M).group(1)

try:
    long_description = read("README.rst")
except:
    long_description = ""


requirements = ["openpyxl>=2.3.0, <3.0",
                "pymarc>=3.0, <4.0",
                "click>=6.6, <7.0",
                "tqdm>=4.7.0, <5.0"
                ]

setup_requirements = ["pytest-runner>=2.6.0, <3.0"]

test_requirements = ["pytest>=2.8.0, <3.0",
                     "pytest-cov>=2.2.0, <3.0",
                     "pytest-mock>=0.11.0, <1.0"]

setup(
    name='marc2excel',
    version=version,
    author='Alexander Mendes',
    author_email='alexanderhmendes@gmail.com',
    url='https://github.com/alexandermendes/marc2excel',
    license="BSD",
    description='Convert MARC files to Excel spreadsheets and vice-versa.',
    long_description=long_description,
    install_requires=requirements,
    setup_requires=setup_requirements,
    tests_require=test_requirements,
    packages=['marc2excel'],
    scripts=["bin/marc2excel_cli.py", "bin/excel2marc_cli.py"],
    test_suite="test",
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: BSD License",
        "Programming Language :: Python"
    ],
)
