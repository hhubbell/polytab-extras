from setuptools import setup
import os
import re


def get_version():
    with open(os.path.join('polytab_extras', '__init__.py'), 'r') as f:
        return re.search('^__version__ = \'(.*)\'$', f.read(), re.M).group(1)

setup(
    name='polytab-extras',
    version=get_version(),
    description='Additional polytab utilities requiring external dependencies.',
    author='Harry Hubbell',
    url='https://github.com/hhubbell/polytab-extras',
    packages=['polytab_extras', 'polytab_extras.parsers'],
    install_requires=['polytab', 'openpyxl'],
    entry_points={
        'polytab.parsers': 'xlsx=polytab_extras.parsers.xlsx:XLSXParser'
    })
