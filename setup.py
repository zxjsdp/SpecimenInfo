from setuptools import setup, find_packages
from codecs import open
from os import path

here = path.abspath(path.dirname(__file__))

with open(path.join(here, 'README.md'), encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='SpecimenInfo',
    version='1.2.0',
    description=('Fetch and format plant specimen informations from data file'
                 ' and web, save outcome to xlsx file and SQLite3 db file.'),
    author='Haofei Jin',
    author_email='zxjsdp@gmail.com',
    url='https://github.com/zxjsdp/SpecimenInfo',
    license='Apache',
    keywords='specimen automated plant format xlsx',
    packages=['specimen_info'],
    install_requires=['requests', 'BeautifulSoup4', 'openpyxl'],
    # $ pip install -e .[dev,test]
    extras_require={
        'dev': ['pytest', 'tox', 'sphinx'],
        'test': ['pytest'],
    },
    long_description=long_description,
    classifiers=[
        'Development Status :: 3 - Alpha',
        'License :: OSI Approved :: Apache Software License',
        'Programming Language :: Python',
        'Programming Language :: Python :: 2.6',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
    ],
)
