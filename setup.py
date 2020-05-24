from setuptools import setup

from os import path
this_directory = path.abspath(path.dirname(__file__))
with open(path.join(this_directory, 'README.md'), encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='excel_templates',
    version='0.1',
    packages=['excel_templates'],
    url='https://github.com/abielr/excel-templates',
    license='MIT',
    author='Abiel Reinhart',
    author_email='abielr@gmail.com',
    description='A package for filling in Excel templates',
    install_requires=[
        'openpyxl>=2.6'
    ],
    long_description=long_description,
    long_description_content_type='text/markdown',
    classifiers = [
        "Programming Language :: Python :: 3"
    ]
)
