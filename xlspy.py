"""
XLSPy.

Simple wrapper for openpyxl::


>>> from xlspy import Book
>>> book = Book('filename.xlsx')
>>> book
<Book:   /home/jugurtha/filename.xlsx>
<Sheets: COOL Sheet 1 | Accounting | Home renovations>

>>> book['Home renovations']
<Worksheet "Home renovations">
"""

import openpyxl


class Book(object):
    """Wrapper for openpyxl."""

    def __init__(self, filename, **kwargs):
        """Initialize the book."""
        self._filename = filename
        self._book = openpyxl.load_workbook(filename, **kwargs)
        self._sheet_names = self._book.get_sheet_names()
        self._sheets = [self._book.get_sheet_by_name(key) for key in self.sheet_names]

    def __getitem__(self, key):
        """Access worksheets in book as you would items in a dictionary.

        >>> sheet = book['sheet_name']
        """
        return self._book.get_sheet_by_name(key)

    @property
    def sheet_names(self):
        """Return sheet names."""
        return self._sheet_names

    @property
    def sheets(self):
        """Return actual sheets."""
        return self._sheets

    def __repr__(self):
        """Nice repr."""
        msg = "<Book:\t{}>\n" \
              "<Sheets:\t{}>".format(
                  self._filename,
                  " | ".join(["{}".format(name) for name in self.sheet_names]))
        return msg


class Sheet(object):
    """Wrapper for openpyxl sheets."""

    def __init__(self, sheet):
        """Initialize the sheet."""
