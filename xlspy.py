"""
XLSPy.

Simple wrapper for openpyxl::

>>> from xlspy import Book
>>> book = Book('filename.xlsx')
>>> book
<Book:   /home/jugurtha/filename.xlsx>
<Sheets: Cool Sheet 1 | Accounting | Home renovations>

>>> book.active
<Sheet "Cool Sheet 1">

# Get a sheet directly:

>>> home = book['Home renovations']
>>> home
<Sheet "Home renovations">

>>> column_a = home['A']
"""

from itertools import chain

import openpyxl


class Sheet(object):
    """Wrapper for openpyxl sheets.

    Basic usage of a Sheet instance:

        >>> s['F']
        # Result: list of cells of column 'F'

        >>> s['F1:F4']
        # Result: tuple of one element tuples (openpyxl)
    """

    # I wonder if I'm going to change the standard way openpyxl returns
    # a cell range (tuple of tuples) as I did with the whole column.
    def __init__(self, sheet):
        """Initialize the sheet."""
        self._sheet = sheet

    def __getitem__(self, key):
        """Make access to "columns" easier. Cell access remains the same."""
        if ':' not in key:
            # Use wants to access a column, example: Sheet['F'].
            # We return Fmin_row:Fmax_row cells under the hood.
            return list(chain.from_iterable(
                self._sheet['{key}{min_row}:{key}{max_row}'.format(
                    key=key,
                    min_row=self._sheet.min_row,
                    max_row=self._sheet.max_row)
                ])
            )

        else:
            # User wants to access specific cells Sheet['F1:F3']
            return self._sheet[key]

    def __getattr__(self, name):
        """Return the original openpyxl arguments for whatever is not ours."""
        return getattr(self._sheet, name)

    def __repr__(self):
        """Nice representation."""
        return "<{}: '{}'>".format(self.__class__.__name__, self._sheet.title)


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
    def active(self):
        """Return the active sheet."""
        return Sheet(self._book.active)

    @property
    def sheet_names(self):
        """Return sheet names."""
        return self._sheet_names

    @property
    def sheets(self):
        """Return actual sheet objects."""
        return [Sheet(sheet) for sheet in self._sheets]

    def __getattr__(self, name):
        """Return original openpyxl Workbook attributes."""
        return getattr(self._book, name)

    def __repr__(self):
        """Nice repr."""
        msg = "<{}:\t{}>\n" \
              "<Sheets:\t{}>".format(
                  self.__class__.__name__,
                  self._filename,
                  " | ".join(["{}".format(name) for name in self.sheet_names]))
        return msg
