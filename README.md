# xlspy
The way I want to use openpyxl

I sometimes use openpyxl which is pretty cool. However, I'd like some things
changed.

xlspy will use openpyxl to do this:

```python

>>> from xlspy import Book
>>> book = Book('filename.xlsx')
>>> book
<Book:   /home/jugurtha/filename.xlsx>
<Sheets: Cool Sheet 1 | Accounting | Home renovations>

>>> book.active
<Worksheet "Cool Sheet 1">

>>> book['Home renovations']
<Worksheet "Home renovations">
```