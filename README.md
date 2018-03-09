# xlspy
The way I want to use openpyxl

[UPDATE]: It seems that all the features implemented here (sheet names being a property, accessing a range of cells with brackets, etc) were implemented in the newer versions of openpyxl, which is great and renders this repo caducous. One reason I didn't submit the code as a pull request is because openpyxl is in bitbucket and I just don't like it.

One other change I'd like to see in openpyxl is in `openpyxl.reader.excel`, there's a function `load_workbook`, and in that function there's a call to another function, `apply_stylesheet`. That call can take up to 2 minutess for weird workbooks (you can't force that on your clients) and it would be good to add a default argument to `load_workbook` to make the call to `apply_stylesheet` optional in case we just care about the cell values.

I had workbooks that took 3 minutes to load and I changed that to a few milliseconds when I removed the call to `apply_stylesheet` because all I cared about was the data.

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
