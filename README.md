# xlspy
The way I want to use openpyxl

**Update**:

> All the features presented here (sheet names as a property instead of `getSheetNames`, accessing a range of cells with brackets, etc) have been implemented in the newer versions of openpyxl, which is great and renders this repo caducous. One reason I didn't submit code as a pull request is because openpyxl is in BitBucket.
>
> One other change I'd like to see in openpyxl is in `openpyxl.reader.excel`: in function `load_workbook`, there's a call to function `apply_stylesheet`. That call can take up to 3 minutess for workbooks with a lot of styling (say, from a client). If all you care about is data and not the styling, which is almost certainly the case if you are using a python library to work with a spread-sheet, then it makes sense to make the styling optional. Thus, it would be good to add a default argument to `load_workbook` to make the call to `apply_stylesheet` optional in case we just care about the cell values.
>
> I had workbooks that took 3 minutes to load and I changed that to a few milliseconds when I removed the call to `apply_stylesheet` because all I cared about was the data.


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
