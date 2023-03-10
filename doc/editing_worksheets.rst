Inserting and deleting rows and columns, moving ranges of cells
===============================================================


Inserting rows and columns
--------------------------

You can insert rows or columns using the relevant worksheet methods:

    * :func:`openpyxl.worksheet.worksheet.Worksheet.insert_rows`
    * :func:`openpyxl.worksheet.worksheet.Worksheet.insert_cols`
    * :func:`openpyxl.worksheet.worksheet.Worksheet.delete_rows`
    * :func:`openpyxl.worksheet.worksheet.Worksheet.delete_cols`

The default is one row or column. For example to insert a row at 7 (before
the existing row 7)::

    >>> ws.insert_rows(7)


Deleting rows and columns
--------------------------

To delete the columns ``F:H``::

    >>> ws.delete_cols(6, 3)

.. note::

    Openpyxl does not manage dependencies, such as formulae, tables, charts,
    etc., when rows or columns are inserted or deleted. This is considered to
    be out of scope for a library that focuses on managing the file format.
    As a result, client code **must** implement the functionality required in
    any particular use case.


Moving ranges of cells
----------------------

You can also move ranges of cells within a worksheet::

    >>> ws.move_range("D4:F10", rows=-1, cols=2)

This will move the cells in the range ``D4:F10`` up one row, and right two
columns. The cells will overwrite any existing cells.

If cells contain formulae you can let openpyxl translate these for you, but
as this is not always what you want it is disabled by default. Also only the
formulae in the cells themselves will be translated. References to the cells
from other cells or defined names will not be updated; you can use the
:doc:`formula` translator to do this::

    >>> ws.move_range("G4:H10", rows=1, cols=1, translate=True)

This will move the relative references in formulae in the range by one row and one column.


Merge / Unmerge cells
---------------------

When you merge cells all cells but the top-left one are **removed** from the
worksheet. To carry the border-information of the merged cell, the boundary cells of the
merged cell are created as MergeCells which always have the value None.
See :ref:`styling-merged-cells` for information on formatting merged cells.

.. :: doctest

>>> from openpyxl.workbook import Workbook
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> ws.merge_cells('A2:D2')
>>> ws.unmerge_cells('A2:D2')
>>>
>>> # or equivalently
>>> ws.merge_cells(start_row=2, start_column=1, end_row=4, end_column=4)
>>> ws.unmerge_cells(start_row=2, start_column=1, end_row=4, end_column=4)
