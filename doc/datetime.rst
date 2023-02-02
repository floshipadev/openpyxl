Dates and Times
===============

Dates and times can be stored in two distinct ways in XLSX files: as an
ISO 8601 formatted string or as a single number. `openpyxl` supports
both representations and translates between them and Python's datetime
module representations when reading files. In either representation,
the maximum date and time precision in XLSX files is millisecond precision.

XLSX files are not suitable for storing historic dates (before 1900),
due to bugs in Excel that cannot be fixed without causing backward
compatibility problems. To discourage users from trying anyway, Excel
deliberately refuses to display such dates. Consequently,
it is not advised to use `openpyxl` for such purposes either, especially
when exchanging files with others.


Timezones
---------

The date and time representations in Excel do not support timezones,
therefore `openpyxl` can only deal with naive datetime/time objects.
Any timezone information attached to Python datetimes must be stripped
off by the user before datetimes can be stored in XLSX files.


Using the ISO 8601 format
-------------------------

To make `openpyxl` store dates and times in the ISO 8601 format on
writing your file, set the workbook's ``iso_dates`` flag to ``True``
This is the default for openpyxl:

    >>> import openpyxl
    >>> wb = openpyxl.Workbook()
    >>> wb.iso_dates = True

The benefit of using this format is that the meaning of the stored
information is not subject to interpretation, as it is with the single
number format [#f1]_.

The Office Open XML standard does not specify a supported subset of the
ISO 8601 duration format for representing time interval durations.
`openpyxl` therefore always uses the single number format for timedelta
values when writing them to file.


The 1900 and 1904 date systems
------------------------------

The 'date system' of an XLSX file determines how dates and times in the
single number representation are interpreted. XLSX files always use one
of two possible date systems:

 * Legacy 1900 date system (Excel's default), the epoch is 1899-12-31,
   the first displayable date is 1900-01-01.
 * Legacy 1904 date system (MacOS legacy) the epoch is 1904-01-01.
 * Standard 1900 date system (Strict OOXML) the epoch is 1899-12-30.

Complications arise not only from the different epochs, but also from
the fact that the legacy date system incorrectly assumes that 1900 was a leap year.

More information on this issue is available from Microsoft:
 * https://docs.microsoft.com/en-us/office/troubleshoot/excel/1900-and-1904-date-system
 * https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year

In workbooks using the legacy system, `openpyxl` has the same dates as
Excel for January and February 1900, except for 1900-02-29, which is invalid.
Cells with the value 60 formatted as a date will be marked as errors to allow
client code to make any necessary adjustments.

Comparison of problematic serial dates
--------------------------------------

=======  ==========  ========
Ordinal  Excel       Standard
-------  ----------  --------
  …          …            …
58       1900-02-27  1900-02-27
59       1900-02-28  1900-02-28
60       1900-02-29  1900-03-01
61       1900-03-01  1900-03-02
62       1900-03-02  1900-03-03
63       1900-03-03  1900-03-04
  …          …            …
=======  ==========  ========

You can get the date system of a workbook like this:

    >>> import openpyxl
    >>> wb = openpyxl.Workbook()
    >>> if wb.epoch == openpyxl.utils.datetime.CALENDAR_WINDOWS_1900:
    ...     print("This workbook is using the 1900 date system.")
    ...
    This workbook is using the 1900 date system.


and set it like this:

    >>> wb.epoch = openpyxl.utils.datetime.CALENDAR_MAC_1904



Handling timedelta values
-------------------------

Excel users can use number formats resembling ``[h]:mm:ss`` or
``[mm]:ss`` to display time interval durations, which `openpyxl`
considers to be equivalent to timedeltas in Python.
`openpyxl` recognizes these number formats when reading XLSX files and
returns datetime.timedelta values for the corresponding cells.

When writing timedelta values from worksheet cells to file, `openpyxl`
uses the ``[h]:mm:ss`` number format for these cells.

.. rubric:: Footnotes

.. [#f1] For example, the serial 1 in an Excel worksheet can be
         interpreted as 00:00, as 24:00, as 1900-01-01, as 1440
         (minutes), etc., depending solely on the formatting applied.
