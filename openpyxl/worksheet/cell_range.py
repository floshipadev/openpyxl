from __future__ import absolute_import

from openpyxl.compat.strings import VER

from openpyxl.utils import (
    range_boundaries,
    range_to_tuple,
    get_column_letter,
    quote_sheetname,
)


### Notes
# focus on range functions


class CellRange(object):
    """
    Represents a range in a sheet: title and coordinates.

    This object is used to perform calculations on ranges, like:

    - shifting to a direction,
    - union/intersection with another sheet range,
    - collapsing to a 1 x 1 range,
    - expanding to a given size.

    We can check whether a range is:

    - equal or not equal to another,
    - disjoint of another,
    - contained in another.

    We can get:

    - the string representation,
    - the coordinates,
    - the size of a range.
    """


    def __init__(self, range_string=None, min_col=None, min_row=None,
                 max_col=None, max_row=None, title=None):
        if range_string is not None:
            try:
                title, (min_col, min_row, max_col, max_row) = range_to_tuple(range_string)
            except ValueError:
                min_col, min_row, max_col, max_row = range_boundaries(range_string)
        # None > 0 is False
        if not all(idx > 0 for idx in (min_col, min_row, max_col, max_row)):
            msg = "Values for 'min_col', 'min_row', 'max_col' *and* 'max_row_' " \
                  "must be strictly positive"
            raise ValueError(msg)
        # Intervals are inclusive
        if not min_col <= max_col:
            fmt = "{max_col} must be greater than {min_col}"
            raise ValueError(fmt.format(min_col=min_col, max_col=max_col))
        if not min_row <= max_row:
            fmt = "{max_row} must be greater than {min_row}"
            raise ValueError(fmt.format(min_row=min_row, max_row=max_row))
        self.min_col = min_col
        self.min_row = min_row
        self.max_col = max_col
        self.max_row = max_row
        self.title = title


    @property
    def bounds(self):
        """
        Vertices of the range as a tuple
        """
        return self.min_col, self.min_row, self.max_col, self.max_row


    @property
    def coord(self):
        """
        Excel style representation of the range
        """
        fmt = "{min_col}{min_row}:{max_col}{max_row}"
        if (self.min_col == self.max_col
            and self.min_row == self.max_row):
            fmt = "{min_col}{min_row}"

        return fmt.format(
            min_col=get_column_letter(self.min_col),
            min_row=self.min_row,
            max_col=get_column_letter(self.max_col),
            max_row=self.max_row
        )


    def __repr__(self):
        fmt = "{coord}"
        if self.title:
            fmt = "{title!r}!{coord}"
        return fmt.format(title=self.title, coord=self.coord)


    def _get_range_string(self):
        fmt = "{coord}"
        if self.title:
            fmt = u"{title}!{coord}"

        return fmt.format(title=quote_sheetname(self.title), coord=self.coord)

    if VER[0] == 3:
        __str__ = _get_range_string

    else:
        __unicode__ = _get_range_string

        def __str__(self):
            title = self.title or ''
            title = title.encode('ascii', errors='backslashreplace')
            if b"'" in title:
                title = b"'" + title.replace(b"'", b"''") + b"'"
            coord = self.coord
            if title:
                return title + b"!" + coord
            else:
                return coord


    def __copy__(self):
        return self.__class__(min_col=self.min_col, min_row=self.min_row,
                              max_col=self.max_col, max_row=self.max_row,
                              title=self.title)


    def shift(self, col_shift=0, row_shift=0):
        """
        Shift the range according to the shift values (*col_shift*, *row_shift*).

        :type col_shift: int
        :param col_shift: number of columns to be moved by, can be negative
        :type row_shift: int
        :param row_shift: number of rows to be moved by, can be negative
        :raise: :class:`ValueError` if any row or column index < 1
        """

        if (self.min_col + col_shift <= 0
            or self.min_row + row_shift <= 0):
            raise ValueError("Invalid shift value: col_shift={0}, row_shift={1}".format(col_shift, row_shift))
        self.min_col += col_shift
        self.min_row += row_shift
        self.max_col += col_shift
        self.max_row += row_shift


    def __iadd__(self, value):
        col_shift, row_shift = value
        self.shift(col_shift, row_shift)
        return self


    def __add__(self, other):
        return self.__copy__().__iadd__(other)


    def __ne__(self, other):
        """
        Test whether the ranges are not equal.

        :type other: CellRange
        :param other: Other sheet range
        :return: ``True`` if *range* != *other*.
        """
        if isinstance(other, CellRange):
            # Test whether sheet titles are different and not empty.
            this_title = self.title
            that_title = other.title
            ne_sheet_title = this_title and that_title and this_title.upper() != that_title.upper()
            return (ne_sheet_title or
                    other.min_row != self.min_row or self.max_row != other.max_row or
                    other.min_col != self.min_col or self.max_col != other.max_col)
        raise TypeError(repr(type(other)))


    def __eq__(self, other):
        """
        Test whether the ranges are equal.

        :type other: CellRange
        :param other: Other sheet range
        :return: ``True`` if *range* == *other*.
        """
        return not self.__ne__(other)


    def issubset(self, other):
        """
        Test whether every element in the range is in *other*.

        :type other: SheetRange
        :param other: Other sheet range
        :return: ``True`` if *range* <= *other*.
        """
        if isinstance(other, CellRange):
            # Test whether sheet titles are equals (or if one of them is empty).
            this_title = self.title
            that_title = other.title
            eq_sheet_title = not this_title or not that_title or this_title.upper() == that_title.upper()
            return (eq_sheet_title and
                    (other.min_row <= self.min_row <= self.max_row <= other.max_row) and
                    (other.min_col <= self.min_col <= self.max_col <= other.max_col))
        raise TypeError(repr(type(other)))

    __le__ = issubset


    def __lt__(self, other):
        """
        Test whether every element in the range is in *other*, but not all.

        :type other: CellRange
        :param other: Other sheet range
        :return: ``True`` if *range* < *other*.
        """
        return self.__le__(other) and self.__ne__(other)


    def issuperset(self, other):
        """
        Test whether every element in *other* is in the range.

        :type other: CellRange or tuple[int, int]
        :param other: Other sheet range or cell index (*row_idx*, *col_idx*).
        :return: ``True`` if *range* >= *other* (or *other* in *range*).
        """
        if isinstance(other, CellRange):
            # Test whether sheet titles are equals (or if one of them is empty).
            this_title = self.title
            that_title = other.title
            eq_sheet_title = not this_title or not that_title or this_title.upper() == that_title.upper()
            return (eq_sheet_title and
                    (self.min_row <= other.min_row <= other.max_row <= self.max_row) and
                    (self.min_col <= other.min_col <= other.max_col <= self.max_col))
        elif isinstance(other, tuple):
            row_idx, col_idx = other  # cell index in worksheet._cells
            return ((self.min_row <= row_idx <= self.max_row) and
                    (self.min_col <= col_idx <= self.max_col))
        raise TypeError(repr(type(other)))

    __contains__ = __ge__ = issuperset


    def __gt__(self, other):
        """
        Test whether every element in *other* is in the range, but not all.

        :type other: CellRange
        :param other: Other sheet range
        :return: ``True`` if *range* > *other*.
        """
        return self.__ge__(other) and self.__ne__(other)


    def isdisjoint(self, other):
        """
        Return ``True`` if the range has no elements in common with other.
        Ranges are disjoint if and only if their intersection is the empty range.

        :type other: CellRange
        :param other: Other sheet range.
        :return: `True`` if the range has no elements in common with other.
        """
        if isinstance(other, CellRange):
            # Test whether sheet titles are different and not empty.
            this_title = self.title
            that_title = other.title
            ne_sheet_title = this_title and that_title and this_title.upper() != that_title.upper()
            return (ne_sheet_title or
                    (not (self.min_row <= other.min_row <= self.max_row) and
                     not (other.min_row <= self.max_row <= other.max_row)) or
                    (not (self.min_col <= other.min_col <= self.max_col) and
                     not (other.min_col <= self.max_col <= other.max_col)))
        raise TypeError(repr(type(other)))


    def intersection(self, *others):
        """
        Return a new range with elements common to the range and all *others*.

        :type others: tuple[CellRange]
        :param others: Other sheet ranges.
        :return: the current sheet range.
        :raise: :class:`ValueError` if an *other* range don't intersect
            with the current range.
        """
        for other in others:
            if isinstance(other, CellRange):
                if self.isdisjoint(other):
                    raise ValueError("Range {0} don't intersect {0}".format(self, other))
                self.min_row = max(self.min_row, other.min_row)
                self.max_row = min(self.max_row, other.max_row)
                self.min_col = max(self.min_col, other.min_col_idx)
                self.max_col = min(self.max_col, other.max_col)
                return self
            raise TypeError(repr(type(other)))
        return self

    __iand__ = intersection


    def __and__(self, other):
        return self.__copy__().__iand__(other)


    def union(self, other):
        """
        Return a new range with elements from the range and all *others*.

        :type others: tuple[CellRange]
        :param others: Other sheet ranges.
        :return: the current sheet range.
        """
        if isinstance(other, CellRange):
            self.min_row = min(self.min_row, other.min_row)
            self.max_row = max(self.max_row, other.max_row)
            self.min_col = min(self.min_col, other.min_col)
            self.max_col = max(self.max_col, other.max_col)
            return self
        raise TypeError(repr(type(other)))

    __ior__ = union


    def __or__(self, other):
        return self.__copy__().__ior__(other)


    def collapse(self, direction="top-left"):
        """
        Collapse the range to the given direction.

        :type direction: str
        :param direction: Collapsing direction:

            - to a single cell: "top-left", "top-right", "bottom-left", "bottom-right",
            - to a column: "left", "right",
            - to a row: "top", bottom".
        """
        parts = direction.split("-")
        # top and bottom are exclusive
        if "top" in parts:
            self.max_row_p = self.min_row_p
            self.max_row = self.min_row
        elif "bottom" in parts:
            self.min_row_p = self.max_row_p
            self.min_row = self.max_row
        # left and right are exclusive
        if "left" in parts:
            self.max_col_p = self.min_col_p
            self.max_col = self.min_col
        elif "right" in parts:
            self.min_col_p = self.max_col_p
            self.min_col = self.max_col


    def expand(self, min_col_idx, min_row_idx, max_col_idx, max_row_idx, direction):
        """
        Expand the range to the given direction in the bounding range.

        :type direction: str
        :param direction: Expansion direction: a combinaison of "left", "right", "top" or "bottom".
        """
        parts = direction.split("-")
        if "top" in parts:
            self.min_row = min_row_idx
        if "bottom" in parts:
            self.max_row = max_row_idx
        if "left" in parts:
            self.min_col = min_col_idx
        if "right" in parts:
            self.max_col = max_col_idx
        if self.min_row > self.max_row or self.min_col > self.max_col:
            raise ValueError("Invalid expanded range: {0}".format(self))


    def get_size(self):
        """ Return the size of the range (*count_cols*, *count_rows*). """
        count_cols = self.max_col + 1 - self.min_col
        count_rows = self.max_row + 1 - self.min_row
        return count_cols, count_rows


    @property
    def top(self):
        """
        Returns the cells for the top row
        """

    @property
    def left(self):
        """
        Returns the cells for the left column
        """

    @property
    def right(self):
        """
        Returns the cells for the right column
        """

    @property
    def bottom(self):
        """
        Returns the cells for the bottom row
        """