# Copyright (c) 2010-2020 openpyxl

from openpyxl.compat import safe_string
from openpyxl.xml.functions import Element, SubElement, whitespace, XML_NS, REL_NS
from openpyxl import LXML
from openpyxl.utils.datetime import to_excel, days_to_time
from datetime import timedelta

from openpyxl.worksheet.datatable import TableFormula


def _set_attributes(cell, styled=None):
    """
    Set coordinate and datatype
    """
    coordinate = cell.coordinate
    attrs = {'r': coordinate}
    if styled:
        attrs['s'] = f"{cell.style_id}"

    if cell.data_type == "s":
        attrs['t'] = "inlineStr"
    elif cell.data_type != 'f':
        attrs['t'] = cell.data_type

    value = cell._value

    if cell.data_type == "d":
        if cell.parent.parent.iso_dates:
            if isinstance(value, timedelta):
                value = days_to_time(value)
            value = value.isoformat()
        else:
            attrs['t'] = "n"
            value = to_excel(value, cell.parent.parent.epoch)

    if cell.hyperlink:
        cell.parent._hyperlinks.append(cell.hyperlink)

    return value, attrs


def etree_write_cell(xf, worksheet, cell, styled=None):

    value, attributes = _set_attributes(cell, styled)

    el = Element("c", attributes)
    if value is None or value == "":
        xf.write(el)
        return

    if cell.data_type == 'f':
        attrib = {}
        coord = cell.coordinate
        if coord in worksheet.formula_attributes:
            attrib = worksheet.formula_attributes[coord]
        elif isinstance(value, TableFormula):
            attrib = dict(value)
            value = None

        formula = SubElement(el, 'f', attrib)
        if value is not None and not attrib.get('t') == "dataTable":
            formula.text = value[1:]
            value = None

    if cell.data_type == 's':
        inline_string = SubElement(el, 'is')
        text = SubElement(inline_string, 't')
        text.text = value
        whitespace(text)


    else:
        cell_content = SubElement(el, 'v')
        if value is not None:
            cell_content.text = safe_string(value)

    xf.write(el)


def lxml_write_cell(xf, worksheet, cell, styled=False):
    value, attributes = _set_attributes(cell, styled)

    if value == '' or value is None:
        with xf.element("c", attributes):
            return

    with xf.element('c', attributes):
        if cell.data_type == 'f':
            attrib = {}
            coord = cell.coordinate
            if coord in worksheet.formula_attributes:
                attrib = worksheet.formula_attributes[coord]
            elif isinstance(value, TableFormula):
                attrib = dict(value)
                value = None

            with xf.element('f', attrib):
                if value is not None and not attrib.get('t') == "dataTable":
                    xf.write(value[1:])
                    value = None

        if cell.data_type == 's':
            with xf.element("is"):
                attrs = {}
                if value != value.strip():
                    attrs["{%s}space" % XML_NS] = "preserve"
                el = Element("t", attrs) # lxml can't handle xml-ns
                el.text = value
                xf.write(el)
                #with xf.element("t", attrs):
                    #xf.write(value)
        else:
            with xf.element("v"):
                if value is not None:
                    xf.write(safe_string(value))


if LXML:
    write_cell = lxml_write_cell
else:
    write_cell = etree_write_cell
