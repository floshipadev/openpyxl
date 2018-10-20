from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

import pytest

import datetime
from io import BytesIO
from zipfile import ZipFile

from lxml.etree import iterparse, fromstring

from openpyxl import load_workbook
from openpyxl.compat import unicode
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.worksheet import Worksheet
from openpyxl.worksheet.pagebreak import Break, PageBreak
from openpyxl.worksheet.scenario import ScenarioList, Scenario, InputCells
from openpyxl.packaging.relationship import Relationship, RelationshipList
from openpyxl.utils.datetime  import CALENDAR_WINDOWS_1900, CALENDAR_MAC_1904
from openpyxl.styles.styleable import StyleArray
from openpyxl.styles import numbers
from openpyxl.formula.translate import Translator


@pytest.fixture
def Workbook():


    class DummyStyle:
        number_format = numbers.FORMAT_GENERAL
        font = ""
        fill = ""
        border = ""
        alignment = ""
        protection = ""

        def copy(self, **kw):
            return self


    class DummyWorkbook:

        data_only = False
        _colors = []
        encoding = "utf8"
        epoch = CALENDAR_WINDOWS_1900

        def __init__(self):
            self._differential_styles = []
            self.shared_strings = IndexedList()
            self.shared_strings.add("hello world")
            self._fonts = IndexedList()
            self._fills = IndexedList()
            self._number_formats = IndexedList()
            self._borders = IndexedList()
            self._alignments = IndexedList()
            self._protections = IndexedList()
            self._cell_styles = IndexedList()
            self.vba_archive = None
            for i in range(29):
                self._cell_styles.add((StyleArray([i]*9)))
            self._cell_styles.add(StyleArray([0,4,6,0,0,1,0,0,0])) #fillId=4, borderId=6, alignmentId=1))
            self.sheetnames = []


        def create_sheet(self, title):
            return Worksheet(self)


    return DummyWorkbook()


@pytest.fixture
def WorkSheetParser(Workbook):
    """Setup a parser instance with an empty source"""
    from .._reader import WorkSheetParser

    styles = IndexedList()
    for i in range(29):
        styles.add((StyleArray([i]*9)))
    styles.add(StyleArray([0,4,6,14,0,1,0,0,0])) #fillId=4, borderId=6, number_format=14 alignmentId=1))
    return WorkSheetParser(None, {0:'a'}, styles=styles)


@pytest.fixture
def WorkSheetParserKeepVBA(Workbook):
    """Setup a parser instance with an empty source"""
    Workbook.vba_archive=True
    from .._reader import WorkSheetParser
    ws = Workbook.create_sheet('sheet')
    return WorkSheetParser(None, {0:'a'}, {})


@pytest.mark.parametrize("filename, expected",
                         [
                             ("dimension.xml", (4, 1, 27, 30)),
                             ("no_dimension.xml", None),
                             ("invalid_dimension.xml", (None, 1, None, 113)),
                          ]
                         )
def test_read_dimension(WorkSheetParser, datadir, filename, expected):
    datadir.chdir()
    parser = WorkSheetParser
    parser.source = filename

    dimension = parser.parse_dimensions()
    assert dimension == expected


def test_col_width(datadir, WorkSheetParser):
    datadir.chdir()
    parser = WorkSheetParser
    with open("complex-styles-worksheet.xml", "rb") as src:
        cols = iterparse(src, tag='{%s}col' % SHEET_MAIN_NS)
        for _, col in cols:
            parser.parse_column_dimensions(col)
    assert set(parser.column_dimensions) == set(['A', 'C', 'E', 'I', 'G'])
    assert dict(parser.column_dimensions['A']) == {'max': '1', 'min': '1',
                                               'customWidth': '1',
                                               'width': '31.1640625'}


def test_hidden_col(WorkSheetParser):
    parser = WorkSheetParser

    src = """<col min="4" max="4" width="0" hidden="1" customWidth="1"/>"""
    element = fromstring(src)

    parser.parse_column_dimensions(element)
    assert 'D' in parser.column_dimensions
    assert dict(parser.column_dimensions['D']) == {'customWidth': '1', 'hidden':
                                               '1', 'max': '4', 'min': '4'}


def test_styled_col(datadir, WorkSheetParser):
    datadir.chdir()
    parser = WorkSheetParser

    with open("complex-styles-worksheet.xml", "rb") as src:
        cols = iterparse(src, tag='{%s}col' % SHEET_MAIN_NS)
        for _, col in cols:
            parser.parse_column_dimensions(col)
    assert 'I' in parser.column_dimensions
    cd = parser.column_dimensions['I']
    assert cd._style == StyleArray([28]*9)
    assert dict(cd) ==  {'customWidth': '1', 'max': '9', 'min': '9', 'width': '25'}


def test_row_dimensions(WorkSheetParser):
    src = """<row r="2" spans="1:6" />"""
    element = fromstring(src)

    parser = WorkSheetParser
    parser.parse_row(element)

    assert 2 not in parser.row_dimensions


def test_hidden_row(WorkSheetParser):
    parser = WorkSheetParser

    src = """
    <row r="2" spans="1:4" hidden="1" />
    """
    element = fromstring(src)

    parser.parse_row(element)
    assert dict(parser.row_dimensions[2]) == {'hidden': '1'}


def test_styled_row(datadir, WorkSheetParser):
    parser = WorkSheetParser
    src = """<row r="23" s="28" spans="1:8" />"""
    element = fromstring(src)

    parser.parse_row(element)
    rd = parser.row_dimensions[23]
    assert rd._style == StyleArray([28]*9)
    assert dict(rd) == {'customFormat':'1'}


def test_sheet_protection(datadir, WorkSheetParser):
    parser = WorkSheetParser

    src = """
    <sheetProtection xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" password="DAA7" sheet="1" formatCells="0" formatColumns="0" formatRows="0" insertColumns="0" insertRows="0" insertHyperlinks="0" deleteColumns="0" deleteRows="0" sort="0" autoFilter="0" pivotTables="0"/>
    """
    element = fromstring(src)
    parser.parse_sheet_protection(element)

    assert dict(parser.protection) == {
        'autoFilter': '0', 'deleteColumns': '0',
        'deleteRows': '0', 'formatCells': '0', 'formatColumns': '0', 'formatRows':
        '0', 'insertColumns': '0', 'insertHyperlinks': '0', 'insertRows': '0',
        'objects': '0', 'password': 'DAA7', 'pivotTables': '0', 'scenarios': '0',
        'selectLockedCells': '0', 'selectUnlockedCells': '0', 'sheet': '1', 'sort':
        '0'
    }


def test_formula(WorkSheetParser):
    parser = WorkSheetParser

    src = """
    <c r="A1" t="str" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <f>IF(TRUE, "y", "n")</f>
        <v>y</v>
    </c>
    """
    element = fromstring(src)

    cell = parser.parse_cell(element)
    assert cell == {'column': 1, 'data_type': 'f', 'row': 1,
                    'style_id':0, 'value': '=IF(TRUE, "y", "n")'}


def test_formula_data_only(WorkSheetParser):
    parser = WorkSheetParser
    parser.data_only = True

    src = """
    <c r="A1" t="str" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <f>1+2</f>
        <v>3</v>
    </c>
    """
    element = fromstring(src)

    cell = parser.parse_cell(element)
    assert cell == {'column': 1, 'data_type': 'n', 'row': 1,
                    'style_id':0, 'value': 3}


def test_string_formula_data_only(WorkSheetParser):
    parser = WorkSheetParser
    parser.data_only = True

    src = """
    <c r="A1" t="str" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <f>IF(TRUE, "y", "n")</f>
        <v>y</v>
    </c>
    """
    element = fromstring(src)

    cell = parser.parse_cell(element)
    assert cell == {'column': 1, 'data_type': 's', 'row': 1,
                    'style_id':0, 'value': 'y'}


def test_number(WorkSheetParser):
    parser = WorkSheetParser

    src = """
    <c r="A1" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <v>1</v>
    </c>
    """
    element = fromstring(src)

    cell = parser.parse_cell(element)
    assert cell == {'column': 1, 'data_type': 'n', 'row': 1,
                    'style_id':0, 'value': 1}



def test_datetime(WorkSheetParser):
    parser = WorkSheetParser

    src = """
    <c r="A1" t="d" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <v>2011-12-25T14:23:55</v>
    </c>
    """
    element = fromstring(src)

    cell = parser.parse_cell(element)
    assert cell == {'column': 1, 'data_type': 'd', 'row': 1,
                    'style_id':0, 'value': datetime.datetime(2011, 12, 25, 14, 23, 55)}


def test_mac_date(WorkSheetParser):
    parser = WorkSheetParser
    parser.epoch = CALENDAR_MAC_1904

    src = """
    <c r="A1" t="n" s="29" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <v>41184</v>
    </c>
    """
    element = fromstring(src)

    cell = parser.parse_cell(element)
    assert cell == {'column': 1, 'data_type': 'd', 'row': 1,
                    'style_id':29, 'value':datetime.datetime(2016, 10, 3, 0, 0)}


def test_string(WorkSheetParser):
    parser = WorkSheetParser

    src = """
    <c r="A1" t="s" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <v>0</v>
    </c>
    """
    element = fromstring(src)

    cell = parser.parse_cell(element)
    assert cell == {'column': 1, 'data_type': 's', 'row': 1,
                    'style_id':0, 'value':'a'}


def test_boolean(WorkSheetParser):
    parser = WorkSheetParser

    src = """
    <c r="A1" t="b" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <v>1</v>
    </c>
    """
    element = fromstring(src)

    cell = parser.parse_cell(element)
    assert cell == {'column': 1, 'data_type': 'b', 'row': 1,
                    'style_id':0, 'value':True}


def test_inline_string(WorkSheetParser):
    parser = WorkSheetParser
    src = """
    <c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="A1" s="0" t="inlineStr">
      <is>
      <t>ID</t>
      </is>
    </c>
    """

    element = fromstring(src)

    cell = parser.parse_cell(element)
    assert cell == {'column': 1, 'data_type': 's', 'row': 1,
                    'style_id':0, 'value':"ID"}


def test_inline_richtext(WorkSheetParser):
    parser = WorkSheetParser
    src = """
    <c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="R2" s="4" t="inlineStr">
    <is>
      <r>
        <rPr>
          <sz val="8.0" />
        </rPr>
        <t xml:space="preserve">11 de September de 2014</t>
      </r>
      </is>
    </c>
    """

    element = fromstring(src)
    cell = parser.parse_cell(element)
    assert cell == {'column': 18, 'data_type': 's', 'row': 2,
                    'style_id':4, 'value':"11 de September de 2014"}


@pytest.mark.xfail
def test_legacy_drawing(datadir):
    datadir.chdir()
    wb = load_workbook("legacy_drawing.xlsm", keep_vba=True)
    sheet1 = wb['Sheet1']
    assert sheet1.legacy_drawing == 'xl/drawings/vmlDrawing1.vml'
    sheet2 = wb['Sheet2']
    assert sheet2.legacy_drawing == 'xl/drawings/vmlDrawing2.vml'


def test_sheet_views(WorkSheetParser):
    parser = WorkSheetParser

    src = b"""
    <sheetViews xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetView tabSelected="1" zoomScale="200" zoomScaleNormal="200" zoomScalePageLayoutView="200" workbookViewId="0">
        <pane xSplit="5" ySplit="19" topLeftCell="F20" activePane="bottomRight" state="frozenSplit"/>
        <selection pane="topRight" activeCell="F1" sqref="F1"/>
        <selection pane="bottomLeft" activeCell="A20" sqref="A20"/>
        <selection pane="bottomRight" activeCell="E22" sqref="E22"/>
      </sheetView>
    </sheetViews>
    """

    parser.source = source = BytesIO(src)
    parser.parse()

    view = parser.views.sheetView[0]

    assert view.zoomScale == 200
    assert len(view.selection) == 3


@pytest.mark.parametrize("keep_vba, result",
                         [
                             ("True", "rId3"),
                             ("False", None),
                         ]
                         )
def test_legacy_document_(WorkSheetParserKeepVBA, keep_vba, result):
    parser = WorkSheetParserKeepVBA
    parser.keep_vba = keep_vba

    src = """<legacyDrawing r:id="rId3" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>"""
    element = fromstring(src)

    legacy = parser.parse_legacy_drawing(element)
    assert legacy == 'rId3'


def test_shared_formula(WorkSheetParser):
    parser = WorkSheetParser
    src = """
    <c r="A9" t="str" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <f t="shared" si="0"/>
      <v>9</v>
    </c>
    """
    element = fromstring(src)
    parser.shared_formulae['0'] = Translator("=A4*B4", "A1")
    formula = parser.parse_formula(element)
    assert formula == "=A12*B12"


def test_array_formula(WorkSheetParser, datadir):
    parser = WorkSheetParser

    src = """
    <c r="C10" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <f t="array" ref="C10:C14">SUM(A10:A14*B10:B14)</f>
    </c>"""
    element = fromstring(src)

    parser.parse_formula(element)
    assert parser.array_formulae['C10']['ref'] == 'C10:C14'


from warnings import simplefilter
simplefilter("always")

def test_extended_conditional_formatting(WorkSheetParser, recwarn):
    parser = WorkSheetParser

    src = """
    <extLst>
        <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
            <x14:conditionalFormattings>
                <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                    <x14:cfRule type="containsText" priority="1" operator="containsText" id="{5F3BF262-1B4E-49BC-812D-8F5B1802CE0A}">
                        <xm:f>NOT(ISERROR(SEARCH($AA$7,A1)))</xm:f>
                        <xm:f>$AA$7</xm:f>
                        <x14:dxf>
                            <fill>
                                <patternFill>
                                    <bgColor rgb="FF6699FF"/>
                                </patternFill>
                            </fill>
                        </x14:dxf>
                    </x14:cfRule>
                </x14:conditionalFormatting>
            </x14:conditionalFormattings>
        </ext>
    </extLst>
    """
    element = fromstring(src)

    parser.parse_extensions(element)
    w = recwarn.pop()

    assert issubclass(w.category, UserWarning)


def test_cell_without_coordinates(WorkSheetParser):
    parser = WorkSheetParser

    src = """
    <row xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <c t="s">
        <v>2</v>
      </c>
      <c t="s">
        <v>4</v>
      </c>
      <c t="s">
        <v>3</v>
      </c>
      <c t="s">
        <v>6</v>
      </c>
      <c t="s">
        <v>9</v>
      </c>
    </row>
    """
    element = fromstring(src)

    parser.shared_strings = ["Whatever"] * 10
    parser.parse_row(element)

    assert parser.max_row == 1
    assert parser.max_column == 5


def test_external_hyperlinks(WorkSheetParser):
    src = b"""
    <hyperlinks xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       display="http://test.com" r:id="rId1" ref="A1"/>
    </hyperlinks>
    """
    element = fromstring(src)
    #from openpyxl.packaging.relationship import Relationship, RelationshipList

    #r = Relationship(type="hyperlink", Id="rId1", Target="../")
    #rels = RelationshipList()
    #rels.append(r)

    parser = WorkSheetParser
    parser.source = BytesIO(src)
    #parser.ws._rels = rels

    parser.parse()
    assert len(parser.hyperlinks.hyperlink) == 1

    #assert parser.ws['A1'].hyperlink.target == "../"


def test_local_hyperlinks(WorkSheetParser):
    src = b"""
     <hyperlinks xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <hyperlink ref="B4:B7" location="'STP nn000TL-10, PKG 2.52'!A1" display="STP 10000TL-10"/>
    </hyperlinks>
    """
    element = fromstring(src)
    parser = WorkSheetParser
    parser.source = BytesIO(src)
    parser.parse()

    assert len(parser.hyperlinks.hyperlink) == 1

    #assert parser.ws['B4'].hyperlink.location == "'STP nn000TL-10, PKG 2.52'!A1"


#@pytest.mark.xfail
def test_merge_cells(WorkSheetParser):
    src = b"""
    <sheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <mergeCells>
        <mergeCell ref="C2:F2"/>
        <mergeCell ref="B19:C20"/>
        <mergeCell ref="E19:G19"/>
      </mergeCells>
    </sheet>
    """

    parser = WorkSheetParser
    parser.source = BytesIO(src)

    parser.parse()

    assert len(parser.merged_cells.mergeCell) == 3

    #assert parser.ws.merged_cells == "C2:F2 B19:C20 E19:G19"


def test_conditonal_formatting(WorkSheetParser):
    src = b"""
    <sheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <conditionalFormatting sqref="S1:S10">
        <cfRule type="top10" dxfId="25" priority="12" percent="1" rank="10"/>
    </conditionalFormatting>
    <conditionalFormatting sqref="T1:T10">
      <cfRule type="top10" dxfId="24" priority="11" bottom="1" rank="4"/>
    </conditionalFormatting>
    </sheet>
    """
    from openpyxl.styles.differential import DifferentialStyle

    parser = WorkSheetParser
    #dxf = DifferentialStyle()
    #parser.differential_styles = [dxf] * 30
    parser.source = BytesIO(src)

    parser.parse()
    assert len(parser.formatting) == 2
    #assert parser.ws.conditional_formatting['T1:T10'][-1].dxf == dxf


def test_sheet_properties(WorkSheetParser):
    src = b"""
    <sheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr codeName="Sheet3">
      <tabColor rgb="FF92D050"/>
      <outlinePr summaryBelow="1" summaryRight="1"/>
      <pageSetUpPr/>
    </sheetPr>
    </sheet>
    """
    parser = WorkSheetParser
    parser.source = BytesIO(src)
    parser.parse()

    assert parser.sheet_properties.tabColor.rgb == "FF92D050"
    assert parser.sheet_properties.codeName == "Sheet3"


def test_sheet_format(WorkSheetParser):

    src = b"""
    <sheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetFormatPr defaultRowHeight="14.25" baseColWidth="15"/>
    </sheet>
    """
    parser = WorkSheetParser
    parser.source = BytesIO(src)
    parser.parse()

    assert parser.sheet_format.defaultRowHeight == 14.25
    assert parser.sheet_format.baseColWidth == 15


def test_tables(WorkSheetParser):
    src = b"""
    <tableParts count="1" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <tablePart r:id="rId1"/>
    </tableParts>
    """
    parser = WorkSheetParser
    parser.source = BytesIO(src)
    parser.parse()

    assert parser.tables.tablePart[0].id == "rId1"


def test_auto_filter(WorkSheetParser):
    src = b"""
    <sheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <autoFilter ref="A1:AK3237">
        <sortState ref="A2:AM3269">
            <sortCondition ref="B1:B3269"/>
        </sortState>
      </autoFilter>
    </sheet>
    """

    parser = WorkSheetParser
    parser.source = BytesIO(src)
    parser.parse()

    assert parser.auto_filter.ref == "A1:AK3237"
    assert parser.auto_filter.sortState.ref == "A2:AM3269"


def test_page_break(WorkSheetParser):
    src = b"""
    <sheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <rowBreaks count="1" manualBreakCount="1">
            <brk id="15" man="1" max="16383" min="0"/>
        </rowBreaks>
    </sheet>
    """
    expected_pagebreak = PageBreak()
    expected_pagebreak.append(Break(id=15))

    parser = WorkSheetParser
    parser.source = BytesIO(src)
    parser.parse()

    assert parser.page_breaks == expected_pagebreak


def test_scenarios(WorkSheetParser):
    src = b"""
    <sheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <scenarios current="0" show="0">
    <scenario name="Worst case" locked="1" user="User" comment="comment">
      <inputCells r="B2" val="50000" />
    </scenario>
    </scenarios>
    </sheet>
    """

    c = InputCells(r="B2", val="50000")
    s = Scenario(name="Worst case", inputCells=[c], locked=True, user="User", comment="comment")
    scenarios = ScenarioList(scenario=[s], current="0", show="0")

    parser = WorkSheetParser
    parser.source = BytesIO(src)
    parser.parse()

    assert parser.scenarios == scenarios
