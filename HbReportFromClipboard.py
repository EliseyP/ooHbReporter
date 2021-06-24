# -*- coding: utf_8 -*-

# import locale
import uno

from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK
from com.sun.star.text.TextContentAnchorType import AS_CHARACTER
from com.sun.star.awt import Size
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
from com.sun.star.style.ParagraphAdjust import RIGHT, CENTER, LEFT
from com.sun.star.uno.TypeClass import STRING
from com.sun.star.util.NumberFormat import CURRENCY
from com.sun.star.awt.FontWeight import NORMAL, BOLD

headers_dic = {
    'date': 'Дата',
    'paymode': 'Платеж',
    'info': 'Сведения',
    'payee': 'Получатель',
    'memo': 'Заметки',
    'amount': 'Сумма',
    'c': 'Статус',
    'category': 'Категория',
    'tags': 'Метка',
    'Total': 'Итого'
}

language = 'en'
need_translate = False


class AppStrings:
    INCOME_EXPENSE = 'Income / Expense'
    CLIPBOARD_NOT_CONTAIN_TABLE = 'Clipboard not contain table!'
    NON_CORRECT_DATA_IN_CLIPBOARD = 'Non correct data in clipboard!'


translate_ru = {
    AppStrings.INCOME_EXPENSE: "Приход / расход",
    AppStrings.CLIPBOARD_NOT_CONTAIN_TABLE: 'Буфер обмена не содержит таблицы!',
    AppStrings.NON_CORRECT_DATA_IN_CLIPBOARD: 'Неверные данные в буфере обмена!',
}

langs_dic = {
    'ru': translate_ru
}


def translate(_string):
    if need_translate:
        translate_dic = langs_dic.get(language)
        if translate_dic:
            return translate_dic.get(_string, _string)
        else:
            return _string
    else:
        return _string


def MsgBox(message, title=''):
    '''MsgBox'''
    # desktop = XSCRIPTCONTEXT.getDesktop()
    # doc = desktop.getCurrentComponent()
    doc = get_current_component()
    parent_window = doc.CurrentController.Frame.ContainerWindow
    box = parent_window.getToolkit().createMessageBox(parent_window, MESSAGEBOX, BUTTONS_OK, title, message)
    box.execute()
    return None


def ErrorBox(message, title=''):
    '''MsgBox'''
    # desktop = XSCRIPTCONTEXT.getDesktop()
    # doc = desktop.getCurrentComponent()
    doc = get_current_component()
    parent_window = doc.CurrentController.Frame.ContainerWindow
    box = parent_window.getToolkit().createMessageBox(parent_window, ERRORBOX, BUTTONS_OK, title, message)
    box.execute()
    return None


def Mri(target):
    ctx = uno.getComponentContext()
    _mri = ctx.ServiceManager.createInstanceWithContext(
        "mytools.Mri", ctx)
    _mri.inspect(target)


def getClipboardText(_ctx):
    _string = ''
    _smgr = _ctx.getServiceManager()
    clip = _smgr.createInstanceWithContext(
        "com.sun.star.datatransfer.clipboard.SystemClipboard", _ctx)
    converter = _smgr.createInstanceWithContext(
        "com.sun.star.script.Converter", _ctx)
    clip_contents = clip.getContents()
    types = clip_contents.getTransferDataFlavors()

    for i, _type in enumerate(types):
        if _type.MimeType == "text/plain;charset=utf-16":
            _string = converter.convertToSimpleType(
                clip_contents.getTransferData(_type), STRING)
            break
    return _string


def insertTextIntoCell(table, cell_name, text, color):
    table_text = table.getCellByName(cell_name)
    cursor = table_text.createTextCursor()

    cursor.setPropertyValue("CharColor", color)
    table_text.setString(text)


def get_current_component():
    _ctx = uno.getComponentContext()
    _smgr = _ctx.getServiceManager()
    _desktop = _smgr.createInstanceWithContext('com.sun.star.frame.Desktop', _ctx)
    _doc = _desktop.getCurrentComponent()
    if _doc:
        return _doc


def get_data_array(_text):
    out_list = []
    # Some validate
    if _text.count('\n') == 0 or _text.count('\t') == 0:
        ErrorBox(translate(AppStrings.CLIPBOARD_NOT_CONTAIN_TABLE))
        return None

    # If data from clipboard, need remove last symbol
    _text = _text[:-1]
    _rows_list = _text.split('\n')
    _rows_lenth_set = set()
    for _row in _rows_list:
        _row = _row[:-1]
        _cols_list = _row.split('\t')
        out_list.append(_cols_list)
        _rows_lenth_set.add(len(_cols_list))
    # Some validate
    if len(_rows_lenth_set) != 1:
        ErrorBox(translate(AppStrings.NON_CORRECT_DATA_IN_CLIPBOARD))
        return None

    return out_list


def get_number_format(_doc, _locale):
    formats = _doc.NumberFormats
    number_format_key = \
        formats.getStandardFormat(CURRENCY, _locale)
    if number_format_key:
        return number_format_key
    else:
        return None


def table_fill(_doc, _table, _data, _locale):
    _rows_amount = len(_data)
    _cols_amount = len(_data[0])
    _number_format = get_number_format(_doc, _locale)
    for _row_ind, _row in enumerate(_data):
        for _col_ind, _column in enumerate(_row):
            _cell = _table.getCellByPosition(_col_ind, _row_ind)
            _data_value = _column
            if _data_value:
                try:
                    _data_float = float(_data_value)
                except ValueError:
                    # String
                    _cell.setString(str(_data_value))
                else:
                    # Float
                    _cell.setValue(_data_float)
                    if _number_format:
                        _cell.NumberFormat = _number_format  # Financial
                    _cell.RightBorderDistance = 500
                    cursor = _cell.createTextCursor()
                    cursor.ParaAdjust = RIGHT
                    if _data_float == 0.0:
                        _cell.setString('')


def table_select_entire(_doc, _table):
    # Select entire Table
    view_cursor = _doc.getCurrentController().getViewCursor()
    _doc.getCurrentController().select(_table)
    view_cursor.gotoEnd(True)  # Move to the end of the current cell
    view_cursor.gotoEnd(True)  # Move to the end of the table


def data_remove_empty_columns(_data, _empty_list_rev):
    _out = []
    for _row in _data:
        for _ind in _empty_list_rev:
            _row = del_by_index_from_list(_row, _ind)
        _out.append(_row)
    return _out


def table_get_remove_empty_columns(_data):
    _rows_amount = len(_data)
    _cols_amount = len(_data[0])
    _empty_columns_list = []
    for _col_ind in range(1, _cols_amount):
        _column_is_empty = True
        for _row_ind in range(1, _rows_amount):
            _value = _data[_row_ind][_col_ind]
            if _value and _value != '0':
                _column_is_empty = False
        if _column_is_empty:
            _empty_columns_list.append(_col_ind)
    return _empty_columns_list


def del_by_index_from_list(_list, _index):
    return _list[:_index] + _list[_index+1:]


def insert_report(*args):
    """Inserts a table from clipboard, and format it.

    """

    def dispatch(_uno_string):
        dispatcher.executeDispatch(frame, f".uno:{_uno_string}", "", 0, ())

    def table_set_cols_optimal_width():
        dispatch('SetOptimalColumnWidth')

    def first_row_format(_table, _translate=False):
        first_cell = _table.getCellByPosition(0, 0)
        first_cell_string = first_cell.getString()
        if _translate:
            if first_cell_string in headers_dic.keys():
                for _col_ind in range(_cols_amount):
                    _cell = _table.getCellByPosition(_col_ind, 0)
                    _cell_string = _cell.getString()
                    _translated_string = headers_dic.get(_cell_string, _cell_string)
                    _cell.setString(_translated_string)

        fist_row_range = _table.getCellRangeByPosition(
            0, 0, _cols_amount - 1, 0)
        fist_row_range.ParaAdjust = CENTER

    def last_row_format(_doc, _table):
        left_bottom_cell = _table.getCellByPosition(0, _rows_amount - 1)
        last_row_range = _table.getCellRangeByPosition(
            0, _rows_amount - 1, _cols_amount - 1, _rows_amount - 1)

        # Check if last row is Total string
        if left_bottom_cell.getString() in ('Total', 'Итого'):
            last_row_range.BackColor = 11388391
        else:
            # Usual string
            # Font weight
            for _col_ind in range(_cols_amount):
                _cell = _table.getCellByPosition(_col_ind, _rows_amount - 1)
                _cursor = _cell.getText().createTextCursor()
                _cursor.gotoEnd(True)
                _cursor.setPropertyValue("CharWeight", NORMAL)  # normal

            # BackColor for odd or even last string.
            _table_template_style = \
                _doc.getStyleFamilies().getByName("TableStyles").getByName(table_template)
            odd_rows_back_color = \
                _table_template_style.getByName("odd-rows").BackColor
            even_rows_back_color = \
                _table_template_style.getByName("even-rows").BackColor
            if _rows_amount - 1 % 2 == 0 or _rows_amount == 2:
                # WHITE
                last_row_range.BackColor = even_rows_back_color
            else:
                # GRAY
                last_row_range.BackColor = odd_rows_back_color

    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    # In current opened document.
    doc = desktop.getCurrentComponent()
    document = XSCRIPTCONTEXT.getDocument()
    frame = document.CurrentController.Frame

    locale = uno.createUnoStruct("com.sun.star.lang.Locale")
    locale.Country = doc.CharLocale.Country
    locale.Language = doc.CharLocale.Language
    global language
    language = locale.Language
    global need_translate
    if language != 'en':
        need_translate = True

    _from_buffer = getClipboardText(ctx)
    data = get_data_array(_from_buffer)
    if not data:
        return None

    empty_columns_list = table_get_remove_empty_columns(data)
    empty_columns_list_rev = empty_columns_list[::-1]
    data = data_remove_empty_columns(data, empty_columns_list_rev)
    _rows_amount = len(data)
    _cols_amount = len(data[0])

    # If need create new writer document:
    # open a writer document
    # doc = desktop.loadComponentFromURL(
    #     "private:factory/swriter", "_blank", 0, ())

    text = doc.Text
    text_cursor = text.createTextCursor()

    # Title string
    title_string = translate(AppStrings.INCOME_EXPENSE)
    # Title style
    text_cursor.ParaStyleName = "Title"
    text_cursor.CharHeight = 16
    text.insertString(text_cursor, title_string, 0)

    # create a text table
    table = doc.createInstance("com.sun.star.text.TextTable")

    # with rows and  columns
    table.initialize(_rows_amount, _cols_amount)

    text.insertTextContent(text_cursor, table, 0)

    # Default Style
    # Academic
    # Box List Blue
    # Box List Green
    # Box List Red
    # Box List Yellow
    # Elegant
    # Financial
    # Simple Grid Columns
    # Simple Grid Rows
    # Simple List Shaded

    table_template = "Box List Blue"
    table.TableTemplateName = table_template

    # Fill table from data array.
    table_fill(doc, table, data, locale)

    first_row_format(table, need_translate)
    last_row_format(doc, table)

    table_select_entire(doc, table)
    table_set_cols_optimal_width()


g_exportedScripts = (
    insert_report,
)

# vim: set shiftwidth=4 softtabstop=4 expandtab:
