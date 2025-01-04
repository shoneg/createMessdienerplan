from constants import BLACK_LIST_PERSONEN, BLACK_LIST_TAGE, DATUM, ENCODING, GOTTESDIENST, TAG, ZEIT
from exit_methods import *

try:
    import datetime

    import pandas as pd
    import regex as re
    import json

    from docx2md import Converter, DocxFile, DocxMedia
except ImportError as e:
    print("Ein Import Fehler ist aufgetreten: " + str(e))
    input("Drücke Enter, um das Programm zu beenden.")
    exit_program()


def docx_table_to_html_md_table(
        src_path: str = "example_input/messplan.docx",
        dst_path: str = "example_input/gottesdienst.html") -> None:
    """
    Converts DOCX with a table to an HTML with table. So we can read it afterwards.
    :param src_path:
    :param dst_path:
    """
    docx = DocxFile(src_path)
    xml_text = docx.document()
    media = DocxMedia(docx)
    converter = Converter(xml_text, media, False)
    md_text = converter.convert()
    with open(dst_path, "w", encoding=ENCODING) as f:
        f.write(md_text)


def get_gottesdienstplan_from_html(path: str = "example_input/gottesdienst.html") -> pd.DataFrame:
    """
    Needs a table with gottesdiensten, then renames and rebuilds columns and return the result
    :param path: path to the gottesdienst table
    :return: a refactored table as DataFrame
    """
    # read source
    gottesdienst_table = pd.read_html(path)[0]
    # delete not used columns
    del gottesdienst_table[2], gottesdienst_table[4], gottesdienst_table[5], gottesdienst_table[6]
    # rename columns
    gottesdienst_table.rename(
        columns={0: DATUM, 1: ZEIT, 3: GOTTESDIENST}, inplace=True)
    # delete first, not used, row
    gottesdienst_table = gottesdienst_table.drop(0)

    # split Date to Date and Day
    day_of_week = []
    dates = []
    for date in gottesdienst_table[DATUM]:
        date_str = re.split(r'(^[^\d]+)', date)[1:]
        day_of_week.append(date_str[0])
        dates.append(date_str[1])
    gottesdienst_table[TAG] = day_of_week
    gottesdienst_table[DATUM] = dates

    # sort table for better debugging
    gottesdienst_table = gottesdienst_table.reindex(
        columns=[TAG, DATUM, ZEIT, GOTTESDIENST])
    return gottesdienst_table


def get_messdiener_from_csv(path: str = "example_input/messdiener.csv") -> pd.DataFrame:
    """
    Reads current messdiener from csv
    :param path: path to messdiener table
    :return: the read table without doing anything
    """
    messdiener_table = pd.read_csv(path, sep=';', encoding=ENCODING)
    for black_list in [BLACK_LIST_TAGE, BLACK_LIST_PERSONEN]:
        messdiener_table[black_list] = [*map(lambda day_str: '' if day_str is None or str(
            day_str) == 'nan' else day_str, messdiener_table[black_list])]
    return messdiener_table


def get_gottesdienst_arten_from_json(path: str = "example_input/gottesdienst-arten.json") -> dict:
    """
    Reads gottesdienste-arten, which need messdiener from json
    :param path: path to gottesdienst-arten
    :return: the read json without doing anything
    """
    with open(path, encoding=ENCODING) as gottesdienst_arten_file:
        gottesdienste_dict = json.load(gottesdienst_arten_file)
    return gottesdienste_dict


def export_table_to_html(export_table: pd.DataFrame, name: str = "table") -> None:
    """
    Exports given DataFrame as HTML table
    :param export_table: table to export
    :param name: name of the document
    """
    export_table.to_html(name, index=False, encoding=ENCODING)


def export_table_to_csv(export_table: pd.DataFrame, name: str = 'table') -> None:
    """
    Exports given DataFrame as CSV table
    :param export_table: table to export
    :param name: path to the new document
    """
    export_table.to_csv(name, sep=';', index=False, encoding=ENCODING)


def export_table_to_excel(export_table: pd.DataFrame, name: str = "table") -> None:
    """
    Exports given DataFrame as styled XLSX
    :param export_table: table to export
    :param name: path to the new document
    """
    sheet_name = f'Messdienerplan Stand {
        datetime.datetime.now().strftime("%d.%m.%Y")}'
    writer = pd.ExcelWriter(name, engine='xlsxwriter')
    export_table.to_excel(excel_writer=writer,
                          sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    (max_row, max_col) = export_table.shape
    worksheet.set_column(0, 2, 12)
    worksheet.set_column(3, 3, 24)
    worksheet.set_column(4, 4, 150)
    worksheet.autofilter(0, 0, max_row, max_col - 1)
    export_table_index = export_table.index
    last_date = None
    length = 0
    for i, day, date in zip(range(len(export_table[TAG])), export_table[TAG], export_table[DATUM]):
        if last_date == date:
            length += 1
        else:
            if length > 1:
                worksheet.merge_range(
                    f'A{i + 2 - length}:A{i + 1}', export_table[TAG][export_table_index[i - 1]])
                worksheet.merge_range(
                    f'B{i + 2 - length}:B{i + 1}', export_table[DATUM][export_table_index[i - 1]])
            last_date = date
            length = 1
        i += 1
    if length > 1:
        worksheet.merge_range(
            f'A{i + 2 - length}:A{i + 1}', export_table[TAG][export_table_index[i - 1]])
        worksheet.merge_range(
            f'B{i + 2 - length}:B{i + 1}', export_table[DATUM][export_table_index[i - 1]])
    writer.close()
