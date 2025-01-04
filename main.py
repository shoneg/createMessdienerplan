from typing import Tuple
from constants import ANZAHL, BLACK_LIST_TAGE, DATUM, EINTEILUNGEN, GOTTESDIENST, GOTTESDIENST_ARTEN_INPUT, GRUPPEN_SPLITTINGS, ID, MESSDIENER, MESSDIENER_INPUT, MESSDIENER_OUTPUT, MESSDIENERZAHL, MESSPLAN_INPUT, MESSPLAN_OUTPUT, MESSPLAN_TMP, TAG, ZEIT
from exit_methods import exit_program
from table_utils import docx_table_to_html_md_table, export_table_to_csv, export_table_to_excel, get_gottesdienst_arten_from_json, get_gottesdienstplan_from_html, get_messdiener_from_csv

try:
    import itertools
    import shutil
    import sys
    import warnings
    from os.path import exists
    from os import mkdir

    import numpy as np
    import pandas as pd
except ImportError as e:
    print("Ein Import Fehler ist aufgetreten: " + str(e))
    input("Drücke Enter, um das Programm zu beenden.")
    exit_program()


class Continue(Exception):
    pass


def remove_not_wanted_columns(gottesdienste: pd.DataFrame) -> pd.DataFrame:
    if MESSDIENERZAHL in gottesdienste:
        del gottesdienste[MESSDIENERZAHL]
    if GRUPPEN_SPLITTINGS in gottesdienste:
        del gottesdienste[GRUPPEN_SPLITTINGS]
    return gottesdienste


def reset_einteilungen(messdiener: pd.DataFrame) -> pd.DataFrame:
    new_einteilungen = []
    min = messdiener[messdiener[EINTEILUNGEN] ==
                     messdiener[EINTEILUNGEN].min()].values[0][3]
    for einteilung in messdiener[EINTEILUNGEN]:
        new_einteilungen.append(einteilung - min)
    messdiener[EINTEILUNGEN] = new_einteilungen
    return messdiener


def prepare_files():
    for file_name in [MESSPLAN_INPUT, MESSDIENER_INPUT, GOTTESDIENST_ARTEN_INPUT]:
        if not exists(file_name):
            print(f'Die Datei {file_name} existiert nicht. Falls du den Namen der Datei ' +
                  'dauerhaft ändern möchtest, kannst du das in der Datei constants.py machen')
            exit_program()
    if not exists('output'):
        mkdir("output")
    shutil.copy2(MESSDIENER_INPUT, MESSDIENER_OUTPUT)
    print(f'Die Datei {MESSDIENER_INPUT} wird während der Ausführung aktualisiere (Einteilungen),' +
          'falls etwas schief läuft wurde die ursprungsdatei nach {MESSDIENER_OUTPUT} kopiert')
    if not exists('tmp'):
        mkdir("tmp")


if __name__ == '__main__':
    warnings.filterwarnings('ignore')

    prepare_files()
    docx_table_to_html_md_table(MESSPLAN_INPUT, MESSPLAN_TMP)
    gottesdienste = get_gottesdienstplan_from_html(MESSPLAN_TMP)
    messdiener = get_messdiener_from_csv(MESSDIENER_INPUT)
    gottesdienst_arten = get_gottesdienst_arten_from_json(
        GOTTESDIENST_ARTEN_INPUT)
    gottesdienste = remove_not_wanted_columns(gottesdienste)
    messdiener = reset_einteilungen(messdiener)
    export_table_to_excel(gottesdienste, MESSPLAN_OUTPUT)
    print(f"Der Messdienerplan wurde nach {MESSPLAN_OUTPUT} exportiert")
    export_table_to_csv(messdiener, MESSDIENER_INPUT)
    print(f'Die Einteilungen in {MESSDIENER_INPUT} wurden aktualisiert')
    try:
        shutil.rmtree("tmp")
    except Exception:
        print(
            'Während der Ordner tmp gelöscht werden sollte ist ein Fehler aufgetreten.' +
            'Falls der Ordner noch existiert, kannst du es manuell versuchen, es ist aber auch ' +
            'nicht schlimm, wenn du nichts tust'
        )
    else:
        print("Temporäre Dateien wurden gelöscht")

    exit_program()
