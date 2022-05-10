import getopt
import logging
import os
import sqlite3
from sqlite3 import Connection, Cursor
from typing import Optional

from pandas import DataFrame

from exit_methods import *
from flask import Flask

try:
    import itertools
    import shutil
    import sys
    import warnings
    from os.path import exists
    from os import mkdir

    import numpy as np

    from table_utils import *
except ImportError as e:
    print("Ein Import Fehler ist aufgetreten: " + str(e))
    input("Drücke Enter, um das Programm zu beenden.")
    exit_programm()


class Continue(Exception):
    pass


def add_messdiener_count(gottesdienste: pd.DataFrame, gottesdienst_arten: dict) -> pd.DataFrame:
    """
    Asks interactive for count of messdiener per gottesdienst and gives as default value from gottesdienst-arten. Returns gottesdienste with extra column with the learned count.
    :param gottesdienste: table with Tag, Datum, Zeit, Gottesdienst, index
    :param gottesdienst_arten: if a gottesdienst not given, we assume zero
    :return: the given table with extra column Messdienerzahl
    """
    print(
        "Tippe die Zahl an Messdienern ein, die du für diesen Gottesdienst einteilen möchtest. Falls du keine Zahl tippst, wird die in den eckigen Klammern verwendet. Wenn du die vorherige Eingabe rückgängig machen möchtest, tippe -1")
    counts = [0] * len(gottesdienste.index)
    ret = gottesdienste

    indexes = gottesdienste.index
    i = 0
    while i < len(indexes):
        gottesdienst = gottesdienste.iloc[i]
        # takes given from dict if existing else 0
        default_count = gottesdienst_arten[gottesdienst[GOTTESDIENST]] if gottesdienst[
                                                                              GOTTESDIENST] in gottesdienst_arten else 0
        # string to ask with
        question = f"{gottesdienst[TAG]}, {gottesdienst[DATUM]} {gottesdienst[ZEIT]} - {gottesdienst[GOTTESDIENST]} [{default_count}]: "
        # takes count for this gottesdienst
        count = int(input(question) or default_count)
        # if <0 we want to go one gottesdienst back
        if count < 0:
            i = max(0, i - 1)
            continue
        counts[i] = count
        i += 1
        if i >= len(indexes) and (str(input('Eingabe beenden und mit Zuteilung starten?[ja]: ')) or 'ja') != 'ja':
            i -= 1
            continue

    # add column
    ret[MESSDIENERZAHL] = counts
    # if 0 Messdiener needed we don't need this gottesdienst anymore, else it's in return table
    ret = ret[ret[MESSDIENERZAHL] > 0]

    return ret


def split_in_summanden(number: int):
    """
    Calculates a list of all possibilities of splitting a number into summanden
    :param number:
    :return:
    """
    # route
    if number == 1:
        return [[1]]
    # take first number iterating and then recursive the rest
    ret = [[number]]
    for i in np.arange(1, number):
        for j in split_in_summanden(number - i):
            ret.append(sorted([i] + j))
    return np.unique(ret)


def group_splittings(gottesdienste: pd.DataFrame) -> pd.DataFrame:
    """
    Adds a column with possible splittings of the Messdienerzahl column
    :param gottesdienste:
    :return: gottesdienste
    """
    splittings = []
    # for every row calculate the splittings
    for count in gottesdienste[MESSDIENERZAHL]:
        splittings.append(split_in_summanden(count))
    gottesdienste[GRUPPEN_SPLITTINGS] = splittings
    return gottesdienste


def messdiener_column_to_string(gottesdienste: pd.DataFrame) -> pd.DataFrame:
    new_col = []
    for i, gottesdienst in gottesdienste.iterrows():
        messdiener = gottesdienst[MESSDIENER]
        names = messdiener[0][1]
        for j in np.arange(1, len(messdiener)):
            names += f', {messdiener[j][1]}'
        new_col.append(names)
    gottesdienste[MESSDIENER] = new_col
    return gottesdienste


def zuteilung(gottesdienste: pd.DataFrame, messdiener: pd.DataFrame):
    # set new column
    gottesdienste[MESSDIENER] = [[]] * len(gottesdienste.index)
    # short handle for later
    gd_index = gottesdienste.index
    # current best option
    best_gottesdienste = None
    best_messdiener = None
    best_max_einteilungen = sys.maxsize
    # iterate all possible group combinations
    for combi in list(itertools.product(*gottesdienste[GRUPPEN_SPLITTINGS].values)):
        # try-catch to continue when we see that there's a problem in this combi
        try:
            # working copies
            tmp_gottesdienste = gottesdienste.copy()
            tmp_messdiener = messdiener.copy()
            # setting one gottesdienst after another
            for i, dienst in enumerate(combi):
                # setting one group after another
                for group in dienst:
                    # query possible messdiener for the group
                    ids_in_use = [*map(lambda item: item[0], tmp_gottesdienste[MESSDIENER].values[i])]
                    take3 = tmp_messdiener[tmp_messdiener[ANZAHL] == group]
                    take2 = take3[~take3[BLACK_LIST_TAGE].str.contains(tmp_gottesdienste[TAG].values[i])]
                    take1 = take2[~take2[ID].isin(ids_in_use)]
                    take0 = take1[take1[EINTEILUNGEN] == take1[EINTEILUNGEN].min()]
                    # if one found take first, else continue with next combi, 'cause this combi doesn't work
                    if len(take0.values) > 0:
                        take = take0.values[0]
                        # add messdiener to gottesdienst and increment its einteilungen
                        tmp_gottesdienste.at[gd_index[i], MESSDIENER] = (tmp_gottesdienste[MESSDIENER].values[
                                                                             i] or []) + [take]
                        tmp_messdiener.at[tmp_messdiener[tmp_messdiener[ID] == take[0]].index[0], EINTEILUNGEN] += 1
                    else:
                        raise Continue()
            # if highest einteilungszahl is smaller than on best option replace
            tmp_max_einteilungen = \
                tmp_messdiener[tmp_messdiener[EINTEILUNGEN] == tmp_messdiener[EINTEILUNGEN].max()].values[0][3]
            if best_max_einteilungen > tmp_max_einteilungen:
                best_gottesdienste = tmp_gottesdienste.copy()
                best_messdiener = tmp_messdiener.copy()
                best_max_einteilungen = tmp_max_einteilungen
        except Continue:
            continue
    return best_gottesdienste, best_messdiener


def remove_not_wanted_columns(gottesdienste: pd.DataFrame) -> pd.DataFrame:
    if MESSDIENERZAHL in gottesdienste:
        del gottesdienste[MESSDIENERZAHL]
    if GRUPPEN_SPLITTINGS in gottesdienste:
        del gottesdienste[GRUPPEN_SPLITTINGS]
    return gottesdienste


def reset_einteilungen(messdiener: pd.DataFrame) -> pd.DataFrame:
    new_einteilungen = []
    min = messdiener[messdiener[EINTEILUNGEN] == messdiener[EINTEILUNGEN].min()].values[0][3]
    for einteilung in messdiener[EINTEILUNGEN]:
        new_einteilungen.append(einteilung - min)
    messdiener[EINTEILUNGEN] = new_einteilungen
    return messdiener


def build_messdienerplanung(gottesdienst_arten: dict, gottesdienste: pd.DataFrame, messdiener: pd.DataFrame) -> (
        pd.DataFrame, pd.DataFrame):
    gottesdienste = add_messdiener_count(gottesdienste, gottesdienst_arten)
    print(
        "Jetzt wird eine optimale Zuteilung gebildet, das kann einige Zeit dauern. Während dieser Zeit sollte die CPU-Auslastung hoch sein")
    gottesdienste = group_splittings(gottesdienste)
    gottesdienste, messdiener = zuteilung(gottesdienste, messdiener)
    gottesdienste = messdiener_column_to_string(gottesdienste)
    return gottesdienste, messdiener


def prepare_files():
    for file_name in [MESSPLAN_INPUT, MESSDIENER_INPUT, GOTTESDIENST_ARTEN_INPUT]:
        if not exists(file_name):
            print(
                f'Die Datei {file_name} existiert nicht. Falls du den Namen der Datei dauerhaft ändern möchtest, kannst du das in der Datei constants.py machen')
            exit_programm()
    if not exists('output'):
        mkdir("output")
    shutil.copy2(MESSDIENER_INPUT, MESSDIENER_OUTPUT)
    print(
        f'Die Datei {MESSDIENER_INPUT} wird während der Ausführung aktualisiere (Einteilungen), falls etwas schief läuft wurde die ursprungsdatei nach {MESSDIENER_OUTPUT} kopiert')
    if not exists('tmp'):
        mkdir("tmp")


def print_help() -> None:
    print('Following arguments are available:')
    print('-c or --create-new-db    creates a new empty database')
    print('-r or --run              creates a new altar server schedule')
    print('-l or --load-worships    loads new worships from a docx file and saves them in the db')


def create_db() -> None:
    if exists('altar_server.db'):
        tmp_input: str = input(
            "There's already a database. Are you sure you want to override it? \
After that you won't be able to get any data back. [y/N] ").lower() or 'n'
        if tmp_input not in ['y', 'n']:
            print("Invalid input.")
            exit_programm()
        elif tmp_input == 'n':
            exit_programm()
        else:
            os.remove('altar_server.db')
    con: Connection = sqlite3.connect('altar_server.db')
    cur: Cursor = con.cursor()
    cur.execute('''CREATE TABLE time_period(
                    id                  INTEGER NOT NULL    PRIMARY KEY     AUTOINCREMENT,
                    start               TEXT    NOT NULL,
                    end                 TEXT    NOT NULL    CHECK(start < end)
                    );''')
    cur.execute('''CREATE TABLE altar_server_group(
                    id                  INTEGER NOT NULL    PRIMARY KEY     AUTOINCREMENT,
                    scheduled_count     INT,
                    high_mass           INT     NOT NULL    CHECK(-1 <= high_mass AND high_mass <= 2),
                        -- 2=only high 1=high, 0=normal, -1=never high
                    holiday             INTEGER,
                    black_list_days     INT     NOT NULL   CHECK(black_list_days >=0 AND black_list_days < 128),
                        --0b0010010=Wed. and Sat. are black listed
                    FOREIGN KEY(holiday)                    REFERENCES time_period(id)
                    );''')
    cur.execute('''CREATE TABLE group_has_black_list_group(
                     id                 INTEGER NOT NULL    PRIMARY KEY     AUTOINCREMENT,
                     first_group        INTEGER NOT NULL,
                     second_group       INTEGER NOT NULL    CHECK(first_group != second_group),
                     UNIQUE(first_group, second_group),
                     FOREIGN KEY(first_group)               REFERENCES altar_server_group(id),
                     FOREIGN KEY(second_group)              REFERENCES altar_server_group(id)
                    );''')
    cur.execute('''CREATE TABLE altar_server(
                    id                  INTEGER NOT NULL    PRIMARY KEY     AUTOINCREMENT,
                    first_name          TEXT    NOT NULL,
                    last_name           TEXT    NOT NULL,
                    gender              TEXT    NOT NULL,
                    role                INT     NOT NULL    CHECK(0 <= role AND role <= 1), -- 0=normal 1=leader
                    birthdate           TEXT    NOT NULL,
                    year_of_acceptance  INT     NOT NULL    CHECK(year_of_acceptance <= STRFTIME('%Y', 'now')+0
                                                                AND year_of_acceptance > 1900),
                    phone               TEXT    NOT NULL,
                    address             INTEGER NOT NULL,
                    server_group        INTEGER,
                    FOREIGN KEY(address)                    REFERENCES address(id),
                    FOREIGN KEY(server_group)               REFERENCES altar_server_group(id)
                    );''')
    cur.execute('''CREATE TABLE mail(
                    id                  INTEGER NOT NULL    PRIMARY KEY     AUTOINCREMENT,
                    mail                TEXT    NOT NULL    UNIQUE    CHECK(mail LIKE '_%@_%._%') 
                    );''')
    cur.execute('''CREATE TABLE server_has_mail(
                    id                  INTEGER NOT NULL    PRIMARY KEY     AUTOINCREMENT,
                    server              INTEGER NOT NULL,
                    mail                TEXT    NOT NULL,
                    UNIQUE(server, mail),
                    FOREIGN KEY(server)                     REFERENCES altar_server(id),
                    FOREIGN KEY(mail)                       REFERENCES mail(id)
                    );''')
    cur.execute('''CREATE TABLE address(
                    id                  INTEGER NOT NULL    PRIMARY KEY,
                    street              TEXT    NOT NULL,
                    street_number       INT     NOT NULL    CHECK(street_number > 0),
                    additional          TEXT,
                    post_code           TEXT,
                    country             TEXT
                    -- osm id / coordinates …
                    );''')
    cur.execute('''CREATE TABLE worship_typ(
                    id                  INTEGER NOT NULL    PRIMARY KEY AUTOINCREMENT ,
                    name                TEXT    NOT NULL    UNIQUE,
                    needed_server       INT     NOT NULL,
                    high_mass           INT     NOT NULL     CHECK(0 <= high_mass AND high_mass <= 1)-- 1=high, 0=normal
                    );''')
    cur.execute('''CREATE TABLE worship(
                    id                  INTEGER NOT NULL    PRIMARY KEY     AUTOINCREMENT,
                    start               TEXT    NOT NULL,
                    place               TEXT,
                    leader              TEXT,
                    contributors        TEXT,
                    info                TEXT,
                    type                INTEGER NOT NULL,
                    FOREIGN KEY(type)                       REFERENCES worship_typ(id)                
                    );''')
    cur.execute('''CREATE TABLE server_group_has_worship(
                    id                  INTEGER NOT NULL    PRIMARY KEY     AUTOINCREMENT,
                    server_group        INTEGER NOT NULL,
                    worship             INTEGER NOT NULL,
                    UNIQUE(server_group, worship),
                    FOREIGN KEY(server_group)               REFERENCES altar_server_group(id),
                    FOREIGN KEY(worship)                    REFERENCES worship(id)
                    );''')
    con.commit()
    con.close()


def load_worships() -> None:
    # read file
    path: str = input(f'Path to import docx: [{MESSDIENER_INPUT}] ') or MESSPLAN_INPUT
    if not exists(MESSPLAN_INPUT):
        print(f'{path} does not exist')
        exit_programm()
    md_table: str = docx_table_to_html_md_tabel(path, None)
    worship_table: DataFrame = pd.read_html(md_table)[0]

    con = sqlite3.connect('altar_server.db')  # make global
    cur = con.cursor()

    for row in worship_table:  # ab 2  zeile
        cur.execute('''INSERT INTO worship VALUES (?, ?)''', row[0], row[1])  # richtig machen

    con.commit()
    con.close()


def create_schedule() -> None:
    pass


if __name__ == '__main__':
    # warnings.filterwarnings('ignore')
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 400)
    logging.basicConfig(  # filename='file.log', level=logging.ERROR
        format='%(asctime)s -- %(levelname)s: [%(filename)s:%(lineno)d] %(message)s')

    try:
        opts, args = getopt.getopt(sys.argv[1:], 'hcrl', ['help', 'create-new-db', 'run', 'load-worships'])
    except getopt.GetoptError:
        print('argument error. For more information run with -h or --help')
        sys.exit(2)
    for opt, arg in opts:
        if opt in ['-h', '--help']:
            exit_programm()
        elif opt in ['-c', '--create-new-db']:
            create_db()
        elif opt in ['-r', '--run']:
            create_schedule()
        elif opt in ['-l', '--load-worships']:
            load_worships()
    exit_programm()

    prepare_files()
    docx_table_to_html_md_tabel(MESSPLAN_INPUT, MESSPLAN_TMP)
    gottesdienste = get_gottesdienstplan_from_html(MESSPLAN_TMP)
    messdiener = get_messdiener_from_csv(MESSDIENER_INPUT)
    gottesdienst_arten = get_gottesdienst_arten_from_json(GOTTESDIENST_ARTEN_INPUT)
    gottesdienste, messdiener = build_messdienerplanung(gottesdienst_arten, gottesdienste, messdiener)
    print("Die Zuteilung wurde erfolgreich abgeschlossen")
    gottesdienste = remove_not_wanted_columns(gottesdienste)
    messdiener = reset_einteilungen(messdiener)
    export_table_to_excel(gottesdienste, MESSPLAN_OUTPUT)
    print(f"Der Messdienerplan wurde nach {MESSPLAN_OUTPUT} exportiert")
    export_table_to_csv(messdiener, MESSDIENER_INPUT)
    print(f'Die Einteilungen in {MESSDIENER_INPUT} wurden aktualisiert')
    try:
        shutil.rmtree("tmp")
    except:
        print(
            'Während der Ordner tmp gelöscht werden sollte ist ein Fehler aufgetreten. Falls der Ordner noch existiert, kannst du es manuell versuchen, es ist aber auch nicht schlimm, wenn du nichts tust')
    else:
        print("Temporäre Dateien wurden gelöscht")

    exit_programm()
