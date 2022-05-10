import logging
import sqlite3
from sqlite3 import Connection, DatabaseError

con: Connection


def setup(path_to_db: str = 'altar_server.db') -> int:
    global con
    try:
        con = sqlite3.connect(path_to_db)
    except DatabaseError:
        logging.critical("Unexpected Error while setting up connection to database")
        return 1
    else:
        return 0


def teardown() -> None:
    global con
    con.close()


def create_tables(creation_script: str) -> None:
    global con
    con.executescript(creation_script)
    con.commit()
