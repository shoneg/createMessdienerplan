import platform

ANZAHL = 'Anzahl'
BLACK_LIST_PERSONEN = 'Black-List-Personen'
BLACK_LIST_TAGE = 'Black-List-Tage'
DATUM = 'Datum'
EINTEILUNGEN = 'Einteilungen'
GOTTESDIENST = 'Gottesdienst'
GRUPPEN_SPLITTINGS = "Gruppen-Splittings"
ID = 'ID'
MESSDIENER = 'Messdiener'
MESSDIENERZAHL = 'Messdienerzahl'
NAMEN = 'Namen'
TAG = 'Tag'
ZEIT = 'Zeit'

GOTTESDIENST_ARTEN_INPUT = 'input/gottesdienst-arten.json'
MESSDIENER_INPUT = 'input/messdiener.csv'
MESSPLAN_INPUT = 'input/messplan.docx'

MESSPLAN_TMP = 'tmp/gottesdienste.html'

MESSDIENER_OUTPUT = 'output/messdiener.csv'
MESSPLAN_OUTPUT = 'output/messdienerplan.xlsx'

ENCODING = 'ansi' if platform.system() == 'Windows' else 'utf-8'
