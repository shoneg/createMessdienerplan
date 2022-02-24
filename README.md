# createMessdienerplan

_createMessdienerplan_ ist ein Python-Skript, das automatisiert aus einer Liste von Messdienern und einem Messplan in einem Word-Dokument einen Messdienerplan erstellt. Dabei werden bestimmte Angaben
zu den Messdiener:innen beachtet und die Anzahl der Einteilungen nachgehalten, sodass auf Dauer alle Dienenden ähnlich oft eingeteilt werden.

## Vorbereitungen

Vor der ersten Ausführung müssen die in Abschnitt [Eingabedateien](#Eingabedateien) angegebenen Dateien erzeugt werden. Außerdem werden folgende Python-Module benötigt:

- docx2md
- itertools
- json
- math
- numpy
- shutil
- os
- pandas
- regex
- sys
- warnings
- xlsxWriter

## Eingabedateien

Als Eingabedateien werden drei Dateien benötigt. Standardmäßig werden diese aus dem Verzeichnis `input` gelesen. Beispiele für die Datein und deren genaue Namen gibt es im Ordner `example_input`.
Wichtig ist, dass die Tabellen genau die angegebene Form haben.

Hier kurze Erläuterungen zu den Dateien:

### Gottesdienstarten

In der Datei [input/gottesdienst-arten.json](example_input/gottesdienst-arten.json) sollten als Keys die Gottesdienstarten aufgeführt sein, die regelmäßig abgehalten werden. Als Wert wird jeweils die
Anzahl der Messdiener:innen angegeben, die im Normalfall für solch eine Messe eingeteilt werden. Die Gottesdienstarten sollten denen aus der Spalte `Gottesdienste` im [Messplan](#Messpan) entsprechen.

### Messdiener

In der Datei [input/messdiener.csv](example_input/messdiener.csv) werden alle Messdiener:innen, die eingeteilt werden können aufgeführt. In einer Zeile können auch mehrere Messdiener:innen eingetragen
werden, die dann immer zusammen eingeteilt werden.

#### ID

Jede Zeile muss eine eindeutige ID zugewiesen bekommen, die IDs können von oben nach unten durchnummeriert werden und können neu vergeben werden, wenn eine alte Zeile, dessen Messdiener:innen nicht
mehr eingeteilt werden sollen, entfernt wurde.

#### Namen

Diese Namen tauchen später im Messdienerplan auf. Bei Gruppen, die zusammen dienen möchten, sollten die Namen durch Kommata getrennt werden, damit die Auflistung im Messdienerplan konsistent ist.

#### Anzahl

Die Anzahl der Namen, die in der Spalte [Namen](#Namen) eingetragen ist.

#### Einteilungen

Diese Zahl gibt an wie oft mehr diese Zeile im Vergleich zu der am seltensten eingeteilt Zeile mehr eingeteilt wurde. Diese Spalte wird bei der Ausführung des Skripts automatisch aktualisiert.

Im Normalfall sollte diese Spalte beim eintragen neuer Messdiener:innen auf `0` gesetzt werden und danach nicht mehr manuell verändert werden.

#### Black-List-Personen

Die Funktion ist noch nicht implementiert. Die Spalte kann leer gelassen werden und hat aktuell keine Auswirkungen.

#### Black-List-Tage

Falls die Personen in dieser Zeile an einem oder mehreren Wochentagen nie eingeteilten werden sollen, kann in dieser Spalte ein String mit den ausgeschriebenen Wochentagen, getrennt durch Leerzeichen,
angegeben werden.

### Messplan

In der Datei [input/messplan.docx](example_input/messplan.docx) muss eine Tabelle angegeben sein, in der `Datum`, `Zeit`, `Ort` und `Gottsdienst` angegeben sind.

Die hier angegebene Datei ist ein beispielhafter Dienstplan. In der Regel gibt es bereits eine Standardformat für Dienstpläne in Kirchengemeinden, die ohne gößeren Aufwand zur Verfügung gestellt
werden können. Ggf. passt das Format dann nicht mehr ganz auf das hier angegebene Schema und es müssen lokal einige Änderungen am Quellcode vorgenommen werden. Gibt es ausschließlich Unterschiede in
der Spaltenbennenung, können die eingelesenen Namen in der Datei [constants.py](constants.py) einfach angepasst werden.

## Verarbeitung

Bei Ausführung des Skripts wird zuerst abgefragt für welche Messe aus dem Messplan wie viele Personen dienen sollen. Dabei sind die Standardwerte die in den [Gottesdienstarten](#Gottensdiensarten)
angegebenen. Danach wird aus den angegebenen Daten ein Messdienerplan erstellt, der eine optimale Einteilung darstellt. Dieser Prozess kann je nach Eingabe längere Zeit in Anspruch nehmen, weil evtl.
alle möglichen Kombinationen ausprobiert werden. Anschließend wird der erstellte Messdienerplan in eine Excel-Datei im Ordner `output` exportiert.

## Changelog

### Geplante Updates

- **Black-List-Personen** sollen in der messdiener.csv eingetragen werden können, sodass dann bestimmte Personen nicht zusammen eingeteilt werdne.
- **Black-List-Daten** sollen in der messdiener.csv eingetragen werder können, sodass z. B. wenn ein Messdierer im Urlaub ist, dieser in der angegebenen Zeit nicht für Messen eingeteilt wird.
- **Entfernen der Anzahl-Spalte**, sodass zukünftig durch die Anzahl der Kommata in der Namen-Spalte die Anzahl ermittelt wird.

### Letzte Änderungen

- Upload des Projekts auf GitHub