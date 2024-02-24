# -*- coding: utf-8 -*-
# LS_python
Python Programm zur Durchführung des Erlanger Ruder-Langstrecken-Tests
basierend auf sqlite3, openpyxl und TriaZeit (ein einfachs ASCII-format)

Erstellt folgende Excel-Dateien:
* für die Meldung:
* für die Startplatzvergabe:
* Ergebnis für den Landestrainer

Erstellt HTML-Seiten für die Rennen im Verzeichnis HTML

Erstellt Melde-Ergebnis und Endergebnis mit LaTeX als pdf.

## Funktionen

**LSglobal.py** die globalen Daten für die Datenbank

### Erstellen der Datenbank
**LS_sqlite_erzeugen.py** erzeugt die Datenbank ohne Daten
**LS_define_Rennen.py**   definiert die Rennen
**LS_add_Vereine.py**     fügt bekannte Vereine hinzu

### Vorbereiten der Ausschreibung
**LS_write_Excel_Meldeformular.py** schreibt das Meldeformular
(*LS_write_MeldeDocx.py* Basis für das Word-Format - noch mit Fehlern: \
                         abgeschnittene Renn-Namen)

### Einlesen der Meldeergebnisse
**lese_Meldungen.py**    Einlesen der Excel Dateien aus *./Meldungen/*
**LS_update.py**         Setzen des Status für die Rennen
**LS_write_HTML.py**     Schreiben der HTML Seite
**LS_write_HTML_V.py**   Schreiben der HTML-Vereins-Seiten

**LS_update_Kader.py**   mit Einlesen aus Excel-File in Kader-Tabelle
**LS_write_Excel_Startreihenfolge.py**

### Vorbereitung der Regatta
lese_Startreihenfolge.py
LS_set_Startzeit.py  ??
LS_write_Zeitprotokolle_tex.py
 
LS_write_Meldeergebnis_tex.py

### Ummelden, Abmelden


### Regatta
* LS_read_TRZ.py
* LS_write_HTML.py
* Check_times.py
* LS_clear_times.py
**LS_update.py**         Setzen des Status für die Rennen


### Ende
* LS_write_Excel_Endergebnis.py
* LS_write_Endergebnis_tex.py
* LS_write_Rechnungen_tex.py
* LS_write_Excel_Endergebnis_BRV.py

### unbekannt
* Check_Startnummer.py

* zz_Test_docx_number.py

## LaTeX
Für die Ausarbeitung wird das Textverarbeitungssystem LaTeX (TeX) verwendet.
Mit diesem Befehl sollte es sofort klappen:
**pdflatex** *Meldungen.tex*


# ToDo
Bereich der auftretenden Änderungen

* _startzeit_ in _rennen_ nutzen: von int auf string: '11:20:45'
* Holen der Zeiten via FTP und auswerten.
* Hochladen der HTML Daten via FTP
* Leerzeichen bei Namen eliminieren
* Vornamen / Nachnamen tauschen, Namen in Datenbank einbauen und vergleichen ?
* FTP-client bei LS_read_TRZ.py
* Fehler-Meldung bei LS_read_TRZ.py, wenn Zeit fehlt (>15 min/3000)
* FTP-client bei LS_write_HTML.py
* HTML-Seite pro Verein => Rennen, Startnr, Name, (Waage), Startzeit, 3000 m, Endzeit
* Vorbereiten auf DRV-Maske

## Fehler und Ergänzungen
* Behandlung der Früh- und Spät-Starter und Grossboote ?!
* Kein oder falscher Jahrgang
* check Doppelte Startnummern => Abmeldungen vor Frist Startnummer entziehen
* AbmeldeZeitpunkt nach Frist notieren (Kommentar?) => "22. 11:22 (a),"
* Einlesen der Meldungen: Vereinskürzel vorhanden oder Telefon-Nummer !!!

## GUI für
* Namen => Vorname / Nachname vertauschen
* Leichtgewicht => mit Änderung von Rennen
* Kommentar bearbeiten
* Boots-Doppelnutzung => Früh-/Spätstarter
