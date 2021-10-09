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
**LS_define_Rennen.py** definiert die Rennen

### Vorbereiten der Ausschreibung
**LS_write_Excel_Meldeformular.py** schreibt das Meldeformular 
**LS_write_MeldeDocx.py** Basis für das Word-Format (noch mit Fehlern: abgeschnittene Renn-Namen)

### Einlesen der Meldeergebnisse 
**lese_Meldungen.py**
**check_Kader.py** ToDo mit Einlesen aus Excel-File in Kader-Tabelle
**LS_write_Excel_Startreihenfolge.py**

### Vorbereitung der Regatta
lese_Startreihenfolge.py
LS_set_Startzeit.py  ??
LS_write_Zeitprotokolle_tex.py
 
LS_write_Meldeergebnis_tex.py

### Ummelden, Abmelden


### Regatta
LS_read_TRZ.py
LS_write_HTML.py		   
Check_times.py
LS_clear_times.py	
	
### Ende
LS_write_Excel_Endergebnis.py
LS_write_Endergebnis_tex.py
LS_write_Rechnungen_tex.py
LS_write_Excel_Endergebnis_BRV.py

### unbekannt
Check_Startnummer.py  
		     			    	     	    	   
zz_Test_docx_number.py
