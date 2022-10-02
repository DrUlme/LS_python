import os, sys, sqlite3

# globale Parameter
import LSglobal

#========================================================================
# Existenz feststellen
if os.path.exists( LSglobal.SQLiteFile ):
    print("Datei bereits vorhanden")
    sys.exit(0)

# Verbindung zur Datenbank erzeugen
connection = sqlite3.connect( LSglobal.SQLiteFile )

# Datensatzcursor erzeugen
cursor = connection.cursor()

# Tabellen erzeugen
sql = "CREATE TABLE rennen(" \
      "nummer INTEGER PRIMARY KEY, " \
      "name TEXT, " \
      "gender TEXT, " \
      "boot TEXT, " \
      "strecke TEXT, " \
      "startzeit TEXT, " \
      "status INTEGER, " \
      "gewicht REAL, " \
      "jahrgangmin INTEGER, " \
      "jahrgangmax INTEGER)"
cursor.execute(sql)

sql = "CREATE TABLE ruderer(" \
      "nummer INTEGER PRIMARY KEY, " \
      "vorname TEXT, " \
      "name TEXT, " \
      "geschlecht TEXT, " \
      "jahrgang INTEGER, " \
      "leichtgewicht INTEGER, " \
      "gewicht REAL, " \
      "verein TEXT, " \
      "kader TEXT, " \
      "form INTEGER)"
cursor.execute(sql)

sql = "CREATE TABLE r2boot(" \
      "nummer INTEGER PRIMARY KEY, " \
      "bootNr INTEGER, " \
      "rudererNr INTEGER, " \
      "platz INTEGER)"
cursor.execute(sql)
 
sql = "CREATE TABLE boote(" \
      "nummer INTEGER PRIMARY KEY, " \
      "startnummer INTEGER, " \
      "rennen INTEGER, " \
      "planstart INTEGER, " \
      "secstart INTEGER, " \
      "sec3000 INTEGER, " \
      "sec6000 INTEGER, " \
      "zeit3000 INTEGER, " \
      "zeit6000 INTEGER, " \
      "zeit INTEGER, " \
      "abgemeldet INTEGER, " \
      "kommentar TEXT )"
cursor.execute(sql)

sql = "CREATE TABLE verein(" \
      "name TEXT, " \
      "kurz TEXT, " \
      "anschrift TEXT, " \
      "anschrift2 TEXT, " \
      "rechnung REAL, " \
      "dabei INTEGER, " \
      "bayrisch INTEGER )"
cursor.execute(sql)

sql = "CREATE TABLE betreuer(" \
      "vorname TEXT, " \
      "name TEXT, " \
      "verein TEXT, " \
      "telefon TEXT, " \
      "email TEXT)"
cursor.execute(sql)


connection.commit()

# Verbindung beenden
connection.close()
