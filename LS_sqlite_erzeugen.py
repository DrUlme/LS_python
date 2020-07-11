import os, sys, sqlite3

# Existenz feststellen
if os.path.exists("LS2020H.db"):
    print("Datei bereits vorhanden")
    sys.exit(0)

# Verbindung zur Datenbank erzeugen
connection = sqlite3.connect("LS2020H.db")

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
      "vorname TEXT, " \
      "name TEXT, " \
      "geschlecht TEXT, " \
      "jahrgang INTEGER, " \
      "leichtgewicht INTEGER, " \
      "gewicht REAL, " \
      "verein TEXT, " \
      "nummer INTEGER PRIMARY KEY, " \
      "boot INTEGER)"
cursor.execute(sql)

sql = "CREATE TABLE boote(" \
      "nummer INTEGER PRIMARY KEY, " \
      "startnummer INTEGER, " \
      "rennen INTEGER, " \
      "vereine TEXT, " \
      "ruderer TEXT, " \
      "planstart INTEGER, " \
      "secstart INTEGER, " \
      "sec3000 INTEGER, " \
      "sec6000 INTEGER, " \
      "zeit3000 INTEGER, " \
      "zeit6000 INTEGER, " \
      "zeit INTEGER, " \
      "abgemeldet INTEGER)"
cursor.execute(sql)

sql = "CREATE TABLE verein(" \
      "name TEXT, " \
      "kurz TEXT, " \
      "anschrift TEXT, " \
      "anschrift2 TEXT, " \
      "rechnung REAL)"
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
