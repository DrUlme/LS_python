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
      "id TEXT, " \
      "name TEXT, " \
      "code TEXT, " \
      "sex TEXT, " \
      "boot TEXT, " \
      "strecke TEXT, " \
      "startzeit INTEGER, " \
      "status INTEGER, " \
      "gewicht REAL, " \
      "jahrgangmin INTEGER, " \
      "jahrgangmax INTEGER, " \
      "cost REAL)"
cursor.execute(sql)

sql = "CREATE TABLE ruderer(" \
      "id TEXT, " \
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
      "meldeid TEXT, " \
      "bootid TEXT, " \
      "rudererid TEXT, " \
      "platz INTEGER)"
cursor.execute(sql)
 
sql = "CREATE TABLE boote(" \
      "id TEXT, " \
      "revision TEXT, " \
      "startnummer INTEGER, " \
      "rennen INTEGER, " \
      "planstart TEXT, " \
      "tStart TEXT, " \
      "t3000 TEXT, " \
      "t6000 TEXT, " \
      "zeit3000 TEXT, " \
      "zeit6000 TEXT, " \
      "zeit TEXT, " \
      "abgemeldet INTEGER, " \
      "alternativ INTEGER, "\
      "kommentar TEXT )"
cursor.execute(sql)

sql = "CREATE TABLE verein(" \
      "id TEXT, " \
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

sql = "CREATE TABLE meta(" \
      "name TEXT, " \
      "wert TEXT)"
cursor.execute(sql)

connection.commit()

# Zeitdeltas für die Meßpunkte
sql = "INSERT INTO meta VALUES( 'sec_Start', '0' )"
cursor.execute(sql)
sql = "INSERT INTO meta VALUES( 'sec_3000m', '0' )"
cursor.execute(sql)
sql = "INSERT INTO meta VALUES( 'sec_6000m', '0' )"
cursor.execute(sql)

connection.commit()

# Verbindung beenden
connection.close()
