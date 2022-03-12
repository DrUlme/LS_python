# define_Rennen

import os, sys, sqlite3

# lade globale Parameter
import LSglobal

#========================================================================

# Existenz feststellen
if os.path.exists( LSglobal.SQLiteFile ):
   print("Datei bereits vorhanden")
else:
   sys.exit(0)

# Verbindung zur Datenbank erzeugen
connection = sqlite3.connect( LSglobal.SQLiteFile )


# Datensatzcursor erzeugen
cursor = connection.cursor()

# Definition des Jahrgangs und der Art (e.g. "Herbst")
RefJahr = LSglobal.RefJahr
TYPE    = LSglobal.Zeit

# sql = "CREATE TABLE rennen(" \
      # "nummer INTEGER PRIMARY KEY, " \
      # "name TEXT, " \
      # "gender TEXT, " \
      # "boot TEXT, " \
      # "strecke TEXT, " \
      # "startzeit TEXT, " \
      # "status INTEGER, " \
      # "gewicht REAL, " \
      # "jahrgangmax INTEGER, " \
      # "jahrgangmin INTEGER)"
      
#  1 Grossboot
#  2 JM 1x B
#  3 JM 1x B LG
#  4 SM 2- A/B
#  5 JM 2- A
#  6 JM 2- B
#  7 SF 2- B/A
#  8 JF 2- B/A
#  9 JM 1x A
# 10 JM 1x A LG
# 11 JF 1x A
# 12 JF 1x A LG
# 13 SM 1x A
# 14 SM 1x B
# 15 SF 1x A/B
# 16 JF 1x B
# 17 JF 1x B LG
Rennen = 0
Rennen = Rennen + 1
sql = "INSERT INTO rennen VALUES(" \
   + str(Rennen) + ", 'Großboot', 'all', 'all', '6000 m', '10:00', " \
   "0, -1.0, 1900, " + str(RefJahr - 10) + " )"
cursor.execute(sql)
connection.commit()

Rennen = Rennen + 1
sql = "INSERT INTO rennen VALUES(" \
   + str(Rennen) + ", 'Frühstarter', 'all', 'all', '6000 m', '10:00'," \
   " 0, -1.0, 1900, " + str(RefJahr - 15) + " )"
cursor.execute(sql)
connection.commit()

Rennen = Rennen + 1
sql = "INSERT INTO rennen VALUES(" \
   + str(Rennen) + ", 'JM 1x B', 'M', '1x', '6000 m', '10:00'," \
   " 0, -1.0, " + str(RefJahr - 16) + ", " + str(RefJahr - 15) + " )"
cursor.execute(sql)
connection.commit()

Rennen = Rennen + 1
sql = "INSERT INTO rennen VALUES(" \
   + str(Rennen) + ", 'JM 1x B LG', 'M', '1x', '6000 m', '10:00'," \
   " 0, 65.0, " + str(RefJahr - 16) + ", " + str(RefJahr - 15) + " )"
cursor.execute(sql)
connection.commit()

Rennen = Rennen + 1
sql = "INSERT INTO rennen VALUES(" \
   + str(Rennen) + ", 'SM 2- A/B', 'M', '2-', '6000 m', '10:00'," \
   " 0, -1.0, " + str(1900) + ", " + str(RefJahr - 19) + " )"
cursor.execute(sql)
connection.commit()

Rennen = Rennen + 1
sql = "INSERT INTO rennen VALUES(" \
   + str(Rennen) + ", 'JM 2- A', 'M', '2-', '6000 m', '10:00'," \
   " 0, -1.0, " + str(RefJahr - 18) + ", " + str(RefJahr - 17) + " )"
cursor.execute(sql)
connection.commit()

Rennen = Rennen + 1
sql = "INSERT INTO rennen VALUES(" \
   + str(Rennen) + ", 'JM 2- B', 'M', '2-', '6000 m', '10:00'," \
   " 0, -1.0, " + str(RefJahr - 16) + ", " + str(RefJahr - 15) + " )"
cursor.execute(sql)
connection.commit()

Rennen = Rennen + 1
sql = "INSERT INTO rennen VALUES(" \
   + str(Rennen) + ", 'SF 2- A/B', 'F', '2-', '6000 m', '10:00'," \
   " 0, -1.0, " + str(1900) + ", " + str(RefJahr - 19) + " )"
cursor.execute(sql)
connection.commit()

Rennen = Rennen + 1
sql = "INSERT INTO rennen VALUES(" \
   + str(Rennen) + ", 'JF 2- A/B', 'F', '2-', '6000 m', '10:00'," \
   " 0, -1.0, " + str(RefJahr - 18) + ", " + str(RefJahr - 15) + " )"
cursor.execute(sql)
connection.commit()

Rennen = Rennen + 1
sql = "INSERT INTO rennen VALUES(" \
   + str(Rennen) + ", 'JM 1x A', 'M', '1x', '6000 m', '10:00'," \
   " 0, -1.0, " + str(RefJahr - 18) + ", " + str(RefJahr - 17) + " )"
cursor.execute(sql)
connection.commit()

Rennen = Rennen + 1
sql = "INSERT INTO rennen VALUES(" \
   + str(Rennen) + ", 'JM 1x A LG', 'M', '1x', '6000 m', '10:00'," \
   " 0, 67.5, " + str(RefJahr - 18) + ", " + str(RefJahr - 17) + " )"
cursor.execute(sql)
connection.commit()

# ============================================

Rennen = Rennen + 1
sql = "INSERT INTO rennen VALUES(" \
   + str(Rennen) + ", 'JF 1x A', 'F', '1x', '6000 m', '10:00'," \
   " 0, -1.0, " + str(RefJahr - 18) + ", " + str(RefJahr - 17) + " )"
cursor.execute(sql)
connection.commit()

Rennen = Rennen + 1
sql = "INSERT INTO rennen VALUES(" \
   + str(Rennen) + ", 'JF 1x A LG', 'F', '1x', '6000 m', '10:00'," \
   " 0, 57.5, " + str(RefJahr - 18) + ", " + str(RefJahr - 17) + " )"
cursor.execute(sql)
connection.commit()

# ============================================

Rennen = Rennen + 1
sql = "INSERT INTO rennen VALUES(" \
   + str(Rennen) + ", 'SM 1x A', 'M', '1x', '6000 m', '10:00'," \
   " 0, -1.0, " + str(1900) + ", " + str(RefJahr - 23) + " )"
cursor.execute(sql)
connection.commit()

Rennen = Rennen + 1
sql = "INSERT INTO rennen VALUES(" \
   + str(Rennen) + ", 'SM 1x B', 'M', '1x', '6000 m', '10:00'," \
   " 0, -1.0, " + str(RefJahr - 22) + ", " + str(RefJahr - 19) + " )"
cursor.execute(sql)
connection.commit()

Rennen = Rennen + 1
sql = "INSERT INTO rennen VALUES(" \
   + str(Rennen) + ", 'SF 1x A/B', 'F', '1x', '6000 m', '10:00'," \
   " 0, -1.0, " + str(1900) + ", " + str(RefJahr - 19) + " )"
cursor.execute(sql)
connection.commit()

Rennen = Rennen + 1
sql = "INSERT INTO rennen VALUES(" \
   + str(Rennen) + ", 'JF 1x B', 'F', '1x', '6000 m', '10:00'," \
   " 0, -1.0, " + str(RefJahr - 16) + ", " + str(RefJahr - 15) + " )"
cursor.execute(sql)
connection.commit()

Rennen = Rennen + 1
sql = "INSERT INTO rennen VALUES(" \
   + str(Rennen) + ", 'JF 1x B LG', 'F', '1x', '6000 m', '10:00'," \
   " 0, 55.0, " + str(RefJahr - 16) + ", " + str(RefJahr - 15) + " )"
cursor.execute(sql)
connection.commit()


Rennen = Rennen + 1
sql = "INSERT INTO rennen VALUES(" \
   + str(Rennen) + ", 'Spätstarter', 'all', 'all', '6000 m', '10:00'," \
   " 0, -1.0, 1900, " + str(RefJahr - 15) + " )"
cursor.execute(sql)
connection.commit()

# Rennen = Rennen + 1
# sql = "INSERT INTO rennen VALUES(" \
   # + str(Rennen) + ", 'Wettkampftest', 'all', 'all', 'Ergometer', '10:00'," \
   # " 0, -1.0, 1900, " + str(RefJahr - 15) + " )"
# cursor.execute(sql)
# connection.commit()

# 11 JF 1x A
# 12 JF 1x A LG
# 13 SM 1x A
# 14 SM 1x B
# 15 SF 1x A/B
# 16 JF 1x B
# 17 JF 1x B LG

if(TYPE == "Herbst"):
   # Jung 2x 13 u. 14 Jahre
   Rennen = 21
   sql = "INSERT INTO rennen VALUES(" \
      + str(Rennen) + ", 'Jungen 2x 13 u. 14 Jahre', 'M', '2x', '3000 m', '10:00'," \
      " 0, -1.0, " + str(RefJahr - 14) + ", " + str(RefJahr - 13) + " )"
   cursor.execute(sql)
   connection.commit()
   # Jung 2x LG 13 u. 14 Jahre
   Rennen = Rennen + 1
   sql = "INSERT INTO rennen VALUES(" \
      "22, 'Jungen 2x LG 13 u. 14 Jahre', 'M', '2x', '3000 m', '10:00'," \
      " 0, 55.0, " + str(RefJahr - 14) + ", " + str(RefJahr - 13) + " )"
   cursor.execute(sql)
   connection.commit()
   # Jung 2x 12 u. 13 Jahre
   Rennen = Rennen + 1
   sql = "INSERT INTO rennen VALUES(" \
      "23, 'Jungen 2x 12 u. 13 Jahre', 'M', '2x', '3000 m', '10:00'," \
      " 0, -1.0, " + str(RefJahr - 13) + ", " + str(RefJahr - 12) + " )"
   cursor.execute(sql)
   connection.commit()
   # Jung 2x LG 12 u. 13 Jahre
   Rennen = Rennen + 1
   sql = "INSERT INTO rennen VALUES(" \
      "24, 'Jungen 2x LG 12 u. 13 Jahre', 'M', '2x', '3000 m', '10:00'," \
      " 0, 50.0, " + str(RefJahr - 13) + ", " + str(RefJahr - 12) + " )"
   cursor.execute(sql)
   connection.commit()
   # Mäd 2x 13 u. 14 Jahre
   Rennen = Rennen + 1
   sql = "INSERT INTO rennen VALUES(" \
      "25, 'Mädchen 2x 13 u. 14 Jahre', 'F', '2x', '3000 m', '10:00'," \
      " 0, -1.0, " + str(RefJahr - 14) + ", " + str(RefJahr - 13) + " )"
   cursor.execute(sql)
   connection.commit()
   # Mäd 2x LG 13 u. 14 Jahre
   Rennen = Rennen + 1
   sql = "INSERT INTO rennen VALUES(" \
      "26, 'Mädchen 2x LG 13 u. 14 Jahre', 'F', '2x', '3000 m', '10:00'," \
      " 0, 52.5, " + str(RefJahr - 14) + ", " + str(RefJahr - 13) + " )"
   cursor.execute(sql)
   connection.commit()
   # Mäd 2x 12 u. 13 Jahre
   Rennen = Rennen + 1
   sql = "INSERT INTO rennen VALUES(" \
      "27, 'Mädchen 2x 12 u. 13 Jahre', 'F', '2x', '3000 m', '10:00'," \
      " 0, -1.0, " + str(RefJahr - 13) + ", " + str(RefJahr - 12) + " )"
   cursor.execute(sql)
   connection.commit()
   # Mäd 2x LG 12 u. 13 Jahre
   Rennen = Rennen + 1
   sql = "INSERT INTO rennen VALUES(" \
      "28, 'Mädchen 2x LG 12 u. 13 Jahre', 'F', '2x', '3000 m', '10:00'," \
      " 0, 50.0, " + str(RefJahr - 13) + ", " + str(RefJahr - 12) + " )"
   cursor.execute(sql)
   connection.commit()
   # Jung 1x 14 Jahre
   Rennen = Rennen + 1
   sql = "INSERT INTO rennen VALUES(" \
      "29, 'Jungen 1x 14 Jahre', 'M', '1x', '3000 m', '10:00'," \
      " 0, -1.0, " + str(RefJahr - 14) + ", " + str(RefJahr - 14) + " )"
   cursor.execute(sql)
   connection.commit()
   # Jung 1x LG 14 Jahre
   Rennen = Rennen + 1
   sql = "INSERT INTO rennen VALUES(" \
      "30, 'Jungen 1x LG 14 Jahre', 'M', '1x', '3000 m', '10:00'," \
      " 0, 55.0, " + str(RefJahr - 14) + ", " + str(RefJahr - 14) + " )"
   cursor.execute(sql)
   connection.commit()
   # Jung 1x 13 Jahre
   Rennen = Rennen + 1
   sql = "INSERT INTO rennen VALUES(" \
      "31, 'Jungen 1x 13 Jahre', 'M', '1x', '3000 m', '10:00'," \
      " 0, -1.0, " + str(RefJahr - 13) + ", " + str(RefJahr - 13) + " )"
   cursor.execute(sql)
   connection.commit()
   # Jung 1x LG 13 Jahre
   Rennen = Rennen + 1
   sql = "INSERT INTO rennen VALUES(" \
      "32, 'Jungen 1x LG 13 Jahre', 'M', '1x', '3000 m', '10:00'," \
      " 0, 50.0, " + str(RefJahr - 13) + ", " + str(RefJahr - 13) + " )"
   cursor.execute(sql)
   connection.commit()
   # Mäd 1x 14 Jahre
   Rennen = Rennen + 1
   sql = "INSERT INTO rennen VALUES(" \
      "33, 'Mädchen 1x 14 Jahre', 'F', '1x', '3000 m', '10:00'," \
      " 0, -1.0, " + str(RefJahr - 14) + ", " + str(RefJahr - 14) + " )"
   cursor.execute(sql)
   connection.commit()
   # Mäd 1x LG 14 Jahre
   Rennen = Rennen + 1
   sql = "INSERT INTO rennen VALUES(" \
      "34, 'Mädchen 1x LG 14 Jahre', 'F', '1x', '3000 m', '10:00'," \
      " 0, 52.5, " + str(RefJahr - 14) + ", " + str(RefJahr - 14) + " )"
   cursor.execute(sql)
   connection.commit()
   # Mäd 1x 13 Jahre
   sql = "INSERT INTO rennen VALUES(" \
      "35, 'Mädchen 1x 13 Jahre', 'F', '1x', '3000 m', '10:00'," \
      " 0, -1.0, " + str(RefJahr - 13) + ", " + str(RefJahr - 13) + " )"
   cursor.execute(sql)
   connection.commit()
   # Mäd 1x LG 13 Jahre
   Rennen = Rennen + 1
   sql = "INSERT INTO rennen VALUES(" \
      "36, 'Mädchen 1x 13 LG Jahre', 'F', '1x', '3000 m', '10:00'," \
      " 0, 50.0, " + str(RefJahr - 13) + ", " + str(RefJahr - 13) + " )"
   cursor.execute(sql)
   connection.commit()
else:
   # 21	Jungen	Athletiktest 	KM 1x
   Rennen = Rennen + 1
   sql = "INSERT INTO rennen VALUES(" \
      "21, 'Jungen Athletiktest', 'M', 'Athletik', 'div.', '10:00'," \
      " 0, -1.0, " + str(RefJahr - 9) + ", " + str(RefJahr - 14) + " )"
   cursor.execute(sql)
   connection.commit()
   # 22	Jungen LG	Athletiktest 	KM 1x Lgw.
   Rennen = Rennen + 1
   sql = "INSERT INTO rennen VALUES(" \
      "22, 'Jungen LG Athletiktest', 'M', 'Athletik', 'div.', '10:00'," \
      " 0, 55.0, " + str(RefJahr - 9) + ", " + str(RefJahr - 14) + " )"
   cursor.execute(sql)
   connection.commit()
   # 23	Mädchen	Athletiktest 	KW 1x
   Rennen = Rennen + 1
   sql = "INSERT INTO rennen VALUES(" \
      "23, 'Mädchen Athletiktest', 'F', 'Athletik', 'div.', '10:00'," \
      " 0, -1.0, " + str(RefJahr - 9) + ", " + str(RefJahr - 14) + " )"
   cursor.execute(sql)
   connection.commit()
   # 24	Mädchen LG	Athletiktest 	KW 1x Lgw.
   Rennen = Rennen + 1
   sql = "INSERT INTO rennen VALUES(" \
      "24, 'Mädchen LG Athletiktest', 'F', 'Athletik', 'div.', '10:00'," \
      " 0, 52.5, " + str(RefJahr - 9) + ", " + str(RefJahr - 14) + " )"
   cursor.execute(sql)
   connection.commit()
   # 25	Ausser Konkurrenz	Athletiktest 	Athletik
   # Rennen = Rennen + 1
   # sql = "INSERT INTO rennen VALUES(" \
   #    "25, 'Athletiktest ausser Konkurrenz', 'all', 'Athletik', 'div.', '10:00'," \
   #    " 0, -1.0, " + str(RefJahr - 9) + ", " + str(RefJahr - 15) + " )"
   # cursor.execute(sql)
   # connection.commit()


# Von Rennen zu Excel in neuem Programm !!
