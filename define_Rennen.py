# define_Rennen

import os, sys, sqlite3

# Existenz feststellen
if os.path.exists("LS2020H.db"):
   print("Datei bereits vorhanden")
else:
   sys.exit(0)

# Verbindung zur Datenbank erzeugen
connection = sqlite3.connect("LS2020H.db")

# Datensatzcursor erzeugen
cursor = connection.cursor()

RefJahr = 2020

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

sql = "INSERT INTO rennen VALUES(" \
   "1, 'Frühstarter / Großboot', 'all', 'all', '6000 m', '10:00', " \
   "0, -1.0, 1900, " + str(RefJahr - 10) + " )"
cursor.execute(sql)
connection.commit()

sql = "INSERT INTO rennen VALUES(" \
   " 2, 'JM 1x B', 'M', '1x', '6000 m', '10:00'," \
   " 0, -1.0, " + str(RefJahr - 16) + ", " + str(RefJahr - 15) + " )"
cursor.execute(sql)
connection.commit()

sql = "INSERT INTO rennen VALUES(" \
   " 3, 'JM 1x B LG', 'M', '1x', '6000 m', '10:00'," \
   " 0, 65.0, " + str(RefJahr - 16) + ", " + str(RefJahr - 15) + " )"
cursor.execute(sql)
connection.commit()

sql = "INSERT INTO rennen VALUES(" \
   " 4, 'SM 2- A/B', 'M', '2-', '6000 m', '10:00'," \
   " 0, -1.0, " + str(1900) + ", " + str(RefJahr - 19) + " )"
cursor.execute(sql)
connection.commit()

sql = "INSERT INTO rennen VALUES(" \
   " 5, 'JM 2- A', 'M', '2-', '6000 m', '10:00'," \
   " 0, -1.0, " + str(RefJahr - 18) + ", " + str(RefJahr - 17) + " )"
cursor.execute(sql)
connection.commit()

sql = "INSERT INTO rennen VALUES(" \
   " 6, 'JM 2- B', 'M', '2-', '6000 m', '10:00'," \
   " 0, -1.0, " + str(RefJahr - 16) + ", " + str(RefJahr - 15) + " )"
cursor.execute(sql)
connection.commit()

sql = "INSERT INTO rennen VALUES(" \
   " 7, 'SF 2- A/B', 'F', '2-', '6000 m', '10:00'," \
   " 0, -1.0, " + str(1900) + ", " + str(RefJahr - 19) + " )"
cursor.execute(sql)
connection.commit()

sql = "INSERT INTO rennen VALUES(" \
   " 8, 'JF 2- A/B', 'F', '2-', '6000 m', '10:00'," \
   " 0, -1.0, " + str(RefJahr - 18) + ", " + str(RefJahr - 15) + " )"
cursor.execute(sql)
connection.commit()

sql = "INSERT INTO rennen VALUES(" \
   " 9, 'JM 1x A', 'M', '1x', '6000 m', '10:00'," \
   " 0, -1.0, " + str(RefJahr - 18) + ", " + str(RefJahr - 17) + " )"
cursor.execute(sql)
connection.commit()

sql = "INSERT INTO rennen VALUES(" \
   "10, 'JM 1x A LG', 'M', '1x', '6000 m', '10:00'," \
   " 0, 67.5, " + str(RefJahr - 18) + ", " + str(RefJahr - 17) + " )"
cursor.execute(sql)
connection.commit()

# ============================================

sql = "INSERT INTO rennen VALUES(" \
   "11, 'SF 1x A', 'F', '1x', '6000 m', '10:00'," \
   " 0, -1.0, " + str(RefJahr - 18) + ", " + str(RefJahr - 17) + " )"
cursor.execute(sql)
connection.commit()

sql = "INSERT INTO rennen VALUES(" \
   "12, 'JF 1x A LG', 'F', '1x', '6000 m', '10:00'," \
   " 0, 57.5, " + str(RefJahr - 18) + ", " + str(RefJahr - 17) + " )"
cursor.execute(sql)
connection.commit()

# ============================================

sql = "INSERT INTO rennen VALUES(" \
   "13, 'SM 1x A', 'M', '1x', '6000 m', '10:00'," \
   " 0, -1.0, " + str(1900) + ", " + str(RefJahr - 23) + " )"
cursor.execute(sql)
connection.commit()

sql = "INSERT INTO rennen VALUES(" \
   "14, 'SM 1x B', 'M', '1x', '6000 m', '10:00'," \
   " 0, -1.0, " + str(RefJahr - 22) + ", " + str(RefJahr - 19) + " )"
cursor.execute(sql)
connection.commit()

sql = "INSERT INTO rennen VALUES(" \
   "15, 'SF 1x A/B', 'F', '1x', '6000 m', '10:00'," \
   " 0, -1.0, " + str(1900) + ", " + str(RefJahr - 19) + " )"
cursor.execute(sql)
connection.commit()

sql = "INSERT INTO rennen VALUES(" \
   "16, 'JF 1x B', 'F', '1x', '6000 m', '10:00'," \
   " 0, -1.0, " + str(RefJahr - 16) + ", " + str(RefJahr - 15) + " )"
cursor.execute(sql)
connection.commit()

sql = "INSERT INTO rennen VALUES(" \
   "17, 'JF 1x B LG', 'F', '1x', '6000 m', '10:00'," \
   " 0, 55.0, " + str(RefJahr - 16) + ", " + str(RefJahr - 15) + " )"
cursor.execute(sql)
connection.commit()

# 11 JF 1x A
# 12 JF 1x A LG
# 13 SM 1x A
# 14 SM 1x B
# 15 SF 1x A/B
# 16 JF 1x B
# 17 JF 1x B LG

# Von Rennen zu Excel in neuem Programm ??
