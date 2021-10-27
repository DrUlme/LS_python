import os, sys, sqlite3
from time import strftime
from time import gmtime
import time
from openpyxl.styles import numbers

import LSglobal

# Excel
#========================================================================
# open existent workbook:
from openpyxl import load_workbook

filename = 'Startreihenfolge_offiziell.xlsx'
# filename = 'Startreihenfolge_H2020.xlsx'
# filename = LSglobal.StartXLS

# read the last used value (not the formula)
wb = load_workbook(filename, data_only=True)

print(wb.sheetnames)
# load worksheets: the last used one, may be wrong ...
#ws = wb.active
# Lese Meldungen
ws = wb["Meldungen"]
# ws = wb.get_sheet_by_name("Rennen")
#========================================================================
# Verbindung zur Datenbank erzeugen
connection = sqlite3.connect( LSglobal.SQLiteFile )
#
# Datensatzcursor erzeugen
cursor  = connection.cursor()
cursorR = connection.cursor()
#
#========================================================================
#
zeile = 7
while zeile > 6:
   #________________________________________________________________
   # integer Werte
   Rennen = ws['A' + str(zeile)].value 
   Positi = ws['B' + str(zeile)].value 
   StNr   = ws['C' + str(zeile)].value
   Boot   = ws['K' + str(zeile)].value
   # Zeit-String
   bZeit  = str(ws['D' + str(zeile)].value)
   #
   #   if(Boot == None):
   #      print(' - ')
   #   elif(Boot < 1):
   #      print('---')
   #   else:
   if(Boot != None and Boot > 0):
      myT = bZeit.split(":")
      mySec = 3600*int(myT[0]) + 60*int(myT[1]) + int(myT[2])
      print("Zeile " + str(zeile) + ": _____________")
      # 
      Vornamen = ws['E' + str(zeile)].value
      Vorname  = Vornamen.split("\n")
      #
      Namen    = ws['F' + str(zeile)].value
      Name     = Namen.split("\n")
      #
      Jahrgang = ws['G' + str(zeile)].value
      Jahre    = Jahrgang.split("\n")
      #
      Vereine  = ws['I' + str(zeile)].value
      Verein   = Vereine.split("\n")
      #
      print("Rennen " + str(Rennen) + "." + str(Positi) + "= #" + str(StNr) + " (" + str(Boot) + ") :  " + str(mySec))      
      sql = "UPDATE rennen SET status = 2  WHERE nummer = " + str(Rennen)
      # print( sql )
      cursor.execute(sql)
      connection.commit()

      #
      # SQL-Abfrage
      sql = "SELECT * FROM boote WHERE nummer = " + str(Boot)
      #
      # Empfang des Ergebnisses
      cursor.execute(sql)
      # _______________________________________________________________________ korrekte Ruderer ?
      # for dsatz in cursor:
      dsatz = cursor.fetchone()
      # print(dsatz)
      # print("Rennen " + str(Rennen) + " = " + str(dsatz[2]) )
      # print("Verein " + Verein[0] + " = " + dsatz[3])
      ruderer = ()
      # dsatz[4].split(',')
      for iR in range(0, (len(ruderer) - 3)):
         #
         sql = "SELECT * FROM ruderer WHERE nummer = " + str(ruderer[iR + 1])
         # Empfang des Ergebnisses
         cursorR.execute(sql)
         Rd = cursorR.fetchone()
         # Vorname
         sqlVorname = Rd[0]
         # Nachname
         sqlName = Rd[1]
         # Jahrgang
         sqlJahrgang = str( Rd[3] )
         if( Vorname[iR] != sqlVorname or Name[iR] != sqlName ):
            print(Vorname[iR] + " " + Name[iR] + " != " + sqlVorname + " " + sqlName)
            # x = raw_input("Ändern? [Y/n]")
            x = input("Ändern? [Y/n] > ")
            if(x == "Y" or x== "y" or x == "j" or x == "J"):
               sql = "UPDATE ruderer SET vorname = " + Vorname[iR] + ", name = " + Name[iR] + " WHERE nummer = " + str(ruderer[iR + 1])
               # print( sql )
               cursor.execute(sql)
               connection.commit()
               print("Würde jetzt ändern!")
         if( Jahre[iR] != sqlJahrgang ):
            print(Jahre[iR] + " != " + sqlJahrgang + "  (" + sqlVorname + " " + sqlName + ")")
            x = input("Ändern? [Y/n] > ")
            if(x == "Y" or x== "y" or x == "j" or x == "J"):
               print("Würde jetzt ändern!")
            # ToDo: Abfrage oder Änderung?
      # ______________________________________________________________________________________
      #
      if( Rennen != dsatz[2] ):
            print( "Boot " + str(Boot) + " mit Startnr " + str(StNr) + " von Rennen " + str(dsatz[2]) + " nach " + str(Rennen) + " ?!" )
            # x = raw_input("Ändern? [Y/n]")
            x = input("Ändern? [Y/n] > ")
            if(x == "Y" or x== "y" or x == "j" or x == "J"):
               sql = "UPDATE boote SET rennen = " + str(Rennen) + " WHERE nummer = " + str(Boot)
               # print( sql )
               cursor.execute(sql)
               connection.commit()
               # dto. für r2boote - entfernen?!
               sql = "UPDATE r2boote SET rennNr = " + str(Rennen) + " WHERE bootNr = " + str(Boot)
               cursor.execute(sql)
               connection.commit()
               print("... geändert !")
         
      # ______________________________________________________________________________________
      #
      sql = "UPDATE boote SET startnummer = " + str(StNr) + ", planstart = " + str(mySec) + " WHERE nummer = " + str(Boot)
      # print( sql )
      cursor.execute(sql)
      connection.commit()
         # dto. mit StNr
   #
   zeile = zeile + 1
   if(bZeit == None):
      zeile = 0
   if(len(bZeit) <= 5):
      zeile = 0
   #_____________________________________________________________


#______________________________________ EOF
