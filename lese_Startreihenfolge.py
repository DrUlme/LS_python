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

# filename = 'Meldungen/01_RVE_Thea.xlsx'
# filename = 'Startreihenfolge_H2020.xlsx'
filename = LSglobal.StartXLS

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
cursor = connection.cursor()
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
   Boot   = ws['J' + str(zeile)].value
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
      Vereine  = ws['H' + str(zeile)].value
      Verein   = Vereine.split("\n")
      
      print("Rennen " + str(Rennen) + "." + str(Positi) + "= #" + str(StNr) + " (" + str(Boot) + ") :  " + str(mySec))
   #
   zeile = zeile + 1
   if(bZeit == None):
      zeile = 0
   if(len(bZeit) <= 5):
      zeile = 0
   #_____________________________________________________________


#______________________________________ EOF
