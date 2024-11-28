#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Oct  6 19:31:42 2020

@author: ulf
"""

import sqlite3
import re 

import LSglobal

#========================================================================
# Verbindung zur Datenbank erzeugen
connection = sqlite3.connect( LSglobal.SQLiteFile )
#
# Datensatzcursor erzeugen
cursor  = connection.cursor()
cursor2 = connection.cursor()
#
#========================================================================
# Unterscheide 3000 und 6000 m - Rennen

is3000 = [0, 0]
sql = "SELECT * FROM rennen "
cursor.execute(sql)
for dsatz in cursor:
   Rennen = dsatz[0]
   if( dsatz[6] == "3000 m" ):
      is3000.insert(Rennen, 1)
   else:
      is3000.insert(Rennen, 0)

print(is3000)
#========================================================================
# Update TRZ-Files
      
for iFile in range(0, len(LSglobal.TrzFiles) ):
   file = LSglobal.TrzFiles[ iFile ]
   TXT = open( LSglobal.TrzDir + "/" + file ).readlines()
   #
   DBname = LSglobal.TrzDBname[iFile ]
   print("# ======= " + file + " =============== == =>> " + DBname )
   #
   for iL in range(0, len(TXT) ):
      PARTS = re.split('\s', TXT[iL] )
      # PARTS = re.split('\s|\:', TXT[iL] )
      # print(PARTS)
      # seconds = 3600 * int(PARTS[1]) + 60 * int(PARTS[2]) + int(PARTS[3]) + LSglobal.Trz_dSec[iFile]
      #
      # get values from SQL:
      if(LSglobal.Zeit ==  "Frühjahr" ):
         sql = "SELECT * FROM boote WHERE startnummer = " + (PARTS[0]) + "  AND rennen < 20"
      else:
         sql = "SELECT * FROM boote WHERE startnummer = " + (PARTS[0]) + " "
      cursor.execute(sql)
      Bt = cursor.fetchone()
      if(Bt == None):
         Part = "-"
         print("# " + str(PARTS[0]) + " ist NICHT in Datenbank !")         
      else:
         print("# " + str(PARTS[0]) + ": " + LSglobal.TrzFiles[ iFile ] + ": " + str(Bt[LSglobal.TrzDBpos[ iFile ]]) + "(" + str(Bt[LSglobal.TrzDBpos[ iFile ]-1]) + ")... in DB" )
         # failback nehme nur kleinste Zahl
         # if( (Bt[LSglobal.TrzDBpos[ iFile ]] <= 0 ) or ( Bt[LSglobal.TrzDBpos[ iFile ]] > seconds) ):
         #    if(Bt[LSglobal.TrzDBpos[ iFile ]-1] <= 0):
         #       print("->" +  str(seconds) + " sec - dürfte aber noch nicht hier sein!" ) 
         #    else:
         sql = "UPDATE boote SET " + DBname + " = '" + PARTS[1] + "' WHERE startnummer = " + PARTS[0]
         # print(sql)
         cursor.execute(sql)
         connection.commit()
         # elif( seconds != Bt[LSglobal.TrzDBpos[ iFile ]]):
         #    print("->" +  str(seconds) + " sec sind anders als in Datenbank: " + str(Bt[LSglobal.TrzDBpos[ iFile ]]))
         # print(Bt)
         print("_________________________________________")
         #
         # if(Bt[LSglobal.TrzDBpos[ 1 ]] > 0):
         #    print("->" +  str( Bt[LSglobal.TrzDBpos[ 1 ]] - Bt[LSglobal.TrzDBpos[ 0 ]]) + " sec für 1. 3000 m")
         # if(Bt[LSglobal.TrzDBpos[ 2 ]] > 0):
         #    print("->" +  str( Bt[LSglobal.TrzDBpos[ 2 ]] - Bt[LSglobal.TrzDBpos[ 1 ]]) + " sec für 2. 3000 m")
         #    print("->" +  str( Bt[LSglobal.TrzDBpos[ 2 ]] - Bt[LSglobal.TrzDBpos[ 0 ]]) + " sec für 6000 m")
        
   #       
#
   print("#========================================================================")
#========================================================================
# Update Zeiten
#
# sql = "SELECT * FROM boote "
# cursor2.execute(sql)
# for dsatz in cursor2:
#    Boot   = dsatz[0]
#    Rennen = dsatz[2]
#    #                für später: LSglobal.TrzDBpos[0]
#    Start  = dsatz[4]
#    m3000  = dsatz[5]
#    m6000  = dsatz[6]
#    if(Start > 0):
#       if(m3000 > 0):
#          sql = "UPDATE boote SET zeit3000 = " + str(m3000 - Start) + " WHERE nummer = " + str(Boot)
#          print(sql)
#          cursor.execute(sql)
#          connection.commit()
#          if(is3000[Rennen] == 1):
#             sql = "UPDATE boote SET zeit = " + str(m3000 - Start) + " WHERE nummer = " + str(Boot)
#             cursor.execute(sql)
#             connection.commit()
#          elif(m6000 > 0):
#             sql = "UPDATE boote SET zeit6000 = " + str(m6000 - m3000) + " WHERE nummer = " + str(Boot)
#             cursor.execute(sql)
#             connection.commit()
#             sql = "UPDATE boote SET zeit = " + str(m6000 - Start) + " WHERE nummer = " + str(Boot)
#             cursor.execute(sql)
#             connection.commit()

# Verbindung beenden
connection.close()
