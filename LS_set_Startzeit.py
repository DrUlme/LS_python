#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Oct  7 16:59:41 2020

@author: ulf
"""

import os, sys, sqlite3
import numpy as np

import LSglobal

#========================================================================
# Verbindung zur Datenbank erzeugen
connection = sqlite3.connect( LSglobal.SQLiteFile )
#
# Datensatzcursor erzeugen
cursor  = connection.cursor()
Bcursor = connection.cursor()
#

# Renn-Abstand in sec
Abstand = 50
#========================================================================
Rennen = 18

# wenn Startnummer > 0 wird sie gesetzt !
Startnummer = 0
Erstes = "13:07:00"

Test = 1

#========================================================================
Hs = int(Erstes[0:2])
Ms = int(Erstes[3:5])
Ss = int(Erstes[6:8])

secStart = 3600*Hs  + 60*Ms + Ss

if(Rennen == 2 or Rennen == 18):
   # Wenn Früh oder Spätstarter:
   sql = "SELECT * FROM boote  WHERE abgemeldet == 0 AND alternativ ==" + str(Rennen) + "  ORDER BY planstart, startnummer "
else:
   # sonst
   sql = "SELECT * FROM boote  WHERE abgemeldet == 0 AND Rennen ==" + str(Rennen) + " AND alternativ == 0  ORDER BY planstart, startnummer "

Bcursor.execute(sql)
n = 0
for Bsatz in Bcursor:
   bootNr = Bsatz[0]
   StNr   = Bsatz[2]
   #
   seconds = secStart + n * Abstand
   LastH = np.floor(seconds/3600)
   LastM = np.floor((seconds-LastH*3600)/60)
   LastS = np.floor((seconds-LastH*3600 - LastM*60))
   Startzeit = str(int(LastH)).rjust(2, '0') + ":" + str(int(LastM)).rjust(2, '0') + ":" + str(int(LastS)).rjust(2, '0')
   #
   sql = "UPDATE boote SET planstart = '" + Startzeit + "' WHERE id = " + str(bootNr)
   if(Test == 1):
      print(sql)
   else:
      cursor.execute(sql)
      connection.commit()
      print("executed: " + sql)
      #
   if(Startnummer > 0):
      # sql = "UPDATE boote SET planstart = '" + str(int(LastH)).rjust(2, '0') + ":" + str(int(LastM)).rjust(2, '0') + ":" + str(int(LastS)).rjust(2, '0') + "' AND SET startnummer = " + str(Startnummer) + " WHERE id = " + str(bootNr)
      sql = "UPDATE boote SET startnummer = " + str(Startnummer) + " WHERE id = " + str(bootNr)
      Startnummer = Startnummer + 1
      #--
      if(Test == 1):
         print(sql)
      else:
         cursor.execute(sql)
         connection.commit()
         print("executed: " + sql)
   #--|--
   n = n + 1
#---
print("Rennen " + str(Rennen) + " OK,\n nächstes mit Startnummer " + str(Startnummer) + " und nach " + Startzeit)
seconds = secStart + n * Abstand
LastH = np.floor(seconds/3600)
LastM = np.floor((seconds-LastH*3600)/60)
LastS = np.floor((seconds-LastH*3600 - LastM*60))
Startzeit_1 = str(int(LastH)).rjust(2, '0') + ":" + str(int(LastM)).rjust(2, '0') + ":" + str(int(LastS)).rjust(2, '0')
seconds = secStart + (n+1) * Abstand
LastH = np.floor(seconds/3600)
LastM = np.floor((seconds-LastH*3600)/60)
LastS = np.floor((seconds-LastH*3600 - LastM*60))
Startzeit_2 = str(int(LastH)).rjust(2, '0') + ":" + str(int(LastM)).rjust(2, '0') + ":" + str(int(LastS)).rjust(2, '0')
print(" => mindestends " + Startzeit_1 + " oder mit 1 Leerstelle: " + Startzeit_2)

connection.close()
