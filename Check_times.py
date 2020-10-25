#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Oct 24 16:58:23 2020

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
#
sql = "SELECT * FROM boote WHERE abgemeldet = 0 ORDER BY startnummer "
cursor.execute(sql)

Rennen = 0
for dsatz in cursor:
   if(Rennen < dsatz[2]):
      print("######################### Rennen " + str(dsatz[2]) + "  ####")
      Rennen = dsatz[2]
   # ReStr  = str(Rennen)
   # Btime = dsatz[6] # Startzeit
   # 
   Btime = dsatz[7] # Zeit 3000 m
   # Btime = dsatz[8] # Zielzeit
   if(Btime > 0):
      BtimH   = math.floor(Btime/3600)
      BtimM   = math.floor(Btime/60 - BtimH*60 )
      ZeitStr =  str(BtimH) + ":" + str(BtimM).rjust(2, '0') + ":" + str(Btime - 3600*BtimH - 60*BtimM).rjust(2, '0') + " "
      print(  str(dsatz[1]) + " = " + ZeitStr )
#   