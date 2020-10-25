#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Oct 25 18:22:31 2020

@author: ulf
"""


import sqlite3
import re 

import LSglobal

nummer = 66

#========================================================================
# Verbindung zur Datenbank erzeugen
connection = sqlite3.connect( LSglobal.SQLiteFile )
#
# Datensatzcursor erzeugen
cursor  = connection.cursor()
#
sql = "SELECT * FROM boote WHERE startnummer = " + str(nummer) + " "
cursor.execute(sql)
Bsatz = cursor.fetchone()

#========================================================================

print("______________________________ # " + str(nummer) + " ______________________________ \n")

#======================================================================== Zeiten
Stime = Bsatz[6]
StimH = math.floor(Stime/3600)
StimM = math.floor(Stime/60 - StimH*60 )
StZeit = " " + str(StimH) + ":" + str(StimM).rjust(2, '0') + ":" + str(Stime - 3600*StimH - 60*StimM).rjust(2, '0') + " "
# ---------------------------------------------------------------------
Time3  = Bsatz[9]
Time3m = math.floor(Time3/60)
# ---------------------------------------------------------------------
Stime = Bsatz[7]
StimH = math.floor(Stime/3600)
StimM = math.floor(Stime/60 - StimH*60 )
Zeit3 = " " + str(StimH) + ":" + str(StimM).rjust(2, '0') + ":" + str(Stime - 3600*StimH - 60*StimM).rjust(2, '0') + " "
# ---------------------------------------------------------------------
if(Bsatz[8] > 0):
   Stime = Bsatz[8]
   StimH = math.floor(Stime/3600)
   StimM = math.floor(Stime/60 - StimH*60 )
   Zeit6 = " " + str(StimH) + ":" + str(StimM).rjust(2, '0') + ":" + str(Stime - 3600*StimH - 60*StimM).rjust(2, '0') + " "
   # ---------------------------------------------------------------------
   Time6 = Bsatz[10]
   Time6m = math.floor(Time6/60)
   # ---------------------------------------------------------------------
   Btime = Bsatz[11]   
   BtimM = math.floor(Btime/60)
   # ---------------------------------------------------------------------
   print( StZeit + " => [" + str(Time3m) + ":" + str(Time3 - 60*Time3m).rjust(2, '0') + "] =>  " \
   + Zeit3 + " => [" + str(Time6m) + ":" + str(Time6 - 60*Time6m).rjust(2, '0') + "] => " + Zeit6 + " = total: " + \
   str(BtimM) + ":" + str(Btime - 60*BtimM).rjust(2, '0') )
else:
   print( StZeit + " => [" + str(Time3m) + ":" + str(Time3 - 60*Time3m).rjust(2, '0') + "] =>  " + Zeit3 )
   
   
   
# ---------------------------------------------------------------------
#         3000 m     Bsatz[9]
#         6000 m     Bsatz[10]
#         Endzeit    Bsatz[11]
# ---------------------------------------------------------------------
