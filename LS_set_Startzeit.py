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
StartRennen = 2
StartSec    = 3600*11  + 60*0 + 0

SecDif = 60

Test = 0

#   [ 00, 01, 02, 03, 04, 05, 06, 07, 08, 09, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37 ]
H = [ 11, 11, 11, 11, 11, 11, 11, 11, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13 ]
M = [ 00,  1,  2, 15, 35, 54, 55, 59,  2,  2,  6, 15, 20, 22, 27, 28, 28, 29, 37, 46, 50, 50, 52, 54, 54, 56, 58,  0,  0,  5,  9, 12, 13, 15, 21, 23, 23, 13 ]
N = [  1,  2,  3,  6, 27, 52, 52, 56, 56, 56, 58, 68, 75, 78, 78,116,116,118,129,143,145,145,148,151,151,155,158,161,161,164,170,174,177,180,188,190,190,200 ]

# cursor.execute( "SELECT MAX(nummer) FROM rennen " )
# RennenMax = cursor.fetchone()

# for iR in range(0, (len(RudInd) - 2)): 
seconds = StartSec    
for Rennen in range(1, 19): # (1, 37):
    seconds = 3600*H[Rennen]  + 60*M[Rennen] + 0
    print("Rennen " + str(Rennen) + " ------------- at " + "{:02d}".format(H[Rennen]) + ":" + "{:02d}".format(M[Rennen]))
    #----------------------------------------------
    # sql = "SELECT * FROM boote  WHERE rennen = " + str(Rennen) + " AND startnummer >= " + str(N[Rennen]) "ORDER BY startnummer "
    # sql = "SELECT * FROM boote  WHERE startnummer >= " + str(N[Rennen]) + " AND startnummer < " + str(N[Rennen+1]) + "  ORDER BY startnummer "
    sql = "SELECT * FROM boote  WHERE startnummer >= " + str(N[Rennen]) + " AND startnummer < " + str(N[Rennen+1]) + " AND rennen < 20 ORDER BY startnummer "
    Bcursor.execute(sql)
    for Bsatz in Bcursor:
      bootNr = Bsatz[0]
      StNr   = Bsatz[1]
      Abmeldung = Bsatz[10]
      if(Abmeldung > 0):
         sql = "UPDATE boote SET planstart = 0 WHERE startnummer = " + str(StNr) + " "
      else:      
         # sql = "UPDATE boote SET planstart = " + str( seconds ) + " WHERE startnummer = " + str(StNr) + " "
         sql = "UPDATE boote SET planstart = " + str( seconds ) + " WHERE nummer = " + str(bootNr)
         seconds = seconds + SecDif
      if(Test == 1):
         print(sql)
      else:
         cursor.execute(sql)
         connection.commit()
         print("executed: " + sql)
    #---
    LastH = np.floor(seconds/3600)
    LastM = np.floor((seconds-LastH*3600)/60)
    LastS = np.floor((seconds-LastH*3600 - LastM*60))
    print("next would be Rennen " + "{:.0f}".format(Rennen + 1) + " at " + "{:2.0f}".format(LastH) + ":" + "{:2.0f}".format(LastM) + ":" + "{:2.0f}".format(LastS))

connection.close()
