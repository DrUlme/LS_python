#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Oct  7 16:59:41 2020

@author: ulf
"""

import os, sys, sqlite3

import LSglobal

#========================================================================
# Verbindung zur Datenbank erzeugen
connection = sqlite3.connect( LSglobal.SQLiteFile )
#
# Datensatzcursor erzeugen
cursor  = connection.cursor()
Bcursor = connection.cursor()
#
StartRennen = 5
StartSec    = 3600*11  + 60*42 + 0

SecDif = 60

cursor.execute( "SELECT MAX(nummer) FROM rennen " )
RennenMax = cursor.fetchone()

# for iR in range(0, (len(RudInd) - 2)): 
seconds = StartSec    
for Rennen in range(StartRennen, (RennenMax[0]+1) ):
    sql = "SELECT * FROM boote  WHERE rennen = " + str(Rennen) + " ORDER BY startnummer "
    Bcursor.execute(sql)
    for Bsatz in Bcursor:
      StNr   = Bsatz[1]
      Abmeldung = Bsatz[12]
      if(Abmeldung > 0):
         sql = "UPDATE boote SET planstart = 0 WHERE startnummer = " + str(StNr)
      else:      
         sql = "UPDATE boote SET planstart = " + str( seconds ) + " WHERE startnummer = " + str(StNr)
         seconds = seconds + SecDif
      #
      print(sql)
      cursor.execute(sql)
      connection.commit()

