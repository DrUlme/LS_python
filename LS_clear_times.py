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
#

removeSet = 0
removeStart = 1

if(removeSet == 1):
   for DBname in LSglobal.TrzDBname:
      sql = "UPDATE boote SET " + DBname + " = 0 "
      cursor.execute(sql)
      connection.commit()

if(removeStart == 1):
      sql = "UPDATE boote SET startnummer = 0 "
      print(sql)
      cursor.execute(sql)
      connection.commit()
      #
      sql = "UPDATE boote SET planstart = 0 "
      print(sql)
      cursor.execute(sql)
      connection.commit()