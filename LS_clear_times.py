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

for DBname in LSglobal.TrzDBname:
   sql = "UPDATE boote SET " + DBname + " = 0 "
   cursor.execute(sql)
   connection.commit()

