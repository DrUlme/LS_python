#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Oct 23 18:05:59 2021

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

sql = "UPDATE boote SET secstart = 0" 
cursor.execute(sql)
connection.commit()

sql = "UPDATE boote SET sec3000 = 0" 
cursor.execute(sql)
connection.commit()

sql = "UPDATE boote SET sec6000 = 0" 
cursor.execute(sql)
connection.commit()

sql = "UPDATE boote SET zeit = 0" 
cursor.execute(sql)
connection.commit()

sql = "UPDATE boote SET zeit3000 = 0" 
cursor.execute(sql)
connection.commit()

sql = "UPDATE boote SET zeit6000 = 0" 
cursor.execute(sql)
connection.commit()
