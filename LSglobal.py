#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Oct  5 14:57:53 2020

@author: ulf
"""

Tag      = 26
Monat    = 10
MonatStr = 'Oktober'
Jahr     = 2024

if(Monat < 7):
   Zeit    = "Frühjahr" 
   Name    = "Erlanger Frühjahrs-Langstrecke " + str(Jahr)
   RefJahr = Jahr
else:
   Zeit    = "Herbst"
   Name    = "Erlanger Herbst-Langstrecke " + str(Jahr)
   # Die Herbst-LS zählt bereits als Vorbereitung auf das folgende Jahr
   RefJahr = Jahr + 1


ZeitK  = Zeit[0]
DBName = "LS" + str(Jahr) + ZeitK + ".db"

JMBriemen = 0



# Referenz für diese Langstrecke


Datum        = str(Tag)    + "." + str(Monat) + "."
Meldeschluss = str(Tag-10) + "." + str(Monat) + "."
DoppeltGeld  = str(Tag-7)  + "." + str(Monat) + "."
Setzdatum    = str(Tag-5)  + "." + str(Monat) + "."
AbmeldungDD  = str(Tag-1)  + "." + str(Monat) + "."
AbmeldungHH  = "14:00 Uhr"

MeldeDir   = "Meldungen"
StartXLS   = 'Startreihenfolge_" + ZeitK + str(Jahr) + ".xlsx'

# SQLiteFile = "LS2022F.db"
SQLiteFile = DBName

# _________________________________________________ Herbst 2019
# TrzDir    = "../H2019/Zeiten"
# 
# _________________________________________________ Herbst 2020
TrzDir    = "./Zeiten"

TrzFiles  = ["Start.trz", "3000m.trz", "Ziel.trz"]
# TrzFiles  = ["Start.trz_gelesen", "3000.trz_gelesen", "Ziel.trz_gelesen"]


TrzDBname = ["tStart",  "t3000",  "t6000", "zeit3000", "zeit6000", "zeit"]
TrzDBpos  = [ 5,    6,    7]
Trz_m     = [ 0, 3000, 6000]

# Zeit-Versatz in TRZ-Dateien ausgleichen:
Trz_dSec  = [ 0, 0, 0]

# zusätzlich erlaubtes Gewicht für die Langstrecke
if(ZeitK == "F"):
   Gewicht = 0.0
else:
   Gewicht = 2.5

Cost_1x = 18
Cost_2x = 18

Cost_Atletiktest = 6

"""
========================================================================
rennen.status = [6]
0: niemand gemeldet
1: gemeldet, aber ohne Startnummer & Zeit
2: Startnummern und Zeiten vorhanden
3: Aktiv
4: (End-) Ergebnis





========================================================================
"""
