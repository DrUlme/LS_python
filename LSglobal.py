#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Oct  5 14:57:53 2020

@author: ulf
"""

Tag     = 18
Monat   =  3
Jahr    = 2023

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





# Referenz für diese Langstrecke


Datum        = str(Tag)    + "." + str(Monat) + "."
Meldeschluss = str(Tag-10) + "." + str(Monat) + "."
DoppeltGeld  = str(Tag-7)  + "." + str(Monat) + "."
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

TrzFiles  = ["Start.trz", "3000.trz", "Ziel.trz"]
# TrzFiles  = ["Start.trz_gelesen", "3000.trz_gelesen", "Ziel.trz_gelesen"]


TrzDBname = ["secstart",  "sec3000",  "sec6000", "zeit3000", "zeit6000", "zeit"]
TrzDBpos  = [ 4,    5,   6]
Trz_m     = [ 0, 3000, 6000]

# Zeit-Versatz in TRZ-Dateien ausgleichen:
Trz_dSec  = [ 0, 0, 0]

# zusätzlich erlaubtes Gewicht für die Langstrecke
if(ZeitK == "F"):
   Gewicht = 0.0
else:
   Gewicht = 2.5


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
