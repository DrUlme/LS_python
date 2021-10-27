#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Oct  5 14:57:53 2020

@author: ulf
"""

Name   = "Erlanger Herbst-Langstrecke 2021"
DBName = "LS2021H.db"

Zeit   = "Herbst"
ZeitK  = "H"

Datum   = "23.10."
Jahr    = 2021

# Referenz für diese Langstrecke
# Die Herbst-LS zählt bereits als Vorbereitung auf das folgende Jahr
RefJahr = 2022

Meldeschluss = "13.10."
DoppeltGeld  = "16.10."
AbmeldungDD  = "22.10."
AbmeldungHH  = "14:00 Uhr"

MeldeDir   = "Meldungen"
StartXLS   = 'Startreihenfolge_H2021.xlsx'

SQLiteFile = "LS2021H.db"

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
Trz_dSec  = [ 0, 68, 0]

# zusätzlich erlaubtes Gewicht für die Langstrecke
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
