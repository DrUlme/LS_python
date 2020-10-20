#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Oct  5 14:57:53 2020

@author: ulf
"""

Name   = "Erlanger Herbst-Langstrecke 2020"
DBName = "LS2020H.db"

Jahr = 2020

MeldeDir   = "Meldungen"
StartXLS   = 'Startreihenfolge_H2020.xlsx'

SQLiteFile = "LS2020H.db"

# _________________________________________________ Herbst 2019
# TrzDir    = "../H2019/Zeiten"
# 
# _________________________________________________ Herbst 2019
TrzDir    = "../Backup/LS2018F"

TrzFiles  = ["Start.trz", "3000.trz", "Ziel.trz"]
# TrzFiles  = ["Start.trz_gelesen", "3000.trz_gelesen", "Ziel.trz_gelesen"]


TrzDBname = ["secstart",  "sec3000",  "sec6000"]
TrzDBpos  = [ 6,    7,   8]
Trz_m     = [ 0, 3000, 6000]
# Zeit-Versatz in TRZ-Dateien ausgleichen:
Trz_dSec  = [ 0, 0, 0]

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
