#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Oct 23 19:57:30 2021

@author: ulf
"""
import os, sys, sqlite3

# import modules
import tkinter as tk
# 
from PIL import Image
#, ImageTk

# globale Parameter

#========================================================================
class bcolors:
    OK = '\033[92m' #GREEN
    WARNING = '\033[93m' #YELLOW
    FAIL = '\033[91m' #RED
    RESET = '\033[0m' #RESET COLOR

#========================================================================

import LSglobal

# Verbindung zur Datenbank erzeugen
connection = sqlite3.connect( LSglobal.SQLiteFile )

# Datensatzcursor erzeugen
cursor  = connection.cursor()
Qcursor = connection.cursor()
Bcursor = connection.cursor()
Rcursor = connection.cursor()


#========================================================================
sql = "SELECT * FROM ruderer WHERE leichtgewicht = 1"   
cursor.execute(sql)
for dsatz in cursor:
   Alter  = LSglobal.RefJahr - dsatz[4]
   Gender = dsatz[3]
   if(Gender == "M"):
      if(dsatz[6] < 1):
         print(bcolors.WARNING + "Ruderer '" + dsatz[1] + " " + dsatz[2] +"' ist noch nicht verwogen worden" + bcolors.RESET)
      else:
         if(Alter < 14):
            Soll = 50
         elif(Alter < 15):
            Soll = 55
         elif(Alter < 17):
            Soll = 65
         elif(Alter < 19):
            Soll = 67.5
         else:
            Soll = 72.5
         if(dsatz[6] <= Soll ):
            print(bcolors.OK + "Ruderer '" + dsatz[1] + " " + dsatz[2] + "' hat das Gewicht für " + str(Alter) + " Jahre (" +  str(dsatz[6]) + " / " + str(Soll) + ")" + bcolors.RESET)
         elif(dsatz[6] <= ( Soll + LSglobal.Gewicht )):
            print(bcolors.WARNING + "Ruderer '" + dsatz[1] + " " + dsatz[2] + "' hat NOCH das Gewicht für " + str(Alter) + " Jahre (" +  str(dsatz[6]) + " / " + str(Soll) + ")" + bcolors.RESET)
         else:
            print(bcolors.FAIL + "Ruderer '" + dsatz[1] + " " + dsatz[2] +"' hat NICHT sein Gewicht: " + str(dsatz[6]) + " > " + str(Soll) + "+" + str(LSglobal.Gewicht) + " kg" + bcolors.RESET)
     
   else:
      if(dsatz[6] < 1):
         print(bcolors.WARNING + "Ruderin '" + dsatz[1] + " " + dsatz[2] +"' ist nicht verwogen worden " + bcolors.RESET)
      else:
         if(Alter < 14):
            Soll = 50
         elif(Alter < 15):
            Soll = 55
         elif(Alter < 17):
            Soll = 65
         elif(Alter < 19):
            Soll = 67.5
         else:
            Soll = 72.5
         #
         if(dsatz[6] <= Soll ):
            print(bcolors.OK + "Ruderin '" + dsatz[1] + " " + dsatz[2] + "' hat das Gewicht für " + str(Alter) + " Jahre (" +  str(dsatz[6]) + " / " + str(Soll) + ")" + bcolors.RESET)
         elif(dsatz[6] <= ( Soll + LSglobal.Gewicht )):
            print(bcolors.WARNING + "Ruderin '" + dsatz[1] + " " + dsatz[2] + "' hat NOCH das Gewicht für " + str(Alter) + " Jahre (" +  str(dsatz[6]) + " / " + str(Soll) + ")" + bcolors.RESET)
         else:
            print(bcolors.FAIL + "Ruderin '" + dsatz[1] + " " + dsatz[2] +"' hat NICHT ihr Gewicht: " + str(dsatz[6]) + " > " + str(Soll) + "+" + str(LSglobal.Gewicht) + " kg" + bcolors.RESET)

   # sql = "SELECT * FROM r2boot  WHERE rudererNr = " + str(dsatz[0])
   # Qcursor.execute(sql)
   # for RBind in Qcursor:
   #    sql = "SELECT * FROM boote  WHERE nummer = " + str(RBind[2]) 
   #    Bcursor.execute(sql)
   #    Boot = Bcursor.fetchone()
   #    # test ob abgemeldet
   #    if(Boot[10] == 0):
   #       # Rennen suchen
   #       sql = "SELECT gewicht FROM rennen WHERE nummer = " + str(Boot[2]) 
   #       Rcursor.execute(sql)
   #       RennGewicht = Rcursor.fetchone()
   #       # Gewicht von Rennen zu Gewicht von Ruderer
   #       if(dsatz[6] < 1):
   #          print(bcolors.WARNING + "Ruderer '" + dsatz[1] + " " + dsatz[2] +"' ist nicht verwogen worden (Re. " + str(Boot[2]) + ")" + bcolors.RESET)
   #       elif(dsatz[6] <= RennGewicht[0] ):
   #          print(bcolors.OK + "Ruderer '" + dsatz[1] + " " + dsatz[2] +"' hat sein Gewicht - " + str(Alter) + " Jahre alt" + bcolors.RESET)
   #       elif(dsatz[6] <= ( RennGewicht[0] + LSglobal.Gewicht )):
   #          print(bcolors.WARNING + "Ruderer '" + dsatz[1] + " " + dsatz[2] +"' hat noch sein Gewicht " + bcolors.RESET)
   #       else:
   #          print(bcolors.FAIL + "Ruderer '" + dsatz[1] + " " + dsatz[2] +"' hat NICHT sein Gewicht: " + str(dsatz[6]) + " > " + str(RennGewicht[0]) + " kg" + bcolors.RESET)
#
