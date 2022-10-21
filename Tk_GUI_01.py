#!/bin/python3
# GUI for a database
#  python3

import os, sys, sqlite3

# import modules
import tkinter as tk
# 
from PIL import Image
#, ImageTk

# globale Parameter
import LSglobal

#========================================================================
# Verbindung zur Datenbank erzeugen
connection = sqlite3.connect( LSglobal.SQLiteFile )
# Datensatzcursor erzeugen
cursor  = connection.cursor()
Qcursor = connection.cursor()
Bcursor = connection.cursor()

rudererInd = 0
bootNr=0

#========================================================================
def searchID():
   global rudererInd
   global bootNr
   print("searchID: ")
   sql = "SELECT * FROM ruderer WHERE "   
   LTEXT = "found: "
   STEXT = ""
   BTEXT = ""
   # setze den globalen Index wider zurück:
   rudererInd = 0
   bootNr = 0
   #
   useNAME = 0
   #
   StartNr  = myNr.get()
   if(len(StartNr) > 0):
      useNR = 1
   else:
      useNR = 0
   print("use Startnummer '" + str(StartNr) +"' (" + str(useNR) + ")")
   #
   Vorname  = myVorName.get()
   if(len(Vorname) > 1):
      # sql = sql + "vorname='" + Vorname +"' "
      sql = sql + "vorname LIKE '" + Vorname +"%' "
      useNAME = 1
   #
   Nachname = myName.get()  
   if(len(Nachname) > 1):
      if(useNAME > 0):
         sql = sql + " and "
      sql = sql + "name LIKE '" + Nachname +"%' "
      useNAME = 1
   #
   mylabel.configure(text=sql)
   # print(sql)
   if(useNAME > 0):
      nR = 0
      cursor.execute(sql)
      for dsatz in cursor:
         nR = nR + 1
         # LTEXT = LTEXT + str(dsatz[0]) + ": '" + dsatz[1] + "' '" + dsatz[2] + "' (" +  + dsatz[7] + ")\n"   
         LTEXT = LTEXT + "'" + dsatz[1] + "' '" + dsatz[2] + "' (" + str(dsatz[0]) + "):  " 
         # print(dsatz)
         RnrStr = str(dsatz[0]) 
         sql = "SELECT * FROM r2boot  WHERE rudererNr = " + RnrStr
         Qcursor.execute(sql)
         anzahl = 0
         for RBind in Qcursor:
            anzahl = anzahl + 1
            sql = "SELECT * FROM boote  WHERE nummer = " + str(RBind[1]) 
            Bcursor.execute(sql)
            Boot = Bcursor.fetchone()
            if(anzahl > 1):
               bootNr = 0
            else:
               bootNr = Boot[0]
            #
            LTEXT = LTEXT + "#" + str(Boot[0]) + ": StNr." + str(Boot[1]) + " in Rennen " + str(Boot[2]) + "\n" 
            BTEXT = BTEXT + "#" + str(Boot[0]) + ": StNr." + str(Boot[1]) + " in Rennen " + str(Boot[2]) 
            if(Boot[10] == 1):
               BTEXT = BTEXT + " - abgemeldet\n"
            else:
               BTEXT = BTEXT + " - ok\n"
            # print("B#" + str(Boot[0]) + ": StartNr." + str(Boot[1]) + " in Rennen " + str(Boot[2]) )
            if(useNR and int(StartNr) == int(Boot[1])):
               STEXT = "'" + dsatz[1] + "' '" + dsatz[2] + "':  Boot #" + str(Boot[0]) + ": StNr. " + str(Boot[1]) + " in Rennen " + str(Boot[2])
               #
               if(Boot[10] == 1):
                  STEXT = STEXT + " - abgemeldet"
         #  #  #
      if(nR == 1):
         rudererInd = dsatz[0]
   elif(useNR > 0):
      sql = "SELECT * FROM boote  WHERE startnummer = " + str(StartNr) 
      if(LSglobal.ZeitK == "F"):
         sql = sql + " and rennen<21"
      Bcursor.execute(sql)
      Boot = Bcursor.fetchone()
      #
      BTEXT = BTEXT + "#" + str(Boot[0]) + ": StNr." + str(Boot[1]) + " in Rennen " + str(Boot[2]) 
      if(Boot[10] == 1):
         BTEXT = BTEXT + " - abgemeldet\n"
      else:
         BTEXT = BTEXT + " - ok\n"
      #
      bootNr = Boot[0]
      sql = "SELECT * FROM r2boot  WHERE bootNr = " + str(Boot[0])
      Qcursor.execute(sql)
      nR = 0
      for RBind in Qcursor:
         sql = "SELECT * FROM ruderer WHERE nummer = " + str(RBind[2])
         cursor.execute(sql)
         dsatz = cursor.fetchone()
         nR = nR + 1
      if(nR == 1):
         rudererInd = dsatz[0]
         STEXT = "'" + dsatz[1] + "' '" + dsatz[2] + "':  Boot #" + str(Boot[0]) + ": StNr. " + str(Boot[1]) + " in Rennen " + str(Boot[2])
   #_______________________________________________________________________________________________________________________________________________
   # mylabel.configure(text=LTEXT)    
   if(len(STEXT) > 3):
      print(STEXT)
      mylabel.configure(text=STEXT) 
   else:
      print(LTEXT)
      mylabel.configure(text=LTEXT)
   #
   if(len(BTEXT) > 3):
      print(BTEXT)
      myStatus.configure(text=BTEXT) 
   else:
      print('BTEXT="' + BTEXT + '"')
      myStatus.configure(text="?") 
   #
   print(" - saved Index for ruderer = " + str(rudererInd) )
   return
#========================================================================
def submit():
   global rudererInd
   Gewicht  = myKG.get()
   if(len(Gewicht) > 0):
      if(rudererInd == 0):
         print("submit " + str(Gewicht) + " kg - with no defined rower ?!")
      else:
         sql = "SELECT * FROM ruderer WHERE nummer = " + str(rudererInd)
         Bcursor.execute(sql)
         Ruderer = Bcursor.fetchone()
         if(Ruderer[5] == 1):
            LGWstr = "Lgw."
         else:
            LGWstr = " ? "
         print("'" + Ruderer[1] + "' '" + Ruderer[2] + "' (" + LGWstr + ") = " + str(Gewicht) + " kg" )
         sql = "UPDATE ruderer SET gewicht = " + str(Gewicht) + " WHERE nummer = " + str(rudererInd)
         cursor.execute(sql)
         connection.commit()
   return
#========================================================================
def abmelden():
   global bootNr
   if(bootNr == 0):
      print("submit abgemeldet  - with no defined boot ?!")
   else:
      sql = "SELECT * FROM boote WHERE nummer = " + str(bootNr)
      Bcursor.execute(sql)
      Boot = Bcursor.fetchone()
      if(Boot[10] == 1):
         abmeldung = "0"
      else:
         abmeldung = "1"
      # 
      sql = "UPDATE boote SET abgemeldet = " + abmeldung + " WHERE nummer = " + str(bootNr)
      cursor.execute(sql)
      connection.commit()
      #
      BTEXT = "#" + str(Boot[0]) + ": StNr." + str(Boot[1]) + " in Rennen " + str(Boot[2]) 
      if(Boot[10] == 0):
         BTEXT = BTEXT + " - abgemeldet\n"
      else:
         BTEXT = BTEXT + " - ok\n"

      myStatus.configure(text=BTEXT) 
   return
#========================================================================
win = tk.Tk()
win.title('my GUI for Langstrecke')
win.geometry("800x400")

#========================================================================

# Databases
mydata=tk.Label(win,text="Langstrecke " + str(LSglobal.Jahr) + " " + LSglobal.ZeitK,font=("Hack",16),fg="blue")
mydata.grid(row=0,column=0)

mylabel=tk.Label(win,text="Kein Name\n kein Verein\nangegeben",font=("Hack",10),fg="blue")
mylabel.grid(row=4,column=0,columnspan=3)

myStatus=tk.Label(win,text="?",font=("Hack",10),fg="blue")
myStatus.grid(row=8,column=2,columnspan=1)
#

butGetEntry = tk.Button(win, text="Hole Einträge", command=searchID)
butGetEntry.grid(row=6,column=0,padx=10,pady=10,ipadx=20)

butSetKG = tk.Button(win, text="Setze Gewicht", command=submit)
butSetKG.grid(row=9,column=0,padx=10,pady=10,ipadx=20)

butAbmeldung = tk.Button(win, text="Abmeldung", command=abmelden)
butAbmeldung.grid(row=9,column=2,padx=10,pady=10,ipadx=20)

H_StartNr=tk.Label(win,text="Startnr",font=("Hack",10),fg="blue")
H_StartNr.grid(row=2,column=0)
myNr = tk.Entry(win)
myNr.grid(row=3,column=0)

H_Vorname=tk.Label(win,text="Vorname",font=("Hack",10),fg="blue")
H_Vorname.grid(row=2,column=1)
myVorName = tk.Entry(win)
myVorName.grid(row=3,column=1)

H_Nachname=tk.Label(win,text="Nachname",font=("Hack",10),fg="blue")
H_Nachname.grid(row=2,column=2)
myName = tk.Entry(win)
myName.grid(row=3,column=2)

H_KG=tk.Label(win,text="[kg]",font=("Hack",10),fg="blue")
H_KG.grid(row=7,column=0)
myKG = tk.Entry(win)
myKG.grid(row=8,column=0)


logo = tk.PhotoImage(file="tiny_RVE_BRV_Flag.png")
w1 = tk.Label(win, image=logo).grid(row=0,column=1,columnspan=2,sticky="E")
#Wuerfel = tk.Label(win, image=imgBRV)
# Wuerfel.grid(row=0,column=1)

win.mainloop()
