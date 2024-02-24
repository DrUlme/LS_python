#!/bin/python3
# GUI for a database
#  python3

# os, sys,
import  sqlite3

# import modules
import tkinter as tk
#
# from datetime import datetime


# globale Parameter
import LSglobal

#========================================================================
# Verbindung zur Datenbank erzeugen
connection = sqlite3.connect( LSglobal.SQLiteFile )
# Datensatzcursor erzeugen
cursor  = connection.cursor()
Qcursor = connection.cursor()
Bcursor = connection.cursor()

rudererID = 0
bootNr=0

#========================================================================
def searchID(event):
   global rudererID
   global bootNr
   print("searchID: ")
   sql = "SELECT * FROM ruderer WHERE "   
   LTEXT = "found: "
   STEXT = ""
   BTEXT = ""
   # setze den globalen Index wieder zurück:
   rudererID = 0
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
         RnrStr = dsatz[0]
         sql = "SELECT * FROM r2boot  WHERE rudererid = '" + RnrStr + "'"
         Qcursor.execute(sql)
         anzahl = 0
         for RBind in Qcursor:
            anzahl = anzahl + 1
            sql = "SELECT * FROM boote  WHERE id = '" + RBind[1] + "'"
            Bcursor.execute(sql)
            Boot = Bcursor.fetchone()
            if(anzahl > 1):
               bootNr = 0
            else:
               bootNr = Boot[0]
            #
            LTEXT = LTEXT + "#" + Boot[0] + ": StNr." + str(Boot[2]) + " in Rennen " + str(Boot[3]) + "\n" 
            BTEXT = BTEXT + "#" + Boot[0] + ": StNr." + str(Boot[2]) + " in Rennen " + str(Boot[3]) 
            if(Boot[11] == 1):
               BTEXT = BTEXT + " - abgemeldet\n"
            else:
               BTEXT = BTEXT + " - ok\n"
            # print("B#" + str(Boot[0]) + ": StartNr." + str(Boot[1]) + " in Rennen " + str(Boot[2]) )
            if(useNR and int(StartNr) == int(Boot[2])):
               STEXT = "'" + dsatz[1] + "' '" + dsatz[2] + "':  Boot #" + str(Boot[0]) + ": StNr. " + str(Boot[2]) + " in Rennen " + str(Boot[3])
               #
               if(Boot[11] == 1):
                  STEXT = STEXT + " - abgemeldet"
         #  #  #
      if(nR == 1):
         rudererID = dsatz[0]
   elif(useNR > 0):
      sql = "SELECT * FROM boote  WHERE startnummer = " + str(StartNr) 
      #if(LSglobal.ZeitK == "F"):
      #   sql = sql + " and rennen<21"
      Bcursor.execute(sql)
      Boot = Bcursor.fetchone()
      #
      BTEXT = BTEXT + "#" + str(Boot[0]) + ": StNr." + str(Boot[2]) + " in Rennen " + str(Boot[3]) 
      if(Boot[11] == 1):
         BTEXT = BTEXT + " - abgemeldet\n"
      else:
         BTEXT = BTEXT + " - ok\n"
      #
      bootNr = Boot[0]
      sql = "SELECT * FROM r2boot  WHERE bootid = '" + Boot[0] + "'"
      Qcursor.execute(sql)
      nR = 0
      for RBind in Qcursor:
         sql = "SELECT * FROM ruderer WHERE id = '" + RBind[2] + "'"
         cursor.execute(sql)
         dsatz = cursor.fetchone()
         nR = nR + 1
      if(nR == 1):
         rudererID = dsatz[0]
         STEXT = "'" + dsatz[1] + "' '" + dsatz[2] + "':  Boot #" + str(Boot[0]) + ": StNr. " + str(Boot[2]) + " in Rennen " + str(Boot[3])
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
   print(" - saved Index for ruderer = " + str(rudererID) )
   return
#========================================================================
def submit():
   global rudererID
   Gewicht  = myKG.get()
   if(len(Gewicht) > 0):
      if(rudererID == 0):
         print("submit " + str(Gewicht) + " kg - with no defined rower ?!")
      else:
         sql = "SELECT * FROM ruderer WHERE id = '" + rudererID + "' "
         Bcursor.execute(sql)
         Ruderer = Bcursor.fetchone()
         if(Ruderer[5] == 1):
            LGWstr = "Lgw."
         else:
            LGWstr = " ? "
         print("'" + Ruderer[1] + "' '" + Ruderer[2] + "' (" + LGWstr + ") = " + str(Gewicht) + " kg" )
         sql = "UPDATE ruderer SET gewicht = " + str(Gewicht) + " WHERE id = '" + rudererID + "' "
         cursor.execute(sql)
         connection.commit()
   return
#========================================================================
def abmelden():
   global bootNr
   if(bootNr == 0):
      print("submit abgemeldet  - with no defined boot ?!")
   else:
      sql = "SELECT * FROM boote WHERE id = '" + bootNr + "'"
      Bcursor.execute(sql)
      Boot = Bcursor.fetchone()
      if(Boot[11] == 1):
         abmeldung = "0"
      else:
         abmeldung = "1"
      # 
      sql = "UPDATE boote SET abgemeldet = " + abmeldung + " WHERE id = '" + bootNr + "'"
      cursor.execute(sql)
      connection.commit()
      #
      BTEXT = "#" + Boot[0] + ": StNr." + str(Boot[2]) + " in Rennen " + str(Boot[3]) 
      if(Boot[10] == 0):
         BTEXT = BTEXT + " - abgemeldet\n"
      else:
         BTEXT = BTEXT + " - ok\n"

      myStatus.configure(text=BTEXT) 
   return
#========================================================================
# Fontsize_huge = 16
# Fontsize_norm = 10
globalSize = "800x400"

Fontsize_huge = 30
Fontsize_norm = 20
globalSize = "1600x800"

win = tk.Tk()
win.title('my GUI for Langstrecke')
win.geometry(globalSize)
win.resizable(True, True)
win.columnconfigure(0, weight=1)
win.columnconfigure(1, weight=1)

#========================================================================

# Databases
mydata=tk.Label(win,text="Langstrecke " + str(LSglobal.Jahr) + " " + LSglobal.ZeitK,font=("Hack",Fontsize_huge),fg="blue")
mydata.grid(row=0,column=0)

mylabel=tk.Label(win,text="Kein Name\n kein Verein\nangegeben",font=("Hack",Fontsize_norm),fg="blue")
mylabel.grid(row=4,column=0,columnspan=3)

myStatus=tk.Label(win,text="?",font=("Hack",Fontsize_norm),fg="blue")
myStatus.grid(row=8,column=2,columnspan=1)
#

butGetEntry = tk.Button(win, text="Hole Einträge", font=("Hack",Fontsize_norm))
butGetEntry.bind('<Button-1>', searchID)
butGetEntry.grid(row=6,column=0,padx=10,pady=10,ipadx=20)

butSetKG = tk.Button(win, text="Setze Gewicht", command=submit, font=("Hack",Fontsize_norm))
butSetKG.grid(row=9,column=0,padx=10,pady=10,ipadx=20)

butAbmeldung = tk.Button(win, text="Abmeldung", command=abmelden, font=("Hack",Fontsize_norm))
butAbmeldung.grid(row=9,column=2,padx=10,pady=10,ipadx=20)

H_StartNr=tk.Label(win,text="Startnr",font=("Hack",Fontsize_norm),fg="blue")
H_StartNr.grid(row=2,column=0)
myNr = tk.Entry(win,font=("Hack",Fontsize_norm))
myNr.grid(row=3,column=0)
myNr.bind('<Return>', searchID)


H_Vorname=tk.Label(win,text="Vorname",font=("Hack",Fontsize_norm),fg="blue")
H_Vorname.grid(row=2,column=1)
myVorName = tk.Entry(win,font=("Hack",Fontsize_norm))
myVorName.grid(row=3,column=1)
myVorName.bind('<Return>', searchID)

H_Nachname=tk.Label(win,text="Nachname",font=("Hack",Fontsize_norm),fg="blue")
H_Nachname.grid(row=2,column=2)
myName = tk.Entry(win,font=("Hack",Fontsize_norm))
myName.grid(row=3,column=2)
myName.bind('<Return>', searchID)

H_KG=tk.Label(win,text="[kg]",font=("Hack",Fontsize_norm),fg="blue")
H_KG.grid(row=7,column=0)
myKG = tk.Entry(win,font=("Hack",Fontsize_norm))
myKG.grid(row=8,column=0)


logo = tk.PhotoImage(file="tiny_RVE_BRV_Flag.png")
w1 = tk.Label(win, image=logo).grid(row=0,column=1,columnspan=2,sticky="E")
#Wuerfel = tk.Label(win, image=imgBRV)
# Wuerfel.grid(row=0,column=1)

win.mainloop()
