#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Oct  8 15:48:30 2020

@author: ulf
"""
import time
import math
import os, sys, sqlite3

import LSglobal

#========================================================================
# Verbindung zur Datenbank erzeugen
connection = sqlite3.connect( LSglobal.SQLiteFile )
#
# Datensatzcursor erzeugen
Rcursor  = connection.cursor()
Bcursor  = connection.cursor()
RBcursor = connection.cursor()
Pcursor  = connection.cursor()
Vcursor  = connection.cursor()
#
# Counter für angemeldete:
Count_Boote   = 0
Count_Ruderer = 0
Count_Athlets = 0

t1 = time.localtime()

TXT = "\\documentclass[a4paper]{article}\n\\usepackage[ngerman]{babel}\n\\usepackage{colortbl,array,booktabs}\n\
\\usepackage[table]{xcolor}\n\\usepackage{tabularx}\n\\usepackage{fancyhdr}\n\\usepackage{graphicx}\n\
\\usepackage{multirow}\n\\usepackage[left=2.5cm, right=2.5cm, top=2.25cm, bottom=2.5cm]{geometry}\n\
% ___________________________________________________ Colors \n\
\\definecolor{cLightGray}{rgb}{.90,.90,.90}\n\\definecolor{cMidGray}{rgb}{.50,.50,.50}\n\n\
% ___________________________________________________ Header\n\\setlength{\\headsep}{30pt}\n\
\\fancyhead[L]{\\textbf{" +  LSglobal.Name + "}\\\\ Meldeergebnis - Status " + str(t1.tm_hour) + ":" + str(t1.tm_min).rjust(2, '0') + "{\\small Uhr } - " \
   + str(t1.tm_mday) + "." + str(t1.tm_mon) + "." + str(t1.tm_year) + " }\n" + \
"\\fancyhead[R]{\\includegraphics[height=1.3cm]{RVE-BRV.png} }\n\\renewcommand{\\headrulewidth}{0pt}\n\n\
% ___________________________________________________ other definitions\n\
\\setlength{\\heavyrulewidth}{1pt}\n\\newcommand\\thickc[1][0.5pt]{\\vrule width #1}\n\
\\newcommand\myMidrule{\\specialrule{1.2pt}{0pt}{0pt}}\n\n\\renewcommand{\\arraystretch}{1.05}\n\
\\newcolumntype{L}[1]{>{\\raggedright\\arraybackslash}p{#1}}\n\\newcolumntype{C}[1]{>{\\centering\\arraybackslash}p{#1}}\
\n\\newcolumntype{R}[1]{>{\\raggedleft\\arraybackslash}p{#1}}\n\n\
% ___________________________________________________  S T A R T   O F   D O C U M E N T   _________________________\
%\n\\begin{document}\n\\pagestyle{fancy}\n\n"

#print(TXT)
TXT = TXT + "\n"
# --------------------------------------------------------------------------------------------------------------- Athletiktest
maxRennen = 100
sql = "SELECT * FROM rennen WHERE boot LIKE 'Athletik%' ORDER BY nummer"
Rcursor.execute(sql)
for Rd in Rcursor:
   if(maxRennen > Rd[0]):
      maxRennen = Rd[0]

LateStart = 0
# --------------------------------------------------------------------------------------------------------------- Spätstarter
sql = "SELECT * FROM rennen WHERE name LIKE 'Spätstarter%'"
Rcursor.execute(sql)
Rd = Rcursor.fetchone()
# print(Rd)
SpätR = Rd[0]
SpätT = Rd[5]
# suche letztes Boot vom Rennen davor (später als spät ...)
sql = "SELECT * FROM boote  WHERE rennen == " + str(SpätR-1) + " and abgemeldet = 0  ORDER BY planstart"
Rcursor.execute(sql)
SpätT = 0
for Rsatz in Rcursor:
   if(SpätT < Rsatz[3]):
      SpätT = Rsatz[3]
print("TimeSpät = " + str(SpätT))

# --------------------------------------------------------------------------------------------------------------- Frühstarter
sql = "SELECT * FROM boote  WHERE rennen = 3 and abgemeldet = 0  ORDER BY planstart"
Rcursor.execute(sql)
TimeFrüh = 0
for Rsatz in Rcursor:
   if(TimeFrüh == 0):
      TimeFrüh = Rsatz[3]
print("TimeFrüh = " + str(TimeFrüh))
# --------------------------------------------------------------------------------------------------------------- Großboot

sql = "SELECT * FROM rennen WHERE nummer = 1"
Rcursor.execute(sql)
for Rsatz in Rcursor:
   Rennen       = Rsatz[0]
   RennenString = Rsatz[1]
   BootTyp      = Rsatz[3]
   #
   NoBoote = 0
   Ngray = 0
   # hole die gemeldeten Boote für die Rennen
   sqlPart = " FROM boote  WHERE rennen = " + str(Rennen) + " and abgemeldet = 0  ORDER BY planstart, startnummer, nummer"
   #
   #----------------------------------- Anzahl der Boote und setzen von Split:
   sql = "SELECT count(*) " + sqlPart
   Bcursor.execute(sql)
   tmp = Bcursor.fetchone()
   myBoote = tmp[0]
   # print("Rennen #" + str(Rennen) + ": " + str(myBoote) + " Boote/Teilnehmer")
   if(myBoote > 29):
      mySplit = 25
   else:
      mySplit = -1
   #
   sql = "SELECT * " + sqlPart
   Bcursor.execute(sql)
   for Bsatz in Bcursor:
      Boot   = Bsatz[0]
      StNr   = Bsatz[1]
      # VBoot  = Bsatz[3]
      # sql = "SELECT rudererNr FROM r2boot  WHERE bootNr = " + str(Boot) 
      sql = "SELECT * FROM r2boot  WHERE bootNr = " + str(Boot) 
      RBcursor.execute(sql)
      #
      Count_Boote   = Count_Boote + 1
      
      Abmeldung = Bsatz[10]
      if(Abmeldung == 0):
         # _______________________________________________________________________________________________________________
         if(NoBoote == 0):
            TXT = TXT + "\n% ============================= Rennen:  " + str(Rennen) + " __________ Start\n\\noindent\n"
            TXT = TXT + "\\begin{tabular}{|m{1.0cm}|m{5.5cm}m{6.5cm}|C{2.0cm}|}\n\
            \\rowcolor{cMidGray} \\small Start- Nr. & \\multicolumn{3}{|c|}{\\color{white}\\parbox[1cm][2em][c]{135mm}{\
            \\textbf{\\Large Rennen " + str(Rennen) + "} \\hfill \\textbf{\\large " + RennenString + "} } } \\\\\n"
            print("Rennen " + str(Rennen) + " : " + RennenString + "      ( " + str(myBoote) + " )")
         elif(NoBoote == mySplit):
            TXT = TXT + "\n% ============================= split  %\n\\end{tabular}\\\\[\\bigskipamount]\nWeiter auf nächster Seite\\\\\n\\noindent\n"
            TXT = TXT + "\\begin{tabular}{|m{1.0cm}|m{5.5cm}m{6.5cm}|C{2.0cm}|}\n\
            \\rowcolor{cMidGray} \\small Start- Nr. & \\multicolumn{3}{|c|}{\\color{white}\\parbox[1cm][2em][c]{135mm}{\
            \\textbf{\\Large Rennen " + str(Rennen) + "} \\hfill \\textbf{\\large " + RennenString + "} Fortsetzung} } \\\\\n"
            mySplit = mySplit + 25
         #
         NoBoote = NoBoote + 1
         #--------------------------------------------------------------------- Ruderer -
         Vorname = ['-']
         Name    = ['-']
         JGNGstr = ['-']
         Platz   = ['-']
         
         nPers   = 0
         iR      = 0
         for RudInd in RBcursor: # for iR in range(0, (len(RudInd) - 2)):         
            # rudererNr: sql = "SELECT * FROM ruderer WHERE nummer = " + str(RudInd[0])
            Count_Ruderer = Count_Ruderer + 1
            sql = "SELECT * FROM ruderer WHERE nummer = " + str(RudInd[2])
            Pcursor.execute(sql)
            Rd = Pcursor.fetchone()
            nPers = nPers + 1
            Vorname.insert(iR, Rd[1])
            Name.insert(iR,  Rd[2])
            JGNGstr.insert(iR, str(Rd[4]))
            if( iR == 0):
               Verein = "{" + Rd[7] +"}"
               VStr   = Rd[7]
            elif(iR>0 and VStr != Rd[7]):
               Verein =  Verein + "\\\\{" + Rd[7] + "}"
            #
            if(RudInd[3] < 0):
               Platz.insert(iR,  "Stm.")
               print("Steuermann...?")
            else:
               Platz.insert(iR,  str(RudInd[3]))               
            #
            if(Rsatz[7] < 1 and Rd[5] == 1):
               # JGNGstr.insert(iR, ("$" + str(Rd[3]) + "^{\\textrm{Lgw}}$"))
               JGNGstr.insert(iR, ("$" + str(Rd[4]) + "^{{Lgw}}$"))
            else:
               JGNGstr.insert(iR, str(Rd[4]))      
            # print(Name[0] + ", '" + Name[1] + "' =>" + Rd[2] )
            iR += 1
         #---------------------------------------------------------------------
         Btime = Bsatz[3]
         # print(Name[0] + " - " + str(len(RudInd)))
         #
         if(Btime == 0):
            if(BootTyp == 'Athletik'):
               if(StNr == 0):
                  StrStNr = "tbd."
               else:
                  StrStNr = "{\\it " + str(StNr) + "}"
               StrZeit = "Halle 10:30"
            else:
               StrStNr = "tbd."
               StrZeit = "tbd."
         else:
            
            StrStNr = str(StNr)
            if(isinstance(Btime, str)):
               StrZeit = Btime
            else:
               BtimH = math.floor(Btime/3600)
               BtimM = math.floor(Btime/60 - BtimH*60 )
               #StrZeit = "$" + str(BtimH) + "$:$" + str(BtimM).rjust(2, '0') + "^{" + str(Btime - 3600*BtimH - 60*BtimM).rjust(2, '0') + "}$"
               StrZeit = "$" + str(BtimH) + "$:$" + str(BtimM).rjust(2, '0') + "$:\\small{\\textcolor{gray}{" + str(Btime - 3600*BtimH - 60*BtimM).rjust(2, '0') + "}}"
         #
         if(Ngray == 1):
            TXT = TXT + "\\rowcolor[gray]{.9}"
            Ngray = 0
         else:
            Ngray = 1
         #
         TXT = TXT + "\\parbox[1cm][" + str(nPers+1) + "em][c]{10mm}{\\textbf{" + StrStNr + "}} & \
         \\parbox[1cm][" + str(nPers+1) + "em][c]{55mm}{"
         #
         for iR in range(0, nPers):
            # Ruderer
            if(Platz[iR] == "Stm."):
               TXT = TXT + "\\textit{Stm. }" 
            TXT = TXT + "\\textbf{" + Vorname[iR] + " " + Name[iR] + "} {\\small(" + JGNGstr[iR] + ")} "
            if(nPers > 1 and iR < (nPers - 1)):
               # wenn mehrere:
               TXT = TXT + "\\\\"
         #
         # Verein
         TXT = TXT + "} & \\parbox[1cm][" + str(nPers+1) + "em][c]{65mm}{ \\small "
         TXT = TXT + Verein
         #for iR in range(0, nPers):
         #   TXT = TXT + "Ruder-Club Aschaffenburg v. 1898 e.V."
         #   # wenn mehrere und ungleich:
         #   if(nPers > (iR + 1)):
         #      TXT = TXT + "\\\\"
         #
         TXT = TXT + "} & \\parbox[1cm][" + str(nPers+1) + "em][c]{20mm}{ " + StrZeit + " }\\\\\myMidrule\n"
      # _______________________________________________________________________________________________________________
   #__________________________________________________________________________________________________________
   if(NoBoote > 0):
      TXT = TXT + "%\n\\end{tabular}\\\\[\\bigskipamount]\n%\n"
   #else:
   #   

# --------------------------------------------------------------------------------------------------------------- Frühstarter
if(TimeFrüh > 39600):    # 11*3600
   sql = "SELECT * FROM boote  WHERE planstart < " + str(TimeFrüh) + " and rennen > 1 and rennen < " + str(maxRennen) + " and abgemeldet = 0  ORDER BY planstart, startnummer, nummer"
   # print(sql)
   Bcursor.execute(sql)
   NoBoote = 0
   Ngray = 0
   for Bsatz in Bcursor:
      Boot   = Bsatz[0]
      StNr   = Bsatz[1]
      print(Boot)
      # VBoot  = Bsatz[3]
      sql = "SELECT rudererNr FROM r2boot  WHERE bootNr = " + str(Boot) 
      # print(sql)
      RBcursor.execute(sql)
      #
      Count_Boote   = Count_Boote + 1
      
      Abmeldung = Bsatz[10]
      if(Abmeldung == 0):
         # _______________________________________________________________________________________________________________
         if(NoBoote == 0):
            TXT = TXT + "\n% ============================= Rennen:  Frühstarter __________ Start\n\\noindent\n"
            TXT = TXT + "\\begin{tabular}{|m{1.0cm}|m{5.5cm}m{6.5cm}|C{2.0cm}|}\n\
            \\rowcolor{cMidGray} \\small Start- Nr. & \\multicolumn{3}{|c|}{\\color{white}\\parbox[1cm][2em][c]{135mm}{\
            \\textbf{\\Large Frühstarter} \\hfill \\textbf{\\large Rennen } } } \\\\\n"
         #
         NoBoote = NoBoote + 1
         #--------------------------------------------------------------------- Ruderer -
         Vorname = ['-']
         Name    = ['-']
         JGNGstr = ['-']
         
         nPers   = 0
         iR      = 0
         for RudInd in RBcursor: # for iR in range(0, (len(RudInd) - 2)):
            Count_Ruderer = Count_Ruderer + 1
            sql = "SELECT * FROM ruderer WHERE nummer = " + str(RudInd[0])
            Pcursor.execute(sql)
            Rd = Pcursor.fetchone()
            nPers = nPers + 1
            Vorname.insert(iR, Rd[1])
            Name.insert(iR,  Rd[2])
            JGNGstr.insert(iR, str(Rd[4]))
            if( iR == 0):
               Verein = "{" + Rd[7] +"}"
               VStr   = Rd[7]
            elif(iR>0 and VStr != Rd[7]):
               Verein =  Verein + "\\\\{" + Rd[7] + "}"
            #
            #
            if(Rsatz[7] < 1 and Rd[5] == 1):
               # JGNGstr.insert(iR, ("$" + str(Rd[3]) + "^{\\textrm{Lgw}}$"))
               JGNGstr.insert(iR, ("$" + str(Rd[4]) + "^{{Lgw}}$"))
            else:
               JGNGstr.insert(iR, str(Rd[4]))      
            # print(Name[0] + ", '" + Name[1] + "' =>" + Rd[2] )
            iR += 1
         #---------------------------------------------------------------------
         Btime = Bsatz[3]
         # print(Name[0] + " - " + str(len(RudInd)))
         #
         if(Btime == 0):
            StrStNr = "tbd."
            StrZeit = "tbd."
         else:
            StrStNr = str(StNr)
            if(isinstance(Btime, str)):
               StrZeit = Btime
            else:
               BtimH = math.floor(Btime/3600)
               BtimM = math.floor(Btime/60 - BtimH*60 )
               #StrZeit = "$" + str(BtimH) + "$:$" + str(BtimM).rjust(2, '0') + "^{" + str(Btime - 3600*BtimH - 60*BtimM).rjust(2, '0') + "}$"
               StrZeit = "$" + str(BtimH) + "$:$" + str(BtimM).rjust(2, '0') + "$:\\small{\\textcolor{gray}{" + str(Btime - 3600*BtimH - 60*BtimM).rjust(2, '0') + "}}"
         #
         if(Ngray == 1):
            TXT = TXT + "\\rowcolor[gray]{.9}"
            Ngray = 0
         else:
            Ngray = 1
         #
         TXT = TXT + "\\parbox[1cm][" + str(nPers+1) + "em][c]{10mm}{\\textbf{" + StrStNr + "}} & \
         \\parbox[1cm][" + str(nPers+1) + "em][c]{55mm}{"
         #
         for iR in range(0, nPers):
            # Ruderer
            TXT = TXT + "\\textbf{" + Vorname[iR] + " " + Name[iR] + "} {\\small(" + JGNGstr[iR] + ")} "
            if(nPers > 1 and iR < (nPers - 1)):
               # wenn mehrere:
               TXT = TXT + "\\\\"
         #
         # Verein
         TXT = TXT + "} & \\parbox[1cm][" + str(nPers+1) + "em][c]{60mm}{ \\small "
         TXT = TXT + Verein
         #for iR in range(0, nPers):
         #   TXT = TXT + "Ruder-Club Aschaffenburg v. 1898 e.V."
         #   # wenn mehrere und ungleich:
         #   if(nPers > (iR + 1)):
         #      TXT = TXT + "\\\\"
         #
         TXT = TXT + "} & \\parbox[1cm][" + str(nPers+1) + "em][c]{20mm}{ " + StrZeit + " }\\\\\myMidrule\n"
      # _______________________________________________________________________________________________________________
   #__________________________________________________________________________________________________________
   if(NoBoote > 0):
      TXT = TXT + "%\n\\end{tabular}\\\\[\\bigskipamount]\n%\n"


sql = "SELECT * FROM rennen WHERE nummer > 2"
Rcursor.execute(sql)
for Rsatz in Rcursor:
   Rennen       = Rsatz[0]
   RennenString = Rsatz[1]
   BootTyp      = Rsatz[3]
   #
   NoBoote = 0
   Ngray = 0
   # hole die gemeldeten Boote für die Rennen
   if(Rennen == SpätR and SpätT > 0):
      sqlPart = " FROM boote  WHERE rennen < " + str(SpätR) + " and planstart > " + str(SpätT) + " and abgemeldet = 0  ORDER BY planstart, startnummer, nummer"
   else:
      sqlPart = " FROM boote  WHERE rennen = " + str(Rennen) + " and abgemeldet = 0  ORDER BY planstart, startnummer, nummer"
      #
   #----------------------------------- Anzahl der Boote und setzen von Split:
   sql = "SELECT count(*) " + sqlPart
   Bcursor.execute(sql)
   tmp = Bcursor.fetchone()
   myBoote = tmp[0]
   # print("Rennen #" + str(Rennen) + ": " + str(myBoote) + " Boote/Teilnehmer")
   if(myBoote > 29):
      mySplit = 25
   else:
      mySplit = -1
   #
   sql = "SELECT * " + sqlPart
   Bcursor.execute(sql)
   for Bsatz in Bcursor:
      Boot   = Bsatz[0]
      StNr   = Bsatz[1]
      # VBoot  = Bsatz[3]
      sql = "SELECT rudererNr FROM r2boot  WHERE bootNr = " + str(Boot) 
      RBcursor.execute(sql)
      #
      if(BootTyp != "Athletik"):
         Count_Boote   = Count_Boote + 1
      #
      Abmeldung = Bsatz[10]
      if(Abmeldung == 0):
         # _______________________________________________________________________________________________________________
         if(NoBoote == 0):
            TXT = TXT + "\n% ============================= Rennen:  " + str(Rennen) + " __________ Start\n\\noindent\n"
            TXT = TXT + "\\begin{tabular}{|m{1.0cm}|m{5.5cm}m{6.5cm}|C{2.0cm}|}\n\
            \\rowcolor{cMidGray} \\small Start- Nr. & \\multicolumn{3}{|c|}{\\color{white}\\parbox[1cm][2em][c]{135mm}{\
            \\textbf{\\Large Rennen " + str(Rennen) + "} \\hfill \\textbf{\\large " + RennenString + "} } } \\\\\n"
            print("Rennen " + str(Rennen) + " : " + RennenString + "      ( " + str(myBoote) + " )")
         elif(NoBoote == mySplit):
            TXT = TXT + "\n% ============================= split  %\n\\end{tabular}\\\\[\\bigskipamount]\nWeiter auf nächster Seite\\\\\n\\noindent\n"
            TXT = TXT + "\\begin{tabular}{|m{1.0cm}|m{5.5cm}m{6.5cm}|C{2.0cm}|}\n\
            \\rowcolor{cMidGray} \\small Start- Nr. & \\multicolumn{3}{|c|}{\\color{white}\\parbox[1cm][2em][c]{135mm}{\
            \\textbf{\\Large Rennen " + str(Rennen) + "} \\hfill \\textbf{\\large " + RennenString + "} Fortsetzung} } \\\\\n"
            mySplit = mySplit + 25
         #
         NoBoote = NoBoote + 1
         #--------------------------------------------------------------------- Ruderer -
         Vorname = ['-']
         Name    = ['-']
         JGNGstr = ['-']
         
         nPers   = 0
         iR      = 0
         for RudInd in RBcursor: # for iR in range(0, (len(RudInd) - 2)):         
            sql = "SELECT * FROM ruderer WHERE nummer = " + str(RudInd[0])
            Pcursor.execute(sql)
            Rd = Pcursor.fetchone()
            nPers = nPers + 1
            Vorname.insert(iR, Rd[1])
            Name.insert(iR,  Rd[2])
            JGNGstr.insert(iR, str(Rd[4]))
            if( iR == 0):
               Verein = "{" + Rd[7] +"}"
               VStr   = Rd[7]
            elif(iR>0 and VStr != Rd[7]):
               Verein =  Verein + "\\\\{" + Rd[7] + "}"
            #
            #
            if(Rsatz[7] < 1 and Rd[5] == 1):
               # JGNGstr.insert(iR, ("$" + str(Rd[3]) + "^{\\textrm{Lgw}}$"))
               JGNGstr.insert(iR, ("$" + str(Rd[4]) + "^{{Lgw}}$"))
            else:
               JGNGstr.insert(iR, str(Rd[4]))      
            # print(Name[0] + ", '" + Name[1] + "' =>" + Rd[2] )
            iR += 1
            #
            if(BootTyp != "Athletik"):
               Count_Ruderer = Count_Ruderer + 1
            else:
               Count_Athlets = Count_Athlets + 1
         #-----
         
         #---------------------------------------------------------------------
         Btime = Bsatz[3]
         # print(Name[0] + " - " + str(len(RudInd)))
         #
         if(Btime == 0):
            if(BootTyp == 'Athletik'):
               if(StNr == 0):
                  StrStNr = "tbd."
               else:
                  StrStNr = "{\\it " + str(StNr) + "}"
               StrZeit = "Halle RVE 9:30"
            else:
               StrStNr = "tbd."
               StrZeit = "tbd."
         else:
            
            if(TimeFrüh > Btime):
               # StrStNr = "$" + str(StNr) + "^{{Früh}}$"
               StrStNr = str(StNr) + "{\it F}"
            elif(Rennen < SpätR and Btime > SpätT):
               # StrStNr = "$" + str(StNr) + "^{{Früh}}$"
               StrStNr = str(StNr) + "{\it S}"
            else:
               StrStNr = str(StNr)
            if(isinstance(Btime, str)):
               StrZeit = Btime
            else:
               BtimH = math.floor(Btime/3600)
               BtimM = math.floor(Btime/60 - BtimH*60 )
               #StrZeit = "$" + str(BtimH) + "$:$" + str(BtimM).rjust(2, '0') + "^{" + str(Btime - 3600*BtimH - 60*BtimM).rjust(2, '0') + "}$"
               StrZeit = "$" + str(BtimH) + "$:$" + str(BtimM).rjust(2, '0') + "$:\\small{\\textcolor{gray}{" + str(Btime - 3600*BtimH - 60*BtimM).rjust(2, '0') + "}}"
         #
         if(Ngray == 1):
            TXT = TXT + "\\rowcolor[gray]{.9}"
            Ngray = 0
         else:
            Ngray = 1
         #
         TXT = TXT + "\\parbox[1cm][" + str(nPers+1) + "em][c]{10mm}{\\textbf{" + StrStNr + "}} & \
         \\parbox[1cm][" + str(nPers+1) + "em][c]{55mm}{"
         #
         for iR in range(0, nPers):
            # Ruderer
            TXT = TXT + "\\textbf{" + Vorname[iR] + " " + Name[iR] + "} {\\small(" + JGNGstr[iR] + ")} "
            if(nPers > 1 and iR < (nPers - 1)):
               # wenn mehrere:
               TXT = TXT + "\\\\"
         #
         # Verein
         TXT = TXT + "} & \\parbox[1cm][" + str(nPers+1) + "em][c]{65mm}{ \\small "
         TXT = TXT + Verein
         #for iR in range(0, nPers):
         #   TXT = TXT + "Ruder-Club Aschaffenburg v. 1898 e.V."
         #   # wenn mehrere und ungleich:
         #   if(nPers > (iR + 1)):
         #      TXT = TXT + "\\\\"
         #
         TXT = TXT + "} & \\parbox[1cm][" + str(nPers+1) + "em][c]{20mm}{ " + StrZeit + " }\\\\\myMidrule\n"
      # _______________________________________________________________________________________________________________
   #__________________________________________________________________________________________________________
   if(NoBoote > 0):
      TXT = TXT + "%\n\\end{tabular}\\\\[\\bigskipamount]\n%\n"
   #else:
   #   print("# " + str(Rennen) )

##############################################################################################################
# Vereine
#
# Count_Boote   = 0
# Count_Ruderer = 0
Count_Verein  = 0
# Athletiktest = "10:30 in Halle - Dorfstraße 21"
Athletiktest = "9:00 -9:30 Halle des RVE " # "am Ruderverein Erlangen"

sql = "SELECT * FROM verein "
Vcursor.execute(sql)
for Vsatz in Vcursor:
   # ------------------------------------- Kurzform mit Langform ersetzen
   Anzahl_Rennen = 0
   VereinStr = "{" + Vsatz[1] +"}"
   TXT = TXT.replace(VereinStr, Vsatz[0])
   
   VTXT = "\n%================================\n\\newpage\n{\\huge "
   VTXT = VTXT + Vsatz[0]
   VTXT = VTXT + " }\\\\\n"
   #___________________________ durchsuche Rennen
   sql = "SELECT * FROM rennen "
   Rcursor.execute(sql)
   for Rsatz in Rcursor:
      Rennen       = Rsatz[0]
      RennenString = Rsatz[1]
      BootTyp      = Rsatz[3]
      #
      # print("checke Rennen " + str(Rennen) + " nach '" + Vsatz[1] + "'")
      #
      NoBoote = 0
      Ngray = 0
      Vrennen = 0
      # alle: 
      sql = "SELECT * FROM boote  WHERE rennen = " + str(Rennen) + " ORDER BY planstart, startnummer, nummer "
      # sql = "SELECT * FROM boote  WHERE rennen = " + str(Rennen) + " AND abgemeldet = 0 ORDER BY planstart, startnummer, nummer "
      Bcursor.execute(sql)
      for Bsatz in Bcursor:
         Boot   = Bsatz[0]
         StNr   = Bsatz[1]
         Abmeldung = Bsatz[10]
         if(Abmeldung == 0):
            # ===========================================================================================
            sql = "SELECT * FROM r2boot  WHERE bootNr = " + str(Boot) 
            RBcursor.execute(sql)
            nPers   = 0
            iR = 0
            for RBind in RBcursor:
               # rudererNr = 3 (4. Eintrag)
               sql = "SELECT * FROM ruderer WHERE nummer = " + str(RBind[2])
               Pcursor.execute(sql)
               Rd = Pcursor.fetchone()
               # gibt nur diesen einen Eintrag pro Ruderer
               if(Rd[7] == Vsatz[1]):
                  nPers = 1
               if(iR == 0):
                  Name = "\\textbf{" + Rd[1] + " } " + Rd[2]
               else:
                  Name = Name + ", \\textbf{ " + Rd[1] + " } " + Rd[2]
               #               
               if(Vsatz[1] != Rd[7]):
                  Name = Name + " \\textcolor{gray}{\\small (" + Rd[7] + ")}"
               #
               iR += 1
            #
            if(nPers > 0):
               if(Anzahl_Rennen == 0):
                  TXT = TXT + VTXT
                  Anzahl_Rennen = Anzahl_Rennen + 1
                  Count_Verein  = Count_Verein  + 1
               #
               if(Vrennen == 0):
                  Vrennen = 1
                  TXT = TXT + "\n{\\textbf Rennen " + str(Rennen) + ": } " + RennenString + "\\\\\n%"
                  TXT = TXT + "\n\\begin{tabular}{m{1.0cm}m{8cm}m{6.0cm}}\n"
                  # TXT = TXT + "\n\\begin{tabular}{m{1.0cm}m{8cm}m{2.0cm}}\n"
               #_______________________________________________________________________________________
               Btime = Bsatz[3]
               #
               if(Btime == 0):
                  if(BootTyp == 'Athletik'):
                     if(StNr == 0):
                        StrStNr = "tbd."
                     else:
                        StrStNr = "{\\it " + str(StNr) + "}"
                     StrZeit = Athletiktest
                  else:
                     StrStNr = "tbd."
                     StrZeit = "tbd."
               else:
                  if(TimeFrüh > Btime):
                     # StrStNr = "$" + str(StNr) + "^{{Früh}}$"
                     StrStNr = str(StNr) + "{\it F}"
                  else:
                     StrStNr = str(StNr)
                  if(isinstance(Btime, str)):
                     StrZeit = Btime
                  else:
                     BtimH = math.floor(Btime/3600)
                     BtimM = math.floor(Btime/60 - BtimH*60 )
                     StrZeit = "$" + str(BtimH) + "$:$" + str(BtimM).rjust(2, '0') + "$:\\small{\\textcolor{gray}{" + str(Btime - 3600*BtimH - 60*BtimM).rjust(2, '0') + "}}"
                     # "^{" + str(Btime - 3600*BtimH - 60*BtimM).rjust(2, '0') + "}$"
               #
               TXT = TXT + " " + StrStNr + " & " + Name + " & " + StrZeit + " \\\\\n"
            # ===========================================================================================
         # elif(Abmeldung > 1):  # verspätet abgemeldet
         #_______________________________________________________________________________________
         #
      #_______________________________________________________________________________________

      if(Vrennen > 0):
         TXT = TXT + "%\n\\end{tabular}\\\\\n%\n%\n"
 

#TXT = TXT + "\\\\{\\bf\\large Rennen " + str(Rennen) + ": " +  RennenString + " } \\\\\n"
# # \null\hfill{\bf Startzeit 11:02\\}
# Vorname = ['Vor', 'und', 'zu']
# Name    = ['Name', 'ist', 'unwichtig']
# #TXT = TXT + "{\\bf " + str(StNo) + " : " + Vorname[0] + " } " + Name[0] + "\\\\\n" 

##############################################################################################################
# Korrektur des 'ß' - sz:
TXT = TXT.replace('ß', '{\ss}')

TXT = TXT + "\n%======================\n\\end{document}\n"
fp = open("LaTeX/Meldungen.tex","w")
fp.write(TXT)
fp.close()

connection.close()

print("-----------\nGemeldet haben " + str(Count_Verein) + " Vereine")
print(str(Count_Ruderer) + " Ruderer in " + str(Count_Boote) + " Booten")
if(Count_Athlets > 0):
   print("Zusätzlich sind " + str(Count_Athlets) + " Kinder für die Lauf und Athletikwettbewerbe gemeldet.")

