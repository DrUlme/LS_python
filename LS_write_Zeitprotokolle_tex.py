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
Pcursor  = connection.cursor()
Qcursor  = connection.cursor()
Vcursor  = connection.cursor()
#

t1 = time.localtime()

Meter = 00

if(Meter == 0):
   POSITION = "S T A R T"
elif(Meter == 3000):
   POSITION = "3000 m"
elif(Meter == 6000):
   POSITION = " Z I E L "

   
# --------------------------------------------------------------------------------------------------------------- Spätstarter
# sql = "SELECT * FROM rennen WHERE name LIKE 'Spätstarter%'"
sql = "SELECT wert FROM meta WHERE name = 'Spätstarter'"
Rcursor.execute(sql)
Rd = Rcursor.fetchone()
Spätstart = Rd[0]
Spät = int(Spätstart)

# --------------------------------------------------------------------------------------------------------------- Frühstarter
# hole Renn-Nummern für Früh und Spät-Starter
sql = "SELECT wert FROM meta WHERE name = 'Frühstarter'"
Rcursor.execute(sql)
Rd = Rcursor.fetchone()
Frühstart = Rd[0]
Früh = int(Frühstart)

# --------------------------------------------------------------------------------------------------------------- Header
TXT = "\\documentclass[a4paper]{article}\n\\usepackage[ngerman]{babel}\n\\usepackage{colortbl,array,booktabs}\n\
\\usepackage[table]{xcolor}\n\\usepackage{tabularx}\n\\usepackage{fancyhdr}\n\\usepackage{graphicx}\n\
\\usepackage{multirow}\n\\usepackage[left=2.5cm, right=2.5cm, top=2.25cm, bottom=2.5cm]{geometry}\n\
% ___________________________________________________ Colors \n\
\\definecolor{cLightGray}{rgb}{.90,.90,.90}\n\\definecolor{cMidGray}{rgb}{.50,.50,.50}\n\n\
% ___________________________________________________ Header\n\\setlength{\\headsep}{30pt}\n\
\\fancyhead[L]{\\textbf{" +  LSglobal.Name + " - - - - - - - - " + POSITION + "}\\\\ Meldeergebnis - Status " + str(t1.tm_hour) + ":" + str(t1.tm_min).rjust(2, '0') + "{\\small Uhr } - " \
   + str(t1.tm_mday) + "." + str(t1.tm_mon) + "." + str(t1.tm_year) + " }\n" + \
"\\fancyhead[R]{\\includegraphics[height=1.3cm]{RVE-Flag.png} }\n\\renewcommand{\\headrulewidth}{0pt}\n\n\
% ___________________________________________________ other definitions\n\
\\setlength{\\heavyrulewidth}{1pt}\n\\newcommand\\thickc[1][0.5pt]{\\vrule width #1}\n\
\\newcommand\myMidrule{\\specialrule{1.2pt}{0pt}{0pt}}\n\n\\renewcommand{\\arraystretch}{1.05}\n\
\\newcolumntype{L}[1]{>{\\raggedright\\arraybackslash}p{#1}}\n\\newcolumntype{C}[1]{>{\\centering\\arraybackslash}p{#1}}\
\n\\newcolumntype{R}[1]{>{\\raggedleft\\arraybackslash}p{#1}}\n\n\
% ___________________________________________________  S T A R T   O F   D O C U M E N T   _________________________\
%\n\\begin{document}\n\\pagestyle{fancy}\n\n"

#print(TXT)
TXT = TXT + "\n"

if(Meter == 6000):
   sql = "SELECT * FROM rennen WHERE strecke = '6000 m' "
else:
   sql = "SELECT * FROM rennen "
print(sql)
   
Rcursor.execute(sql)
for Rsatz in Rcursor:
   Rennen       = Rsatz[0]
   RennenString = Rsatz[1]
   #
   NoBoote = 0
   Ngray = 0
   if(Rennen == Früh or Rennen == Spät):
      sql = "SELECT * FROM boote  WHERE alternativ = " + str(Rennen) + " ORDER BY planstart, startnummer "
   else:
      sql = "SELECT * FROM boote  WHERE rennen = " + str(Rennen) + " ORDER BY planstart, startnummer "
   Bcursor.execute(sql)
   for Bsatz in Bcursor:
      Boot   = Bsatz[0]
      StNr   = Bsatz[2]
      VBoot  = Bsatz[4]
      LGW    = Bsatz[9]
      #
      Abmeldung = Bsatz[11]
      if(Abmeldung == 0):
         # _______________________________________________________________________________________________________________
         if(NoBoote == 0):
            TXT = TXT + "\n% ============================= Rennen:  " + str(Rennen) + " __________ Start\n\\noindent\n"
            TXT = TXT + "\\begin{tabular}{|m{1.0cm}|m{5.5cm}m{2.0cm}|C{2.0cm}|C{4.0cm}|}\n\
            \\rowcolor{cMidGray} \\small Start- Nr. & \\multicolumn{3}{|c|}{\\color{white}\\parbox[1cm][2em][c]{95mm}{\
            \\textbf{\\Large Rennen " + str(Rennen) + "} \\hfill \\textbf{\\large " + RennenString + "} } } & \
            \\parbox[1cm][2em][c]{40mm}{\\color{white}\\textbf{\\Large " + POSITION + "}} \\\\\n"
            print("Rennen " + str(Rennen) + " : " + RennenString)
         #
         NoBoote = NoBoote + 1
         #--------------------------------------------------------------------- Ruderer -
         Vorname = ['-']
         Name    = ['-']
         JGNGstr = ['-']
         #
         # ===========================================================================================
         sql = "SELECT * FROM r2boot  WHERE bootid = '" + str(Boot) + "'"
         Qcursor.execute(sql)
         nPers   = 0
         iR = 0
         #
         for RBind in Qcursor:
            # rudererNr = 3 (4. Eintrag)
            sql = "SELECT * FROM ruderer WHERE id = '" + RBind[2] + "'"
            Pcursor.execute(sql)
            Rd = Pcursor.fetchone()
            #
            nPers = nPers + 1
            Vorname.insert(iR, Rd[1])
            Name.insert(iR,  Rd[2])
            # JGNGstr.insert(iR, str(Rd[3]))
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
         #---------------------------------------------------------------------
         Btime = Bsatz[4]
         # print(Name[0] + " - " + str(len(RudInd)))
         #
         if(len(Btime) < 8):
            StrStNr = "tbd."
            StrZeit = "tbd."
         else:
            StrStNr = str(StNr)
            StrZeit = Btime
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
         TXT = TXT + "} & \\parbox[1cm][" + str(nPers+1) + "em][c]{20mm}{ \\small "
         TXT = TXT + Verein
         #for iR in range(0, nPers):
         #   TXT = TXT + "Ruder-Club Aschaffenburg v. 1898 e.V."
         #   # wenn mehrere und ungleich:
         #   if(nPers > (iR + 1)):
         #      TXT = TXT + "\\\\"
         #
         TXT = TXT + "} & \\parbox[1cm][" + str(nPers+1) + "em][c]{20mm}{ " + StrZeit + " } & \\\\\myMidrule\n"
      # _______________________________________________________________________________________________________________
   #__________________________________________________________________________________________________________
   if(NoBoote > 0):
      TXT = TXT + "%\n\\end{tabular}\\\\[\\bigskipamount]\n%\n"
   #else:
   #   print("# " + str(Rennen) )


##############################################################################################################
# Korrektur des 'ß' - sz:
TXT = TXT.replace('ß', '{\ss}')

TXT = TXT + "\n%======================\n\\end{document}\n"
fp = open("LaTeX/Zeitprotokoll.tex","w")
fp.write(TXT)
fp.close()

# print("Gemeldet haben " + str(Count_Verein) + " Vereine")