#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Oct  8 15:48:30 2020

@author: ulf
"""
import time
import math
import os, sys, sqlite3

#==============================================================================
from termcolor import colored
#
# Text colors:
#   grey, red, green, yellow, blue, magenta, cyan, white
# Text highlights: 
#   on_grey,on_red, on_green, on_yellow, on_blue, on_magenta, on_cyan, on_white
# Attributes:
#   bold, dark, underline, blink, reverse, concealed
#========================================================================

import LSglobal

#========================================================================
# Verbindung zur Datenbank erzeugen
connection = sqlite3.connect( LSglobal.SQLiteFile )
#
# Datensatzcursor erzeugen
Rcursor  = connection.cursor()
Bcursor  = connection.cursor()
Qcursor  = connection.cursor()
Pcursor  = connection.cursor()
Vcursor  = connection.cursor()
#

# DRV Zeiten von 2020
#DRV_velo =    [ 0, 1, 2,  3,   4,    5,    6,    7,    8,    9,    10,   11,   12,   13,   14,   15,   16,   17,   18,   19, 
#            20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36 ]
DRV_velo =     [ 0, 0, 0, 4.85, 4.71, 5.43, 5.15, 5.12, 4.87, 4.61, 4.86, 4.74, 4.42, 4.28, 5.09, 5.09, 4.68, 4.39, 4.28, 0, 
            0, 0, 0, 0, 0,  0,  0,  0,  0,  0, 0,  0, 0, 0,  0, 0, 0 ]
# gesondert für Leichtgewichte:
# DRV_velo_Lgw = [ 0, 0, 0, 4.85, 4.71, 5.43, 5.15, 5.12, 4.87, 4.61, 4.86, 4.74, 4.42, 4.28, 4.98, 4.98, 4.52, 4.39, 4.28, 0, 
DRV_velo_Lgw = [ 0, 0, 0, 4.85, 4.71, 5.24, 5.02, 4.98, 4.87, 4.61, 4.86, 4.74, 4.42, 4.28, 4.98, 4.98, 4.52, 4.39, 4.28, 0,
            0, 0, 0, 0, 0,  0,  0,  0,  0,  0, 0,  0, 0, 0,  0, 0, 0 ]

Gewichtsklassen = dict([('SM A', 70.0), ('SM B', 70.0), ('JM A', 65.0), ('JM B', 62.5), ('SF A', 57.0), ('SF B', 57.0), ('JF A', 55.0), ('JF B', 52.5)])
Gewichtsgrenzen = dict([('SM A', 72.5), ('SM B', 72.5), ('JM A', 67.5), ('JM B', 65.0), ('SF A', 59.0), ('SF B', 59.0), ('JF A', 57.5), ('JF B', 55.0)])

DRV_velo = dict([('JF A 1x', 4.425), ('JF A 2x', 4.765), ('JF A 2-', 4.613), ('JF AL 1x', 4.283), ('JF AL 2x', 4.587), ('JF AL 2-', 4.465),
                 ('JF B 1x', 4.392), ('JF B 2x', 4.759), ('JF B 2-', 4.528), ('JF BL 1x', 4.283), ('JF BL 2x', 4.587), ('JF BL 2-', 4.465),
                 ('SF A 1x', 4.676), ('SF A 2x', 5.034), ('SF A 2-', 4.871), ('SF AL 1x', 4.525), ('SF AL 2x', 4.907), ('SF AL 2-', 4.748),
                 ('SF B 1x', 4.676), ('SF B 2x', 5.034), ('SF B 2-', 4.871), ('SF BL 1x', 4.525), ('SF BL 2x', 4.907), ('SF BL 2-', 4.748),
                 ('JM A 1x', 4.861), ('JM A 2x', 5.342), ('JM A 2-', 5.155), ('JM AL 1x', 4.738), ('JM AL 2x', 5.180), ('JM AL 2-', 5.023),
                 ('JM B 1x', 4.847), ('JM B 2x', 5.323), ('JM B 2-', 5.125), ('JM BL 1x', 4.710), ('JM BL 2x', 5.126), ('JM BL 2-', 5.000),
                 ('SM A 1x', 5.085), ('SM A 2x', 5.560), ('SM A 2-', 5.427), ('SM AL 1x', 4.975), ('SM AL 2x', 5.475), ('SM AL 2-', 5.236),
                 ('SM B 1x', 5.085), ('SM B 2x', 5.560), ('SM B 2-', 5.427), ('SM BL 1x', 4.975), ('SM BL 2x', 5.475), ('SM BL 2-', 5.236) ])

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

# --------------------------------------------------------------------------------------------------------------- Frühstarter
sql = "SELECT * FROM boote  WHERE rennen = 3 ORDER BY planstart"
Rcursor.execute(sql)
Rcursor.execute(sql)
Rsatz = Rcursor.fetchone()
TimeFrüh = Rsatz[3]


#================================================================================================
myTab1 = "7mm"

sql = "SELECT * FROM rennen WHERE strecke LIKE '%000 m' "
Rcursor.execute(sql)
for Rsatz in Rcursor:
   Rennen       = Rsatz[0]
   RennenString = Rsatz[1]
   # definiere das Durchschnittsgewicht / max.-Gewicht für das Rennen ---------
   Gewicht      = Rsatz[7]
   if(Gewicht > 0):
      maxGewicht = Gewicht
   else:
      Gewicht      = -Rsatz[7]
      if(Rsatz[3] == "1x"):
         maxGewicht = Gewicht
      elif(Gewicht == 57.0):
         maxGewicht = 59.0
      else:
         maxGewicht = Gewicht + 2.5
   #---------------------------------------------------------------------------
   Strecke      = int(Rsatz[4][0:4])
   #
   NoBoote = 0
   Platz = 0
   ZeitP = 0
   Ngray = 0
   sql = "SELECT * FROM boote  WHERE rennen = " + str(Rennen) + " AND abgemeldet = 0 AND zeit > 0 ORDER BY zeit, startnummer "
   Bcursor.execute(sql)
   for Bsatz in Bcursor:
      Boot    = Bsatz[0]
      StNr    = Bsatz[1]
      StrStNr = str(StNr)
      #
      # _______________________________________________________________________________________________________________
      if(NoBoote == 0):
         TXT = TXT + "\n% ============================= Rennen:  " + str(Rennen) + " __________ Start\n\\noindent\n"
         TXT = TXT + "\\begin{tabular}{|C{" + myTab1 + "}|C{1.0cm}|m{5.5cm}m{6.0cm}|C{2.0cm}|}\n\
         \\rowcolor{cMidGray} \\small Platz &  \\small Start- Nr. & \\multicolumn{2}{|c|}{\\color{white}\\parbox[1cm][2em][c]{115mm}{\
         \\textbf{\\Large Rennen " + str(Rennen) + "} \\hfill \\textbf{\\large " + RennenString + "} } } & \\small Zeiten\\\\\n"
         print("Rennen " + str(Rennen) + " : " + RennenString)
      #
      NoBoote = NoBoote + 1
      #--------------------------------------------------------------------- Ruderer -
      Vorname = ['-']
      Name    = ['-']
      JGNGstr = ['-']
      
      nPers   = 0
      bootGewicht = 0
      maxGewicht = 0
      MdA = 0
      #
      sql = "SELECT * FROM r2boot  WHERE bootNr = " + str(Boot) 
      Qcursor.execute(sql)
      iR = 0
      for RudInd in Qcursor:         
         sql = "SELECT * FROM ruderer WHERE nummer = " + str(RudInd[3])
         Pcursor.execute(sql)
         Rd = Pcursor.fetchone()
         #
         nPers = nPers + 1
         Vorname.insert(iR, Rd[1])
         Name.insert(iR,  Rd[2])
         JGNGstr.insert(iR, str(Rd[4]))
         MdA = MdA + LSglobal.RefJahr - Rd[4]
         if( iR == 0):
            Verein = "{" + Rd[7] +"}"
            VStr   = Rd[7]
         elif(iR>0 and VStr != Rd[7]):
            Verein =  Verein + "\\\\{" + Rd[7] + "}"
         # mit Jahrgang
         JGNGstr.insert(iR, str(Rd[4]))
         #
         if(Rd[5] == 1 and Rd[6] > 0.0):
            bootGewicht = bootGewicht + Rd[6]
            if(maxGewicht < Rd[6]):
               maxGewicht = Rd[6]
         #_____________________________________________
      # Check Altersklasse:
      MdA = MdA / nPers
      if(MdA > 22):
         ReDefNew = "S" + Rd[3] + " A"
      elif(MdA > 18):
         ReDefNew = "S" + Rd[3] + " B"
      elif(MdA > 16):
         ReDefNew = "J" + Rd[3] + " A"
      elif(MdA > 14):
         ReDefNew = "J" + Rd[3] + " B"
      elif(Rd[3] == "F"):
         ReDefNew = "Mädchen"
      else:
         ReDefNew = "Jungen"
      # Check Leichtgewicht:
      if(maxGewicht > 0):
         maxGewStr = ""
         bootGewicht = bootGewicht / nPers - LSglobal.Gewicht
         maxGewicht = maxGewicht - LSglobal.Gewicht
         if(nPers == 1):
            if(maxGewicht > Gewichtsgrenzen.get(ReDefNew)):
               print(StrStNr + " (" + ReDefNew + "): " +  colored(str(bootGewicht), 'red', attrs=['bold']) + " kg > " + str(Gewichtsgrenzen.get(ReDefNew)))
               maxGewStr = "^+"
               StrStNr = "$" + str(StNr) + "^{{Lgw^+}}$"
            else:
               StrStNr = "$" + str(StNr) + "^{{Lgw}}$"
         elif(maxGewicht > Gewichtsgrenzen.get(ReDefNew)):
            print(StrStNr + " (" + ReDefNew + "): " + str(maxGewicht) + " kg > " + str(Gewichtsgrenzen.get(ReDefNew)))
            maxGewStr = "^+"
            StrStNr = "$" + str(StNr) + "^{{Lgw^+}}$"
         elif(bootGewicht > Gewichtsklassen.get(ReDefNew) ):
            print(StrStNr + " (" + ReDefNew + "): " + str(maxGewicht) + " kg > " + str(Gewichtsklassen.get(ReDefNew)))
            maxGewStr = "^+"
         #
         if(Rsatz[7] > 1):
            StrStNr = "$" + str(StNr) + maxGewStr + "$"
         else:
            StrStNr = "$" + str(StNr) + "^{{Lgw" + maxGewStr + "}}$"
         ReDefNew = ReDefNew + "L"
      # Ergänze die Bootsklasse - Rsatz kann unbekannt sein
      ReDefNew = ReDefNew + " " + Rsatz[3]
      RefV = DRV_velo.get( ReDefNew )
      # ______________________________________________________________________________       # Zeiten: 
      # --------------------------------    Startzeit: Bsatz[6]
      Stime = Bsatz[4]
      StimH = math.floor(Stime/3600)
      StimM = math.floor(Stime/60 - StimH*60 )
      StZeit = "$" + str(StimH) + "$:$" + str(StimM).rjust(2, '0') + "$:" + str(Stime - 3600*StimH - 60*StimM).rjust(2, '0') + "}"
      # ---------------------------------------------------------------------
      #         3000 m     Bsatz[9]
      #         6000 m     Bsatz[10]
      #         Endzeit    Bsatz[11]
      # ---------------------------------------------------------------------
      Btime = Bsatz[9]   
      BtimM = math.floor(Btime/60)
      EZeit = "\\textbf{ " + str(BtimM) + ":" + str(Btime - 60*BtimM).rjust(2, '0') + "}" 
      if(RefV > 0):
         Percent = "%4.1f"% (100 * Strecke / Btime / RefV)
         EZeit = EZeit + "$^{\\textrm{ }" + Percent + "\\%}$"
      #
      Time6 = Bsatz[8]
      if(Time6 > 0):
         Time6m = math.floor(Time6/60)
         Time3  = Bsatz[7]
         Time3m = math.floor(Time3/60)
         #
         EZeit = EZeit + "\\\\ \\tiny{  " + str(Time3m) + ":" + str(Time3 - 60*Time3m).rjust(2, '0') \
         + ",  " + str(Time6m) + ":" + str(Time6 - 60*Time6m).rjust(2, '0') + "}"

      # ______________________________________________________________________________ 
      if(Ngray == 1):
         TXT = TXT + "\\rowcolor[gray]{.9}"
         Ngray = 0
      else:
         Ngray = 1
      #____________________________ Platz
      if(Btime > ZeitP):
         Platz = NoBoote
         ZeitP = Btime
      # 
      #
      if(nPers == 1):
         myHeight = str(nPers+1.5)
      else:
         myHeight = str(nPers+1)
      TXT = TXT + "\\parbox[" + myTab1 + "][" + myHeight  + "em][c]{10mm}{\\textbf{" + str(Platz) + "}} & "
      TXT = TXT + "\\parbox[1cm][" + str(nPers+1) + "em][c]{10mm}{\\textcolor{gray}{\\textbf{" + StrStNr + "}\\\\ \\small{" + StZeit + "}} & \
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
      TXT = TXT + "} & \\parbox[1cm][" + str(nPers+1) + "em][c]{20mm}{ " + EZeit + " }\\\\\myMidrule\n"
      # _______________________________________________________________________________________________________________
   #__________________________________________________________________________________________________________
   if(NoBoote > 0):
      TXT = TXT + "%\n\\end{tabular}\\\\[\\bigskipamount]\n%\n"
   #else:
   #   print("# " + str(Rennen) )


##############################################################################################################
TXT = TXT + "%\n%======================\n%\n"
TXT = TXT + "%\nDie \\textbf{2. Spalte} gibt die \\textbf{Startnummer} und die \\textbf{Startzeit} an. \n\
   Wenn ein Leichtgewicht im offenen Rennen startet erscheint ein \\textbf{ Lgw } hinter der Startnummer. \n\
   Wenn die Gewichtslimits nur knapp gerissen werden erscheint zusätzlich ein $^+$ dahinter.\\\\\n\
   \\\\\n%======================\n%\n"
TXT = TXT + "Die fette Zahl bei den \\textbf{ Zeiten } ist die finale Zeit über die jeweilige Wettkampfstrecke.\\\\\n\\\\\n\
Wenn \\textbf{ DRV-Prognosezahlen von 2020 } für die Altersklasse existieren, ist die Prozentzahl der Geschwindigkeit zum Referenzwert \
über die Standardstrecke (2000 m, bei Junioren B 1500m) dahinter zu sehen. \\\\\n\\\\\n\
Bei \\textbf{ 6000 m Streckenlänge } sind die Zeiten für die ersten 3000 m und für die zweiten 3000 m in der Zeile darunter angegeben.\n%\n"


TXT = TXT + "\n%======================\n\\end{document}\n"

##############################################################################################################

sql = "SELECT * FROM verein "
Vcursor.execute(sql)
for Vsatz in Vcursor:
   # ------------------------------------- Kurzform mit Langform ersetzen
   VereinStr = "{" + Vsatz[1] + "}"
   TXT = TXT.replace(VereinStr, ("{" + Vsatz[0] + "}"))
##############################################################################################################
connection.close()

fp = open("LaTeX/Endergebnis.tex","w")
fp.write(TXT)
fp.close()

