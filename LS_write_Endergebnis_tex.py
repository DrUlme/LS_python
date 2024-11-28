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
DRV_Prognose_Str = "9/2022" # "12.9.2022" - [m/s] aus Zeiten neu gewonnen !!!

Gewichtsklassen = dict([('SM A', 70.0), ('SM B', 70.0), ('JM A', 65.0), ('JM B', 62.5), ('SF A', 57.0), ('SF B', 57.0), ('JF A', 55.0), ('JF B', 52.5)])
Gewichtsgrenzen = dict([('SM A', 72.5), ('SM B', 72.5), ('JM A', 67.5), ('JM B', 65.0), ('SF A', 59.0), ('SF B', 59.0), ('JF A', 57.5), ('JF B', 55.0)])

DRV_velo = dict([('JF A 1x', 4.430), ('JF A 2x', 4.765), ('JF A 2-', 4.613), ('JF AL 1x', 4.329), ('JF AL 2x', 4.696), ('JF AL 2-', 4.465),
                 ('JF B 1x', 4.372), ('JF B 2x', 4.704), ('JF B 2-', 4.552), ('JF BL 1x', 4.283), ('JF BL 2x', 4.587), ('JF BL 2-', 4.465),
                 ('SF A 1x', 4.676), ('SF A 2x', 5.034), ('SF A 2-', 4.909), ('SF AL 1x', 4.525), ('SF AL 2x', 4.984), ('SF AL 2-', 4.695),
                 ('SF B 1x', 4.676), ('SF B 2x', 5.034), ('SF B 2-', 4.909), ('SF BL 1x', 4.525), ('SF BL 2x', 4.984), ('SF BL 2-', 4.695),
                 ('JM A 1x', 4.933), ('JM A 2x', 5.342), ('JM A 2-', 5.181), ('JM AL 1x', 4.773), ('JM AL 2x', 5.240), ('JM AL 2-', 5.023),
                 ('JM B 1x', 4.869), ('JM B 2x', 5.272), ('JM B 2-', 5.114), ('JM BL 1x', 4.710), ('JM BL 2x', 5.126), ('JM BL 2-', 5.000),
                 ('SM A 1x', 5.119), ('SM A 2x', 5.560), ('SM A 2-', 5.427), ('SM AL 1x', 5.000), ('SM AL 2x', 5.475), ('SM AL 2-', 5.263),
                 ('SM B 1x', 5.119), ('SM B 2x', 5.560), ('SM B 2-', 5.427), ('SM BL 1x', 5.000), ('SM BL 2x', 5.475), ('SM BL 2-', 5.263) ])
ohne_aktuelle_Daten=[ "JF AL2- ohne", "JF/JM B Lg" ]

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
   RennenString = Rsatz[2]
   Bootstyp     = Rsatz[5]
   # definiere das Durchschnittsgewicht / max.-Gewicht für das Rennen ---------
   Gewicht      = Rsatz[8]
   if(Gewicht > 0):
      maxGewicht = Gewicht
   else:
      Gewicht      = -Rsatz[8]
      if(Rsatz[5] == "1x"):
         maxGewicht = Gewicht
      elif(Gewicht == 57.0):
         maxGewicht = 59.0
      else:
         maxGewicht = Gewicht + 2.5
   #---------------------------------------------------------------------------
   Strecke      = int(Rsatz[6][0:4])
   #
   NoBoote = 0
   Platz = 0
   ZeitP = 0
   Ngray = 0
   sql = "SELECT * FROM boote  WHERE rennen = " + str(Rennen) + " AND abgemeldet = 0 AND zeit > 0 ORDER BY zeit, startnummer "
   Bcursor.execute(sql)
   for Bsatz in Bcursor:
      Boot    = Bsatz[0]
      StNr    = Bsatz[2]
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
      sql = "SELECT * FROM r2boot  WHERE bootid = '" + Boot + "'" 
      Qcursor.execute(sql)
      iR = 0
      for RudInd in Qcursor:         
         sql = "SELECT * FROM ruderer WHERE id = '" + RudInd[2] + "'"
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
         iR = iR + 1
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
      #__________________________________ ToDo: Großboot
      #if(StNr == 1):
      #   ReDefNew = "SF B"
      #   Bootstyp = "2x"
      #elif(StNr == 2):
      #   ReDefNew = "JM B"
      #   Bootstyp = "2x"
      
      # ' hole Gewichtslimit pro Rennen'
      RennenMaxGewicht = Gewichtsgrenzen.get(ReDefNew)
      RennenGewichtsKlasse = Gewichtsklassen.get(ReDefNew)
      #
      # Check Leichtgewicht:
      if(RennenMaxGewicht is not None and RennenGewichtsKlasse is not None and maxGewicht > 0):
         maxGewStr = ""
         bootGewicht = bootGewicht / nPers - LSglobal.Gewicht
         maxGewicht = maxGewicht - LSglobal.Gewicht
         
         if(nPers == 1):
            if(maxGewicht > RennenMaxGewicht):
               print(StrStNr + " (" + ReDefNew + "): " +  colored(str(bootGewicht), 'red', attrs=['bold']) + " kg > " + str(RennenMaxGewicht))
               maxGewStr = "^+"
               StrStNr = "$" + str(StNr) + "^{{Lgw^+}}$"
            else:
               StrStNr = "$" + str(StNr) + "^{{Lgw}}$"
         elif(maxGewicht > RennenMaxGewicht):
            print(StrStNr + " (" + ReDefNew + "): " + str(maxGewicht) + " kg > " + str(RennenMaxGewicht))
            maxGewStr = "^+"
            StrStNr = "$" + str(StNr) + "^{{Lgw^+}}$"
         elif(bootGewicht > RennenGewichtsKlasse ):
            print(StrStNr + " (" + ReDefNew + "): " + str(maxGewicht) + " kg > " + str(RennenGewichtsKlasse))
            maxGewStr = "^+"
         #
         if(Rsatz[7] > 1):
            StrStNr = "$" + str(StNr) + maxGewStr + "$"
         else:
            StrStNr = "$" + str(StNr) + "^{{Lgw" + maxGewStr + "}}$"
         ReDefNew = ReDefNew + "L"
      # Ergänze die Bootsklasse - Rsatz kann unbekannt sein
      ReDefNew = ReDefNew + " " + Bootstyp
      RefV = DRV_velo.get( ReDefNew )
      #if(StNr == 4):
      #   print(str(StNr) + ": '" + ReDefNew + "' -> " + str(RefV) )
      # ______________________________________________________________________________       # Zeiten: 
      # --------------------------------    Startzeit: Bsatz[6]
      # Stime = Bsatz[4]
      # StimH = math.floor(Stime/3600)
      # StimM = math.floor(Stime/60 - StimH*60 )
      # StZeit = "$" + str(StimH) + "$:$" + str(StimM).rjust(2, '0') + "$:" + str(Stime - 3600*StimH - 60*StimM).rjust(2, '0') + "}"
      StZeit = Bsatz[5] + "}"
      # ---------------------------------------------------------------------
      #         3000 m     Bsatz[9]
      #         6000 m     Bsatz[10]
      #         Endzeit    Bsatz[11]
      # ---------------------------------------------------------------------
      # Btime = Bsatz[9]   
      # BtimM = math.floor(Btime/60)
      # EZeit = "\\textbf{ " + str(BtimM) + ":" + str(Btime - 60*BtimM).rjust(2, '0') + "}" 
      EZeit = "\\textbf{ " + Bsatz[10] + "}"
      # Zeit in sec.
      Btime = 60 * float(Bsatz[10][0:2]) + float(Bsatz[10][3:5])
      #
      if(RefV is not None and RefV > 0):
         Percent = "%4.1f"% (100 * Strecke / Btime / RefV)
         EZeit = EZeit + "$^{\\textrm{ }" + Percent + "\\%}$"
         # check in einem Rennen
         if(Rennen == 4):
            print("#" + StrStNr + ": " + ReDefNew + " => " + str(RefV) + "[m/s], " + str(Btime) + "s = " + Percent )
      #
      Time6 = Bsatz[9]
      if(len(Time6) == 5):
         EZeit = EZeit + "\\\\ \\tiny{  " + Bsatz[8] + ",  " + Bsatz[9] + "}"
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
Wenn \\textbf{ DRV-Prognosezahlen von " + DRV_Prognose_Str + " } für die Altersklasse existieren, ist die Prozentzahl der Geschwindigkeit zum Referenzwert \
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

