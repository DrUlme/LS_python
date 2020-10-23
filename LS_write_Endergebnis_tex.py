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
Vcursor  = connection.cursor()
#


t1 = time.localtime()

TXT = "\\documentclass[a4paper]{article}\n\\usepackage[ngerman]{babel}\n\\usepackage{colortbl,array,booktabs}\n\
\\usepackage[table]{xcolor}\n\\usepackage{tabularx}\n\\usepackage{fancyhdr}\n\\usepackage{graphicx}\n\
\\usepackage{multirow}\n\\usepackage[left=2.5cm, right=2.5cm, top=2.25cm, bottom=2.5cm]{geometry}\n\
% ___________________________________________________ Colors \n\
\\definecolor{cLightGray}{rgb}{.90,.90,.90}\n\\definecolor{cMidGray}{rgb}{.50,.50,.50}\n\n\
% ___________________________________________________ Header\n\\setlength{\\headsep}{30pt}\n\
\\fancyhead[L]{\\textbf{" +  LSglobal.Name + "}\\\\ Meldeergebnis - Status " + str(t1.tm_hour) + ":" + str(t1.tm_min).rjust(2, '0') + "{\\small Uhr } - " \
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

sql = "SELECT * FROM rennen "
Rcursor.execute(sql)
for Rsatz in Rcursor:
   Rennen       = Rsatz[0]
   RennenString = Rsatz[1]
   Gewicht      = Rsatz[7]
   Strecke      = Rsatz[4]
   #
   NoBoote = 0
   Ngray = 0
   sql = "SELECT * FROM boote  WHERE rennen = " + str(Rennen) + " abgemeldet = 0 ORDER BY zeit, startnummer "
   Bcursor.execute(sql)
   for Bsatz in Bcursor:
      Boot    = Bsatz[0]
      StNr    = Bsatz[1]
      StrStNr = str(StNr)
      VBoot   = Bsatz[3]
      RudInd  = Bsatz[4].split(',')
      # _______________________________________________________________________________________________________________
      if(NoBoote == 0):
         TXT = TXT + "\n% ============================= Rennen:  " + str(Rennen) + " __________ Start\n\\noindent\n"
         TXT = TXT + "\\begin{tabular}{|m{1.0cm}|m{1.0cm}|m{5.5cm}m{6.0cm}|C{2.0cm}|}\n\
         \\rowcolor{cMidGray} \\small Start- Nr. & \\multicolumn{3}{|c|}{\\color{white}\\parbox[1cm][2em][c]{135mm}{\
         \\textbf{\\Large Rennen " + str(Rennen) + "} \\hfill \\textbf{\\large " + RennenString + "} } } \\\\\n"
         print("Rennen " + str(Rennen) + " : " + RennenString)
      #
      NoBoote = NoBoote + 1
      #--------------------------------------------------------------------- Ruderer -
      Vorname = ['-']
      Name    = ['-']
      JGNGstr = ['-']
      
      nPers   = 0
      for iR in range(0, (len(RudInd) - 2)):         
         sql = "SELECT * FROM ruderer WHERE nummer = " + str(RudInd[iR + 1])
         Pcursor.execute(sql)
         Rd = Pcursor.fetchone()
         nPers = nPers + 1
         Vorname.insert(iR, Rd[0])
         Name.insert(iR,  Rd[1])
         JGNGstr.insert(iR, str(Rd[3]))
         if( iR == 0):
            Verein = "{" + Rd[6] +"}"
            VStr   = Rd[6]
         elif(iR>0 and VStr != Rd[6]):
            Verein =  Verein + "\\\\{" + Rd[6] + "}"
         #
         # ______________________________________________________________________________ Lgw ? 
         if(Rsatz[7] < 1 and Rd[4] == 1):
            if(Rd[5] < 0 or Rd[5] > (Gewicht + 5) ):
               JGNGstr.insert(iR, ("$" + str(Rd[3]) + "^{\\textrm{Lgw}}$"))
            else:
               JGNGstr.insert(iR, ("$" + str(Rd[3]) + "^{{Lgw}}$"))
         else:
            JGNGstr.insert(iR, str(Rd[3]))      
         # print(Name[0] + ", '" + Name[1] + "' =>" + Rd[2] )
      #---------------------------------------------------------------------
      # Zeiten: Startzeit: Bsatz[6]
      #         3000 m     Bsatz[9]
      #         6000 m     Bsatz[10]
      #         Endzeit    Bsatz[11]
      Btime = Bsatz[6]

         
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
   #else:
   #   print("# " + str(Rennen) )

##############################################################################################################

##############################################################################################################

TXT = TXT + "\n%======================\n\\end{document}\n"
fp = open("LaTeX/Endergebnis.tex","w")
fp.write(TXT)
fp.close()
