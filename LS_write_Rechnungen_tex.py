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
cursor  = connection.cursor()
#


t1 = time.localtime()

TXT_1 = "\\documentclass[DIN,\n	fromalign=right,\n	fromurl=on,\n	fromemail=on,\n	fromrule=on,\n	fromlogo=on\n]{scrlttr2}\n\n\
% Handle internationalization\n\\usepackage[utf8]{inputenc}\n\\usepackage[T1]{fontenc}\n\
\\usepackage[right]{eurosym}\n\\usepackage[ngerman]{babel}\n\n% Handle graphics and tables\n\
\\usepackage{graphicx}\n\\usepackage{booktabs}\n\\usepackage{multirow}\n%\n% Handle hyperlinks\n\
\\usepackage[hidelinks]{hyperref}\n%\n% Use gray text for contact data\n\\usepackage{color}\n%\n"

# \\setkomavar{fromname}{RVE $\\cdot $ Habichtstra{\\ss}e 12 $\\cdot $ }\n\
# %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
#   RVE-Flag.png: 692 x 252 pixel
TXT_2 = "% Begin KOMA variables - make changes here\n\
\\newkomavar{gigdate}\n\
\\setkomavar{gigdate}{24. Oktober 2020}\n\
\\newkomavar{ausrichter}\n\
\\setkomavar{ausrichter}{Ruderverein Erlangen 1922 e.V.}\n\
\\setkomavar{fromname}{\\textbf{RVE}}\n\
\\setkomavar{fromaddress}{Habichtstra{\\ss}e 12\\\\91072 Erlangen}\n\
\\setkomavar{subject}{Rechnung:	Meldegebühr für die Herbst-Langstrecke }\n\
\\setkomavar{fromurl}{\\url{http://www.ruderverein-erlangen.de/}}\n\
\\setkomavar{fromemail}{\\href{mailto:langstrecke@ruderverein-erlangen.de}{langstrecke@ruderverein-erlangen.de}}\n\
\\setkomavar{fromlogo}{\\includegraphics[height=1.5cm,width=4.12cm]{RVE-Flag.png}}\n\
\\setkomavar{yourref}[Ihre Teilnahme an der Langstrecke am ]{\\usekomavar{gigdate}}\n\
\n\
% Invoice data\n\
\\newkomavar{price}\n\
\\setkomavar{price}{\\EUR{105,00}}\n\
\\newkomavar{invoicedate}\n\
\\setkomavar{invoicedate}{FÄLLIGKEITSDATUM}\n\
\\newkomavar{iban}\n\
\\setkomavar{iban}{DE00000000000000000000}\n\
\\newkomavar{bic}\n\
\\setkomavar{bic}{ABCADEXY000}\n\n"

# %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
TXT_3 = "% Begin header\n\
\\firsthead{\\null\\hfill \n\
 \\parbox[t][\\headheight][t]{10cm}{%\n\
  \\flushright\n\
  {\\small\\textsf{\n\
  \\color[gray]{.5}%\n\
  \\usekomavar{fromlogo}\\\\\n\
  \\textbf{\\usekomavar{ausrichter}}\\\\\n\
  \\usekomavar{fromaddress}\\\\\n\
  \\footnotesize{\\usekomavar{fromemail}}\\\\\n\
  \\footnotesize{\\usekomavar{fromurl}}\\\\\n\
  [\\baselineskip]\n\
  }}}%\n\
}\n\n\
% Begin footer\n\
\\firstfoot{\n\
\\parbox[t]{\\textwidth}{\n\
	\\footnotesize\n\
	\\color[gray]{.5}%\n\
	\\begin{tabular}[t]{l | r l | r l}\n\
		\\textbf{\\usekomavar{ausrichter}} &\n\
		\\multicolumn{2}{l|}{\\textbf{Kontakt}} &\n\
		\\multicolumn{2}{l}{\\textbf{Bankverbindung}}\\\\\n"
#		\\multirow{2}{*}{\\usekomavar{fromaddress}} &\n\
TXT_3 = TXT_3 + "		Habichtstra{\\ss}e 12 & \n\
		Web: & \\usekomavar{fromurl} &\n\
		IBAN: & \\usekomavar{iban}\\\\\n\
		91072 Erlangen &\n\
		E-Mail: & \\usekomavar{fromemail} &\n\
		BIC: & \\usekomavar{bic}\\\\\n\
	\\end{tabular}\n\
	}\n\
}\n\n"

TXT_4 = "% Begin main part - make changes here\n\
\\begin{document}\n\
\\begin{letter}{EMPFÄNGER\\\\\n\
	EMPFÄNGERADRESSE\\\\\n\
	12345 MUSTERSTADT}\n\n\
\\opening{Sehr geehrte Damen und Herren,}\n\n\
Für die Teilnahme an der Langstrecke vom Bayrischen Ruderverband und dem \\usekomavar{ausrichter} am \n\
\\usekomavar{gigdate} dürfen wir Ihnen folgendes berechnen:\n\n\\vspace{10pt}\n\n\
\\begin{tabular}{p{0.8\\textwidth} r}\n\
	\\toprule\n\
	\\textbf{Bezeichnung} & \\textbf{Betrag} \\\\\n\
	\\midrule\n%\n"

# \n\n\\closing{Mit rudersportlichen Grüßen,}\n\n\n\

# \\\\fancyhead[L]{\\\\textbf{" +  LSglobal.Name + "}\\\\\\\\ Meldeergebnis - Status " + str(t1.tm_hour) + ":" + str(t1.tm_min).rjust(2, '0') + "{\\\\small Uhr } - " \\

##############################################################################################################
# Vereine
#
Count_Boote   = 0
Count_Ruderer = 0
Count_Verein  = 0

EUR_total = 0
sql = "SELECT * FROM verein WHERE kurz != 'RVE' "

Meldegeld = 15
Bugnummer = 10

fehlende_Bugnummern = [ 63, 101]

Vcursor.execute(sql)
for Vsatz in Vcursor:
   # ------------------------------------- Kurzform mit Langform ersetzen
   Anzahl_Rennen = 0
   # ____________ Meldegeld
   EURO = 0
   # ____________ verspätete Abmeldung
   EURA = 0
   # ____________ verlorene Bugnummer
   EURB = 0
   #
   TXT = TXT_1 + TXT_2 + TXT_3 
   TXT = TXT +  "% Begin main part - make changes here\n\\begin{document}\n\\begin{letter}{"
   TXT = TXT +  Vsatz[0] + "\\\\\n"
   TXT = TXT +  Vsatz[2] + "\\\\\n"
   TXT = TXT +  Vsatz[3] + "}\n\n"
   TXT = TXT +  "\\opening{Sehr geehrte Damen und Herren,}\n\n"
   TXT = TXT +  "Für die Teilnahme an der Langstrecke vom Bayrischen Ruderverband und dem \\usekomavar{ausrichter} am \n"
   TXT = TXT +  "\\usekomavar{gigdate} dürfen wir Ihnen folgendes berechnen:\n\n\\vspace{10pt}\n\n"
   TXT = TXT +  "\\begin{tabular}{p{0.8\\textwidth} r}\n	\\toprule\n	\\textbf{Bezeichnung} & \\textbf{Betrag} \\\\\n"
   TXT = TXT +  "\\midrule\n%\n"
   #
   TXTA = "Verspätete Abmeldungen:\\small{"
   #
   TXTB = "Verlorene Bugnummern: \\newline\\small{"
   #
   #
   #___________________________ durchsuche Rennen
   sql = "SELECT * FROM rennen "
   Rcursor.execute(sql)
   for Rsatz in Rcursor:
      Rennen       = Rsatz[0]
      RennenString = Rsatz[1]
      #
      # print("checke Rennen " + str(Rennen) + " nach '" + Vsatz[1] + "'")
      #
      NoBoote = 0
      Ngray = 0
      Vrennen = 0
      sql = "SELECT * FROM boote  WHERE rennen = " + str(Rennen) + " ORDER BY startnummer, vereine "
      Bcursor.execute(sql)
      for Bsatz in Bcursor:
         StNr   = Bsatz[1]
         Verein = Bsatz[3]
         RudInd = Bsatz[4].split(',')
         Boot   = len(RudInd) - 2
         Abmeldung = Bsatz[12]
         # ===========================================================================================
         nPers   = 0
         for iR in range(0, (Boot)):         
            sql = "SELECT * FROM ruderer WHERE nummer = " + str(RudInd[iR + 1])
            Pcursor.execute(sql)
            Rd = Pcursor.fetchone()
            if(Rd[6] == Vsatz[1]):
               nPers = nPers + 1
            if(iR == 0):
               # Name = "\\textbf{" + Rd[0] + " } " + Rd[1]
               Name = "\\textsf{" + Rd[0] + "} " + Rd[1]
            else:
               # Name = Name + ", \\textbf{ " + Rd[0] + " } " + Rd[1]
               Name = Name + ", \\textsf{" + Rd[0] + "} " + Rd[1]
            #               
            if(Vsatz[1] != Rd[6]):
               Name = Name + " \\textcolor{gray}{\\footnotesize (" + Rd[6] + ")}"
            #
         #
         if(nPers > 0):
            if(Abmeldung == 0):
               if(Boot == 2 and nPers == 1):
                  EURO = EURO + Meldegeld/2
               else:
                  EURO = EURO + Meldegeld
               TXT = TXT + VTXT
               Anzahl_Rennen = Anzahl_Rennen + 1
               #
               if(Vrennen == 0):
                  Vrennen = 1
                  TXT = TXT +  "Meldegeld für folgende Mannschaften: \\newline\n	\\small\\textsf{\n"
                  TXT = TXT + "\n\\newline{\\textbf Rennen " + str(Rennen) + ": } " + RennenString + ": "
                  TXT = TXT + Name + " "
               else:
                  TXT = TXT + ", " + Name + " "
            # ===========================================================================================
            elif(Abmeldung > 1):  # verspätet abgemeldet
               if(Vrennen == 0):
                  Vrennen = 1
                  TXTA = TXTA + "\n\\newline{\\textbf Rennen " + str(Rennen) + ": } " + RennenString + ": "
                  TXTA = TXTA + "\\textcolor{red}{\\footnotesize " + Name + "}"
               else:
                  TXTA = TXTA + ", \\textcolor{red}{\\footnotesize" + Name + "}"
               #
               if(Boot == 2 and nPers == 1):
                  EURA = EURA + Meldegeld/2
               else:
                  EURA = EURA + Meldegeld
            #
            if(StNr in fehlende_Bugnummern):
               TXTB = TXTB +  "	(" + str(StNr) + ") " + Name + "\\newline"
               EURB = EURB + 10
         #_______________________________________________________________________________________
         #
      #_______________________________________________________________________________________

   #_______________________________________________________________________________________
   #if(Vrennen > 0):
   #   TXT = TXT + "%\\n\\\\end{tabular}\\\\\\\\\\n%\\n%\\n"
   # Korrektur des 'ß' - sz:
   if(EURO > 0):
      euronen = "%6.2f"% (EURO)
      TXT = TXT + "\n\\newline}	& " + euronen + " \\\\\n"
   if(EURA > 0):
      euronen = "%6.2f"% (EURA)
      TXT = TXT + TXTA +  "\\newline}	& " + euronen + " \\\\\n"
      TXT = TXT + "%\nKulanz für die verspäteten Abmeldungen: \\newline"
      TXT = TXT + "\\footnotesize\\textsf{Durch das Ansteigen der Corona-Pandemie wurden Abmeldungen bis zum 23.10. um 14 Uhr kostenfrei gestellt. \n"
      TXT = TXT + "Krankmeldungen waren danach auch noch kostenfrei.\\newline} & \color{red}{-" + euronen + "} \\\\\n%\n"
      TXT = TXT + "%\nAufwandspauschale der entstandenen Kosten & 20,00 \\\\\n%"
      EURA = 20
   #           
   if(EURB > 0):
      euronen = "%6.2f"% (EURB)
      TXT = TXT + TXTB +  "}\n	& " + euronen + " \\\\\n"
   TXT = TXT + "%\n\\midrule\n"
   #
   EUR_total = EUR_total + EURO + EURA + EURB
   #
   Count_Verein  = Count_Verein  + 1
   #                  #
   sql = "UPDATE verein SET rechnung = " + str(EURO + EURA + EURB) + " WHERE kurz = '" + Vsatz[1] + "' "
   cursor.execute(sql)
   connection.commit()
   #
   euronen = "%6.2f"% (EURO + EURA + EURB)
   TXT = TXT + "\\textbf{Gesamtbetrag (brutto):} & " + euronen + " \\\\\n"
   TXT = TXT + "\\bottomrule\n\\end{tabular}\n%\n\\vspace{10pt}\\\\\n%\n"
   TXT = TXT + "\\small{ Bitte beachten Sie, dass der \\usekomavar{ausrichter} nicht umsatzsteuerpflichtig\n" + \
         "ist. Daher ist im ausgewiesenen Betrag gemä{\\ss} \\textsection 19 UStG keine Umsatzsteuer enthalten. \\\\}\n"
   TXT = TXT + "Mit rudersportlichen Grüßen,\\\\\n\n\n"
   TXT = TXT + "Dr. Ulf Meerwald, 2. Vorsitzender \n"
   TXT = TXT + "\\end{letter}\n\\end{document}\n"

   TXT = TXT.replace('ß', '{\\ss}')
   fp = open("LaTeX/Rechnung_" + Vsatz[1] + ".tex","w")
   fp.write(TXT)
   fp.close()


#TXT = TXT_1 + TXT_2 + TXT_3 + TXT_4

##############################################################################################################
# Korrektur des 'ß' - sz:
#TXT = TXT.replace('ß', '{\\ss}')

# TXT = TXT + "\\n%======================\\n\\\\end{document}\\n"

print("Gemeldet haben " + str(Count_Verein) + " Vereine und wir stellen " + str(EUR_total) + " EUR in Rechnung")