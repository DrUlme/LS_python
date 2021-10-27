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
cursor  = connection.cursor()
#


t1 = time.localtime()

TXT_1 = "\\documentclass[DIN,\n	fromalign=right,\n	fromurl=on,\n	fromemail=on,\n	fromrule=on,\n	fromlogo=on\n]{scrlttr2}\n\n\
% Handle internationalization\n\\usepackage[utf8]{inputenc}\n\\usepackage[T1]{fontenc}\n\
\\usepackage[right]{eurosym}\n\\usepackage[ngerman]{babel}\n\n% Handle graphics and tables\n\
\\usepackage{graphicx}\n\\usepackage{booktabs}\n\\usepackage{multirow}\n%\n% Handle hyperlinks\n\
\\usepackage[hidelinks]{hyperref}\n%\n% Use gray text for contact data\n\\usepackage{color}\n%\n"
#
# mehr Platz nach unten:
TXT_1 = TXT_1 + "%\n%\n\\setlength{\\textheight}{26cm}\n\\setlength{\\footskip}{0mm}\n\\setlength{\\footheight}{0mm}\n%\n"
#
# \\setkomavar{fromname}{RVE $\\cdot $ Habichtstra{\\ss}e 12 $\\cdot $ }\n\
#
# %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
#   RVE-Flag.png: 692 x 252 pixel
TXT_2 = "% Begin KOMA variables - make changes here\n\
\\newkomavar{gigdate}\n\
\\setkomavar{gigdate}{" + LSglobal.Datum + str(LSglobal.Jahr) + "}\n\
\\newkomavar{ausrichter}\n\
\\setkomavar{ausrichter}{Ruderverein Erlangen e.V.}\n\
\\setkomavar{fromname}{\\textbf{RVE}}\n\
\\setkomavar{fromaddress}{Habichtstra{\\ss}e 12\\\\91056 Erlangen}\n\
\\setkomavar{subject}{Rechnung:	Meldegebühr für die Herbst-Langstrecke }\n\
\\setkomavar{fromurl}{\\url{http://www.ruderverein-erlangen.de}}\n\
\\setkomavar{fromemail}{\\href{mailto:langstrecke@ruderverein-erlangen.de}{langstrecke@ruderverein-erlangen.de}}\n\
\\setkomavar{fromlogo}{\\includegraphics[height=1.76cm,width=2cm]{RVE-Flag.png}}\n\
\\setkomavar{yourref}[Ihre Teilnahme an der Langstrecke am ]{\\usekomavar{gigdate}}\n\
\n\
% Invoice data\n\
\\newkomavar{price}\n\
\\setkomavar{price}{\\EUR{105,00}}\n\
\\newkomavar{invoicedate}\n\
\\setkomavar{invoicedate}{FÄLLIGKEITSDATUM}\n\
\\newkomavar{iban}\n\
\\setkomavar{iban}{DE68 7635 0000 0000 0076 01}\n\
\\newkomavar{bic}\n\
\\setkomavar{bic}{BYLADEM1ERH}\n\n"

# %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
# \\footnotesize => \\scriptsize
TXT_3 = "% Begin header\n\
\\firsthead{\\null\\hfill \n\
 \\parbox[t][\\headheight][t]{10cm}{%\n\
  \\flushright\n\
  {\\small\\textsf{\n\
  \\color[gray]{.5}%\n\
  \\begin{tabular}[t]{r l}\n\
	\\textbf{\\usekomavar{ausrichter}} & \\multirow{5}{*}{\\usekomavar{fromlogo}} \\\\\n\
	Dr. Ulf Meerwald & \\\\\n\
	2.Vorsitzender &\\\\\n\
  \\usekomavar{fromaddress} & \\\\\n\
  \\end{tabular}\n\
  \\footnotesize{\\usekomavar{fromemail}}\\\\\n\
  \\footnotesize{\\usekomavar{fromurl}}\\\\\n\
  [\\baselineskip]\n\
  }}}%\n\
}\n\n\
% Begin footer\n\
\\firstfoot{\n\
\\parbox[t]{\\textwidth}{\n\
	\\scriptsize\n\
	\\color[gray]{.5}%\n\
	\\begin{tabular}[t]{l | r l | r l}\n\
		\\textbf{\\usekomavar{ausrichter}} &\n\
		\\multicolumn{2}{l|}{\\textbf{Kontakt}} &\n\
		\\multicolumn{2}{l}{\\textbf{Bankverbindung}}\\\\\n"

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


####################################################################################################
# Vereine
#
Count_Boote   = 0
Count_Ruderer = 0
Count_Verein  = 0

EUR_total = 0
sql = "SELECT * FROM verein WHERE kurz != 'RVE' "

Meldegeld = 15
Bugnummer = 10
Deckelung = 250

fehlende_Bugnummern = [ 138 ]

Vcursor.execute(sql)
for Vsatz in Vcursor:
   # ------------------------------------- Kurzform mit Langform ersetzen
   Anzahl_Rennen = 0
   NoBoote = 0
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
   TXT = TXT +  Vsatz[3] + "}\n%\n"
   TXT = TXT +  "\\opening{Sehr geehrte Damen und Herren,}\n\n"
   TXT = TXT +  "Für die Teilnahme an der Langstrecke vom Bayrischen Ruderverband und dem \\usekomavar{ausrichter} am \n"
   TXT = TXT +  "\\usekomavar{gigdate} dürfen wir Ihnen folgendes berechnen:\n\n\\vspace{10pt}\n\n"
   TXT = TXT +  "\\begin{tabular}{p{0.8\\textwidth} r}\n\\toprule\n\\textbf{Bezeichnung} & \\textbf{Betrag} \\\\\n"
   TXT = TXT +  "\\midrule\n%\n"
   #
   TXT_Fontsize = "\\footnotesize"
   TXTA = "Verspätete Abmeldungen: " +TXT_Fontsize + "{"
   #
   TXTB = "Verlorene Bugnummern: \\newline" +TXT_Fontsize + "{"
   #
   TXTM = "Meldegeld für folgende Mannschaften: " +TXT_Fontsize + "{\\newline\n%"
   #
   #___________________________ durchsuche Rennen
   sql = "SELECT * FROM rennen WHERE status >= 1"
   Rcursor.execute(sql)
   for Rsatz in Rcursor:
      Rennen       = Rsatz[0]
      RennenString = Rsatz[1]
      # 
      #
      Ngray = 0
      Vrennen  = 0
      VArennen = 0
      sql = "SELECT * FROM boote  WHERE rennen = " + str(Rennen) + " ORDER BY startnummer"
      Bcursor.execute(sql)
      for Bsatz in Bcursor:
         StNr   = Bsatz[1]
         Boot   = Bsatz[0]
         #
         Abmeldung = Bsatz[10]
         # ===========================================================================================
         nPers   = 0
         sql = "SELECT * FROM r2boot  WHERE bootNr = " + str(Boot) 
         Qcursor.execute(sql)
         iR = 0
         for RudInd in Qcursor:         
            sql = "SELECT * FROM ruderer WHERE nummer = " + str(RudInd[3])
            Pcursor.execute(sql)
            Rd = Pcursor.fetchone()
            #
            if(Rd[7] == Vsatz[1]):
               nPers = nPers + 1
            if(iR == 0):
               # Name = "\\textbf{" + Rd[0] + " } " + Rd[1]
               Name =  Rd[1] + " " + Rd[2]
            else:
               # Name = Name + ", \\textbf{ " + Rd[0] + " } " + Rd[1]
               Name = "(" + Name + ", " + Rd[1] + " " + Rd[2] + ")"
            #               
            if(Vsatz[1] != Rd[7]):
               # Name = Name + " \\textcolor{gray}{\\scriptsize (" + Rd[6] + ")}"
               Name = Name + "$ ^{(" + Rd[7] + ")}$"
            #
            iR = iR + 1
         #
         if(nPers > 0):
            #
            if(Abmeldung == 0):
               NoBoote = NoBoote + 1
               if(iR == 2 and nPers == 1):
                  EURO = EURO + Meldegeld/2
               else:
                  EURO = EURO + Meldegeld
               # TXT = TXT + VTXT
               Anzahl_Rennen = Anzahl_Rennen + 1
               #
               if(Vrennen == 0):
                  Vrennen = 1
                  TXTM = TXTM + "\n\\newline\\textbf{ Rennen " + str(Rennen) + ": " + RennenString + ":} "
                  TXTM = TXTM + Name + " "
               else:
                  TXTM = TXTM + ", " + Name + " "
            # ===========================================================================================
            elif(Abmeldung > 1):  # verspätet abgemeldet
               if(VArennen == 0):
                  VArennen = 1
                  TXTA = TXTA + "\n\\newline\\textbf{Rennen " + str(Rennen) + ": " + RennenString + ": } "
                  TXTA = TXTA + "\\textcolor{red}{ " + Name + "}"
               else:
                  TXTA = TXTA + ", \\textcolor{red}{ " + Name + "}"
               #
               if(iR == 2 and nPers == 1):
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
      TXT = TXT + TXTM + "\n\\newline}	& " + euronen + " \\\\\n"
      TXT = TXT + "%\n% ... Meldegeld für " + str(NoBoote) + " Boote (genauer auf nächster Seite) ...\n%\n%\n"
      if(EURO > Deckelung):
         euronen = "%6.2f"% (EURO - Deckelung)
         TXT = TXT + "%\nDeckelung der Meldegebühren auf EUR " + str(Deckelung) + ": \\newline"
         TXT = TXT + " & \color{red}{-" + euronen + "} \\\\\n%\n"
         EURO = Deckelung
         
   if(EURA > 0):
      euronen = "%6.2f"% (EURA)
      TXT = TXT + TXTA +  "\\newline}	& " + euronen + " \\\\\n"
      # TXT = TXT + "%\nKulanz für die verspäteten Abmeldungen: \\newline"
      #TXT = TXT + "\\footnotesize\\textsf{Auf Grund der Corona-Pandemie waren Abmeldungen bis zum 22.10. um 14 Uhr kostenfrei. \n"
      #TXT = TXT + "Danach haben wir für das Boot Start- und Bugnummern eingetütet.\\newline} & \color{red}{" + euronen + "} \\\\\n%\n"
      # TXT = TXT + "%\nAufwandspauschale der entstandenen Kosten & 20,00 \\\\\n%"
      # EURA = 20
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
   TXT = TXT + "{\\scriptsize Die Rechnungsstellung erfolgt ohne Ausweis der Umsatzsteuer nach \\textsection19 UStG.\\vspace{5pt}\\\\}\n"
   TXT = TXT + "Bitte benutzen Sie bei der Überweisung das Kennwort:\n\\\\ '\\textbf{Meldegebühr Langstreckentest " + Vsatz[1] + "'} \\vspace{1.2cm}\\\\\n"
   #TXT = TXT + "\\vspace{5pt}\\\\\nMit rudersportlichen Grü{\ss}en,\\vspace{-10pt}\\\\\n\\includegraphics[height=1.71cm,width=4.13cm]{UlfMeerwald.png}"
   TXT = TXT + "Vielen Dank im voraus, \\\\mit rudersportlichen Grü{\ss}en \\vspace{-2.1cm}\\\\ \\color{white}{.}\\hspace{7cm}\\color{white}{.}\n"
   TXT = TXT + "\\includegraphics[height=2.55cm,width=6.2cm]{UlfMeerwald.png}\n"
   #Vielen Dank im voraus, \vspace{0.2cm}\\mit rudersportlichen Grü{\ss}en \vspace{-2.1cm}\\ \color{white}{.}\hspace{7cm}\color{white}{.}
   TXT = TXT + "\n\\end{letter}\n\\end{document}\n"
   #
   TXT = TXT.replace('ß', '{\\ss}')
   fp = open("LaTeX/Rechnung_" + Vsatz[1] + ".tex","w")
   fp.write(TXT)
   fp.close()


#TXT = TXT_1 + TXT_2 + TXT_3 + TXT_4

##############################################################################################################
# Korrektur des 'ß' - sz:
#TXT = TXT.replace('ß', '{\\ss}')

# TXT = TXT + "\\n%======================\\n\\\\end{document}\\n"
connection.close()

print("Gemeldet haben " + str(Count_Verein) + " Vereine und wir stellen " + str(EUR_total) + " EUR in Rechnung")