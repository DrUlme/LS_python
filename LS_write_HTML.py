#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Oct 12 18:35:57 2020

@author: ulf
"""

import time
import math
import os, sys, sqlite3
import fnmatch

import LSglobal

#===============================================================================================================================
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
# --------------------------------------------------------------------------------------------------------------- Frühstarter
# hole Renn-Nummern für Früh und Spät-Starter
sql = "SELECT wert FROM meta WHERE name = 'Frühstarter'"
Rcursor.execute(sql)
Rd = Rcursor.fetchone()
Frühstart = Rd[0]

sql = "SELECT wert FROM meta WHERE name = 'Spätstarter'"
Rcursor.execute(sql)
Rd = Rcursor.fetchone()
Spätstart = Rd[0]

# --------------------------------------------------------------------------------------------------------------- Athletiktest
sql = "SELECT * FROM rennen WHERE name LIKE 'Spätstarter%'"
Rcursor.execute(sql)
Rd = Rcursor.fetchone()
print(Rd)
SpätR = Rd[0]
# TimeSpät = int(Rd[5])


#===============================================================================================================================
t1 = time.localtime()

d_M_string  = str(t1.tm_mday).zfill(2) + '.' + str(t1.tm_mon).zfill(2) + '.'

# HTMSel = "\n\n <select name=\"forma\" onchange=\"location = this.value;\" target=\"iframe_a\"; >\n"
HTMSel = "\n\n <select name=\"forma\" onchange=\"location = this.value;\" >\n"
HTMSel = HTMSel + "   <option value=\"\" selected disabled hidden>zu Rennen</option>\n"
HTMSel = HTMSel + "   <option value=\"index.html\">HOME </option>\n"
# HTMSel = HTMSel + " <option value=\"aktuell.html\">aktuell </option>\n"

#print(TXT)
HTXT = "<!DOCTYPE html>\n\n<html>\n<head>\n <meta charset=\"utf-8\">\n <title>RVE-Langstreckenregatta</title>"
HTXT = HTXT + " <link rel=\"shortcut icon\" type=\"image/x-icon\" href=\"/Langstrecke/favicon.ico\">\n\n"
# HTXT = HTXT + " <link rel=\"stylesheet\" href=\"Rennen.css\">\n"

HTXT = HTXT + " <style>\n  p {\n    display: block;\n    margin-top: 0em;\n    margin-bottom: 0em;\n"
HTXT = HTXT + "    margin-left: 0;\n    margin-right: 3%;\n  }\n"
# HTXT = HTXT + "  iframe { height:100%; width:80%; position:absolute; right:0px; }\n"
HTXT = HTXT + "   a:link { color: black; }\n   a:visited { color: darkblue; }\n    a:hover {   color: hotpink; }\n\n"
# 	/* selected link */
# 	a:active {
# 	  color: red;
# 	}
# 	aside { width: 20%; position: absolute; top: 0px; left: 0px;}
# 	
HTXT = HTXT + "   select {\n     font-size: 1.9rem;\n     padding: 2px 5px;\n     width: 400px;\n     max-width: 400px;\n   }\n"

HTXT = HTXT + " </style>\n</head>\n\n<body>\n\n"
# 
#------------------------------------------------------------------------------------
HTXT = HTXT + '<p style="font-size:20px; color:blue">Herbst-Langstrecke <b>2023 </b>\n<br>\n'
HTXT = HTXT + '<img height="10%" width="10%" src="RVE-Flag.png">\n<br>\n<br>\n</p>\n\n'

HTXT = HTXT + '<p style="font-size:18px; color:blue">\nDie Langstrecke wird am <b> '
HTXT = HTXT + str(LSglobal.Tag) + '. ' + LSglobal.MonatStr + ' ' + str(LSglobal.Jahr) + ' </b> durchgeführt.<br>\n<br>\n</p>  \n'
HTXT = HTXT + ' <p style="font-size:18px; color:red">\n<br>Hier ist der <a HREF="Meldebogen_' + LSglobal.ZeitK + '_' + str(LSglobal.Jahr) + '.xlsx"> Meldebogen </a> für die Langstrecke.</p>\n'
HTXT = HTXT + ' <p style="font-size:15px; color:blue">Jetzt auch mit Beispiel für die Eingabe (Daten bitte überschreiben).</p>\n'

# ToDO: Ausschreibung
HTXT = HTXT + ' <p style="font-size:18px; color:blue"><br>Die offizielle <a HREF="2023_09_25_BRV_HLS_Ausschreibung.pdf"> Ausschreibung </a> mit weiteren Daten.<br>\nDas wichtigste hier direkt:</p>\n\n'

HTXT = HTXT + '<p style="font-size:15px; color:blue">\n<br>\n<br>\n'
HTXT = HTXT + ' <b>Meldeschluss: </b> Mittwoch, den <b> ' + LSglobal.Meldeschluss + ' </b>' + str(LSglobal.Jahr) + ', 22:00 Uhr<br>\n'
HTXT = HTXT + ' Meldungen an: <a href="mailto:langstrecke@ruderverein-erlangen.de"><i>langstrecke@ruderverein-erlangen.de</i></a><br>\n<br>\n'
HTXT = HTXT + ' Die Startreihenfolge wird am Montag ' + LSglobal.Setzdatum + str(LSglobal.Jahr) + ' festgelegt.<br>\n<br>\n'
HTXT = HTXT + ' Bei Meldungen nach Samstag den <b>' + LSglobal.DoppeltGeld + str(LSglobal.Jahr) + '</b> ist das doppelte Meldegeld zu zahlen!<br>\n<br>\n'
HTXT = HTXT + ' Um- und Abmeldungen sind nur bis Freitag den ' + LSglobal.AbmeldungDD + ' bis ' + LSglobal.AbmeldungHH + ' Uhr möglich<br> - danach wird das Meldegeld in Rechnung gestellt. <br>\n<br>\n<br>\n'
if(LSglobal.ZeitK == 'H'):
   HTXT = HTXT + ' Die Junioren B m&uuml;ssen f&uuml;r die DOSB Bewertungen mindestens einmal während der Saison im 1x einen Test &ge; 5 km absolvieren.<br>\n'
   HTXT = HTXT + ' Die entsprechenden Riemen-Zweier sind daher aus dem Programm für den Herbst gestrichen worden.<br>\n<br>\n'
   HTXT = HTXT + ' Die Altersgrenzen für 2024 gelten bereits für diesen Langstreckentest. <br>\n<br>\n'
   HTXT = HTXT + ' <b>Kinder, die dadurch als Junioren-B starten, </b>\nm&uuml;ssen jetzt bereits das &auml;rztlichen Attest f&uuml;r ' + str(LSglobal.RefJahr) + ' vorweisen - <br>\n<b>bitte einen Termin beim Arzt zwischen dem 1. und 20. Oktober ausmachen!</b>\n\n'

HTXT = HTXT + ' <br><br>\n</p>\n\n<hr>\n'
HTXT = HTXT + '\n<p style="font-size:19px; color:blue"><a href="Hinweise_Fahrordnung_Ruderstrecke.pdf">Hinweise zur Fahrordnung auf der Ruderstrecke währen der Regatta</a>.\n<br>\n</p>\n'


#-------------------------------------------------------------------
HTXT = HTXT + "<br><font size=\"4\"><a href=\"Meldungen.pdf\" >aktuelles Melde-PDF</a></font><br>\n\n<p>\n"
#
#HTXT = HTXT + "<font size=\"4\"><a href=\"aktuell.html\" target=\"iframe_a\">Aktuell</a></font> \n"
#HTXT = HTXT + " <font size=\"3\" color=\"gray\"> aktuell laufende Rennen</font><br>\n<br>--------<br>\n<br>\n"
# Buttons
#HTXT = HTXT + "<div>\n <input type=\"radio\" name=\"biframe\" value=\"aktuell.html\" onclick=\"document.querySelector('iframe.reload').setAttribute('src', this.value);\" checked=\"checked\">Aktuell<br>\n"
#HTXT = HTXT + " <input type=\"radio\" name=\"biframe\" value=\"aktuell.html\" onclick=\"window.location = 'aktuell.html';\" >Mobile Version<br>\n <br>\n"

Count_Boote   = 0
Count_Ruderer = 0


#===============================================================================================================================
# sql = "SELECT * FROM rennen WHERE status > 0 "
# Rcursor.execute(sql)
# for Rsatz in Rcursor:
#    Rennen       = Rsatz[0]   # Nummer
#    RennenString = Rsatz[2]   # ausgeschrieben
#    RennenCode   = Rsatz[3]   # SM 2x A Lgw.
#    RLgw         = Rsatz[9]
#    Status       = Rsatz[8]
#    RStr         = str(Rennen)
#    HTMSel = HTMSel + "   <option value=\"Rennen_" + RStr + ".html\">Rennen  " + RStr + ": " + RennenString + "</option>\n"
# HTMSel = HTMSel + " </select>\n"
#===============================================================================================================================
HTMLsel_V = '\n <select name="formV" onchange="location = this.value;" >\n'
HTMLsel_V = HTMLsel_V + '   <option value="" selected disabled hidden>zu Verein</option>\n'
HTMLsel_V = HTMLsel_V + '   <option value="index.html">HOME </option>\n'
sql = "SELECT * FROM verein WHERE dabei > 0 "
Rcursor.execute(sql)
for Rsatz in Rcursor:
   HTMLsel_V = HTMLsel_V + '   <option value="' + Rsatz[2] + '.html">' + Rsatz[1] + '   (' + Rsatz[2] + ')</option>\n'
HTMLsel_V = HTMLsel_V + ' </select>'
#   
#   
#   <option value="aktuell.html">aktuell </option>
#   <option value="Rennen_1.html">Rennen  1: Großboot</option>
# 
 
#===============================================================================================================================

sql = "SELECT * FROM rennen WHERE status > 0 and strecke != 'Athletik'"  # Athletik-Wettbewerb! (ohne: 37)
Rcursor.execute(sql)
for Rsatz in Rcursor:
   Rennen       = Rsatz[0]   # Nummer
   RennenString = Rsatz[2]   # ausgeschrieben
   RennenCode   = Rsatz[3]   # SM 2x A Lgw.
   RLgw         = Rsatz[9]
   Status       = Rsatz[8]
   RStr         = str(Rennen)
   #
   HTMLfile = "HTML/Rennen_" + RStr + ".html"
   # print(RStr)
   #
   # ___________________ option <a href>: _________________
   # long format
   # HTXT = HTXT + "<font size=\"4\"><a href=\"Rennen_" + RStr + ".html\" target=\"iframe_a\">Rennen " + RStr + " </a> </font> " + RennenString  
   # short format
   # HTXT = HTXT + "<font size=\"4\"><a href=\"Rennen_" + RStr + ".html\" target=\"iframe_a\"># <b>" + RStr + ": </b> <i> " + RennenString + "</i></a> </font> "   
   # HTXT = HTXT + "<a href=\"Rennen_" + RStr + ".html\" target=\"iframe_a\"># <b>" + RStr + ": </b> <i> " + RennenString + "</i></a>"   
   # HTXT = HTXT + "<br>\n<br>\n"
   #
   # ___________________ option <a href>: _________________
   # HTXT = HTXT + "<a href=\"Rennen_" + RStr + ".html\" target=\"iframe_a\"># <b>" + RStr + ": </b> <i> " + RennenString + "</i></a>"   
   # HTXT = HTXT + "<br>\n<br>\n"
   #
   # ___________________ option <radio button>: _________________
   #HTXT = HTXT + " <input type=\"radio\" name=\"biframe\" value=\"Rennen_"+ RStr + ".html\" onclick=\"document.querySelector('iframe.reload').setAttribute('src', this.value);\" >Rennen " + RStr + " : <small> " + RennenString +"</small><br>\n"
   #
   HTMSel = HTMSel + "   <option value=\"Rennen_" + RStr + ".html\">Rennen  " + RStr + ": " + RennenString + "</option>\n"
   #
   #=====================================================================================================================
   # Header für 'Rennen X'
   TXT = "<!DOCTYPE html>\n<html lang=\"de\">\n  <head>\n    <meta charset=\"utf-8\">\n    <title> Rennen "
   TXT = TXT + RStr + "</title>\n    <link rel=\"stylesheet\" href=\"Rennen.css\">\n"
   TXT = TXT + "   <link rel=\"shortcut icon\" type=\"image/x-icon\" href=\"/Langstrecke/favicon.ico\">\n\n"
   #
   # ToDo: zeitabhängig reload abfragen:
   # TXT = TXT + "   <meta http-equiv=\"refresh\" content=\"120\"/>\n"
   #
   TXT = TXT + " <link rel=\"stylesheet\" href=\"Rennen.css\">\n"
   TXT = TXT + " <style>\n  p {\n    display: block;\n    margin-top: 0em;\n    margin-bottom: 0em;\n"
   TXT = TXT + "    margin-left: 0;\n    margin-right: 3%;\n  }\n"
   TXT = TXT + "    body {\n     max-width: 60em;\n }\n"
   TXT = TXT + " </style>\n"
   #
   TXT = TXT + "  </head>\n\n<body>\n"
   TXT = TXT + "<table id=\"coll\">\n <thead>\n  <tr>\n      <th style=\"width:10%\">Start-<br>Nummer</th>\n      <th colspan=\"3\"><b> Rennen "
   # TXT = TXT + RStr +" - " + RennenString + " - </b> "
   TXT = TXT + RStr +" - " + RennenCode + " - </b> "
   if(Status == 1):
      TXT = TXT + "- <iH> vorl&auml;ufige Meldungen</iH>"
      Bemerkung = "noch keine Zeitplanung"
      myOrder = " ORDER BY id "
   elif(Status == 2):
      TXT = TXT + "- <iH> Melde-Ergebnis</iH>"
      Bemerkung = "Startzeit"
      myOrder = " ORDER BY zeit3000 DESC, zeit DESC, planstart, startnummer "
   elif(Status == 5):
      TXT = TXT + "- <iH> Endergebnis </iH>"
      Bemerkung = "Endergebnis"
      myOrder = " ORDER BY zeit"   
   else:
      TXT = TXT + "- <iH> vorl&auml;ufiges Ergebnis</iH>"
      Bemerkung = "vorläufig"
      myOrder = " ORDER BY zeit, zeit3000, tStart, planstart " 
   #
   if(RStr == Frühstart):
      # Frühstarter besonders gehandelt
      sql = "SELECT * FROM boote  WHERE abgemeldet = 0 and alternativ = " + Frühstart + "  " + myOrder
      print(sql)
   elif(RStr == Spätstart):
      # Spätstarter besonders gehandelt
      sql = "SELECT * FROM boote  WHERE abgemeldet = 0 and alternativ = " + Spätstart + "  " + myOrder
   else:
      sql = "SELECT * FROM boote  WHERE rennen = " + RStr + " and abgemeldet = 0 " + myOrder
   TXT = TXT + "</th>\n  </tr>\n </thead>\n <tbody>\n"
   # comment
   TXT = TXT + " <!-- ------------------------------------------------------------------------------------- -->\n"
   #
   # print(sql)
   ###############################################################################################################
   Bcursor.execute(sql)
   for Bsatz in Bcursor:
      Boot   = Bsatz[0]
      StNr   = Bsatz[2]
      #         
      if(len(Bsatz[6]) == 8):
         gestartet = 1
      else:
         gestartet = 0
      # _______________________________________________________________________________________________________________
      #
      Count_Boote = Count_Boote + 1
      #
      sql = "SELECT * FROM r2boot  WHERE bootid = " + str(Boot) 
      Qcursor.execute(sql)
      iR = 0
      for RudInd in Qcursor:         
         sql = "SELECT * FROM ruderer WHERE id = " + str(RudInd[2])
         Pcursor.execute(sql)
         Rd = Pcursor.fetchone()
         # 
         Count_Ruderer = Count_Ruderer + 1
         #
         if(iR == 0):
            Name   = "<b>" + Rd[1] + " " + Rd[2] + " </b> "
            Verein = "<i>" + Rd[7] + "</i>"
         else:
            Name = Name + ", <b>" + Rd[1] + " " + Rd[2] + " </b> "
            if(("<i>" + Rd[7] + "</i>") != Verein):
               Verein = Verein + "<br><i>" + Rd[7] + "</i>"
         if(RLgw < 0 and Rd[5] == 0):
            Name = Name + " (" + str(Rd[4]) + ") "
         elif(RLgw > 0 and Rd[6] < 0):
            Name = Name + " (" + str(Rd[4]) + ") "
         elif(Rd[6] <= 0 and gestartet == 1):
            Name = Name + " (" + str(Rd[4]) + "<sup><s> Lgw. </s></sup>) "
         elif(Rd[6] <= RLgw):
            Name = Name + " (" + str(Rd[4]) + "<sup> Lgw.</sup>) "
         elif(Rd[6] <= (RLgw + 2) ):
            Name = Name + " (" + str(Rd[4]) + "<sup> Lgw +2</sup>) "
         else:
            Name = Name + " (" + str(Rd[4]) + "<sup> Lgw?</sup>) "
            
         iR = iR + 1
         #
      #
      if(StNr == 0):
         NrStr = "-"
         Btime = "-"
      else:
         NrStr = str(StNr)
         # ToDo - Früh-, Spätstarter
         if(str(Bsatz[12]) == Frühstart):
            NrStr = NrStr + "<sup><i>F</i></sup>"
         elif(Bsatz[12] == Spätstart ):
            NrStr = NrStr + "<sup><i>S</i></sup>"
         #---
         #
         # (End-) Zeit gesetzt ?
         if(len(Bsatz[10]) == 5):
            Btime = Bsatz[10]
            if(len(Bsatz[9]) == 5):
               # Pfeil Mitte &#8614; oder &#8696;
               zBemerkung = "<small>&#8614; " + Bsatz[8] + " &#8696; "   + Bsatz[9] + " &#8677;</small>"                
            else:
               zBemerkung = "Endzeit"
         elif(len(Bsatz[8]) == 5):
            Btime = Bsatz[8]
            zBemerkung = "3000 m"
         elif(len(Bsatz[5]) == 8):
            Btime = Bsatz[5]
            zBemerkung = "gestartet"
         else:
            Btime = Bsatz[4]
            zBemerkung = "Plan"
            #
      #-----
      # BtimH   = math.floor(Btime/3600)
      # BtimM   = math.floor(Btime/60 - BtimH*60 )
      # ZeitStr =  str(BtimH) + ":" + str(BtimM).rjust(2, '0') + ":" + str(Btime - 3600*BtimH - 60*BtimM).rjust(2, '0') + " "
      ZeitStr = Btime
      #
      TXT = TXT + " <tr>\n   <td><b>" + NrStr + "</b></td>\n   <td>"
      # <b> Ben Kentersack </b>(2002<sup><nolgw>Lgw.</nolgw></sup>)
      TXT = TXT + Name
      TXT = TXT + "   </td>\n   <td> " + Verein + " </td>\n"
      if(Status == 1):
         TXT = TXT + "<td> <zeitbemerkung>" + Bemerkung +"</zeitbemerkung></td>\n"
      elif(len(Bsatz[10]) == 5):
         TXT = TXT + "   <td><b>" + ZeitStr + " </b><br><zeitbemerkung>" + zBemerkung +"</zeitbemerkung></td>\n"
      elif(len(Bsatz[8]) == 5):
         TXT = TXT + "   <td><font color=\"blue\"><b>" + ZeitStr + " </b><br><zeitbemerkung>" + zBemerkung +"</zeitbemerkung></font></td>\n"
      elif(len(Bsatz[5]) == 8):
         TXT = TXT + "   <td><font color=\"red\"><i>" + ZeitStr + " </i><br><zeitbemerkung>" + zBemerkung +"</zeitbemerkung></font></td>\n"
      else:
         TXT = TXT + "   <td><font color=\"green\"><i>" + ZeitStr + " </i><br><zeitbemerkung>" + zBemerkung +"</zeitbemerkung></font></td>\n"
      TXT = TXT + "  </tr>\n\n"
   #-----
   # comment
   TXT = TXT + " <!-- ------------------------------------------------------------------------------------- -->\n"
   #
   TXT = TXT + "\n\n </tbody>\n</table>\n\n"
   TXT = TXT + "<p style=\"font-size:15px; color:blue\"><br>Status vom " + str(t1.tm_mday) + "." + str(t1.tm_mon) + \
      "." + str(t1.tm_year) + " um " + str(t1.tm_hour) + ":" + str(t1.tm_min).rjust(2, '0') + "  Uhr  \n"
   TXT = TXT + "<br>\n"
   # closing of file now in appended part "</p>\n\n</body>\n</html>\n"
   # replacing obsolet if <meta charset="utf-8"> in header
   #   TXT = TXT.replace("ß", "&szlig;")
   #   TXT = TXT.replace("ä", "&aumlg;")
   #   TXT = TXT.replace("ö", "&ouml;")
   #   TXT = TXT.replace("ü", "&uuml;")  
   #   TXT = TXT.replace("Ö", "&Ouml;")
   #
   sql = "SELECT * FROM verein "
   Vcursor.execute(sql)
   for Vsatz in Vcursor:
      # ------------------------------------- Kurzform mit Langform ersetzen
      VereinStr = "<i>" + Vsatz[1] +"</i>"
      TXT = TXT.replace(VereinStr, Vsatz[0])
   #
   fp = open(HTMLfile,"w")
   fp.write(TXT)
   fp.close()
#
HTMSel = HTMSel + " </select>\n\n\n  " + HTMLsel_V + "\n </p>\n"
# 
#
HTXT = HTXT + "<br>\n" + HTMSel + "<br>\n"
HTXT = HTXT + "Es starten aktuell <b>" + str(Count_Ruderer) + "</b> Ruderer in <b>" + str(Count_Boote) + "</b> Booten. <br>\n"
HTXT = HTXT + "<br>\n<hr>\n<br>\n"
#
HTXT = HTXT + "\n  <p><br><a HREF=\"F2023_Endergebnis.pdf\">Ergebnis der Fr&uuml;hjahrs-Langstrecke 2023</a></p>\n"
HTXT = HTXT + "\n  <p><br><a HREF=\"H2022_Endergebnis.pdf\">Ergebnis der Herbst-Langstrecke 2022</a></p>\n"
HTXT = HTXT + "\n  <p><br><a HREF=\"F2022_Endergebnis.pdf\">Ergebnis der Fr&uuml;hjahrs-Langstrecke 2022</a></p>\n"
HTXT = HTXT + "\n  <p><br><a HREF=\"H2021_Endergebnis.pdf\">Ergebnis der Herbst-Langstrecke 2021</a></p>\n"
HTXT = HTXT + "\n  <p><br><a HREF=\"H2020_Endergebnis.pdf\">Ergebnis der Herbst-Langstrecke 2020</a></p>\n"
#
HTXT = HTXT + "\n  </p>\n\n</body>\n</html>\n"
fp = open("HTML/index.html","w")
fp.write(HTXT)
fp.close()

# =================================================== append Auswahlliste an "Rennen_*.html"
HTMSel = HTMSel + "\n\n</body>\n</html>\n"
FILES=os.listdir( "./HTML" )
for filename in FILES:
   #if filename.endswith('.html'):
   if fnmatch.fnmatch(filename, 'Rennen_*.html'):
      # print("append to '" + filename + "'" )
      fp = open("HTML/" + filename,"a")
      fp.write(HTMSel)
      fp.close()

print("__________________________________________________________________\n\nEs starten <b>" + str(Count_Ruderer) + "</b> Ruderer in <b>" + str(Count_Boote) + "</b> Booten ")

# Verbindung beenden
connection.close()
