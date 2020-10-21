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

# HTMSel = "\n\n <select name=\"forma\" onchange=\"location = this.value;\" target=\"iframe_a\"; >\n"
HTMSel = "\n\n <select name=\"forma\" onchange=\"location = this.value;\" >\n"
HTMSel = HTMSel + " <option value=\"\" selected disabled hidden>Choose here</option>\n"
HTMSel = HTMSel + " <option value=\"index.html\">HOME </option>\n"
HTMSel = HTMSel + " <option value=\"aktuell.html\">aktuell </option>\n"

#print(TXT)
HTXT = "<!DOCTYPE html>\n\n<html>\n<head>\n <meta charset=\"utf-8\">\n <title>RVE-Langstreckenregatta</title>"
HTXT = HTXT + " <link rel=\"shortcut icon\" type=\"image/x-icon\" href=\"/Langstrecke/favicon.ico\">\n\n"
HTXT = HTXT + " <link rel=\"stylesheet\" href=\"Rennen.css\">\n"
HTXT = HTXT + " <style>\n  p {\n    display: block;\n    margin-top: 0em;\n    margin-bottom: 0em;\n"
HTXT = HTXT + "    margin-left: 0;\n    margin-right: 80%;\n  }\n"
# HTXT = HTXT + "  iframe { height:100%; width:80%; position:absolute; right:0px; }\n"
HTXT = HTXT + " </style>\n</head>\n\n<body>\n\n"
# 
HTXT = HTXT + "<iframe class=\"reload\" align=\"right\" height=\"900\" width=\"80%\" src=\"aktuell.html\" name=\"iframe_a\"></iframe>\n\n"
# HTXT = HTXT + "<iframe align=\"right\" height=\"100%\" width=\"80%\" src=\"aktuell.html\" name=\"iframe_a\"></iframe>\n\n"
# HTXT = HTXT + "<iframe src=\"aktuell.html\" name=\"iframe_a\"></iframe>\n\n"
HTXT = HTXT + "<img height=\"15%\" width=\"15%\" src=\"RVE-Flag.png\"><br>\n\n"
HTXT = HTXT + "<p style=\"font-size:30px; color:blue\">Langstrecke</p>\n"
HTXT = HTXT + "<p style=\"font-size:25px; color:blue\">Herbst <b>2020</b><br></p>\n"
HTXT = HTXT + "<br>\n<br>\n"
HTXT = HTXT + "<font size=\"4\">Bitte beachtet <a href=\"Hygienekonzept_RVE_Langstrecke_24.10.20.pdf\"> unser Hygienekonzept </a>!<br></font>\n<br>\n"
HTXT = HTXT + "<font size=\"4\"><a href=\"Meldungen.pdf\" >aktuelles Melde-PDF</a></font><br>\n\n<p><br>"
#
#HTXT = HTXT + "<font size=\"4\"><a href=\"aktuell.html\" target=\"iframe_a\">Aktuell</a></font> \n"
#HTXT = HTXT + " <font size=\"3\" color=\"gray\"> aktuell laufende Rennen</font><br>\n<br>--------<br>\n<br>\n"
# Buttons
HTXT = HTXT + "<div>\n <input type=\"radio\" name=\"biframe\" value=\"aktuell.html\" onclick=\"document.querySelector('iframe.reload').setAttribute('src', this.value);\" checked=\"checked\">Aktuell<br>\n"
HTXT = HTXT + " <input type=\"radio\" name=\"biframe\" value=\"aktuell.html\" onclick=\"window.location = 'aktuell.html';\" >Mobile Version<br>\n <br>\n"

Count_Boote   = 0
Count_Ruderer = 0

sql = "SELECT * FROM rennen WHERE status > 0"
Rcursor.execute(sql)
for Rsatz in Rcursor:
   Rennen       = Rsatz[0]
   RennenString = Rsatz[1]
   Status       = Rsatz[6]
   RStr         = str(Rennen)
   #
   HTMLfile = "HTML/Rennen_" + RStr + ".html"
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
   HTXT = HTXT + " <input type=\"radio\" name=\"biframe\" value=\"Rennen_"+ RStr + ".html\" onclick=\"document.querySelector('iframe.reload').setAttribute('src', this.value);\" >Rennen " + RStr + " : <small> " + RennenString +"</small><br>\n"
   #
   HTMSel = HTMSel + " <option value=\"Rennen_" + RStr + ".html\">Rennen  " + RStr + ": " + RennenString + "</option>\n"
   #
   TXT = "<!DOCTYPE html>\n<html lang=\"de\">\n  <head>\n    <meta charset=\"utf-8\">\n    <title> Rennen "
   TXT = TXT + RStr + "</title>\n    <link rel=\"stylesheet\" href=\"Rennen.css\">\n"
   TXT = TXT + "   <link rel=\"shortcut icon\" type=\"image/x-icon\" href=\"/Langstrecke/favicon.ico\">\n\n"
   TXT = TXT + "   <meta http-equiv=\"refresh\" content=\"120\"/>\n"
   TXT = TXT + "  </head>\n\n<body>\n"
   TXT = TXT + "<table id=\"coll\">\n <thead>\n  <tr>\n      <th>Start-<br>Nummer</th>\n      <th colspan=\"3\"><b> Rennen "
   TXT = TXT + RStr +" - " + RennenString + " - </b> "
   if(Status == 1):
      TXT = TXT + "- <iH> vorl&auml;ufige Meldungen</iH>"
      Bemerkung = "noch keine Zeitplanung"
      sql = "SELECT * FROM boote  WHERE rennen = " + RStr + " ORDER BY vereine, nummer "
   elif(Status == 2):
      TXT = TXT + "- <iH> Melde-Ergebnis</iH>"
      Bemerkung = "Startzeit"
      sql = "SELECT * FROM boote  WHERE rennen = " + RStr + " and abgemeldet = 0 ORDER BY startnummer, planstart "
   elif(Status == 5):
      TXT = TXT + "- <iH> Endergebnis </iH>"
      Bemerkung = "Endergebnis"
      sql = "SELECT * FROM boote  WHERE rennen = " + RStr + " and abgemeldet = 0  ORDER BY zeit"   
   else:
      TXT = TXT + "- <iH> vorl&auml;ufiges Ergebnis</iH>"
      Bemerkung = "vorläufig"
      sql = "SELECT * FROM boote  WHERE rennen = " + RStr + " and abgemeldet = 0  ORDER BY zeit, zeit3000, secstart, planstart " 
   #
   TXT = TXT + "</th>\n  </tr>\n </thead>\n <tbody>\n"
   # comment
   TXT = TXT + " <!-- ------------------------------------------------------------------------------------- -->\n"
   #
   ###############################################################################################################
   Bcursor.execute(sql)
   for Bsatz in Bcursor:
      Boot   = Bsatz[0]
      StNr   = Bsatz[1]
      Verein = Bsatz[3]
      RudInd = Bsatz[4].split(',')
      Abmeldung = Bsatz[12]
      if(Abmeldung == 0):
         # _______________________________________________________________________________________________________________
         #
         Count_Boote = Count_Boote + 1
         #
         for iR in range(0, (len(RudInd) - 2)):         
            sql = "SELECT * FROM ruderer WHERE nummer = " + str(RudInd[iR + 1])
            Pcursor.execute(sql)
            Rd = Pcursor.fetchone()
            # 
            Count_Ruderer = Count_Ruderer + 1
            #
            if(iR == 0):
               Name   = "<b>" + Rd[0] + " " + Rd[1] + " </b> ( " + str(Rd[3]) + ") "
               Verein = "<i>" + Rd[6] + "</i>"
            else:
               Name = Name + ", <b>" + Rd[0] + " " + Rd[1] + " </b> ( " + str(Rd[3]) + ") "
               if(("<i>" + Rd[6] + "</i>") != Verein):
                  Verein = Verein + "<br><i>" + Rd[6] + "</i>"
            #
         #
         if(StNr == 0):
            NrStr = "-"
            ZeitStr = "-"
         else:
            NrStr = str(StNr)
            #
            if(Bsatz[11] > 0):
               Btime = Bsatz[11]
               ZeitStr =  str(BtimH) + "$:$" + str(BtimM).rjust(2, '0') + "$:\\small{\\textcolor{gray}{" + str(Btime - 3600*BtimH - 60*BtimM).rjust(2, '0') + "}}"
               zBemerkung = "Endzeit"
            elif(Bsatz[9] > 0):
               Btime = Bsatz[9]
               zBemerkung = "3000 m"
            elif(Bsatz[6] > 0):
               Btime = Bsatz[6]
               zBemerkung = "gestartet"
            else:
               Btime = Bsatz[5]
               zBemerkung = "Plan"
               
                              
         BtimH   = math.floor(Btime/3600)
         BtimM   = math.floor(Btime/60 - BtimH*60 )
         ZeitStr =  str(BtimH) + ":" + str(BtimM).rjust(2, '0') + ":" + str(Btime - 3600*BtimH - 60*BtimM).rjust(2, '0') + " "
                  
         #
         TXT = TXT + " <tr>\n   <td><b>" + NrStr + "</b></td>\n   <td>"
         # <b> Ben Kentersack </b>(2002<sup><nolgw>Lgw.</nolgw></sup>)
         TXT = TXT + Name
         TXT = TXT + "   </td>\n   <td> " + Verein + " </td>\n"
         if(Status == 1):
            TXT = TXT + "<td> <zeitbemerkung>" + Bemerkung +"</zeitbemerkung></td>\n"
         else:
            TXT = TXT + "   <td><b>" + ZeitStr + " </b><br><zeitbemerkung>" + zBemerkung +"</zeitbemerkung></td>\n"
         TXT = TXT + "  </tr>\n\n"

   # comment
   TXT = TXT + " <!-- ------------------------------------------------------------------------------------- -->\n"
   #
   TXT = TXT + "\n\n </tbody>\n</table>\n\n"
   TXT = TXT + "<p style=\"font-size:15px; color:blue\"><br>Status vom " + str(t1.tm_mday) + "." + str(t1.tm_mon) + \
      "." + str(t1.tm_year) + " um " + str(t1.tm_hour) + ":" + str(t1.tm_min).rjust(2, '0') + "  Uhr  \n"
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
HTMSel = HTMSel + " </select>\n\n\n  </p>\n\n</body>\n</html>\n"
#
HTXT = HTXT + "<br>\n" + HTMSel + "<br>\n"
#
HTXT = HTXT + "\n  </p>\n\n</body>\n</html>\n"
fp = open("HTML/index.html","w")
fp.write(HTXT)
fp.close()

# =================================================== append Auswahlliste an "Rennen_*.html"
FILES=os.listdir( "./HTML" )
for filename in FILES:
   #if filename.endswith('.html'):
   if fnmatch.fnmatch(filename, 'Rennen_*.html'):
      # print("append to '" + filename + "'" )
      fp = open("HTML/" + filename,"a")
      fp.write(HTMSel)
      fp.close()

print("__________________________________________________________________\n\nEs starten <b>" + str(Count_Ruderer) + "</b> Ruderer in <b>" + str(Count_Boote) + "</b> Booten ")
