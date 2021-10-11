#  python3
#
# needs OpenPYXL:
#   /usr/bin/python3 -c "import openpyxl; print(openpyxl)" 
# 
# has issues with SoftMaker Office !!! 
# XML support:
# from lxml import etree
import os, sys, sqlite3

# Excel
#========================================================================
from openpyxl import load_workbook

# globale Parameter
import LSglobal

Pfad_LS        = os.getcwd()
print("Starte aus Pfad '" + Pfad_LS + "'")

#========================================================================
# Verbindung zur Datenbank erzeugen
connection = sqlite3.connect( LSglobal.SQLiteFile )

# Datensatzcursor erzeugen
cursor = connection.cursor()

print("__________________________________________________________\n")
Pfad_Meldungen = Pfad_LS + '/' + LSglobal.MeldeDir
print("Suche Files in Pfad '" + Pfad_Meldungen + "'")

#os.chdir( Pfad_Meldungen )
#FILES=os.listdir( '.' )
FILES=os.listdir( Pfad_Meldungen )
for filename in FILES:
   print(filename)
   nFN = len(filename)
   Endung = filename[nFN-4:nFN]
   if(Endung.find("xlsx")>=0):
      print("__________________________________________________________\n")
      
      # open existent workbook:
      print("_________________________________________________________ " + filename )
      wb = load_workbook(Pfad_Meldungen + "/" + filename)
      # wb = load_workbook(filename = '../F2020/Meldungen/01__RVE_Thea.xlsx')
      # create a new Workbook:
      # from openpyxl import Workbook
      print(wb.sheetnames)
      ws = wb.active
      
      
      #========================================================================
      Vereinsname = ws['A3'].value
      Verein      = ws['G3'].value
      
      Adresse_1 =  ws['A5'].value
      Adresse_2 =  ws['A6'].value
      if( Adresse_2 == None): 
         Adresse_2 = ""
      
      sql = "SELECT EXISTS(SELECT 1 FROM verein WHERE kurz='" + Verein + "' LIMIT 1)"
      cursor.execute(sql)
      record = cursor.fetchone()
      
      if record[0] == 1:
         print (Verein +  " ist schon in der Datenbank!" )
      else:
         sql = "INSERT INTO verein VALUES(" \
            "'" + Vereinsname + "', '" + Verein + "', '" \
            + Adresse_1 + "', '" + Adresse_2 + "', " \
            "0.0)"
         # print(sql)
         cursor.execute(sql)
      
      sqlite_select_query = """SELECT count(*) from verein"""
      cursor.execute(sqlite_select_query)
      totalRows = cursor.fetchone()
      print("Total rows are:  ", totalRows[0])
      
      print(' ____________________________________ ')
      print(' ')
      
      # ========================================================================
      
      Vorname  = ws['A8'].value
      Name     = ws['C8'].value
      Phone    = ws['E8'].value
      email    = ws['H8'].value
      
      #   "name" TEXT, "vorname" TEXT, "telefon" TEXT, "verein" TEXT, "email" TEXT
      # And this is the named style:
      # cursor.execute("select * from betreuer where vorname=:Vorname and name=:Name and verein=:Verein ", {"Vorname": Vorname, "Name": Name, "Verein": Verein})
      # print( cursor.fetchone())
      
      #
      if(Name == None):
         Hname = Vorname.split()
         Hlen = len(Hname)
         Vorname = Hname[0]
         Name = Hname[Hlen-1]
         
      sql = "SELECT EXISTS(SELECT 1 FROM betreuer WHERE verein='" + Verein + "' and name='" + Name + "' LIMIT 1)"
      cursor.execute(sql)
      record = cursor.fetchone()
      
      if record[0] == 1:
         print ("Betreuer: " + Vorname + " " + Name + " (" + Verein + ") ist schon in der Datenbank!" )
      else:
         sql = "INSERT INTO betreuer VALUES( " \
            "'" + Vorname + "', '" + Name + "', '" + Verein + "', " \
            "'" + Phone + "', '" + email + "' )"
         cursor.execute(sql)
         connection.commit()
         print(sql)
      
      print(' ____________________________________ ')
      print(' ')
      
      # ========================================================================
      
      sqlite_select_query = """SELECT count(*) from ruderer"""
      cursor.execute(sqlite_select_query)
      tmp = cursor.fetchone()
      nRuderer = tmp[0]
      print("Anzahl der Ruderer in Datenbank:  ", nRuderer)
      
      sqlite_select_query = """SELECT count(*) from boote"""
      cursor.execute(sqlite_select_query)
      tmp = cursor.fetchone()
      nBoote = tmp[0]
      print("Anzahl der Boote in Datenbank:  ", nBoote)
      
      sqlite_select_query = """SELECT count(*) from r2boot"""
      cursor.execute(sqlite_select_query)
      tmp = cursor.fetchone()
      nR2Boot = tmp[0]
      print("Anzahl der Ruderer in Booten in der Datenbank:  ", nR2Boot)
      
      # indices für Ruderer (max. 4+)
      Ruderer = [None] * 5
      # =======================================================================   Loop über die Zeilen der Meldung   ==============
      Position = 11
      while Position > 7:
         Names = ","
         NR = 0
         BootNr = -1
         # ____________________________________________ Rennen und Bootsgattung
         Rennen   = ws['B' + str(Position)].value
         if(Rennen == None):
            Rennen = 0;
         print("Starte mit Eintrag für Rennen " + str(Rennen) + " ..." )
         if Rennen > 0:
            Comment  = ws['I' + str(Position)].value
            if( Comment == None):
               Comment = "-"
            #
            sql = "SELECT * FROM rennen WHERE nummer='" + str(Rennen) + "' "
            cursor.execute(sql)
            record = cursor.fetchone()
            # ToDo: check für 3 Rennen ohne Zuordnung
            Gender = record[2]
            # print(record[7])
            if(record[7] > 0):
               LGWi = 1
               Gewicht = -1.0;
            else:
               LGWi = 0
               Gewicht = 0.0;
               anz = Comment.find('LGW') + Comment.find('Lgw')
               # print(anz)
            #
            if(record[3] == 'All'):
               print("Undefiniertes Rennen, muß Comment auswerten!")
               Boot  = ws['H' + str(Position)].value
            else:
               Boot  = record[3]
               
            #____________________________________ Boot -> Anzahl Ruderer (NR)
            if   Boot == '1x':
               NR = 1
            elif Boot == '2x':
               NR = 2
            elif Boot == '2-':
               NR = 2
            elif Boot == '4x':
               NR = 4
            elif Boot == '4x+':
               NR = 5
            else:
               NR = 0
            
            BootNr = 0
            
            #Gender   = ws['E' + str(Position + iP)].value
            #LGW      = ws['D' + str(Position + iP)].value
            
            # setze status-Flag für das Rennen - es existiert hiermit eine Meldung
            if(record[6] == 0):
               sql = "UPDATE rennen SET status = 1 WHERE nummer='" + str(Rennen) + "' "
               cursor.execute(sql)
               connection.commit()
            #
            # =================================================================   Loop über die Ruderer   ==============
            for iP in range(NR):
              Vorname  = ws['C' + str(Position + iP)].value
              Name     = ws['D' + str(Position + iP)].value
              Jahr     = ws['E' + str(Position + iP)].value
              Verein2  = ws['F' + str(Position + iP)].value
              
              
              if( Jahr == None): 
                print( "in '" + filename + "' fehlt in Zeile " + str(Position) + " das Jahr")
              
              if(iP == 0):
                JahrBoot = Jahr
              elif(JahrBoot > Jahr):
                JahrBoot = Jahr
              
              if Gender == "w":
                 Gender = "F"
              elif Gender == "W":
                 Gender = "F"
              elif Gender == "f":
                 Gender = "F"
              elif Gender == "m":
                 Gender = "M"
              
              # wenn 
              if( Verein2 == None): 
                Verein2 = Verein
                
              if( Vorname == None): 
                print( "in '" + filename + "' fehlt in Zeile " + str(Position) + " der Vorname")
              if( Name == None): 
                print( "in '" + filename + "' fehlt in Zeile " + str(Position) + " der Name")
              
              sql = "SELECT EXISTS(SELECT 1 FROM ruderer WHERE verein='" + Verein2 + "' and name='" + Name + \
              "' and vorname='" + Vorname + "' LIMIT 1)"
              # print(sql)
              cursor.execute(sql)
              record = cursor.fetchone()
              
              if record[0] == 1:
                sql = "SELECT nummer FROM ruderer WHERE verein='" + Verein2 + "' and name='" + Name + \
                     "' and vorname='" + Vorname + "' "
                cursor.execute(sql)                
                RHelp= cursor.fetchone()
                Ruderer[iP] = RHelp[0]
                
                print ("Ruderer: " + Vorname + " " + Name + " (" + str(Jahr) + ", " + \
                   Verein + ") ist schon in der Datenbank! " + str(Ruderer[iP]) )
                   
                # nun Suche 
                sql = "SELECT bootNr FROM r2boot WHERE rudererNr=" + str(Ruderer[iP]) + " "
                #sql = "SELECT boot FROM ruderer WHERE verein='" + Verein2 + "' and name='" + Name + \
                #     "' and vorname='" + Vorname + "' "
                print(sql)
                cursor.execute(sql)
                tupleNr = cursor.fetchone()
                
                if(tupleNr != None  and  tupleNr[0] > 0):
                   BootNr = tupleNr[0]
                   print("Ruderer ist auch in Boot " + str(BootNr) + " gemeldet")
                   # ToDo: Check ob Rennen gleich ist!
                
                sql = "SELECT nummer FROM ruderer WHERE verein='" + Verein2 + "' and name='" + Name + \
                     "' and vorname='" + Vorname + "' "
                # print(sql)
                cursor.execute(sql)
                # myCol = cursor.fetchone()
                Names = Names + str(cursor.fetchone())
                Names = Names.replace("(", "")
                Names = Names.replace(")", "")
                # 
                # print( Names )
              else:
                nRuderer = nRuderer + 1
                Names = Names + str(nRuderer) + ","
                #
                #    sql = "INSERT INTO ruderer VALUES( " \
                #       + "'" + Vorname + "', '" + Name + "', '" + Gender + "', " + str(Jahr) + ", " \
                #      + str(LGWi) + ", -1.0, '" + Verein2 + "', " + str(nRuderer) + ", -1, 0 )"
                sql = "INSERT INTO ruderer VALUES( " + str(nRuderer) \
                   + ", '" + Vorname + "', '" + Name + "', '" + Gender + "', " + str(Jahr) + ", " \
                  + str(LGWi) + ", -1.0, '" + Verein2 + "', '', -1 )"
                print(sql)
                #  "( vorname , name , geschlecht, jahrgang, leichtgewicht, gewicht, verein, nummer, boot )"
                #+ "'" + Vorname + "', '" + Name + "', '" + Gender + "', " + str(Jahr) + ", " 
                cursor.execute(sql)
                connection.commit()
                # definiere den Ruderer-index für den Ruderer auf Platz iP
                Ruderer[iP] = nRuderer
              
            # print("Boot: '" + Boot + "': " + Names + " (" + str(JahrBoot) + ") in Rennen " + str(Rennen))
            
         if(NR > 0):
            if isinstance(Rennen, int):
               print("Rennen aus Excel: " + str(Rennen) )
            else:
               if(LGWi == 1):
                  sql = "SELECT nummer FROM rennen WHERE boot='" + Boot + "' and jahrgangmin <= " + str(JahrBoot) + \
                       " and jahrgangmax >= " + str(JahrBoot) + " and gewicht > 0 and gender = '" + Gender + "'"
               else:
                  sql = "SELECT nummer FROM rennen WHERE boot='" + Boot + "' and jahrgangmin <= " + str(JahrBoot) + \
                       " and jahrgangmax >= " + str(JahrBoot) + " and gewicht <= 0 and gender = '" + Gender + "'"
               # print(sql)
               cursor.execute(sql)
               tupleNr = cursor.fetchone()
               if tupleNr is None:
                  Rennen  = 1
               else:
                  Rennen = tupleNr[0]
               print("Rennen aus Syntax: " + str(Rennen) )
               
            
            if(BootNr == 0):
                print("_______________________________________________________")
                nBoote = nBoote + 1
                #    sql = "INSERT INTO boote VALUES( " \
                #      + str(nBoote) + ", 0, " + str(Rennen) + ", '" + Verein + "', '" + Names + "', " \
                #      + "0,   0, 0, 0,   0, 0,   0,   0, '" + Comment + "')"
                sql = "INSERT INTO boote VALUES( " \
                  + str(nBoote) + ", 0, " + str(Rennen) + ", " \
                  + "0,   0, 0, 0,   0, 0,   0,   0, '" + Comment + "')"
                cursor.execute(sql)
                connection.commit()
                #  nummer, startnummer, rennen, [-- vereine(TEXT), ruderer(TEXT-index),--] " \
                #  planstart   secstart sec3000 sec6000 zeit3000 zeit6000 zeit    abgemeldet (alles INTEGER)"
                #____________________________________ BootsNummer (nBoote) zu Ruderern
                print(Names)
                #
                for iP in range(NR):
                   # print(iP)
                   # sql = "UPDATE ruderer SET boot = " + str(nBoote) + " WHERE nummer = " + str(Ruderer[1 + iP])
                   nR2Boot = nR2Boot + 1
                   sql = "INSERT INTO r2boot VALUES( " \
                     + str(nR2Boot) + ", " + str(Rennen) + ", " + str(nBoote) + ", "  \
                     + str(Ruderer[iP]) + ", " + str(iP) + " )"
                   # r2boot:  "nummer, rennNr, bootNr, verein, rudererNr, platz INTEGER"
                   cursor.execute(sql)
                   connection.commit()
                # sql = "SELECT * FROM ruderer WHERE geschlecht LIKE '%," + Ruderer[1] + ",%' "
                #print(sql)
                #cursor.execute(sql)
                #for dsatz in cursor:
                #   print( dsatz[0], dsatz[1], dsatz[2] )
                print ("++++++")
            #      print("Rennen: '" + Rennen + "': " + Vorname + " " + Name + " (" + str(Jahr) + ")" + Comment)
            #print(ws[cellA].value + ' ' + ws['F' + str(Position)].value + ' ' + ws['G' + str(Position)].value + ' (' \
            #+ ws['E' + str(Position)].value + ') ' + str(ws['H' + str(Position)].value) )
         print(' --- ')
         if Position > 50:
            Position = 0
         elif NR == 0:
            Position = 0
         else:
            Position = Position + NR
            
         
      #sheet_ranges = wb['range names']
      #print(sheet_ranges['A1'].value)
      # ==============================================================================================================
      wb.close()
      print("=========================================================================================================\n")
   else:
      print("'" + Endung + "' ist kein Excel-Format !")
      print("=========================================================================================================\n")


# import xlrd
#
# book = xlrd.open_workbook('../Meldeliste_RVE.xlsx')
# sheet = book.sheet_by_index(1)
# cell = sheet.cell(1,1)
# print(cell)

