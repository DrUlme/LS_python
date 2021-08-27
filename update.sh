#!/bin/bash
#
# rclone ls dropbox:Langstrecke/H2020
# automatisiertes Holden der Zeiten
#rclone copy dropbox:Langstrecke/H2020/Start.trz Zeiten
# rclone copy dropbox:Langstrecke/H2020/3000.trz Zeiten
#rclone copy dropbox:Langstrecke/H2020/Ziel.trz Zeiten
#dropbox_uploader.sh download /Langstrecke/H2020/Start.trz Zeiten/Start.trz
#dropbox_uploader.sh download /Langstrecke/H2020/3000.trz Zeiten/3000.trz
#dropbox_uploader.sh download /Langstrecke/H2020/Ziel.trz Zeiten/Ziel.trz
# 
# Update Datenbank
# python3 LS_read_TRZ.py
# python3 LS_write_HTML.py
# 
# Erstelle Backup.zip
DBname=`grep DBName LSglobal.py | awk -F'"' '{print $2}'`
Name=`echo $DBname | awk -F'.' '{print $1}'`
DATESTR=`date +"_%d.%m._%H-%M"`

zip "Backup/"$Name"_"$DATESTR".zip" $DBname Zeiten/*

# zip -c 
