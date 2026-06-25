# BMW XML Analyzer

Dit script analyseert BMW ISTA XML-bestanden (RG_META en RG_TRANS) en zet de relevante voertuig-,
ECU- en foutcode-informatie om naar een overzichtelijk Excel-bestand.

Het script is bedoeld voor gebruik via **Windows PowerShell**.

---

#Bestanden in deze repository

- `XML_BMW_EXE.py`  
  De hoofdcode die de XML-bestanden analyseert en een Excel-bestand genereert.

- `run_bmw_xml.ps1`  
  PowerShell-script dat de EXE of Python-code aanroept met de juiste parameters.

- `XML_BMW_EXE.spec` (optioneel)  
  PyInstaller-configuratie om zelf een EXE te bouwen.


#Vereisten
### Gebruik met EXE 
- Windows 10 / 11
- PowerShell

#EXE bouwen 
Maak (indien nog niet aanwezig) een map aan:
C:\BMW_XML

Plaats in deze map:
XML_BMW_EXE.pyrun_bmw_xml.ps1
RG_META_<VIN>.xml
RG_TRANS_<VIN>.xml

Open PowerShell in de map XML_BMW en voer uit:
- pip install pyinstaller
- pyinstaller --onefile XML_BMW_EXE.py

PowerShell-instelling (eenmalig):
- Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned

Script uitvoeren
Voer in PowerShell het volgende commando uit:
- .\run_bmw_xml.ps1 .\RG_META_<VIN>.xml .\RG_TRANS_<VIN>.xml

Na afloop wordt automatisch een Excel-bestand aangemaakt in:
C:\BMW_XML\BMW_XML_<VIN>.xlsx

Dit bestand bevat onder andere:
Voertuig- en metadata
ECU-overzicht
Foutcodes (DTC’s)
Context- en tijdinformatie
Samenvattende tabellen

Toekomstige uitbreidingen
Grafische gebruikersinterface (GUI)
Extra filter- en exportopties
Ondersteuning voor aanvullende XML-typen