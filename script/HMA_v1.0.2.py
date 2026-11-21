#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# SPDX-License-Identifier: MIT
# Copyright (c) 2026 HMA contributors
"""
HMA – Harris Matrix Adapter v1.0.2

Read-only data preparer: turns tables (CSV, Excel .xlsx/.xls, LibreOffice Calc .ods, dBase .dbf, Access,
LibreOffice Base, SQLite) or database queries (Airtable, PostgreSQL, MySQL/MariaDB, FileMaker, Google Sheets)
into files that Harris Matrix programs can open: HMC project files (.hmcx) or HME files (.json).
The column mapping ("mapping schema") is chosen in the GUI and stored in hma_config.json.
HMA never writes to a database. The preview only serves to check the data; the matrix itself is drawn by HMC / HME.

Start (GUI):       python HMA_v1.0.2.py      (Windows: HMA.bat, macOS: HMA.command)
Start (no GUI):    python HMA_v1.0.2.py --cli --schema "Name" [--out FOLDER] file1.csv [file2.xlsx ...]
Language:          --lang de|en, environment variable HMA_LANG, or the EN | GER switch in the window.

Deutsch: Reiner Datenaufbereiter für Tabellen und Datenbank-Abfragen; erzeugt HMC-Projektdateien (.hmcx) oder
HME-Dateien (.json). HMA schreibt nie in eine Datenbank. Die Voransicht dient nur der Kontrolle der Daten.

Output: one <name>.hmcx (matrix.xml + project.xml) or <name>.json (HME) per table. The base skeleton of the .hmcx
(TEMPLATE_XML) holds only the required basic structure; HMA creates nodes and edges itself. HMA contains no program
code of HMC. For the HME output, the function hme_layout() re-implements the Auto-Layout procedure of HME (MIT
licence, Copyright (c) 2026 oeai-dac); the licence notice is given in front of that function and in the file
THIRD_PARTY_NOTICES.md.

License: MIT (see file LICENSE). No warranty: the software is provided "as is"; check generated matrices before
use. HMA is not affiliated with the makers of HMC, HME, Airtable, FileMaker, Microsoft Access, Google, PostgreSQL
or MySQL; all names mentioned belong to their respective owners.
"""

import base64
import csv
import io
import json
import os
import re
import sys
import traceback
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape

APP_NAME = "HMA – Harris Matrix Adapter"
APP_SHORT = "HMA"
APP_VERSION = "1.0.2"
SCRIPT_DIR = Path(__file__).resolve().parent

# ---------------------------------------------------------------------------
# Sprache / Language: Quelltexte sind deutsch (msgid); tr() liefert bei "de" den Text unverändert,
# bei "en" die Übersetzung aus _EN (ohne Eintrag: den deutschen Text).
# Reihenfolge: --lang / --lang=xx  >  Umgebungsvariable HMA_LANG  >  Einstellung in hma_config.json (ui.language)
#              >  Systemsprache (deutsch -> de, sonst en).  Ein Wechsel in der Oberfläche gilt ab dem nächsten Start.
# ---------------------------------------------------------------------------
LANGS = ("de", "en")


def _argv_language(argv):
    for i, a in enumerate(argv):
        if a.startswith("--lang="):
            return a.split("=", 1)[1].strip().lower()
        if a == "--lang" and i + 1 < len(argv):
            return argv[i + 1].strip().lower()
    return ""


def _argv_config(argv):
    for i, a in enumerate(argv):
        if a.startswith("--config="):
            return Path(a.split("=", 1)[1])
        if a == "--config" and i + 1 < len(argv):
            return Path(argv[i + 1])
    return None


def _config_language(argv):
    """Liest ui.language aus der Konfiguration (nur lesend; dieselben Orte wie pick_config_path())."""
    in_skript = SCRIPT_DIR.name.lower() in ("skript", "script")
    cands = [_argv_config(argv)]
    cands.append((SCRIPT_DIR.parent / "config" if in_skript else SCRIPT_DIR) / "hma_config.json")
    cands.append(SCRIPT_DIR / "hma_config.json")
    base = os.environ.get("APPDATA") or os.environ.get("XDG_CONFIG_HOME") or str(Path.home() / ".config")
    cands.append(Path(base) / "HMA" / "hma_config.json")
    for p in cands:
        try:
            if p and p.exists():
                v = json.loads(p.read_text(encoding="utf-8")).get("ui", {}).get("language", "")
                return str(v).strip().lower()
        except Exception:
            pass
    return ""


def _system_language():
    try:
        if os.name == "nt":
            import ctypes
            buf = ctypes.create_unicode_buffer(85)
            if ctypes.windll.kernel32.GetUserDefaultLocaleName(buf, 85):
                return buf.value.lower()
    except Exception:
        pass
    for var in ("LC_ALL", "LC_MESSAGES", "LANGUAGE", "LANG"):
        v = os.environ.get(var, "")
        if v:
            return v.lower()
    return ""


def detect_language(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    for cand in (_argv_language(argv), os.environ.get("HMA_LANG", "").strip().lower(), _config_language(argv)):
        if cand in LANGS:
            return cand
    return "de" if _system_language().startswith("de") else "en"


UI_LANG = detect_language()


def tr(text):
    """Übersetzt einen (deutschen) Quelltext in die Sprache der Oberfläche."""
    return text if UI_LANG == "de" else _EN.get(text, text)


_EN = {
    "ID Teil 1":
        "ID part 1",
    "ID Teil 2":
        "ID part 2",
    "ID Teil 3":
        "ID part 3",
    "ID Teil 4":
        "ID part 4",
    "Funde / Beschreibung":
        "Finds / description",
    "Kategorie":
        "Category",
    "Übergeordnete Gruppe":
        "Parent group",
    "Untergeordnete Gruppe":
        "Sub-group",
    "liegt über":
        "lies above",
    "liegt unter":
        "lies below",
    "zeitgleich mit":
        "contemporary with",
    "ident mit":
        "identical to",
    "Zeilenfilter":
        "Row filter",
    "_FUNDE":
        "_FINDS",
    "Datenbank_Export":
        "Database_Export",
    "DPAPI-Fehler":
        "DPAPI error",
    "Der API-Schlüssel ist ungültig oder abgelaufen. Bitte im Tab 'Datenbank' einen gültigen Personal Access Token eintragen.":
        "The API key is invalid or expired. Please enter a valid personal access token in the 'Database' tab.",
    "Keine Berechtigung. Der Token braucht das Recht 'data.records:read' und muss Zugriff auf diese Base haben (airtable.com/create/tokens).":
        "No permission. The token needs the scope 'data.records:read' and must have access to this base (airtable.com/create/tokens).",
    "Base oder Tabelle nicht gefunden. Bitte Base ID (app…) und Table ID (tbl…) prüfen.":
        "Base or table not found. Please check base ID (app…) and table ID (tbl…).",
    "Airtable lehnt die Anfrage ab (z. B. unbekannter Feldname oder View).":
        "Airtable rejects the request (e.g. unknown field name or view).",
    "Airtable hat die Anfragen begrenzt (Rate Limit). Bitte später erneut versuchen.":
        "Airtable has rate-limited the requests. Please try again later.",
    "Airtable meldet einen Fehler (HTTP {0}).":
        "Airtable reports an error (HTTP {0}).",
    "\n  Antwort von Airtable: {0}":
        "\n  Airtable response: {0}",
    "Tipp (macOS): Im Ordner 'Programme/Python 3.x' die Datei 'Install Certificates.command' doppelklicken und HMA danach neu starten.":
        "Tip (macOS): In the folder 'Applications/Python 3.x' double-click the file 'Install Certificates.command', then restart HMA.",
    "Mögliche Ursachen: Firmen-Proxy oder Virenschutz, der verschlüsselte Verbindungen untersucht, oder eine falsch gehende Systemuhr.":
        "Possible causes: a company proxy or antivirus program that inspects encrypted connections, or a wrong system clock.",
    "Es ist kein API-Schlüssel (Personal Access Token) eingetragen. Bitte im Tab 'Datenbank' eintragen.":
        "No API key (personal access token) has been entered. Please enter it in the 'Database' tab.",
    "Base ID und Table ID sind erforderlich. Bitte im Tab 'Datenbank' eintragen.":
        "Base ID and table ID are required. Please enter them in the 'Database' tab.",
    "Airtable-Rate-Limit erreicht – warte 30 Sekunden …":
        "Airtable rate limit reached – waiting 30 seconds …",
    "Das Sicherheitszertifikat von Airtable konnte nicht geprüft werden: {0}":
        "The security certificate of Airtable could not be verified: {0}",
    "Keine Verbindung zu Airtable möglich: {0}\nBitte Internetverbindung, Proxy/Firewall prüfen.":
        "No connection to Airtable possible: {0}\nPlease check internet connection, proxy/firewall.",
    "Fehler bei der Abfrage von Airtable: {0}":
        "Error querying Airtable: {0}",
    "  {0} Datensätze abgerufen …":
        "  {0} records retrieved …",
    "Die Abfrage hat keine Datensätze geliefert. Bitte Base, Tabelle und View prüfen.":
        "The query returned no records. Please check base, table and view.",
    "Airtable: {0} Datensätze":
        "Airtable: {0} records",
    "; ohne Werte in Airtable (als leere Spalte ergänzt): ":
        "; without values in Airtable (added as an empty column): ",
    ", Verknüpfungen: {0} aufgelöst":
        ", links: {0} resolved",
    ", {0} nicht auflösbar (Verweis in andere Tabelle)":
        ", {0} not resolvable (link to another table)",
    " – Achtung: Verknüpfte Datensätze (recXXXX) konnten nicht aufgelöst werden. Bitte Spalten zuordnen (ID) bzw. unter 'Datenbank' eine Auflösungs-Spalte wählen.":
        " – Warning: linked records (recXXXX) could not be resolved. Please map the columns (ID) or choose a resolution column under 'Database'.",
    "Bsp.: ":
        "e.g.: ",
    "Die Excel-Datei enthält keine Tabellenblätter.":
        "The Excel file contains no sheets.",
    "Zeitüberschreitung bei der Verbindung zu {0} ({1}:{2}). Server nicht erreichbar, Adresse/Port oder Firewall/VPN prüfen.":
        "Timeout while connecting to {0} ({1}:{2}). Server unreachable; check address/port or firewall/VPN.",
    "Der Servername '{0}' wurde nicht gefunden. Bitte Adresse prüfen (und ggf. VPN/Netzwerk).":
        "The server name '{0}' was not found. Please check the address (and VPN/network if applicable).",
    "{0} auf {1}:{2} lehnt die Verbindung ab. Läuft der Server, stimmen Adresse und Port, und erlaubt er Verbindungen von diesem Computer?":
        "{0} at {1}:{2} refuses the connection. Is the server running, are address and port correct, and does it allow connections from this computer?",
    "Die verschlüsselte Verbindung (SSL/TLS) zu {0} ist fehlgeschlagen: {1}":
        "The encrypted connection (SSL/TLS) to {0} failed: {1}",
    "Verbindung zu {0} ({1}:{2}) fehlgeschlagen: {3}":
        "Connection to {0} ({1}:{2}) failed: {3}",
    "Ungültiger Port: '{0}'.":
        "Invalid port: '{0}'.",
    "Die Verbindung zu {0} wurde unterbrochen oder lieferte eine unerwartete Antwort ({1}).":
        "The connection to {0} was interrupted or returned an unexpected response ({1}).",
    "Unerwartete Antwort des Servers ({0}). Ist das wirklich ein PostgreSQL-Server?":
        "Unexpected response from the server ({0}). Is this really a PostgreSQL server?",
    "Verbindung vom Server beendet":
        "Connection closed by the server",
    "Der Server unterstützt keine verschlüsselte Verbindung (SSL), die Einstellung verlangt sie aber. Verschlüsselung auf 'bevorzugt' oder 'aus' stellen.":
        "The server does not support an encrypted connection (SSL), but the setting requires it. Set encryption to 'preferred' or 'off'.",
    "Unerwartete Antwort des Servers. Ist das wirklich ein PostgreSQL-Server?":
        "Unexpected response from the server. Is this really a PostgreSQL server?",
    "Unbekannter Fehler":
        "Unknown error",
    "\nBenutzername oder Passwort stimmen nicht (oder der Server erlaubt diesen Zugang nicht, pg_hba.conf).":
        "\nUsername or password is incorrect (or the server does not allow this access, pg_hba.conf).",
    "\nDie angegebene Datenbank existiert nicht.":
        "\nThe specified database does not exist.",
    "\nDem Benutzer fehlt die Leseberechtigung für diese Tabelle.":
        "\nThe user lacks read permission for this table.",
    "\nDie Tabelle/View wurde nicht gefunden.":
        "\nThe table/view was not found.",
    "Der Server verlangt ein Anmeldeverfahren (SCRAM-SHA-256-PLUS), das HMA nicht kann.":
        "The server requires an authentication method (SCRAM-SHA-256-PLUS) that HMA cannot do.",
    "Ungültige Antwort des Servers bei der Anmeldung.":
        "Invalid response from the server during login.",
    "Der Server konnte sich bei der Anmeldung nicht ausweisen (SCRAM-Prüfung).":
        "The server could not identify itself during login (SCRAM check).",
    "Der Server verlangt ein Anmeldeverfahren, das HMA nicht unterstützt (Code {0}; z. B. Kerberos/SSPI). Bitte einen Benutzer mit Passwort verwenden.":
        "The server requires an authentication method that HMA does not support (code {0}; e.g. Kerberos/SSPI). Please use a user with a password.",
    "Bitte Server und Benutzer eintragen (Tab 'Datenbank').":
        "Please enter server and user (tab 'Database').",
    "Bitte eine Tabelle oder View wählen (Tab 'Datenbank', Knopf 'Tabellen laden').":
        "Please choose a table or view (tab 'Database', button 'Load tables').",
    "PostgreSQL: lese {0} …":
        "PostgreSQL: reading {0} …",
    "PostgreSQL, Tabelle {0}":
        "PostgreSQL, table {0}",
    "Die Abfrage hat keine Datensätze geliefert (leere Tabelle?).":
        "The query returned no records (empty table?).",
    "Unerwartete Antwort des Servers ({0}). Ist das wirklich ein MySQL/MariaDB-Server?":
        "Unexpected response from the server ({0}). Is this really a MySQL/MariaDB server?",
    "\nBenutzername oder Passwort stimmen nicht (oder der Benutzer darf sich von diesem Computer nicht anmelden).":
        "\nUsername or password is incorrect (or the user may not log in from this computer).",
    "MySQL/MariaDB: {0}{1}":
        "MySQL/MariaDB: {0}{1}",
    "Der Server verlangt das Anmeldeverfahren '{0}', das HMA nicht unterstützt. Bitte einen Benutzer mit 'mysql_native_password' oder 'caching_sha2_password' verwenden.":
        "The server requires the authentication method '{0}', which HMA does not support. Please use a user with 'mysql_native_password' or 'caching_sha2_password'.",
    "Unerwartete Antwort des Servers. Ist das wirklich ein MySQL/MariaDB-Server?":
        "Unexpected response from the server. Is this really a MySQL/MariaDB server?",
    "Der Server bietet keine verschlüsselte Verbindung (SSL) an, die Einstellung verlangt sie aber. Verschlüsselung auf 'bevorzugt' oder 'aus' stellen.":
        "The server does not offer an encrypted connection (SSL), but the setting requires it. Set encryption to 'preferred' or 'off'.",
    "Der Server verlangt für diese Anmeldung eine verschlüsselte Verbindung (caching_sha2_password), sie ist aber nicht möglich. SSL am Server aktivieren oder einen Benutzer mit 'mysql_native_password' verwenden.":
        "For this login the server requires an encrypted connection (caching_sha2_password), but it is not possible. Enable SSL on the server or use a user with 'mysql_native_password'.",
    "Unerwartete Antwort bei der Anmeldung.":
        "Unexpected response during login.",
    "MySQL/MariaDB: lese {0} …":
        "MySQL/MariaDB: reading {0} …",
    "MySQL/MariaDB, Tabelle {0}":
        "MySQL/MariaDB, table {0}",
    "Das Zertifikat des Servers wird nicht als vertrauenswürdig erkannt (z. B. selbstsigniert). Wenn du dem Server vertraust, 'Zertifikat nicht prüfen' aktivieren.":
        "The server's certificate is not recognised as trustworthy (e.g. self-signed). If you trust the server, enable 'Do not verify certificate'.",
    "Keine Verbindung möglich: {0}\nBitte Adresse, Internetverbindung, VPN, Proxy/Firewall prüfen.":
        "Connection not possible: {0}\nPlease check address, internet connection, VPN, proxy/firewall.",
    "Fehler bei der Abfrage: {0}":
        "Error during the query: {0}",
    "Bitte die Serveradresse eintragen (z. B. fm.beispiel.at).":
        "Please enter the server address (e.g. fm.example.org).",
    "Eine unverschlüsselte Verbindung (http://) ist nicht erlaubt, weil dabei das Passwort im Klartext übertragen würde. Bitte die Adresse ohne 'http://' angeben (FileMaker verwendet https).":
        "An unencrypted connection (http://) is not allowed because the password would be sent in clear text. Please enter the address without 'http://' (FileMaker uses https).",
    "Bitte den Datenbanknamen eintragen (Name der FileMaker-Datei ohne .fmp12).":
        "Please enter the database name (name of the FileMaker file without .fmp12).",
    "Der Server hat keine gültige FileMaker-Antwort geliefert (HTTP {0}). Stimmen Adresse und Datenbankname, und ist die Data API am Server aktiviert?":
        "The server did not return a valid FileMaker response (HTTP {0}). Are address and database name correct, and is the Data API enabled on the server?",
    "Benutzername oder Passwort stimmen nicht.":
        "Username or password is incorrect.",
    "Das Layout wurde nicht gefunden.":
        "The layout was not found.",
    "Die Datenbank konnte nicht geöffnet werden (Name prüfen).":
        "The database could not be opened (check the name).",
    "Die Data API ist für diese Datenbank/diesen Benutzer nicht aktiviert (Erweiterte Zugriffsrechte 'fmrest' im Rechtesatz des Benutzers).":
        "The Data API is not enabled for this database/user (extended privilege 'fmrest' in the user's privilege set).",
    "Die Sitzung ist ungültig geworden.":
        "The session has become invalid.",
    "Keine Datensätze gefunden.":
        "No records found.",
    "FileMaker (Fehler {0}): {1}":
        "FileMaker (error {0}): {1}",
    "Bitte den Benutzer eintragen.":
        "Please enter the user.",
    "Bitte ein Layout wählen (Tab 'Datenbank', Knopf 'Layouts laden').":
        "Please choose a layout (tab 'Database', button 'Load layouts').",
    "FileMaker: lese {0}, ab Datensatz {1} …":
        "FileMaker: reading {0}, from record {1} …",
    "FileMaker, Layout {0}":
        "FileMaker, layout {0}",
    "Bitte den Link zur Google-Tabelle eintragen.":
        "Please enter the link to the Google spreadsheet.",
    "Aus dem Link konnte keine Google-Tabelle erkannt werden. Bitte die Adresse aus dem Browser kopieren (…/spreadsheets/d/…).":
        "No Google spreadsheet could be recognised from the link. Please copy the address from the browser (…/spreadsheets/d/…).",
    "Google Sheets: lade Tabelle …":
        "Google Sheets: loading table …",
    "Die Google-Tabelle ist nicht öffentlich lesbar. In Google Sheets 'Teilen' → Allgemeiner Zugriff → 'Jeder, der den Link hat' → 'Betrachter' einstellen (oder 'Datei → Im Web veröffentlichen'). Private Tabellen kann HMA nicht abrufen; dort bitte als .xlsx/.csv herunterladen.":
        "The Google spreadsheet is not publicly readable. In Google Sheets set 'Share' → General access → 'Anyone with the link' → 'Viewer' (or 'File → Share → Publish to web'). HMA cannot fetch private spreadsheets; download those as .xlsx/.csv instead.",
    "Die Google-Tabelle wurde nicht gefunden. Bitte den Link prüfen.":
        "The Google spreadsheet was not found. Please check the link.",
    "Google Sheets antwortet mit Fehler {0}.":
        "Google Sheets responded with error {0}.",
    "Unbekannte Datenquelle: {0}":
        "Unknown data source: {0}",
    "API-Schlüssel":
        "API key",
    "Benutzer":
        "User",
    "Tabelle":
        "Table",
    "Datenbank":
        "Database",
    "Die Datei konnte nicht gelesen werden: {0}":
        "The file could not be read: {0}",
    "Die ODS-Datei ist passwortgeschützt und kann nicht gelesen werden.":
        "The ODS file is password-protected and cannot be read.",
    "Die ODS-Datei konnte nicht gelesen werden: {0}":
        "The ODS file could not be read: {0}",
    "Die Datei enthält keine Tabellen.":
        "The file contains no tables.",
    "Die ODS-Datei enthält keine Tabellenblätter.":
        "The ODS file contains no sheets.",
    "Die ODS-Tabelle ist ungewöhnlich groß (über 20 Millionen Zellen) und wird nicht gelesen. Bitte als .csv oder .xlsx speichern.":
        "The ODS table is unusually large (over 20 million cells) and will not be read. Please save as .csv or .xlsx.",
    "Die Excel-Datei (.xls) ist beschädigt oder hat einen nicht lesbaren Aufbau. Bitte in Excel oder LibreOffice öffnen und als .xlsx oder .csv speichern.":
        "The Excel file (.xls) is damaged or has an unreadable structure. Please open it in Excel or LibreOffice and save as .xlsx or .csv.",
    "Das ist kein Excel-Dokument im Format .xls (OLE2). Sehr alte Excel-Versionen (vor 5.0) werden nicht unterstützt – bitte in Excel als .xlsx oder .csv speichern.":
        "This is not an Excel document in .xls format (OLE2). Very old Excel versions (before 5.0) are not supported – please save as .xlsx or .csv in Excel.",
    "In der .xls-Datei wurde kein Tabellenstream gefunden (evtl. ältere Excel-Version oder verschlüsselt). Bitte als .xlsx oder .csv speichern.":
        "No table stream was found in the .xls file (possibly an older Excel version or encrypted). Please save as .xlsx or .csv.",
    "Diese Excel-Datei ist zu alt (vor Excel 5.0). Bitte als .xlsx oder .csv speichern.":
        "This Excel file is too old (before Excel 5.0). Please save as .xlsx or .csv.",
    "Die Excel-Datei ist passwortgeschützt/verschlüsselt und kann nicht gelesen werden.":
        "The Excel file is password-protected/encrypted and cannot be read.",
    "Die DBF-Datei ist leer oder beschädigt: {0}":
        "The DBF file is empty or damaged: {0}",
    "Die DBF-Datei enthält keine Spalten: {0}":
        "The DBF file contains no columns: {0}",
    "In {0} wurde keine Tabelle gefunden.":
        "No table was found in {0}.",
    "Die Datenbank konnte nicht gelesen werden ({0}): {1}":
        "The database could not be read ({0}): {1}",
    "Tabelle '{0}' konnte nicht gelesen werden: {1}":
        "Table '{0}' could not be read: {1}",
    "Access-Dateien können nur unter Windows gelesen werden. Bitte die Tabelle in Access als CSV exportieren.":
        "Access files can only be read on Windows. Please export the table from Access as CSV.",
    "Das Lesen der Access-Datei hat zu lange gedauert (Zeitüberschreitung).":
        "Reading the Access file took too long (timeout).",
    "Exit-Code {0}":
        "Exit code {0}",
    "\n\nDie Access-Datenbank-Engine (OLE DB) wurde nicht gefunden bzw. passt nicht zur Version von Office/Python (32 oder 64 Bit). Mit installiertem Microsoft Office ist sie normalerweise vorhanden; sonst 'Microsoft Access Database Engine 2016 Redistributable' installieren. Alternativ die Tabelle in Access als CSV exportieren (Externe Daten > Exportieren > Textdatei).":
        "\n\nThe Access database engine (OLE DB) was not found or does not match the version of Office/Python (32 or 64 bit). It is normally present when Microsoft Office is installed; otherwise install 'Microsoft Access Database Engine 2016 Redistributable'. Alternatively, export the table from Access as CSV (External Data > Export > Text file).",
    "Die Access-Datei konnte nicht gelesen werden ({0}):\n{1}{2}":
        "The Access file could not be read ({0}):\n{1}{2}",
    "Tabelle '{0}' ist leer.":
        "Table '{0}' is empty.",
    "Die Base-Datei konnte nicht geöffnet werden ({0}): {1}":
        "The Base file could not be opened ({0}): {1}",
    "Diese Base-Datei verwendet die eingebettete Firebird-Datenbank, die ohne Zusatzsoftware nicht gelesen werden kann. Bitte die Tabelle in LibreOffice als CSV exportieren (Tabelle öffnen, Daten kopieren in Calc, als CSV speichern) oder in Base unter Extras > SQL die Daten als Textdatei exportieren.":
        "This Base file uses the embedded Firebird database, which cannot be read without additional software. Please export the table from LibreOffice as CSV (open the table, copy the data into Calc, save as CSV) or export the data as a text file in Base under Tools > SQL.",
    "Diese Base-Datei ist nur eine Verknüpfung auf eine externe Datenbank. Bitte die zugrunde liegende Datenbank-Datei (z. B. Access oder SQLite) direkt auswählen.":
        "This Base file is only a link to an external database. Please select the underlying database file (e.g. Access or SQLite) directly.",
    "In dieser Base-Datei wurde keine lesbare eingebettete Datenbank gefunden.":
        "No readable embedded database was found in this Base file.",
    "In dieser Base-Datei wurden keine Tabellen gefunden.":
        "No tables were found in this Base file.",
    "Diese Base-Datei speichert die Daten in einem Format, das ohne Zusatzsoftware nicht gelesen werden kann (CACHED-Tabellen). Bitte die Tabelle als CSV exportieren.":
        "This Base file stores its data in a format that cannot be read without additional software (CACHED tables). Please export the table as CSV.",
    "Datei ist leer: {0}":
        "File is empty: {0}",
    "Trennzeichen: ":
        "Delimiter: ",
    "DBF-Datei konnte nicht gelesen werden: {0}":
        "The DBF file could not be read: {0}",
    "Tabelle ist leer: {0}":
        "Table is empty: {0}",
    "{0}, Tabelle {1}":
        "{0}, table {1}",
    "{0} ist keine echte Excel-Datei, sondern eine Webseite (HTML) mit der Endung .xls. Bitte in Excel öffnen und als .xlsx oder .csv speichern.":
        "{0} is not a real Excel file but a web page (HTML) with the extension .xls. Please open it in Excel and save as .xlsx or .csv.",
    "Die Datei konnte nicht gelesen werden ({0}): {1}":
        "The file could not be read ({0}): {1}",
    "ID: mindestens Teil 1 wählen":
        "ID: choose at least part 1",
    "'liegt über': Spalte wählen":
        "'lies above': choose a column",
    "Verweise bestehen aus":
        "References consist of",
    "{0} → '{1}' nicht in der Tabelle":
        "{0} → '{1}' not in the table",
    "Die Spalte '{0}' ist mehrfach zugeordnet ({1}) – jede Aufgabe braucht ihre eigene Spalte":
        "The column '{0}' is mapped more than once ({1}) – each task needs its own column",
    " und ":
        " and ",
    "ZIP enthält nicht genau matrix.xml und project.xml":
        "ZIP does not contain exactly matrix.xml and project.xml",
    "Datei nicht lesbar: {0}":
        "File not readable: {0}",
    "hmcnode passt nicht zu node '{0}'":
        "hmcnode does not match node '{0}'",
    "Node ohne ID":
        "Node without ID",
    "{0} doppelte Node-ID(s), z. B. {1}":
        "{0} duplicate node ID(s), e.g. {1}",
    "Kante {0} verweist auf fehlenden Node":
        "Edge {0} refers to a missing node",
    "Knoten-/Kantenzahl im Graph-Kopf stimmt nicht":
        "Node/edge count in the graph header does not match",
    "JSON nicht lesbar: {0}":
        "JSON not readable: {0}",
    "Liste '{0}' fehlt":
        "List '{0}' is missing",
    "doppelte Node-IDs":
        "duplicate node IDs",
    "doppelte Labels":
        "duplicate labels",
    "unbekannter Typ bei {0}":
        "unknown type at {0}",
    "Koordinaten fehlen bei {0}":
        "coordinates missing at {0}",
    "ungültige Kante {0}":
        "invalid edge {0}",
    "doppelte Kanten-ID {0}":
        "duplicate edge ID {0}",
    "Phase von {0} unbekannt":
        "phase of {0} unknown",
    "Objekt {0} verweist auf unbekannte Nodes":
        "object {0} refers to unknown nodes",
    "ohne Periode ":
        "without period ",
    "{0}\nliegt_unter: {1}\nliegt_ueber: {2}\nzeitgleich_mit: {3}\nident_mit: {4}":
        "{0}\nlies_below: {1}\nlies_above: {2}\ncontemporary_with: {3}\nident_with: {4}",
    "Automatisch":
        "Automatic",
    "Perioden-Blöcke":
        "Period blocks",
    "Strikt nach Tiefe":
        "Strictly by depth",
    "Ungültiger regulärer Ausdruck in den erweiterten Einstellungen: {0}":
        "Invalid regular expression in the advanced settings: {0}",
    "Zuordnungsfehler in {0}: ":
        "Mapping error in {0}: ",
    "Der Zeilenfilter lässt keine einzige Zeile übrig. Bitte die Filterwerte prüfen.":
        "The row filter leaves no rows. Please check the filter values.",
    "Die Tabelle enthält {0} verschiedene Projekte/Grabungen (ID-Präfixe)":
        "The table contains {0} different projects/excavations (ID prefixes)",
    "{0}: {1} Zeilen":
        "{0}: {1} rows",
    "Eine Matrix gilt immer für EINE Grabung. SE-Nummern wiederholen sich von Grabung zu Grabung; werden sie vermischt, entstehen doppelte IDs, falsche Verweise und Kreise. Lösung: im Tab 'Werte' beim Zeilenfilter 'Projekt-Präfix der ID' wählen und EIN Projekt anhaken.":
        "A matrix always applies to ONE excavation. SU numbers repeat from excavation to excavation; if they are mixed, duplicate IDs, wrong references and cycles result. Solution: in the 'Values' tab choose 'Project prefix of the ID' under row filter and tick ONE project.",
    "<graph id='g1'> im Template nicht gefunden.":
        "<graph id='g1'> not found in the template.",
    "Parsing-Fehler: Zeile {0} hat zusätzliche Felder (Spaltenverschiebung). Ursache sehr wahrscheinlich: Anführungszeichen im Export oder falsches Trennzeichen. Auszug: {1}":
        "Parsing error: row {0} has additional fields (column shift). Most likely cause: quotation marks in the export or wrong delimiter. Excerpt: {1}",
    "Verweis-Wert '{0}' gehört zu mehreren SEs ({1}, {2})":
        "Reference value '{0}' belongs to several SUs ({1}, {2})",
    "Doppelte IDs: {0} ID(s) kommen mehrfach vor":
        "Duplicate IDs: {0} ID(s) occur more than once",
    "'{0}' in den Zeilen {1}":
        "'{0}' in rows {1}",
    "Häufigste Ursache: die Tabelle enthält mehrere Grabungen/Projekte, und der Projekt-Präfix wurde von der ID entfernt. Lösung: im Tab 'Werte' den Zeilenfilter auf EIN Projekt setzen (oder den ID-Präfix im Tab 'Erweitert' leeren / ein weiteres ID-Feld zuordnen).":
        "Most common cause: the table contains several excavations/projects and the project prefix was removed from the ID. Solution: in the 'Values' tab set the row filter to ONE project (or clear the ID prefix in the 'Advanced' tab / map an additional ID field).",
    "Mehrdeutige Verweise: {0} Verweis-Wert(e) gehören zu mehreren SEs":
        "Ambiguous references: {0} reference value(s) belong to several SUs",
    "Die Spalte, auf die die Beziehungen verweisen, ist nicht eindeutig (z. B. SE-Nummern aus mehreren Grabungen). Lösung: Zeilenfilter auf EIN Projekt setzen.":
        "The column the relations refer to is not unique (e.g. SU numbers from several excavations). Solution: set the row filter to ONE project.",
    "zeitgleich":
        "contemporary",
    "{0} ident {1} (Selbstreferenz) – ignoriert":
        "{0} ident {1} (self-reference) – ignored",
    "{0} ident {1} – Ziel-SE existiert nicht – ignoriert":
        "{0} ident {1} – target SU does not exist – ignored",
    "{0} ident {1}, aber {2} nennt {3} nicht (bitte in der Datenbank prüfen)":
        "{0} ident {1}, but {2} does not mention {3} (please check in the database)",
    "{0} SEs werden als identisch zusammengeführt (ab {1}) – sehr ungewöhnlich, bitte die Spalte 'ident mit' prüfen":
        "{0} SUs are merged as identical (from {1}) – very unusual, please check the column 'identical to'",
    "{0} – Suffixe {1}; verwende Suffix von kleinster ID ({2})":
        "{0} – suffixes {1}; using suffix of smallest ID ({2})",
    "Ident-Gruppe {0} hat unterschiedliche Perioden: {1} – kleinste ID gewinnt ({2}: {3})":
        "Ident group {0} has different periods: {1} – smallest ID wins ({2}: {3})",
    "Ident-Gruppe {0} hat unterschiedliche Phasen: {1} – kleinste ID gewinnt ({2}: {3})":
        "Ident group {0} has different phases: {1} – smallest ID wins ({2}: {3})",
    "[SE {0}]":
        "[SU {0}]",
    "{0} zeitgleich {1}, aber {2} nicht zeitgleich {3}":
        "{0} contemporary {1}, but {2} not contemporary {3}",
    "{0} ({1}) gibt es im HME nicht – wird für die HME-Ausgabe weggelassen.":
        "{0} ({1}) does not exist in HME – omitted from the HME output.",
    "Eine SE hat die ID '{0}' – das kollidiert mit dem Node '{1}' ({2}). Bitte T/G/U deaktivieren oder die SE-ID ändern.":
        "An SU has the ID '{0}' – this collides with the node '{1}' ({2}). Please deactivate T/G/U or change the SU ID.",
    "{0} ({1}): Die gewählten SE-IDs kommen in der Tabelle nicht vor.":
        "{0} ({1}): The chosen SU IDs do not occur in the table.",
    "{0} 'zeitgleich'-Beziehung(en) gibt es im HME nicht – sie werden für die HME-Ausgabe weggelassen.":
        "{0} 'contemporary' relation(s) do not exist in HME – they are omitted from the HME output.",
    "{0} SEs im Kreis, z. B.: {1}   [betroffen: {2}]":
        "{0} SUs in a cycle, e.g.: {1}   [affected: {2}]",
    "Zyklische Stratigraphie: {0} Kreis-Gruppe(n), insgesamt {1} SEs":
        "Cyclic stratigraphy: {0} cycle group(s), {1} SUs in total",
    "A liegt über B, B über C und C wieder über A ist stratigraphisch unmöglich; HMC kann so eine Matrix nicht öffnen. Ursachen: Fehler in den Daten, falsch zugeordnete Spalten, oder mehrere Grabungen vermischt (Zeilenfilter!).":
        "A lies above B, B above C and C above A again is stratigraphically impossible; HMC cannot open such a matrix. Causes: errors in the data, wrongly mapped columns, or several excavations mixed (row filter!).",
    "Es wurde KEINE {0}-Datei erzeugt, weil die Daten logische Fehler enthalten, mit denen {1} die Datei nicht öffnen könnte:":
        "NO {0} file was created because the data contain logical errors that would prevent {1} from opening the file:",
    "   Hinweis: {0}":
        "   Note: {0}",
    "{0} liegt über {1}, ist aber (direkt oder über Dritte) auch zeitgleich – Widerspruch in den Daten":
        "{0} lies above {1}, but is also contemporary (directly or via third units) – contradiction in the data",
    "{0} Nodes – sehr große Matrix, HMC kann langsam werden. Besser einzeln nach Projekt/Zeitraum filtern.":
        "{0} nodes – very large matrix, HMC may become slow. Better to filter by project/period separately.",
    "  HME-Node-Typen: ":
        "  HME node types: ",
    "  HME: {0} Units ohne übergeordnete Gruppe -> {1} Phase(n) 'unknown-N' (je nach stratigraphischer Lage zwischen den datierten Gruppen)":
        "  HME: {0} units without parent group -> {1} phase(s) 'unknown-N' (depending on stratigraphic position between the dated groups)",
    "{0} Unit(s) ohne Gruppe liegen zwischen Gruppen, die sich stratigraphisch widersprechen; die Zuordnung zu 'unknown-N' ist dort nur ungefähr.":
        "{0} unit(s) without a group lie between groups that contradict each other stratigraphically; the assignment to 'unknown-N' is only approximate there.",
    "Units ohne übergeordnete Gruppe stehen im HME in einem eigenen Band am Ende. Option 'Undatierte Units in unknown-Phasen einordnen' aktivieren.":
        "Units without a parent group appear in HME in a separate band at the end. Enable the option 'Place undated units in unknown phases'.",
    "Die übergeordneten Gruppen (HME: Phasen) überlappen sich stratigraphisch (Zyklus zwischen den Gruppen); die Reihenfolge der Phasen folgt der mittleren stratigraphischen Tiefe. Bitte im HME unter 'Validate' prüfen.":
        "The parent groups (HME: phases) overlap stratigraphically (cycle between groups); the order of the phases follows the mean stratigraphic depth. Please check in HME under 'Validate'.",
    "Die Datei kann nicht in diesem Ordner gespeichert werden:\n  {0}\nGrund: {1}":
        "The file cannot be saved in this folder:\n  {0}\nReason: {1}",
    "Die erzeugte Datei hat die Selbstprüfung nicht bestanden und wurde verworfen:\n  - ":
        "The generated file failed the self-check and was discarded:\n  - ",
    "Template hat nicht die erwartete Struktur (graph g1 / data d3).":
        "Template does not have the expected structure (graph g1 / data d3).",
    "Die Datei kann nicht in diesem Ordner gespeichert werden:\n  {0}\nGrund: {1}\nMeist fehlt die Schreibberechtigung (z. B. Ordner unter 'C:\\Program Files') oder die Datei ist in einem anderen Programm (HMC) geöffnet. Bitte oben bei 'Speicherort (.hmcx)' einen Ordner wählen, in dem du schreiben darfst (z. B. Dokumente oder Desktop).":
        "The file cannot be saved in this folder:\n  {0}\nReason: {1}\nUsually write permission is missing (e.g. a folder under 'C:\\Program Files') or the file is open in another program (HMC). Please choose a folder you may write to (e.g. Documents or Desktop) under 'Save location (.hmcx)' above.",
    "keine Fehler":
        "no errors",
    "{0} Fehler gefunden":
        "{0} error(s) found",
    "Projekt: {0}":
        "Project: {0}",
    "Tabelle: {0}":
        "Table: {0}",
    "Blatt":
        "Sheet",
    "erste":
        "first",
    "Zeitpunkt: {0}":
        "Time: {0}",
    "Erzeugt mit: HMA v{0}":
        "Created with: HMA v{0}",
    "Spaltenzuordnung:":
        "Column mapping:",
    "  ID: ":
        "  ID: ",
    " (mit _ verbunden)":
        " (joined with _)",
    "Funde/Beschreibung":
        "Finds/description",
    "  Beziehungen verweisen auf: Spalte '{0}'":
        "  Relations refer to: column '{0}'",
    "  Interface-Werte (Kategorie): ":
        "  Interface values (category): ",
    "  Ignorierte Werte (übergeordnete Gruppe): ":
        "  Ignored values (parent group): ",
    "  {0} ({1}): aktiv":
        "  {0} ({1}): active",
    ", SEs: {0}":
        ", SUs: {0}",
    " (Standardverbindung)":
        " (default connection)",
    "  Ignorierte Werte (untergeordnete Gruppe): ":
        "  Ignored values (sub-group): ",
    "Projekt-Präfix der ID":
        "Project ID prefix",
    "  Zeilenfilter: {0} = {1}  ({2} von {3} Zeilen)":
        "  Row filter: {0} = {1}  ({2} of {3} rows)",
    "Ergebnis: {0} Nodes, {1} übergeordnete Gruppen, {2} untergeordnete Gruppen, {3} Kanten  ({4})":
        "Result: {0} nodes, {1} parent groups, {2} sub-groups, {3} edges  ({4})",
    "Hinweis: Verweise auf SEs, die nicht in der Tabelle stehen, wurden übersprungen (":
        "Note: references to SUs that are not in the table were skipped (",
    ") – normal bei Teilauszügen.":
        ") – normal for partial extracts.",
    "1. Node-Type Fehler (DEPOSIT / SURFACE)":
        "1. Node type errors (DEPOSIT / SURFACE)",
    "{0} Konflikt(e) gefunden (für jeden Konflikt wurde der Typ der numerisch kleinsten SE übernommen).":
        "{0} conflict(s) found (for each conflict the type of the numerically smallest SU was adopted).",
    "Konflikt {0}":
        "Conflict {0}",
    "Merged Node: {0}":
        "Merged node: {0}",
    "Ident-Gruppe (Base-IDs): {0}":
        "Ident group (base IDs): {0}",
    "Übernommener Typ (kleinste SE): {0}":
        "Adopted type (smallest SU): {0}",
    "Typen je SE:":
        "Types per SU:",
    "2. Gleichsetzung (ident)":
        "2. Equating (ident)",
    "3. Zeitgleich-Setzung":
        "3. Contemporaneity",
    "4. Gruppen (übergeordnet / untergeordnet)":
        "4. Groups (parent / sub-group)",
    "5. Stratigraphie und Größe":
        "5. Stratigraphy and size",
    "  {0} Nodes, {1} + {2} Gruppen, {3} Kanten":
        "  {0} nodes, {1} + {2} groups, {3} edges",
    "(nicht verwenden)":
        "(do not use)",
    "(kein Filter)":
        "(no filter)",
    "(Projekt-Präfix der ID, z. B. XTESTX)":
        "(project prefix of the ID, e.g. XTESTX)",
    "(zusammengesetzte ID)":
        "(combined ID)",
    "(automatisch)":
        "(automatic)",
    "Suchen:":
        "Search:",
    "Auswahl leeren":
        "Clear selection",
    "keine Auswahl":
        "no selection",
    "{0} ausgewählt":
        "{0} selected",
    "ausgewählt: ":
        "selected: ",
    "Quelle und Speicherort":
        "Source and save location",
    "Datenbank / Online":
        "Database / Online",
    "Datei (Tabelle / Datenbank):":
        "File (table / database):",
    "Blatt / Tabelle:":
        "Sheet / table:",
    "Datei wählen…":
        "Choose file…",
    "Speicherort:":
        "Save location:",
    "Ordner wählen…":
        "Choose folder…",
    "leer = neben der Tabelle":
        "empty = next to the table",
    "Mapping-Schema":
        "Mapping schema",
    "Schema:":
        "Schema:",
    "Speichern":
        "Save",
    "Speichern unter…":
        "Save as…",
    "Löschen":
        "Delete",
    "Export…":
        "Export…",
    "Import…":
        "Import…",
    "Gespeichert in: {0}  ·  der zuletzt verwendete Stand wird beim Beenden automatisch gemerkt":
        "Saved in: {0}  ·  the last used state is remembered automatically on exit",
    "Matrix erzeugen":
        "Create matrix",
    "Datei öffnen":
        "Open file",
    "Ordner öffnen":
        "Open folder",
    "Ausgabe:":
        "Output:",
    "HMC – Harris Matrix Composer":
        "HMC – Harris Matrix Composer",
    "HME – Harris Matrix Editor":
        "HME – Harris Matrix Editor",
    "bevorzugt (empfohlen)":
        "preferred (recommended)",
    "erforderlich":
        "required",
    "aus":
        "off",
    "erforderlich, Zertifikat prüfen":
        "required, verify certificate",
    "Datenquelle":
        "Data source",
    "Anbieter:":
        "Provider:",
    "Alle Anbieter werden nur gelesen – HMA verändert nie etwas in der Datenbank.":
        "All providers are read-only – HMA never changes anything in the database.",
    "Ausgabe":
        "Output",
    "Dateiname (ohne Endung):":
        "File name (without extension):",
    "Ordner: siehe 'Speicherort' oben (leer = Dokumente)":
        "Folder: see 'Save location' above (empty = Documents)",
    "Vorhandene Datei ohne Nachfrage überschreiben":
        "Overwrite an existing file without asking",
    "Verbindung testen / Daten abrufen":
        "Test connection / fetch data",
    "Beim Klick auf 'Matrix erzeugen' werden die Daten immer frisch aus der Datenquelle abgerufen. 'Daten abrufen' lädt sie vorab, damit du im Tab 'Spalten' die Zuordnung treffen kannst.":
        "When you click 'Create matrix' the data are always fetched fresh from the data source. 'Fetch data' loads them in advance so that you can make the mapping in the 'Columns' tab.",
    "Airtable-Verbindung":
        "Airtable connection",
    "PostgreSQL-Server":
        "PostgreSQL server",
    "MySQL-/MariaDB-Server":
        "MySQL/MariaDB server",
    "FileMaker-Server (Data API)":
        "FileMaker server (Data API)",
    "Speichern: verschlüsselt auf diesem Computer (nur für dieses Windows-Benutzerkonto lesbar, nicht im Schema)":
        "Save: encrypted on this computer (readable only by this Windows user account, not stored in the schema)",
    "Speichern ist nur unter Windows möglich (wird für diese Sitzung gemerkt)":
        "Saving is only possible on Windows (remembered for this session)",
    "Verschlüsselung:":
        "Encryption:",
    "SSL/TLS; 'bevorzugt' verschlüsselt, wenn der Server es anbietet":
        "SSL/TLS; 'preferred' encrypts if the server offers it",
    "Airtable-Link einfügen:":
        "Paste Airtable link:",
    "Übernehmen":
        "Apply",
    "Adresse aus dem Browser kopieren (…/appXXX/tblXXX/viwXXX): Base, Tabelle und View werden erkannt.":
        "Copy the address from the browser (…/appXXX/tblXXX/viwXXX): base, table and view are recognised.",
    "beginnt mit 'app'":
        "starts with 'app'",
    "beginnt mit 'tbl' (oder Tabellenname)":
        "starts with 'tbl' (or table name)",
    "View ID":
        "View ID",
    "optional, beginnt mit 'viw'; leer = alle Datensätze der Tabelle":
        "optional, starts with 'viw'; empty = all records of the table",
    "Felder":
        "Fields",
    "optional, Namen kommagetrennt wie in Airtable; leer = alle Felder":
        "optional, names separated by commas as in Airtable; empty = all fields",
    "Token erstellen: airtable.com/create/tokens, Berechtigung 'data.records:read', Zugriff auf die Base geben. Der Token wird nur einmal angezeigt.":
        "Create a token: airtable.com/create/tokens, scope 'data.records:read', give access to the base. The token is shown only once.",
    "Verknüpfungen als Text abrufen (Primärfeld der verknüpften Tabelle statt recXXXX)":
        "Fetch links as text (primary field of the linked table instead of recXXXX)",
    "Alle Werte kommen dann so an, wie Airtable sie anzeigt. Andere Spalten der verknüpften Tabelle: Lookup-Feld in Airtable anlegen. Ist das Häkchen aus, werden Verknüpfungen nur über die Auflösungs-Spalte (unten) ersetzt.":
        "All values then arrive as Airtable displays them. For other columns of the linked table, create a lookup field in Airtable. If the box is unticked, links are replaced only via the resolution column (below).",
    "Verknüpfungen auflösen mit:":
        "Resolve links with:",
    "Airtable liefert bei verknüpften Feldern (z. B. 'liegt über') interne Kennungen (recXXXX…). HMA ersetzt sie durch den Wert dieser Spalte. Automatisch = die im Tab 'Spalten' zugeordnete ID (bzw. 'Verweise bestehen aus'). Das funktioniert bei Verknüpfungen innerhalb derselben Tabelle.":
        "For linked fields (e.g. 'lies above') Airtable returns internal identifiers (recXXXX…). HMA replaces them with the value of this column. Automatic = the ID mapped in the 'Columns' tab (or 'References consist of'). This works for links within the same table.",
    "Name oder IP-Adresse, z. B. db.beispiel.at":
        "Name or IP address, e.g. db.example.org",
    "Port":
        "Port",
    "leer = Standard (":
        "empty = default (",
    "Name der Datenbank":
        "Name of the database",
    "am besten ein Benutzer mit reiner Leseberechtigung":
        "preferably a user with read-only permission",
    "Passwort":
        "Password",
    "Tabelle / View":
        "Table / view",
    "Tabellen laden":
        "Load tables",
    "HMA liest ausschließlich (schreibgeschützte Sitzung) und liest die gewählte Tabelle bzw. View vollständig. Für Auswertungen über mehrere Tabellen am besten in der Datenbank eine View anlegen und diese wählen.":
        "HMA only reads (read-only session) and reads the chosen table or view completely. For analyses across several tables it is best to create a view in the database and choose that.",
    "Hinweis: MySQL/MariaDB-Anbindung bisher nur gegen einen Nachbau getestet, nicht gegen einen echten Server.":
        "Note: the MySQL/MariaDB connection has so far only been tested against a mock-up, not against a real server.",
    "z. B. fm.beispiel.at (FileMaker Server mit aktivierter Data API)":
        "e.g. fm.example.org (FileMaker Server with Data API enabled)",
    "leer = Standard (443)":
        "empty = default (443)",
    "Name der FileMaker-Datei, ohne .fmp12":
        "Name of the FileMaker file, without .fmp12",
    "Konto mit dem Zugriffsrecht 'fmrest'":
        "Account with the access privilege 'fmrest'",
    "Layouts laden":
        "Load layouts",
    "Zertifikat des Servers nicht prüfen (bei selbstsigniertem Zertifikat)":
        "Do not verify the server's certificate (for a self-signed certificate)",
    "Gelesen werden die Felder des gewählten Layouts (alle Datensätze). Die Data API muss am Server aktiviert sein. Lokale FileMaker-Dateien (.fmp12) können nicht direkt gelesen werden – dort die Tabelle als CSV/Excel exportieren oder die Datei über FileMaker Server freigeben. Hinweis: bisher nur gegen einen Nachbau der Schnittstelle getestet.":
        "The fields of the chosen layout are read (all records). The Data API must be enabled on the server. Local FileMaker files (.fmp12) cannot be read directly – export the table as CSV/Excel or share the file via FileMaker Server. Note: so far only tested against a mock-up of the interface.",
    "Adresse der Tabelle aus dem Browser kopieren":
        "Copy the table address from the browser",
    "Die Tabelle muss für 'Jeden mit dem Link' lesbar sein (Teilen → Allgemeiner Zugriff → Betrachter) oder im Web veröffentlicht sein. Ist die Adresse mit '#gid=…' kopiert, wird genau dieses Blatt gelesen, sonst das erste. Private Tabellen kann HMA nicht abrufen – dort bitte als .xlsx oder .csv herunterladen.":
        "The spreadsheet must be readable for 'Anyone with the link' (Share → General access → Viewer) or be published to the web. If the address is copied with '#gid=…', exactly that sheet is read, otherwise the first one. HMA cannot fetch private spreadsheets – download those as .xlsx or .csv.",
    "Lese Liste …":
        "Reading list …",
    "Fehler":
        "Error",
    "{0} gefunden – bitte aus der Liste wählen":
        "{0} found – please choose from the list",
    "Spalten":
        "Columns",
    "Spalten automatisch erkennen":
        "Detect columns automatically",
    "Zuordnung leeren":
        "Clear mapping",
    "Noch keine Datei geladen.":
        "No file loaded yet.",
    "Tipps: In die Felder tippen filtert die Spaltenliste (z. B. \"liegt\"), dann Enter oder Pfeil-runter. Mauszeiger über eine Bezeichnung = Erklärung.":
        "Tips: typing into the fields filters the column list (e.g. \"lies\"), then Enter or arrow down. Mouse pointer over a label = explanation.",
    "Knoten":
        "Nodes",
    "Eindeutige ID der Stratigraphischen Einheit. Bis zu vier Spalten werden mit _ verbunden (z. B. SE-Nr + Typ + Datierung → 125_Vf_MA). Teil 1 ist Pflicht; Teil 2 wird erst nach Teil 1 freigeschaltet usw.":
        "Unique ID of the stratigraphic unit. Up to four columns are joined with _ (e.g. SU no. + type + dating → 125_Vf_MA). Part 1 is mandatory; part 2 is unlocked only after part 1, and so on.",
    "Anzeigename des Knotens, z. B. der SE-Typ.":
        "Display name of the node, e.g. the SU type.",
    "Node-Name":
        "Node name",
    "Der Inhalt wird zur Beschreibung (description) des Knotens.":
        "The content becomes the description of the node.",
    "Suffix an ID, wenn gefüllt:":
        "Suffix appended to the ID when filled:",
    "Welche Werte Interface (SURFACE) sind, wählst du im Tab 'Werte'. Alle anderen → DEPOSIT.":
        "Which values are Interface (SURFACE) you choose in the 'Values' tab. All others → DEPOSIT.",
    "Node-Type":
        "Node type",
    "Gruppierung":
        "Grouping",
    "Äußere Gruppe, z. B. Periode oder Phase. Einzelne Werte (z. B. 'unklar') kannst du im Tab 'Werte' ignorieren.":
        "Outer group, e.g. period or phase. You can ignore individual values (e.g. 'unclear') in the 'Values' tab.",
    "Innere Gruppe, z. B. Phase oder Befund; liegt in der übergeordneten Gruppe.":
        "Inner group, e.g. phase or feature; lies within the parent group.",
    "Stratigraphische Beziehungen":
        "Stratigraphic relations",
    "Mehrere Verweise in einer Zelle, getrennt durch , oder ;.":
        "Several references in one cell, separated by , or ;.",
    "ABOVE-Kante":
        "ABOVE edge",
    "Optional, umgekehrte Richtung: 'liegt unter X' ⇒ X liegt über diese SE.":
        "Optional, reverse direction: 'lies below X' ⇒ X lies above this SU.",
    "CONTEMPORARY-Kante":
        "CONTEMPORARY edge",
    "Identische SEs werden zu einem Node zusammengeführt (z. B. 3=45_Ps). Nicht dieselbe Spalte wie 'zeitgleich mit' verwenden!":
        "Identical SUs are merged into one node (e.g. 3=45_Ps). Do not use the same column as 'contemporary with'!",
    "Node-Merge":
        "Node merge",
    "Standard: die oben zusammengesetzte ID. Enthalten die Beziehungs-Spalten nur z. B. die SE-Nummer, hier diese Spalte wählen (muss innerhalb der Tabelle eindeutig sein).":
        "Default: the combined ID above. If the relation columns contain only e.g. the SU number, choose that column here (must be unique within the table).",
    "* erforderlich":
        "* required",
    "Werte":
        "Values",
    "Kategorie: welche Werte sind ein Interface (SURFACE)?":
        "Category: which values are an interface (SURFACE)?",
    "Übergeordnete Gruppe: welche Werte werden ignoriert?":
        "Parent group: which values are ignored?",
    "Untergeordnete Gruppe: welche Werte werden ignoriert?":
        "Sub-group: which values are ignored?",
    "Zeilenfilter: nur bestimmte Zeilen verwenden":
        "Row filter: use only certain rows",
    "Filtern nach Spalte:":
        "Filter by column:",
    "Enthält die Tabelle mehrere Grabungen/Projekte, genau EIN Projekt anhaken (sonst sind IDs und Verweise nicht eindeutig). Nichts angehakt = alle Zeilen.":
        "If the table contains several excavations/projects, tick exactly ONE project (otherwise IDs and references are not unique). Nothing ticked = all rows.",
    "Optional: T, G und U einfügen (Top surface, Interface to Geology, Unexcavated)":
        "Optional: insert T, G and U (Top surface, Interface to Geology, Unexcavated)",
    "Standardmäßig aus. Ohne Auswahl von Werten gilt: T über U und U über G.":
        "Off by default. Without choosing values: T above U and U above G.",
    "Top surface (T) einfügen":
        "Insert Top surface (T)",
    "Direkt unter T liegen diese SEs (ID wie in der Knoten-Gruppe):":
        "These SUs lie directly below T (ID as in the node group):",
    "Interface to Geology (G) einfügen":
        "Insert Interface to Geology (G)",
    "Direkt über G liegen diese SEs:":
        "These SUs lie directly above G:",
    "Unexcavated (U) einfügen":
        "Insert Unexcavated (U)",
    "Direkt über U liegen diese SEs:":
        "These SUs lie directly above U:",
    "Erste Zeilen der zugeordneten Spalten (die Matrix siehst du im Tab 'Voransicht')":
        "First rows of the mapped columns (you see the matrix in the 'Preview' tab)",
    "Voransicht":
        "Preview",
    "Aktualisieren":
        "Refresh",
    "Liest die Daten neu ein: bei 'Datenbank / Online' wird die Datenquelle neu abgerufen, bei einer Tabelle wird die Datei neu gelesen. Danach wird die Voransicht neu berechnet.":
        "Re-reads the data: for 'Database / Online' the data source is queried again, for a table the file is re-read. Afterwards the preview is recalculated.",
    "Alles anzeigen":
        "Show all",
    "Anordnung:":
        "Arrangement:",
    "Perioden-Blöcke: Die Perioden (HME: Phasen) liegen als Blöcke übereinander, innerhalb eines Blocks zählt die stratigraphische Tiefe.\nStrikt nach Tiefe: alle Units nach ihrer Tiefe in der gesamten Matrix.\nAutomatisch: bei HME Blöcke (wie der HME-Auto-Layout), bei HMC strikt nach Tiefe.\nDie Voransicht dient nur der Kontrolle der Daten; die Anordnung im Zielprogramm kann abweichen.":
        "Period blocks: The periods (HME: phases) are stacked as blocks; within a block the stratigraphic depth counts.\nStrictly by depth: all units by their depth in the whole matrix.\nAutomatic: blocks for HME (like the HME auto layout), strictly by depth for HMC.\nThe preview is only for checking the data; the arrangement in the target program may differ.",
    "Gruppen":
        "Groups",
    "Units (Klick, Pfeiltasten oder Mausrad)":
        "Units (click, arrow keys or mouse wheel)",
    "Karteikarte":
        "Index card",
    "Scrollen = Verschieben · ⌘ + Scrollen = Zoom · Ziehen = Verschieben · Hover = Beziehungen":
        "Scroll = pan · ⌘ + scroll = zoom · drag = pan · hover = relations",
    "Mausrad = Zoom · Ziehen = Verschieben · Hover = Beziehungen":
        "Mouse wheel = zoom · drag = pan · hover = relations",
    " · Datei hat sich geändert, neu eingelesen ":
        " · file has changed, re-read ",
    "rufe Daten aus der Datenbank ab …":
        "fetching data from the database …",
    " · Datenbank neu abgerufen ":
        " · database queried again ",
    " · Datei neu gelesen ":
        " · file re-read ",
    "Noch keine Tabelle geladen. Bitte im Tab 'Spalten' bzw. 'Datenbank' eine Tabelle laden.":
        "No table loaded yet. Please load a table in the 'Columns' or 'Database' tab.",
    "Die Spaltenzuordnung ist noch nicht vollständig:\n\n":
        "The column mapping is not complete yet:\n\n",
    "berechne …":
        "calculating …",
    "Zu viele Units für die Voransicht ({0}). Bitte im Tab 'Werte' einen Zeilenfilter auf ein Projekt setzen.":
        "Too many units for the preview ({0}). Please set a row filter to one project in the 'Values' tab.",
    "Voransicht nicht möglich:\n\n":
        "Preview not possible:\n\n",
    "Unerwarteter Fehler in der Voransicht:\n\n":
        "Unexpected error in the preview:\n\n",
    " · {0} Beziehung(en) gegen die Reihenfolge (orange)":
        " · {0} relation(s) against the order (orange)",
    " · {0} zeitgleich-Paar(e) widersprechen der Stratigraphie":
        " · {0} contemporary pair(s) contradict the stratigraphy",
    "{0} Units · {1} Beziehungen · {2} · Ausgabe {3}{4}{5}":
        "{0} units · {1} relations · {2} · output {3}{4}{5}",
    "Eine Unit in der Liste oder in der Matrix auswählen.":
        "Select a unit in the list or in the matrix.",
    "Typ":
        "Type",
    "Beschreibung":
        "Description",
    "{0}\nzeitgleich mit\n{1}":
        "{0}\ncontemporary with\n{1}",
    "{0}\nliegt über\n{1}":
        "{0}\nlies above\n{1}",
    "Erweitert":
        "Advanced",
    "CSV-Trennzeichen:":
        "CSV delimiter:",
    "auto = aus der Kopfzeile erkennen":
        "auto = detect from the header row",
    "Zeichenkodierung:":
        "Character encoding:",
    "auto = UTF-8, sonst Windows-1252 (Umlaute/ß/Sonderzeichen)":
        "auto = UTF-8, otherwise Windows-1252 (umlauts/ß/special characters)",
    "ID-Präfix entfernen (Regex):":
        "Remove ID prefix (regex):",
    "Leer = IDs unverändert übernehmen. Beispiele siehe Tabelle unten.":
        "Empty = keep IDs unchanged. See the examples in the table below.",
    "Trenner zwischen Verweisen (Regex):":
        "Separator between references (regex):",
    "Standard [;,]  – mehrere IDs in einer Beziehungszelle":
        "Default [;,]  – several IDs in one relation cell",
    "HME: Units ohne übergeordnete Gruppe in Phasen 'unknown-N' einordnen (nach stratigraphischer Lage)":
        "HME: place units without a parent group into 'unknown-N' phases (by stratigraphic position)",
    "Beispiele für 'ID-Präfix entfernen (Regex)'":
        "Examples for 'Remove ID prefix (regex)'",
    "Regex":
        "Regex",
    "Entfernt":
        "Removes",
    "Beispiel":
        "Example",
    "genau 6 Ziffern + _":
        "exactly 6 digits + _",
    "beliebig viele Ziffern + _":
        "any number of digits + _",
    "Großbuchstaben + _":
        "capital letters + _",
    "alles bis zum ersten _":
        "everything up to the first _",
    "Protokoll":
        "Log",
    "Protokoll speichern…":
        "Save log…",
    "Kopieren":
        "Copy",
    "Leeren":
        "Clear",
    "Tabelle(n) wählen":
        "Choose table(s)",
    "Tabellen und Datenbanken":
        "Tables and databases",
    "CSV / Text":
        "CSV / text",
    "dBase":
        "dBase",
    "Alle Dateien":
        "All files",
    "Speicherort wählen":
        "Choose save location",
    "Datenbank wird gelesen …":
        "Reading database …",
    "Datenbank-Datei":
        "Database file",
    "Excel-Datei":
        "Excel file",
    "Tabellen bzw. Blätter konnten nicht gelesen werden:\n{0}":
        "Tables or sheets could not be read:\n{0}",
    "Datenbank ({0}) – die Daten werden beim Erzeugen abgerufen.":
        "Database ({0}) – the data are retrieved when the matrix is created.",
    "leer = Ordner Dokumente":
        "empty = Documents folder",
    "Airtable-Link":
        "Airtable link",
    "In dem Text wurde keine Base ID (app…) gefunden.":
        "No base ID (app…) was found in the text.",
    "Übernommen: ":
        "Applied: ",
    "Rufe Daten ab …":
        "Fetching data …",
    "OK – {0} Datensätze, {1} Spalten":
        "OK – {0} records, {1} columns",
    "Tabelle konnte nicht gelesen werden":
        "Table could not be read",
    "{0} Zeilen, {1} Spalten":
        "{0} rows, {1} columns",
    " – {0} Zuordnung(en) automatisch erkannt":
        " – {0} mapping(s) detected automatically",
    " – {0} Dateien (gleiche Zuordnung für alle)":
        " – {0} files (same mapping for all)",
    "Konnte die Konfigurationsdatei nicht schreiben (Schreibschutz?).":
        "Could not write the configuration file (write-protected?).",
    "Schema '{0}' gespeichert.":
        "Schema '{0}' saved.",
    "Mapping-Schema speichern unter":
        "Save mapping schema as",
    "Name des Schemas (z. B. Name der Datenbank):":
        "Name of the schema (e.g. name of the database):",
    "Schema überschreiben":
        "Overwrite schema",
    "Schema '{0}' existiert bereits. Überschreiben?":
        "Schema '{0}' already exists. Overwrite?",
    "Schema löschen":
        "Delete schema",
    "Schema '{0}' wirklich löschen?":
        "Really delete schema '{0}'?",
    "Schema exportieren":
        "Export schema",
    "Schema exportiert: {0}":
        "Schema exported: {0}",
    "Schema importieren":
        "Import schema",
    "Import":
        "Import",
    "Datei konnte nicht gelesen werden:\n{0}":
        "File could not be read:\n{0}",
    "Name des Schemas:":
        "Name of the schema:",
    "Automatisch erkennen":
        "Detect automatically",
    "Bitte zuerst eine Datei wählen.":
        "Please choose a file first.",
    "{0} Spalten erkannt – bitte prüfen.":
        "{0} columns detected – please check.",
    "im HME nicht verwendet":
        "not used in HME",
    "Kategorie: jedem Wert einen HME-Node-Typ zuordnen":
        "Category: assign an HME node type to each value",
    "T, G und U einfügen – im HME nicht verwendet":
        "Insert T, G and U – not used in HME",
    "✗ Spalte doppelt vergeben":
        "✗ Column assigned twice",
    "● erforderlich":
        "● required",
    "(keine Datei geladen)":
        "(no file loaded)",
    "✓ gefunden":
        "✓ found",
    "✗ nicht in dieser Tabelle":
        "✗ not in this table",
    "✗ Dieselbe Spalte ist mehrfach als ID-Teil gewählt.":
        "✗ The same column is chosen several times as an ID part.",
    "Beispiel (erste Zeile): ":
        "Example (first row): ",
    "Mindestens Teil 1 wählen; danach wird Teil 2 freigeschaltet usw.":
        "Choose at least part 1; part 2 is unlocked afterwards, and so on.",
    "⚠ Die Tabelle enthält {0} Projekte (ID-Präfixe). Bitte im Tab 'Werte' einen Zeilenfilter auf EIN Projekt setzen.":
        "⚠ The table contains {0} projects (ID prefixes). Please set a row filter to ONE project in the 'Values' tab.",
    "Spalte '{0}': Es befinden sich keine Werte in dieser Spalte.":
        "Column '{0}': there are no values in this column.",
    "Die Spalte '{0}' hat {1} verschiedene Werte – dafür ungeeignet.":
        "The column '{0}' has {1} different values – unsuitable for this.",
    "Spalte '{0}' – {1}":
        "Column '{0}' – {1}",
    "Noch keine Tabelle geladen – gespeicherte Werte werden angezeigt.":
        "No table loaded yet – saved values are shown.",
    "{0}   ({1} Zeilen)":
        "{0}   ({1} rows)",
    "{0}   (nicht in dieser Tabelle)":
        "{0}   (not in this table)",
    "Deposit (Standard)":
        "Deposit (default)",
    "Bitte im Tab 'Spalten' eine Kategorie-Spalte wählen. Ohne Auswahl sind alle Units Deposit.":
        "Please choose a category column in the 'Columns' tab. Without a selection all units are Deposit.",
    "jedem Wert einen HME-Node-Typ zuordnen (nicht zugeordnet = Deposit)":
        "assign an HME node type to each value (unassigned = Deposit)",
    "Bitte im Tab 'Spalten' eine Kategorie-Spalte wählen. Ohne Auswahl sind alle SEs DEPOSIT.":
        "Please choose a category column in the 'Columns' tab. Without a selection all SUs are DEPOSIT.",
    "angehakt = Interface (SURFACE), sonst Deposit":
        "ticked = interface (SURFACE), otherwise Deposit",
    "Bitte im Tab 'Spalten' eine übergeordnete Gruppe wählen.":
        "Please choose a parent group in the 'Columns' tab.",
    "angehakt = ignorieren (z. B. 'unklar'), die SE landet dann in keiner Gruppe":
        "ticked = ignore (e.g. 'unclear'), the SU then ends up in no group",
    "Bitte im Tab 'Spalten' eine untergeordnete Gruppe wählen.":
        "Please choose a sub-group in the 'Columns' tab.",
    "angehakt = ignorieren, die SE landet dann in keiner untergeordneten Gruppe":
        "ticked = ignore, the SU then ends up in no sub-group",
    "Kein Filter gewählt – alle Zeilen werden verwendet.":
        "No filter selected – all rows are used.",
    "angehakt = diese Zeilen verwenden":
        "ticked = use these rows",
    "FEHLER":
        "ERROR",
    "Es wurde KEINE":
        "NO",
    "Noch kein Protokoll vorhanden.":
        "No log available yet.",
    "Protokoll speichern":
        "Save log",
    "Protokoll_{0}.txt":
        "Log_{0}.txt",
    "Textdatei":
        "Text file",
    "Protokoll in die Zwischenablage kopiert.":
        "Log copied to the clipboard.",
    "Öffnen":
        "Open",
    "HMC: Datei wird neu geöffnet …":
        "HMC: file is being reopened …",
    "HMC: Neu öffnen nicht möglich ({0}). Bitte die Datei selbst öffnen.":
        "HMC: reopening not possible ({0}). Please open the file yourself.",
    "Bitte im Tab 'Datenbank' ausfüllen: ":
        "Please fill in the 'Database' tab: ",
    "OK – {0} Datensätze":
        "OK – {0} records",
    "Bitte zuerst eine Tabelle wählen.":
        "Please choose a table first.",
    "Die gewählte Tabelle existiert nicht mehr am gemerkten Ort.\nBitte die Tabelle neu wählen.":
        "The chosen table no longer exists at the remembered location.\nPlease choose the table again.",
    "Keine Daten geladen":
        "No data loaded",
    "Zuordnung prüfen":
        "Check mapping",
    "Datei existiert bereits":
        "File already exists",
    "{0} existiert bereits in\n{1}\nund ist gerade in Harris Matrix Composer geöffnet.\n\nJa = überschreiben; HMC-Fenster wird OHNE Speichern geschlossen und die neue Datei an derselben Stelle wieder geöffnet\nNein = unter neuem Namen speichern\nAbbrechen = diese Datei überspringen":
        "{0} already exists in\n{1}\nand is currently open in Harris Matrix Composer.\n\nYes = overwrite; the HMC window is closed WITHOUT saving and the new file is reopened in the same place\nNo = save under a new name\nCancel = skip this file",
    "{0} existiert bereits in\n{1}\n\nJa = überschreiben\nNein = unter neuem Namen speichern\nAbbrechen = diese Datei überspringen":
        "{0} already exists in\n{1}\n\nYes = overwrite\nNo = save under a new name\nCancel = skip this file",
    "HMC wird geschlossen (ohne Speichern) …":
        "HMC is being closed (without saving) …",
    "HMC konnte nicht automatisch geschlossen werden. Die Datei wird trotzdem erzeugt; bitte in HMC danach schließen und neu öffnen.":
        "HMC could not be closed automatically. The file is created anyway; please close and reopen it in HMC afterwards.",
    "Verarbeite: ":
        "Processing: ",
    "Datenbank ({0})":
        "Database ({0})",
    "→ gespeichert: {0}":
        "→ saved: {0}",
    "FEHLER bei ":
        "ERROR at ",
    "FEHLER bei {0}\n{1}":
        "ERROR at {0}\n{1}",
    "FEHLER (unerwartet):":
        "ERROR (unexpected):",
    "Keine Datei erzeugt für: {0} – Details im Protokoll":
        "No file created for: {0} – see the log for details",
    "Es wurde keine Datei erzeugt, weil die Daten logische Fehler enthalten.\n\nDetails und Lösungsvorschläge stehen im Tab 'Protokoll'.":
        "No file was created because the data contain logical errors.\n\nDetails and suggested solutions are in the 'Log' tab.",
    "✓ {0} von {1} Datei(en) verarbeitet":
        "✓ {0} of {1} file(s) processed",
    " – {0} Warnung(en), siehe Protokoll":
        " – {0} warning(s), see log",
    " – keine Warnungen":
        " – no warnings",
    "Sprache der Meldungen (sonst: Umgebungsvariable HMA_LANG, Konfiguration, Systemsprache)":
        "Language of the messages (otherwise: environment variable HMA_LANG, configuration, system language)",
    "Name eines gespeicherten Mapping-Schemas (sonst: zuletzt verwendeter Stand)":
        "Name of a saved mapping schema (otherwise: last used state)",
    "Speicherordner für die Ausgabedatei":
        "Output folder for the generated file",
    "Ausgabe: hmc = Harris Matrix Composer (.hmcx), hme = Harris Matrix Editor (.json)":
        "Output: hmc = Harris Matrix Composer (.hmcx), hme = Harris Matrix Editor (.json)",
    "Schema nicht gefunden. Vorhanden:":
        "Schema not found. Available:",
    "Verarbeite: {0}":
        "Processing: {0}",
    "FEHLER:":
        "ERROR:",
    "Die Oberfläche benötigt tkinter (gehört zur Standard-Python-Installation von python.org).":
        "The user interface requires tkinter (part of the standard Python installation from python.org).",
    "Alternativ: python {0} --cli --help":
        "Alternatively: python {0} --cli --help",
}


# ---------------------------------------------------------------------------
# Grundgerüst der .hmcx (Schlüsseldefinitionen + Resources-Block).
# Der Inhalt von <graph id="g1"> (Knoten, Kanten) wird beim Export vollständig neu erzeugt.
# ---------------------------------------------------------------------------
TEMPLATE_XML = r'''<?xml version="1.0" encoding="UTF-8" standalone="no"?>
<graphml xmlns="http://graphml.graphdrawing.org/xmlns/graphml" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://graphml.graphdrawing.org/xmlns/graphml http://www.yworks.com/xml/schema/graphml/1.0/graphml-attributes.xsd">
  <key extension.type="hmc" for="node" id="d0"/>
  <key extension.type="hmc" for="edge" id="d1"/>
  <key extension.type="hmc" for="graph" id="d2"/>
  <key for="graphml" id="d3" yfiles.type="resources"/>
  <graph edgedefault="directed" id="g1" parse.edges="0" parse.nodes="0" parse.order="free">
    <data key="d2"/>
  </graph>
  <data key="d3">
    <Resources xmlns="http://www.yworks.com/xml/graphml"/>
  </data>
</graphml>
'''

ICON_64_B64 = "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAMh0lEQVR42u2be3RV1Z3HP/ucc9/JzfOSmASTQIAkQEIwWC1ipfhAqghqxWpb/aMza7R0ZtaMXbNmLVjFyuq0dfUvnXZNx9VZq60MKlarUJRlFYuiAqEkhEDCI4l5QMjj3jzu49x7z9nzxw2RR4Dcm3MtLuf3z13rnr3Pb+/v2b/Xd+8NX3IRyTQuf2RzqRo3aqQirhOInGtpIhLpF6Y8bWhqU/uWDZ2WAVDx6CYvhrIeyYNA3Rfkw/4VwTZU8/kTL24aSREAKWav27xeIDcCvi/oCu9HiKdPbN3wSxByygDUfOdZTzAa/q2A+y21NyEoLS4i25sx6fPAyBidPb1IKa0G4lW33fVY0+9+GLwqADXfedYTioZ2g6i3ehQPrPw6c8pLsSmTL7yYKWk71ckf3n43Hc5uv8vuWn4xCNrFyz4Ue+Y36Zi80+lgTnkZ99T5KPDaJ23TNxJlOwKn00EkolvsJFkSjIZ/C/LB881BOb/R7HWb1yN5KB3G6LAlJu2xq5dtc+7ZubZpWAX3V6x75onz/5sAYN7qn2UKKTd8CUL/popHN3kvAcBw6z9AMCNdWvVYFIBg1Lhsm6BuXNA2TeIjrjx5qQ9I09I/J5GITtupDrYjsSsCoQgUoYAA0zSRpiRqSlpPdlpu/5MYwzeBnwKo5zI8RcqfpnvtBUZHyfR4yMjIwETgHw0SjOiomo1o3KC9q4e9BxsZC4bSPZTCnIUrfuNvfndEA1BNWZsuTaqqMKeslEXV8ygrKSIWi3G49QQHm49ydnAIgNzsbGoq57Coeh6PPbCaM/0DHGpp5cjxk8Ri8bQsAQWjBujSAITkOqtTj2xvJouqK6mtmoPL6WQoMMzuj/dz6GjbJUt8KBBg98f72XvwENUVs1i8oIqVX1vK8ptv5OiJU+xvOsKgP2Dp+ExE8Wc+QEpfcmXRlTK961hUPY+5s8oAaP+0m/1NR8YzvCv3j0ZjHGpp5VBLK4W+fOprqqmpnEtt1Vw6e05zqKWV1vZOpGlOe6yKmXD4GoBUpIZMHQG3y0VNZQV11ZVkeTMJhsLsO9TMweYWRsaCKb3zTP8A2//8F97bu5+FlRXULahizZ3LGQuGaG47ScPhI4xOw1dIRWqTZIJTs+nbv/oV5s66HlVREULgdjkTodQwiUZ12to72bO/AcOY/pcKhsPsbzpCVmYm8+eUk+F2sfSGGpbeUEMoHEFKiWEatJ36lHf2fpK0zqQBqJtfyZKaKm6cnY2YZNGYEtxOOwP+AA2HWyyx17r5lSxZWDmJzqwJnVkZbgYCw0nrTBqAgrxcSn1uqoo8l21zdiRKYX6uZQ4rnTqVFDwdylXchSIS7Syso9OmM3kApMS8ijc3ZaKddaVc+nQmbQJ9g0N09Ido8don/SqmhM7+IH0Dg5bNP506kwbgr0eOkZuVxcjY9WiKhtPpQFEEUkoiEZ2YEaf11KccPHLMOoLvPJ3nRx7TNIlEosTN1HUmDYBhmOza8xF7Dzby0DfuQlFV9hw4yFdqF2K3aeza/QmtpzoszdrO6dy156OJ/8qKi3jg7hUEI2G2bt/FyOhYaglRKp1ys7P57tp78GZ4eHnH2+w71Mz/bPsj/UMB1tz5dW67aQnCSic4iXT09LLlzZ24nC4eu/9eZuTlfD4AFPry+faaVSiqypY3dtJ1ug+AsWCIF/+4g8ajrdxUt5AHV92B02FPKwin+wZ48fUdmFLyyOpVFBX40gvA9UWFfGv1SmLxGFte38HZi5yOYZi89f6HvPX+h5QVF/HYA6vx5aZ3/2TAH+D3r+0gpOt8a/XdlM8sTi6zBchdsPw2ELddXNiUlRRTVlJEoS+fhZVzueOWmwnrOg2HW4gbJsOT2J0QAqfDQSgSYWZRIYsXVOFxu/BmeCj05eN0OCftN50CrNCXz2gwyIz8XOqqK3FPRZ/g/aHm3bsv6wTvv2v5BRS2qiTYG6fm4a6lN16Wwj6/nwAUVWHpogUY4xWc1dT35PrmT1mfZiWF/XlT31boU6yksD9v6tsKfQpfclGspLA/b+rbCn3a1SjsK+3jXUxhp9oPC6j2VPVdNgq8tutdSouKyM7KnPR5YHiUzt7eKfeTUqJHwwwODdLT14OUII0E4ytUDSEFQrWhKDZU1TblsxupjvOqAEiZSDfpSbpypb2rC/1EgFh0hHg8SCwaxDCiiS3KKe7iqaodm92Dpnmw2b04HNkIRbVsnCkXQ5efuEk41E841EdMH0ZePFkhUNxOlEwnwuVI/OW0JfpGYonfsI45GsEM6RiGjhHWgaHx7go2uxe3pwCnKx8hVEvGrVmx7xwM9TI23IFpxifoGa3Eh73UhzYjC7UwByXLjVCnFnSkYWIOhzDO+ImfHSba2U+8e4CoHiCqB1DUU2R6y3G7C5kuna9N82ASgaEWIuFETWArycOxeDaOyhKEI/VXC1VBzc1Azc3AXj0TNyD1OPqxbvSGk8R6Bhn2t6FHhsjJq2I6KEwLgLHhDiLhQRSPA8836nHMKyZdIhwaztoynLVlRI91M7ajgUhogNHhDjKzyq3NA6YMwFg3AN6Hl6V18heLvbKEzIdvSYxhtAfLE6GkXP54iPvbifk3BGBcRrd+QLSl63ObcrSli9GtH1wjUQAwQzojr36E7eM2HItnJZzgeIizLMxGYujHuok0nCTeO2TZey0B4Cf/8n2efeF3+HsGifUMEtxxALU4D3uZD60gB3WGFyU7I7kw6A9inA0Q7wsQ7ezH6BlEjm8O5GZ5+eH3vsu//+L5awOAh+6+nftW3Mqf3v+QV3e9x4Hmo8S7Boh3DVxobxlOFLcT4bEDAuGyjydAUZASGYpiBsOYwUvzdk3TqK+p4pt3rWDlrTfjsNuvHQAAHHY7a+9Yzto7lhMKR/iosZnGY620nerkWHsHfQN+YmMRzLHIVd9l0zQK8nOoLC9j3qwyairncnPtgoldaCtFs4qYzM/J5rPzAk5W3FTPipvqLyiGBgIBBvzDjIyOETeMibNAGR43mqrizcwgPyeL/Ozsq9Lq/X7/tQPA32/8Cc9tfIrighlXJC99OTn4cqbPEnef6ecHz/zcOgCEKeLTOCBCU+tx7v7eP/Lt+1bx8Ko7ub6oMC3hr7PnDFv/9Da/f2Mn4WnyCcIU8QkATIWzYhq5zH23f403/vwXfv3Sa7zwyussmFPBLfW13LhwPlWzy8nLzkrpvYOBYY6ebGdfUzN7DjTSfPwkUkqEEKy5/TZef2d36mFVcOa8FSBPT2c//xf/9s88vvYetrz5Ftt3f0BT63GaWo/zS7YB4MvJYWZRAQV5uRTk5ZLhceNyOnHYE7mCHo0RjkQYC4boGxyib3CI7tN9nB260M5dTgf3Ll/GI/euZMGc2dMCwJSidwIAQ1ObVMOU0ymrFs6t4D/+dT0/Wv937D/cwp6GRhqPtdHa3kG/35+S08r0uJlbXkpt5VyW3VDLkoXVOB0OS/IqoYjDEwC0b9nQWbHux4dI8kqMEApSmkRjMey2xNd0Ohwsq69jWX3dBU7rTH8/ZwYGGfAHCIYjhMJh9PFDkA6bhtvlwuNy4svNoTA/j0Jf3hWdqh6NTowhhfk3nNy6oeuis8LiFYRMCgCbzUM0OsrTz/2ap//pH9DUyVmakkIfJYXW3bqJGwZPP//f42PISOWc6LZLiiGHNJ4D+pJ5jze7AqFovLTzHdY++RS79zVgmumrDE1T8t4nDax54ile3vkOQlHx5lQky+KcRTV+NemVmdnrfrxewHNJVWb6MAF/G0Y8DMDMwhmsvPWrLF1cS/2CqmnbbETXOdB8lA8aGnlrz4d0n+kHQNVcZOfOw273Jvv5nzzx0sZfXebOkBQV6zb/L8h1yRKioWAvwbFejPhnqa6mqpSVFDGvvJTS4kJm5OZSkJ+H2+kg0+NGUZTxL2syGgwRiuj0DQzSP+Sno+c0re2ddHT3Ejc+2/hQNSceTxHujKKk7V/CH06+tPGCKzOTXpoKR8PvSViSCksY1UfQI0PoET/xWPBSdjjpay4CzebB4czB4czF7vCmFqyE3BcKyuW9b24KXfXaXNnjm5xaWH0B5KPTI01N4rEQ8ViQeFxHGgm6W0qJlAZSmhOeXAgVoSgoqgNFsaFpLjSbG83mRkyXt5Fscztcj0/p2tyF5vDME0h+lM6rNGmWPiHZdPzljf+V1MXJ82Xe6p9lGi79+4lrJnLxF2DSEsFBEK84DOM/j7yyacyyy9OzH948U5Gy1pTmdUKhGBMn18Yed0Sa9ChCOW0K0Xguyfl/mYL8H85jsx1NNWgfAAAAAElFTkSuQmCC"

# ---------------------------------------------------------------------------
# Schema: welche Spalte wird wofür verwendet
# ---------------------------------------------------------------------------
# HME-Node-Typen, die zusätzlich zu "interface" (= Kategorie-Werte für SURFACE) wählbar sind; Rest = deposit
HME_EDIT_TYPES = ("layer", "fill", "structure")
HME_TYPES = ("deposit", "interface", "layer", "fill", "structure")

MAP_KEYS = ["id1", "id2", "id3", "id4", "name", "descr", "kategorie", "group_outer", "group_inner",
            "above", "below", "contemp", "ident", "filter_col"]

PREFIX_COL = "@prefix"  # virtuelle Spalte: der von der ID abgeschnittene Projekt-Präfix (z. B. XTESTX)

ROLE_LABELS = {
    "id1": tr("ID Teil 1"), "id2": tr("ID Teil 2"), "id3": tr("ID Teil 3"), "id4": tr("ID Teil 4"), "name": "Name",
    "descr": tr("Funde / Beschreibung"), "kategorie": tr("Kategorie"), "group_outer": tr("Übergeordnete Gruppe"),
    "group_inner": tr("Untergeordnete Gruppe"), "above": tr("liegt über"),
    "below": tr("liegt unter"), "contemp": tr("zeitgleich mit"), "ident": tr("ident mit"), "filter_col": tr("Zeilenfilter"),
}
# Innerhalb dieser Gruppen darf keine Spalte doppelt vergeben sein
DISTINCT_GROUPS = [("id1", "id2", "id3", "id4"), ("above", "below", "contemp", "ident"),
                   ("group_outer", "group_inner")]

DEFAULT_SETTINGS = {
    "descr_suffix_on": "1",           # Suffix an Node-ID hängen, wenn Funde/Beschreibung gefüllt
    "descr_suffix": tr("_FUNDE"),
    "ref_key_col": "",                # leer = Beziehungen verweisen auf die zusammengesetzte ID
    "id_prefix_regex": "",            # Regex für einen Präfix, der von IDs entfernt wird (leer = nichts entfernen)
    "list_split_regex": r"[;,]",      # Trenner zwischen mehreren IDs in einer Beziehungszelle
    "delimiter": "auto",              # auto | , | ; | tab | |
    "tug_t_on": "0", "tug_g_on": "0", "tug_u_on": "0",   # T / G / U Nodes (optional, standardmäßig aus)
    "hme_undated": "1",               # HME: Units ohne Periode in Phasen "unknown-N" einordnen
    "out_format": "hmc",              # hmc = Harris Matrix Composer (.hmcx) | hme = Harris Matrix Editor (.json)
    "encoding": "auto",               # auto | utf-8-sig | utf-8 | cp1252 | latin-1 | utf-16
}

# Frühere mitgelieferte Beispiel-Schemata: werden beim Laden einer alten Konfiguration entfernt
REMOVED_SCHEMA_NAMES = ("Beispiel: Grobdatierung", "Beispiel: Detailmatrix", "Airtable-Standard (Grobdatierung)",
                        "Airtable-Standard (Detailmatrix)", "Beispiel: Airtable – Grobdatierung",
                        "Beispiel: Airtable – Detailmatrix")

ASSIGN_ALIASES = {   # deutsche und englische Spaltenbezeichnungen (Auto-Erkennung)
    "id1": ["id", "seid", "suid", "nodeid", "senr", "sunr", "nr", "nummer", "number", "context"],
    "name": ["setyp", "sutype", "typ", "type", "name", "bezeichnung", "label"],
    "descr": ["fundgattung", "funde", "fund", "finds", "find", "beschreibung", "description"],
    "kategorie": ["kategorie", "sekategorie", "category", "sucategory"],
    "group_outer": ["grobdatierung", "periode", "datierung", "period", "dating"],
    "group_inner": ["phase", "bauphase", "grabungsphase", "objekt", "object", "feature"],
    "above": ["liegtueber", "liegtüber", "liesabove", "above", "ueber", "over"],
    "below": ["liegtunter", "liesbelow", "below", "unter", "under"],
    "contemp": ["zeitgleich", "gleichzeitig", "contemporary", "simultaneous"],
    "ident": ["identmit", "identwith", "identical", "ident", "gleichgesetzt"],
}


class HMAError(Exception):
    """Fehler, der den Export abbricht."""


def default_schema():
    return normalize_schema({})


DEFAULT_DB = {"provider": "airtable", "base_id": "", "table_id": "", "view_id": "", "fields": "",
              "resolve_col": "", "out_name": tr("Datenbank_Export"), "overwrite": "0", "as_text": "1",
              # Server-Datenbanken / FileMaker / Google Sheets (Passwörter stehen nie hier, siehe ConfigStore.set_secret)
              "host": "", "port": "", "dbname": "", "user": "", "sslmode": "prefer", "table": "", "tls_verify": "1",
              "gs_url": ""}


def normalize_schema(p):
    """Bringt ein (evtl. altes) Schema in die aktuelle Form."""
    out = {"mapping": {k: "" for k in MAP_KEYS}, "settings": dict(DEFAULT_SETTINGS),
           "kategorie_interface": [], "group_outer_ignore": [], "group_inner_ignore": [], "filter_values": [],
           "hme_types": {}, "tug_t_values": [], "tug_g_values": [], "tug_u_values": [],
           "files": [], "sheet": "", "out_dir": "", "source": "file", "db": dict(DEFAULT_DB)}
    if not isinstance(p, dict):
        return out
    if p.get("source") in ("file", "db"):
        out["source"] = p["source"]
    if isinstance(p.get("db"), dict):
        for k in DEFAULT_DB:
            if k in p["db"]:
                out["db"][k] = str(p["db"][k] if p["db"][k] is not None else "")
        if out["db"]["provider"] not in ("airtable", "postgres", "mysql", "filemaker", "gsheets"):
            out["db"]["provider"] = "airtable"
        if out["db"]["sslmode"] not in ("prefer", "require", "disable", "verify-full"):
            out["db"]["sslmode"] = "prefer"
    if isinstance(p.get("files"), list):
        out["files"] = [str(v) for v in p["files"] if v]
    out["sheet"] = str(p.get("sheet", "") or "")
    out["out_dir"] = str(p.get("out_dir", "") or "")
    m = p.get("mapping") if isinstance(p.get("mapping"), dict) else {}
    s = p.get("settings") if isinstance(p.get("settings"), dict) else {}
    for k in MAP_KEYS:
        out["mapping"][k] = str(m.get(k, "") or "")
    for k in ("kategorie_interface", "group_outer_ignore", "group_inner_ignore", "filter_values",
              "tug_t_values", "tug_g_values", "tug_u_values"):
        if isinstance(p.get(k), list):
            out[k] = [str(v) for v in p[k]]
    if isinstance(p.get("hme_types"), dict):
        out["hme_types"] = {str(v): str(t) for v, t in p["hme_types"].items() if t in HME_EDIT_TYPES}
    for k in DEFAULT_SETTINGS:
        if k in s:
            out["settings"][k] = str(s[k])
    return out


def _appdata_config():
    base = os.environ.get("APPDATA") or str(Path.home())
    d = Path(base) / "HMA"
    d.mkdir(parents=True, exist_ok=True)
    legacy = Path(base) / "HMCA" / "hmca_config.json"   # Konfiguration der früheren Namensgebung übernehmen
    if not (d / "hma_config.json").exists() and legacy.exists():
        try:
            (d / "hma_config.json").write_bytes(legacy.read_bytes())
        except Exception:
            pass
    return d / "hma_config.json"


def _writable_file(p):
    try:
        with open(p, "a"):
            pass
        return True
    except Exception:
        return False


def _writable_dir(d):
    try:
        probe = d / ".hma_write_test"
        probe.write_text("x")
        probe.unlink()
        return True
    except Exception:
        return False


def pick_config_path():
    """hma_config.json: im HMA-Ordner. Liegt das Skript im Unterordner 'skript', ist das 'config' neben
    diesem (wird bei Bedarf angelegt), sonst der Ordner des Skripts. Nur wenn dort nicht geschrieben werden darf
    (z. B. 'C:\\Program Files' ohne Administratorrechte), wird %APPDATA%\\HMA verwendet; eine vorhandene
    Konfiguration wird dorthin kopiert. Eine Datei hmca_config.json aus der früheren Namensgebung wird beim
    ersten Start als hma_config.json übernommen (kopiert, die alte bleibt für ältere Versionen bestehen)."""
    name, legacy = "hma_config.json", "hmca_config.json"
    in_skript = SCRIPT_DIR.name.lower() in ("skript", "script")
    cfg_dir = SCRIPT_DIR.parent / "config" if in_skript else SCRIPT_DIR
    old = SCRIPT_DIR / name
    if not (cfg_dir / name).exists() and not old.exists():
        for src in (cfg_dir / legacy, SCRIPT_DIR / legacy):
            if src.exists():
                try:
                    cfg_dir.mkdir(parents=True, exist_ok=True)
                    (cfg_dir / name).write_bytes(src.read_bytes())
                except Exception:
                    pass
                break
    if in_skript and old.exists() and not (cfg_dir / name).exists():  # Ablage neben dem Skript übernehmen
        try:
            cfg_dir.mkdir(parents=True, exist_ok=True)
            (cfg_dir / name).write_bytes(old.read_bytes())
            old.unlink()
        except Exception:
            pass
    for cand in (cfg_dir / name, old):
        if cand.exists():
            if _writable_file(cand):
                return cand
            try:
                target = _appdata_config()
                if not target.exists():
                    target.write_bytes(cand.read_bytes())
                return target
            except Exception:
                return cand
    try:
        cfg_dir.mkdir(parents=True, exist_ok=True)
        if _writable_dir(cfg_dir):
            return cfg_dir / name
    except Exception:
        pass
    return _appdata_config()


class ConfigStore:
    def __init__(self, path):
        self.path = Path(path)
        self.data = {"version": 2, "last": {}, "schemas": {}, "secrets": {}, "ui": {}}
        self.load()

    def load(self):
        if self.path.exists():
            try:
                raw = json.loads(self.path.read_text(encoding="utf-8"))
                if isinstance(raw, dict):
                    if isinstance(raw.get("last"), dict):
                        self.data["last"] = raw["last"]
                    if isinstance(raw.get("secrets"), dict):
                        self.data["secrets"] = raw["secrets"]
                    if isinstance(raw.get("ui"), dict):
                        self.data["ui"] = raw["ui"]
                    schemas = raw.get("schemas", raw.get("profiles", {}))  # "profiles" = erste Version
                    if isinstance(schemas, dict):
                        for name, sc in schemas.items():
                            if name in REMOVED_SCHEMA_NAMES:
                                continue
                            self.data["schemas"][name] = normalize_schema(sc)
            except Exception:
                try:
                    self.path.replace(self.path.with_suffix(".json.bak"))
                except Exception:
                    pass

    def get_secret(self, name):
        return secret_decode(self.data.get("secrets", {}).get(name, ""))

    def set_secret(self, name, value):
        """Speichert einen Wert verschlüsselt (Windows-Benutzerkonto). None/leer löscht. False, wenn nicht möglich."""
        if not value:
            self.data.setdefault("secrets", {}).pop(name, None)
            return True
        enc = secret_encode(value)
        if not enc:
            return False
        self.data.setdefault("secrets", {})[name] = enc
        return True

    def _write(self, path):
        tmp = path.with_suffix(".json.tmp")
        tmp.write_text(json.dumps(self.data, ensure_ascii=False, indent=2), encoding="utf-8")
        tmp.replace(path)

    def save(self):
        try:
            self._write(self.path)
            return True
        except Exception:
            try:  # Schreibschutz: auf den Benutzerordner ausweichen
                alt = _appdata_config()
                self._write(alt)
                self.path = alt
                return True
            except Exception:
                return False


# ---------------------------------------------------------------------------
# Geheimnisse (API-Schlüssel): Windows-DPAPI, nur für das aktuelle Benutzerkonto lesbar
# ---------------------------------------------------------------------------
def _dpapi(data, encrypt):
    import ctypes
    from ctypes import wintypes

    class BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    buf = ctypes.create_string_buffer(data, len(data))
    inb = BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    outb = BLOB()
    crypt32, kernel32 = ctypes.windll.crypt32, ctypes.windll.kernel32
    fn = crypt32.CryptProtectData if encrypt else crypt32.CryptUnprotectData
    if not fn(ctypes.byref(inb), None, None, None, None, 0, ctypes.byref(outb)):
        raise OSError(tr("DPAPI-Fehler"))
    try:
        return ctypes.string_at(outb.pbData, outb.cbData)
    finally:
        kernel32.LocalFree(ctypes.cast(outb.pbData, ctypes.c_void_p))


def wheel_units(e):
    """Mausrad-Ereignis -> Scroll-Einheiten (Windows: Vielfache von 120, macOS: kleine Werte, X11: Button-4/5)."""
    d = getattr(e, "delta", 0)
    if d:
        if sys.platform == "darwin":
            return max(-5, min(5, -int(d)))
        return int(-d / 120) or (-1 if d > 0 else 1)
    return 1 if getattr(e, "num", 0) == 5 else -1


def secret_available():
    return os.name == "nt"


def secret_encode(text):
    if not secret_available() or not text:
        return ""
    try:
        return "dpapi:" + base64.b64encode(_dpapi(text.encode("utf-8"), True)).decode("ascii")
    except Exception:
        return ""


def secret_decode(token):
    if not token or not str(token).startswith("dpapi:") or not secret_available():
        return ""
    try:
        return _dpapi(base64.b64decode(token[6:]), False).decode("utf-8")
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# Datenbank-Anbindung (Airtable-API)
# ---------------------------------------------------------------------------
REC_RE = re.compile(r"^rec[A-Za-z0-9]{14}$")


def parse_airtable_url(text):
    """Liest Base-, Tabellen- und View-ID aus einer Airtable-Adresse."""
    out = {}
    for key, pat in (("base_id", r"\b(app[A-Za-z0-9]{14})\b"), ("table_id", r"\b(tbl[A-Za-z0-9]{14})\b"),
                     ("view_id", r"\b(viw[A-Za-z0-9]{14})\b")):
        mm = re.search(pat, text or "")
        if mm:
            out[key] = mm.group(1)
    return out


def _airtable_error(code, body):
    msg = ""
    try:
        j = json.loads(body)
        e = j.get("error")
        msg = (e.get("message") or e.get("type") or "") if isinstance(e, dict) else str(e or "")
    except Exception:
        msg = (body or "")[:200]
    base = {
        401: tr("Der API-Schlüssel ist ungültig oder abgelaufen. Bitte im Tab 'Datenbank' einen gültigen "
             "Personal Access Token eintragen."),
        403: tr("Keine Berechtigung. Der Token braucht das Recht 'data.records:read' und muss Zugriff auf diese "
             "Base haben (airtable.com/create/tokens)."),
        404: tr("Base oder Tabelle nicht gefunden. Bitte Base ID (app…) und Table ID (tbl…) prüfen."),
        422: tr("Airtable lehnt die Anfrage ab (z. B. unbekannter Feldname oder View)."),
        429: tr("Airtable hat die Anfragen begrenzt (Rate Limit). Bitte später erneut versuchen."),
    }.get(code, tr('Airtable meldet einen Fehler (HTTP {0}).').format(code))
    return base + (tr('\n  Antwort von Airtable: {0}').format(msg) if msg else "")


_SECURITY_BIN = "/usr/bin/security"


def _ssl_ctx():
    """Standard-SSL-Kontext. Auf dem Mac hat Python von python.org oft keine Zertifikate installiert
    ('Install Certificates.command' nicht ausgeführt); dann werden die Wurzelzertifikate des Systems
    (Schlüsselbund) nachgeladen. Nur Standardbibliothek, nur lesend."""
    import ssl
    ctx = ssl.create_default_context()
    if sys.platform == "darwin" and not os.environ.get("SSL_CERT_FILE"):
        try:
            if ctx.cert_store_stats().get("x509_ca", 0) == 0:
                import subprocess
                pem = ""
                for kc in (["/System/Library/Keychains/SystemRootCertificates.keychain"], ["/Library/Keychains/System.keychain"]):
                    r = subprocess.run([_SECURITY_BIN, "find-certificate", "-a", "-p"] + kc, capture_output=True,
                                       text=True, timeout=30)
                    if "BEGIN CERTIFICATE" in r.stdout:
                        pem += r.stdout
                if pem:
                    ctx.load_verify_locations(cadata=pem)
        except Exception:
            pass
    return ctx


def _cert_hint():
    """Zusatzhinweis bei Zertifikatsfehlern (je nach System)."""
    if sys.platform == "darwin":
        return "\n" + tr("Tipp (macOS): Im Ordner 'Programme/Python 3.x' die Datei 'Install Certificates.command' doppelklicken "
                         "und HMA danach neu starten.")
    return "\n" + tr("Mögliche Ursachen: Firmen-Proxy oder Virenschutz, der verschlüsselte Verbindungen untersucht, "
                     "oder eine falsch gehende Systemuhr.")


def airtable_fetch(db, api_key, log=None, timeout=30):
    """Holt alle Datensätze einer Tabelle/eines Views über die Airtable-REST-API (nur Standardbibliothek)."""
    import time
    import urllib.error
    import urllib.parse
    import urllib.request
    api_key = (api_key or "").strip()
    base, table = (db.get("base_id") or "").strip(), (db.get("table_id") or "").strip()
    if not api_key:
        raise HMAError(tr("Es ist kein API-Schlüssel (Personal Access Token) eingetragen. Bitte im Tab 'Datenbank' "
                        "eintragen."))
    if not base or not table:
        raise HMAError(tr("Base ID und Table ID sind erforderlich. Bitte im Tab 'Datenbank' eintragen."))
    url = f"https://api.airtable.com/v0/{urllib.parse.quote(base, safe='')}/{urllib.parse.quote(table, safe='')}"
    params0 = [("fields[]", f.strip()) for f in (db.get("fields") or "").split(",") if f.strip()]
    if (db.get("view_id") or "").strip():
        params0.append(("view", db["view_id"].strip()))
    if db.get("as_text", "1") == "1":  # Verknüpfungen als Text (Primärfeld der Zieltabelle) statt recXXXX
        params0 += [("cellFormat", "string"), ("timeZone", "UTC"), ("userLocale", "de")]
    records, offset, waits = [], None, 0
    while True:
        params = list(params0) + ([("offset", offset)] if offset else [])
        req = urllib.request.Request(url + ("?" + urllib.parse.urlencode(params) if params else ""),
                                     headers={"Authorization": f"Bearer {api_key}", "User-Agent": "HMA"})
        try:
            with urllib.request.urlopen(req, timeout=timeout, context=_ssl_ctx()) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            try:
                body = e.read().decode("utf-8", errors="replace")
            except Exception:
                body = ""
            if e.code == 429 and waits < 5:
                waits += 1
                if log:
                    log(tr("Airtable-Rate-Limit erreicht – warte 30 Sekunden …"))
                time.sleep(30)
                continue
            raise HMAError(_airtable_error(e.code, body))
        except urllib.error.URLError as e:
            import ssl
            if isinstance(getattr(e, "reason", None), ssl.SSLCertVerificationError):
                raise HMAError(tr("Das Sicherheitszertifikat von Airtable konnte nicht geprüft werden: {0}").format(e.reason) + _cert_hint())
            raise HMAError(tr('Keine Verbindung zu Airtable möglich: {0}\nBitte Internetverbindung, Proxy/Firewall prüfen.').format(getattr(e, 'reason', e)))
        except (OSError, ValueError) as e:
            raise HMAError(tr('Fehler bei der Abfrage von Airtable: {0}').format(e))
        records.extend(data.get("records", []))
        offset = data.get("offset")
        if log:
            log(tr('  {0} Datensätze abgerufen …').format(len(records)))
        if not offset:
            break
        time.sleep(0.22)  # Airtable erlaubt höchstens 5 Anfragen pro Sekunde
    return records


def _flat_value(v):
    """Airtable-Wert -> Text bzw. Liste von Texten."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "True" if v else ""
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else str(v)
    if isinstance(v, (int, str)):
        return str(v)
    if isinstance(v, dict):
        for k in ("name", "text", "url", "email", "label", "value", "error"):
            if k in v and isinstance(v[k], (str, int, float)):
                return str(v[k])
        return json.dumps(v, ensure_ascii=False)
    if isinstance(v, list):
        out = []
        for x in v:
            fx = _flat_value(x)
            out.extend(fx if isinstance(fx, list) else [fx])
        return out
    return str(v)


def records_to_table(records, db, schema):
    """Airtable-Datensätze -> (headers, rows, info) im gleichen Format wie load_table."""
    if not records:
        raise HMAError(tr("Die Abfrage hat keine Datensätze geliefert. Bitte Base, Tabelle und View prüfen."))
    cols = [f.strip() for f in (db.get("fields") or "").split(",") if f.strip()]
    if not cols:
        for r in records:
            for k in (r.get("fields") or {}):
                if k not in cols:
                    cols.append(k)
    absent = []
    if not (db.get("fields") or "").strip():
        # Airtable liefert Felder, die in KEINEM Datensatz einen Wert haben, gar nicht mit. Zugeordnete Spalten, die
        # deshalb fehlen, werden als leere Spalten ergänzt (sonst gäbe es einen Zuordnungsfehler, obwohl das Feld existiert).
        for _k, c in used_columns(schema):
            if c not in cols and c not in absent:
                absent.append(c)
        cols = cols + absent
    flat = [{c: _flat_value((r.get("fields") or {}).get(c, "")) for c in cols} for r in records]
    m, st = schema["mapping"], schema["settings"]
    resolve = (db.get("resolve_col") or "").strip()
    if resolve:
        label_cols = [resolve]
    elif st.get("ref_key_col"):
        label_cols = [st["ref_key_col"]]
    else:
        label_cols = [m[k] for k in ("id1", "id2", "id3", "id4") if m.get(k)]
    label_cols = [c for c in label_cols if c in cols]
    has_rec = any(isinstance(v, list) and any(REC_RE.match(x) for x in v) for row in flat for v in row.values())
    n_ok = n_open = 0
    labels = {}
    if has_rec and label_cols:
        for r, row in zip(records, flat):
            parts = [", ".join(v) if isinstance(v, list) else v for v in (row[c] for c in label_cols)]
            labels[r.get("id", "")] = "_".join(p for p in parts if p)
    rows = []
    for i, row in enumerate(flat):
        vals = []
        for c in cols:
            v = row[c]
            if isinstance(v, list):
                items = []
                for x in v:
                    if REC_RE.match(x):
                        if x in labels and labels[x]:
                            items.append(labels[x])
                            n_ok += 1
                        else:
                            items.append(x)
                            n_open += 1
                    else:
                        items.append(x)
                v = ", ".join(items)
            vals.append(v)
        rows.append((i + 2, vals))
    info = tr('Airtable: {0} Datensätze').format(len(records))
    if absent:
        info += (tr("; ohne Werte in Airtable (als leere Spalte ergänzt): ") + ", ".join(absent))
    if has_rec:
        if label_cols:
            info += tr(', Verknüpfungen: {0} aufgelöst').format(n_ok)
            if n_open:
                info += tr(', {0} nicht auflösbar (Verweis in andere Tabelle)').format(n_open)
        else:
            info += (tr(" – Achtung: Verknüpfte Datensätze (recXXXX) konnten nicht aufgelöst werden. "
                     "Bitte Spalten zuordnen (ID) bzw. unter 'Datenbank' eine Auflösungs-Spalte wählen."))
    return cols, rows, info


def example_values(headers, rows, col, limit=40, lead=tr("Bsp.: "), etc="..., etc."):
    """Beispielwerte einer Spalte: verschiedene, nicht leere Werte der Reihe nach, so viele wie in 'limit' Zeichen
    passen; gibt es weitere, folgt 'etc'. Leer, wenn die Spalte fehlt oder keine Werte hat."""
    if col not in headers:
        return ""
    i = headers.index(col)
    out, used, seen, more = [], 0, set(), False
    for _line, vals in rows:
        v = (vals[i] if i < len(vals) else "").strip()
        if not v or v in seen:
            continue
        seen.add(v)
        add = len(v) + (2 if out else 0)
        if used + add > limit:
            if not out:
                out.append(v[:limit - 1] + "…")
            more = True
            break
        out.append(v)
        used += add
    if not out:
        return ""
    return lead + ", ".join(out + ([etc] if more else []))


# ---------------------------------------------------------------------------
# Harris Matrix Composer: offenes Fenster erkennen, schließen, an alter Position neu öffnen (nur Windows)
# ---------------------------------------------------------------------------
HMC_TITLE = "Harris Matrix Composer"


def _top_windows():
    """Alle sichtbaren Top-Level-Fenster als Liste (hwnd, titel, pid)."""
    import ctypes
    from ctypes import wintypes
    u = ctypes.windll.user32
    out = []
    proto = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def cb(hwnd, _):
        if u.IsWindowVisible(hwnd):
            n = u.GetWindowTextLengthW(hwnd)
            if n:
                buf = ctypes.create_unicode_buffer(n + 1)
                u.GetWindowTextW(hwnd, buf, n + 1)
                pid = wintypes.DWORD()
                u.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                out.append((hwnd, buf.value, pid.value))
        return True

    u.EnumWindows(proto(cb), 0)
    return out


def hmc_find_window(filename):
    """Fenster von HMC, in dem genau diese Datei geöffnet ist (Titel: 'Harris Matrix Composer - <Datei>')."""
    if os.name != "nt":
        return None
    try:
        for hwnd, title, _pid in _top_windows():
            if title.startswith(HMC_TITLE) and filename.lower() in title.lower():
                return hwnd
    except Exception:
        pass
    return None


def _placement_struct():
    import ctypes
    from ctypes import wintypes

    class POINT(ctypes.Structure):
        _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]

    class RECT(ctypes.Structure):
        _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG), ("right", wintypes.LONG),
                    ("bottom", wintypes.LONG)]

    class WP(ctypes.Structure):
        _fields_ = [("length", wintypes.UINT), ("flags", wintypes.UINT), ("showCmd", wintypes.UINT),
                    ("ptMinPosition", POINT), ("ptMaxPosition", POINT), ("rcNormalPosition", RECT)]

    return WP


def hmc_get_state(hwnd):
    """Merkt Größe und Position des Fensters, wie sie gerade sichtbar sind (auch an Bildschirmkanten
    angedockt/gesnappt), sowie den Maximiert-Zustand. Gibt (rect, maximiert, placement) zurück."""
    import ctypes
    from ctypes import wintypes
    u = ctypes.windll.user32
    wp = _placement_struct()()
    wp.length = ctypes.sizeof(wp)
    if not u.GetWindowPlacement(hwnd, ctypes.byref(wp)):
        return None
    rc = wintypes.RECT()
    u.GetWindowRect(hwnd, ctypes.byref(rc))
    return (rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top), wp.showCmd == 3, wp


def hmc_set_state(hwnd, state):
    import ctypes
    if not state:
        return False
    rect, maximized, wp = state
    u = ctypes.windll.user32
    if maximized:
        return bool(u.SetWindowPlacement(hwnd, ctypes.byref(wp)))
    x, y, w, h = rect
    # SWP_NOZORDER | SWP_NOACTIVATE: tatsächliche Position/Größe, daher auch bei gesnappten Fenstern korrekt
    return bool(u.SetWindowPos(hwnd, 0, x, y, w, h, 0x0004 | 0x0010))


def _press(vk):
    import ctypes
    u = ctypes.windll.user32
    u.keybd_event(vk, 0, 0, 0)
    u.keybd_event(vk, 0, 2, 0)


def hmc_close(hwnd, timeout=20.0):
    """Schließt das HMC-Fenster ohne zu speichern. HMC fragt je nach Zustand nacheinander
    'Save changes to ...?' (Yes / No / Cancel, Standard Yes) und 'Quit Harris Matrix Composer?' (Yes / No,
    Standard No). In beiden Dialogen führt Tab + Enter zur gewünschten Antwort: 'No' (nicht speichern) bzw.
    'Yes' (beenden). Gibt True zurück, wenn das Fenster danach weg ist."""
    import ctypes
    import time
    u = ctypes.windll.user32
    pid = None
    for h, title, p in _top_windows():
        if h == hwnd:
            pid = p
    u.PostMessageW(hwnd, 0x0010, 0, 0)  # WM_CLOSE
    t0, handled = time.time(), set()
    while time.time() - t0 < timeout:
        time.sleep(0.3)
        if not u.IsWindow(hwnd):
            return True
        for h, title, p in _top_windows():
            if h != hwnd and p == pid and title.strip() == HMC_TITLE and h not in handled and len(handled) < 3:
                handled.add(h)
                u.SetForegroundWindow(h)
                time.sleep(0.4)
                _press(0x09)  # Tab
                time.sleep(0.15)
                _press(0x0D)  # Eingabe
                break
    return not u.IsWindow(hwnd)


# ---------------------------------------------------------------------------
# Tabellen einlesen (CSV / TXT / XLSX)
# ---------------------------------------------------------------------------
def _decode(raw: bytes, encoding: str) -> str:
    if encoding == "auto":
        if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
            return raw.decode("utf-16")
        try:
            return raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            return raw.decode("cp1252", errors="replace")
    return raw.decode(encoding, errors="replace")


def _detect_delimiter(text: str) -> str:
    first = text.split("\n", 1)[0]
    counts = {",": first.count(","), ";": first.count(";"), "\t": first.count("\t"), "|": first.count("|")}
    best = max(counts, key=counts.get)
    return best if counts[best] > 0 else ","


def _unique_headers(headers):
    seen = Counter()
    out = []
    for h in headers:
        seen[h] += 1
        out.append(h if seen[h] == 1 else f"{h} [{seen[h]}]")
    return out


_NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def xlsx_sheet_names(path):
    with zipfile.ZipFile(path) as zf:
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        return [s.get("name") for s in wb.iter(_NS_MAIN + "sheet")]


def _xlsx_read(path, sheet):
    with zipfile.ZipFile(path) as zf:
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        targets = {r.get("Id"): r.get("Target") for r in rels}
        sheets = [(s.get("name"), s.get(_NS_REL + "id")) for s in wb.iter(_NS_MAIN + "sheet")]
        if not sheets:
            raise HMAError(tr("Die Excel-Datei enthält keine Tabellenblätter."))
        rid = next((r for n, r in sheets if n == sheet), sheets[0][1])
        target = targets[rid].lstrip("/")
        sheet_path = target if target.startswith("xl/") else "xl/" + target
        shared = []
        if "xl/sharedStrings.xml" in zf.namelist():
            for si in ET.fromstring(zf.read("xl/sharedStrings.xml")).iter(_NS_MAIN + "si"):
                shared.append("".join(t.text or "" for t in si.iter(_NS_MAIN + "t")))
        root = ET.fromstring(zf.read(sheet_path))
        datexf, date1904 = set(), False
        try:
            pr = wb.find(_NS_MAIN + "workbookPr")
            date1904 = pr is not None and pr.get("date1904") in ("1", "true")
            if "xl/styles.xml" in zf.namelist():
                st = ET.fromstring(zf.read("xl/styles.xml"))
                codes = {int(n.get("numFmtId")): n.get("formatCode", "") for n in st.iter(_NS_MAIN + "numFmt")}
                cx = st.find(_NS_MAIN + "cellXfs")
                for i, xf in enumerate(cx.findall(_NS_MAIN + "xf") if cx is not None else []):
                    fid = int(xf.get("numFmtId", "0"))
                    if fid in _BUILTIN_DATE_FMT or (fid in codes and _is_date_fmt(codes[fid])):
                        datexf.add(i)
        except (ET.ParseError, ValueError, KeyError):
            datexf = set()

    def col_index(ref):
        letters = re.match(r"[A-Z]+", ref).group(0)
        n = 0
        for ch in letters:
            n = n * 26 + (ord(ch) - 64)
        return n - 1

    table = []
    for row in root.iter(_NS_MAIN + "row"):
        vals = {}
        for c in row.findall(_NS_MAIN + "c"):
            t = c.get("t")
            v = c.find(_NS_MAIN + "v")
            if t == "inlineStr":
                txt = "".join(x.text or "" for x in c.iter(_NS_MAIN + "t"))
            elif v is None or v.text is None:
                txt = ""
            elif t == "s":
                txt = shared[int(v.text)]
            elif t in ("str", "e", "d"):
                txt = v.text
            elif t == "b":
                txt = "TRUE" if v.text == "1" else "FALSE"
            else:
                try:
                    f = float(v.text)
                    if c.get("s", "").isdigit() and int(c.get("s")) in datexf:
                        txt = _excel_date(f, date1904)
                    else:
                        txt = str(int(f)) if f.is_integer() else format(f, ".15g")
                except ValueError:
                    txt = v.text
            vals[col_index(c.get("r"))] = txt
        width = (max(vals) + 1) if vals else 0
        table.append((int(row.get("r", len(table) + 1)), [vals.get(i, "") for i in range(width)]))
    return table


# ---------------------------------------------------------------------------
# Datenbank-Server (nur lesend, nur Standardbibliothek): PostgreSQL, MySQL/MariaDB, FileMaker, Google Sheets
# ---------------------------------------------------------------------------
DB_PROVIDERS = [("airtable", "Airtable"), ("postgres", "PostgreSQL"), ("mysql", "MySQL / MariaDB"),
                ("filemaker", "FileMaker (Data API)"), ("gsheets", "Google Sheets")]
DB_PROVIDER_NAME = dict(DB_PROVIDERS)
DB_DEFAULT_PORT = {"postgres": "5432", "mysql": "3306"}


def _net_error(e, what, host, port):
    import socket
    import ssl
    if isinstance(e, socket.timeout) or isinstance(e, TimeoutError):
        return HMAError(tr('Zeitüberschreitung bei der Verbindung zu {0} ({1}:{2}). Server nicht erreichbar, Adresse/Port oder '
                           'Firewall/VPN prüfen.').format(what, host, port))
    if isinstance(e, socket.gaierror):
        return HMAError(tr("Der Servername '{0}' wurde nicht gefunden. Bitte Adresse prüfen (und ggf. VPN/Netzwerk).").format(host))
    if isinstance(e, ConnectionRefusedError):
        return HMAError(tr('{0} auf {1}:{2} lehnt die Verbindung ab. Läuft der Server, stimmen Adresse und Port, und erlaubt er '
                           'Verbindungen von diesem Computer?').format(what, host, port))
    if isinstance(e, ssl.SSLError):
        return HMAError(tr('Die verschlüsselte Verbindung (SSL/TLS) zu {0} ist fehlgeschlagen: {1}').format(what, e))
    return HMAError(tr('Verbindung zu {0} ({1}:{2}) fehlgeschlagen: {3}').format(what, host, port, e))


def _db_port(db, prov):
    p = (db.get("port") or "").strip() or DB_DEFAULT_PORT.get(prov, "")
    if not p.isdigit() or not 0 < int(p) < 65536:
        raise HMAError(tr("Ungültiger Port: '{0}'.").format(p))
    return int(p)


# --- PostgreSQL ---
def _pg_array(s):
    """PostgreSQL-Array-Literal '{a,"b c",NULL}' -> 'a, b c' (leer bei NULL-Elementen)."""
    if not (s.startswith("{") and s.endswith("}")):
        return s
    items, cur, q, esc, depth, was_q = [], [], False, False, 0, False
    for ch in s[1:-1]:
        if esc:
            cur.append(ch)
            esc = False
        elif ch == "\\" and q:
            esc = True
        elif ch == '"':
            q = not q
            was_q = True
        elif not q and ch == "{":
            depth += 1
            cur.append(ch)
        elif not q and ch == "}":
            depth -= 1
            cur.append(ch)
        elif ch == "," and not q and depth == 0:
            items.append(("".join(cur), was_q))
            cur, was_q = [], False
        else:
            cur.append(ch)
    if cur or was_q or items:
        items.append(("".join(cur), was_q))
    return ", ".join("" if (t == "NULL" and not w) else t for t, w in items)


_PG_ARRAY_OIDS = {1000, 1001, 1005, 1006, 1007, 1009, 1014, 1015, 1016, 1021, 1022, 1115, 1182, 1183, 1231, 1270}
_PG_FLOAT_OIDS = {700, 701, 1700}


class _SslFallback(Exception):
    """Interner Hinweis: Verschlüsselung ließ sich nicht aushandeln; bei 'bevorzugt' ohne SSL weitermachen."""


def _guard_query(name):
    """Wandelt Verbindungsabbrüche und Protokollfehler beim Abfragen in verständliche Meldungen um."""
    import functools
    import struct as _st

    def deco(fn):
        @functools.wraps(fn)
        def wrapper(self, *a, **k):
            try:
                return fn(self, *a, **k)
            except HMAError:
                raise
            except OSError as e:
                raise _net_error(e, name, getattr(self, "host", "?"), getattr(self, "port", "?"))
            except (EOFError, ValueError, IndexError, _st.error) as e:
                raise HMAError(tr('Die Verbindung zu {0} wurde unterbrochen oder lieferte eine unerwartete Antwort ({1}).').format(name, e))
        return wrapper
    return deco


class PgConnection:
    """Minimaler PostgreSQL-Client (Protokoll 3.0, Simple Query). Nur lesender Zugriff."""

    def __init__(self, host, port, user, password, database, sslmode="prefer", timeout=20):
        import socket
        self.sock = None
        self.timeout = timeout
        self.host, self.port = host, port
        if sslmode not in ("disable", "prefer", "require", "verify-full"):
            sslmode = "prefer"
        for mode in (["prefer", "disable"] if sslmode == "prefer" else [sslmode]):
            try:
                self.sock = socket.create_connection((host, port), timeout=timeout)
            except OSError as e:
                raise _net_error(e, "PostgreSQL", host, port)
            try:
                self._negotiate_ssl(mode)
                self._startup(user, password, database)
                return
            except _SslFallback:  # "bevorzugt": TLS-Aushandlung scheiterte -> wie libpq ohne Verschlüsselung erneut versuchen
                self.close()
            except HMAError:
                self.close()
                raise
            except OSError as e:
                self.close()
                raise _net_error(e, "PostgreSQL", host, port)
            except (EOFError, ValueError, IndexError) as e:
                self.close()
                raise HMAError(tr('Unerwartete Antwort des Servers ({0}). Ist das wirklich ein PostgreSQL-Server?').format(e))

    # -- Grundfunktionen --
    def _recv(self, n):
        buf = bytearray()
        while len(buf) < n:
            chunk = self.sock.recv(n - len(buf))
            if not chunk:
                raise EOFError(tr("Verbindung vom Server beendet"))
            buf += chunk
        return bytes(buf)

    def _read_msg(self):
        import struct
        t = self._recv(1)
        ln = struct.unpack(">I", self._recv(4))[0]
        return t, self._recv(ln - 4) if ln > 4 else b""

    def _send(self, t, payload=b""):
        import struct
        self.sock.sendall(t + struct.pack(">I", len(payload) + 4) + payload)

    def _negotiate_ssl(self, mode):
        import ssl
        import struct
        if mode == "disable":
            return
        self.sock.sendall(struct.pack(">II", 8, 80877103))
        ans = self._recv(1)
        if ans == b"S":
            if mode == "verify-full":
                ctx = _ssl_ctx()
            else:
                ctx = ssl.create_default_context()
                ctx.check_hostname = False
                ctx.verify_mode = ssl.CERT_NONE
            try:
                self.sock = ctx.wrap_socket(self.sock, server_hostname=self.host)
            except ssl.SSLError as e:
                if mode == "prefer":
                    raise _SslFallback()
                raise _net_error(e, "PostgreSQL", self.host, self.port)
        elif mode in ("require", "verify-full"):
            raise HMAError(tr("Der Server unterstützt keine verschlüsselte Verbindung (SSL), die Einstellung verlangt "
                           "sie aber. Verschlüsselung auf 'bevorzugt' oder 'aus' stellen."))
        elif ans != b"N":
            raise HMAError(tr("Unerwartete Antwort des Servers. Ist das wirklich ein PostgreSQL-Server?"))

    @staticmethod
    def _err(payload):
        f = {}
        for part in payload.split(b"\0"):
            if part:
                f[chr(part[0])] = part[1:].decode("utf-8", errors="replace")
        return f

    def _raise_server_error(self, payload):
        f = self._err(payload)
        code, msg = f.get("C", ""), f.get("M", tr("Unbekannter Fehler"))
        hint = ""
        if code in ("28P01", "28000"):
            hint = tr("\nBenutzername oder Passwort stimmen nicht (oder der Server erlaubt diesen Zugang nicht, pg_hba.conf).")
        elif code == "3D000":
            hint = tr("\nDie angegebene Datenbank existiert nicht.")
        elif code == "42501":
            hint = tr("\nDem Benutzer fehlt die Leseberechtigung für diese Tabelle.")
        elif code == "42P01":
            hint = tr("\nDie Tabelle/View wurde nicht gefunden.")
        raise HMAError(f"PostgreSQL: {msg}{hint}")

    def _startup(self, user, password, database):
        import base64
        import hashlib
        import hmac
        import os
        import struct
        params = b"".join(k.encode() + b"\0" + v.encode("utf-8") + b"\0" for k, v in
                          (("user", user), ("database", database or user), ("client_encoding", "UTF8"),
                           ("application_name", "HMA")))
        body = struct.pack(">I", 196608) + params + b"\0"
        self.sock.sendall(struct.pack(">I", len(body) + 4) + body)
        scram = None
        while True:
            t, p = self._read_msg()
            if t == b"E":
                self._raise_server_error(p)
            elif t == b"R":
                code = struct.unpack(">I", p[:4])[0]
                if code == 0:
                    continue
                if code == 3:
                    self._send(b"p", password.encode("utf-8") + b"\0")
                elif code == 5:
                    inner = hashlib.md5(password.encode("utf-8") + user.encode("utf-8")).hexdigest()
                    self._send(b"p", b"md5" + hashlib.md5(inner.encode() + p[4:8]).hexdigest().encode() + b"\0")
                elif code == 10:
                    if b"SCRAM-SHA-256\0" not in p[4:] and b"SCRAM-SHA-256" not in p[4:]:
                        raise HMAError(tr("Der Server verlangt ein Anmeldeverfahren (SCRAM-SHA-256-PLUS), das HMA nicht kann."))
                    import unicodedata
                    pw = unicodedata.normalize("NFKC", password).encode("utf-8")
                    nonce = base64.b64encode(os.urandom(18)).decode()
                    bare = f"n=,r={nonce}"
                    first = ("n,," + bare).encode()
                    self._send(b"p", b"SCRAM-SHA-256\0" + struct.pack(">I", len(first)) + first)
                    scram = {"pw": pw, "nonce": nonce, "bare": bare}
                elif code == 11 and scram:
                    srv = p[4:].decode()
                    kv = dict(x.split("=", 1) for x in srv.split(","))
                    if not kv["r"].startswith(scram["nonce"]):
                        raise HMAError(tr("Ungültige Antwort des Servers bei der Anmeldung."))
                    salted = hashlib.pbkdf2_hmac("sha256", scram["pw"], base64.b64decode(kv["s"]), int(kv["i"]))
                    ck = hmac.new(salted, b"Client Key", "sha256").digest()
                    final0 = f"c=biws,r={kv['r']}"
                    auth = f"{scram['bare']},{srv},{final0}".encode()
                    sig = hmac.new(hashlib.sha256(ck).digest(), auth, "sha256").digest()
                    proof = base64.b64encode(bytes(a ^ b for a, b in zip(ck, sig))).decode()
                    scram["server_sig"] = hmac.new(hmac.new(salted, b"Server Key", "sha256").digest(), auth,
                                                   "sha256").digest()
                    self._send(b"p", f"{final0},p={proof}".encode())
                elif code == 12 and scram:
                    v = dict(x.split("=", 1) for x in p[4:].decode().split(",")).get("v", "")
                    if base64.b64decode(v) != scram["server_sig"]:
                        raise HMAError(tr("Der Server konnte sich bei der Anmeldung nicht ausweisen (SCRAM-Prüfung)."))
                else:
                    raise HMAError(tr('Der Server verlangt ein Anmeldeverfahren, das HMA nicht unterstützt (Code {0}; z. B. '
                                      'Kerberos/SSPI). Bitte einen Benutzer mit Passwort verwenden.').format(code))
            elif t == b"Z":
                return

    @_guard_query("PostgreSQL")
    def query(self, sql):
        """Führt EINE Abfrage innerhalb einer READ-ONLY-Transaktion aus. Gibt (Spalten, Zeilen) zurück."""
        import struct
        self.sock.sendall(self._msg(b"Q", b"BEGIN READ ONLY\0"))
        self._drain()
        try:
            self._send(b"Q", sql.encode("utf-8") + b"\0")
            cols, oids, rows = [], [], []
            err = None
            while True:
                t, p = self._read_msg()
                if t == b"T":
                    n = struct.unpack(">H", p[:2])[0]
                    off, cols, oids = 2, [], []
                    for _ in range(n):
                        end = p.index(b"\0", off)
                        cols.append(p[off:end].decode("utf-8", errors="replace"))
                        oids.append(struct.unpack(">I", p[end + 7:end + 11])[0])
                        off = end + 19
                elif t == b"D":
                    n = struct.unpack(">H", p[:2])[0]
                    off, vals = 2, []
                    for i in range(n):
                        ln = struct.unpack(">i", p[off:off + 4])[0]
                        off += 4
                        if ln < 0:
                            vals.append("")
                            continue
                        raw = p[off:off + ln].decode("utf-8", errors="replace")
                        off += ln
                        oid = oids[i] if i < len(oids) else 0
                        if oid == 16:
                            raw = "TRUE" if raw == "t" else "FALSE"
                        elif oid in _PG_ARRAY_OIDS:
                            raw = _pg_array(raw)
                        elif oid in _PG_FLOAT_OIDS:
                            try:
                                from decimal import Decimal
                                raw = format(Decimal(raw).normalize(), "f") if raw[-1:].isdigit() else raw
                            except Exception:
                                pass
                        vals.append(raw)
                    rows.append(vals)
                elif t == b"E":
                    err = p
                elif t == b"Z":
                    break
            if err:
                self._raise_server_error(err)
            return cols, rows
        finally:
            try:
                self.sock.sendall(self._msg(b"Q", b"ROLLBACK\0"))
                self._drain()
            except Exception:
                pass

    @staticmethod
    def _msg(t, payload):
        import struct
        return t + struct.pack(">I", len(payload) + 4) + payload

    def _drain(self):
        while True:
            t, p = self._read_msg()
            if t == b"E":
                self._raise_server_error(p)
            if t == b"Z":
                return

    def close(self):
        try:
            if self.sock:
                try:
                    self._send(b"X")
                except Exception:
                    pass
                self.sock.close()
        finally:
            self.sock = None


def _pg_ident(name):
    return '"' + name.replace('"', '""') + '"'


def _pg_connect(db, password):
    host = (db.get("host") or "").strip()
    user = (db.get("user") or "").strip()
    if not host or not user:
        raise HMAError(tr("Bitte Server und Benutzer eintragen (Tab 'Datenbank')."))
    return PgConnection(host, _db_port(db, "postgres"), user, password or "", (db.get("dbname") or "").strip(),
                        (db.get("sslmode") or "prefer").strip())


def pg_list_tables(db, password):
    conn = _pg_connect(db, password)
    try:
        _c, rows = conn.query(
            "SELECT table_schema, table_name FROM information_schema.tables "
            "WHERE table_schema NOT IN ('pg_catalog','information_schema') AND table_schema NOT LIKE 'pg\\_toast%' "
            "UNION SELECT schemaname, matviewname FROM pg_matviews ORDER BY 1, 2")
    finally:
        conn.close()
    return [(a if a != "public" else "") + ("." if a != "public" else "") + b for a, b in rows]


def pg_read_table(db, password, log=None):
    table = (db.get("table") or "").strip()
    if not table:
        raise HMAError(tr("Bitte eine Tabelle oder View wählen (Tab 'Datenbank', Knopf 'Tabellen laden')."))
    if "." in table:
        sch, name = table.split(".", 1)
        ref = _pg_ident(sch) + "." + _pg_ident(name)
    else:
        ref = _pg_ident(table)
    conn = _pg_connect(db, password)
    try:
        if log:
            log(tr('PostgreSQL: lese {0} …').format(table))
        cols, rows = conn.query("SELECT * FROM " + ref)
    finally:
        conn.close()
    return _server_table(cols, rows, tr('PostgreSQL, Tabelle {0}').format(table))


def _server_table(cols, rows, info):
    headers = _unique_headers([str(c).strip() for c in cols])
    out = [(i, (list(r) + [""] * len(headers))[:len(headers)]) for i, r in enumerate(rows, start=2)
           if any(str(x).strip() for x in r)]
    if not out:
        raise HMAError(tr("Die Abfrage hat keine Datensätze geliefert (leere Tabelle?)."))
    return headers, out, info


# --- MySQL / MariaDB ---
class MySqlConnection:
    """Minimaler MySQL/MariaDB-Client (Textprotokoll). Nur lesender Zugriff (START TRANSACTION READ ONLY)."""
    CAP_LONG_PASSWORD, CAP_CONNECT_WITH_DB, CAP_PROTOCOL_41 = 0x1, 0x8, 0x200
    CAP_SSL, CAP_SECURE_CONNECTION, CAP_PLUGIN_AUTH = 0x800, 0x8000, 0x80000
    CAP_PLUGIN_AUTH_LENENC = 0x200000

    def __init__(self, host, port, user, password, database, sslmode="prefer", timeout=20):
        import socket
        self.sock, self.seq, self.host, self.port = None, 0, host, port
        sslmode = sslmode if sslmode in ("disable", "prefer", "require") else "prefer"
        for mode in (["prefer", "disable"] if sslmode == "prefer" else [sslmode]):
            try:
                self.sock = socket.create_connection((host, port), timeout=timeout)
            except OSError as e:
                raise _net_error(e, "MySQL/MariaDB", host, port)
            try:
                self._handshake(user, password, database, mode)
                return
            except _SslFallback:
                self.close()
            except HMAError:
                self.close()
                raise
            except OSError as e:
                self.close()
                raise _net_error(e, "MySQL/MariaDB", host, port)
            except (EOFError, ValueError, IndexError) as e:
                self.close()
                raise HMAError(tr('Unerwartete Antwort des Servers ({0}). Ist das wirklich ein MySQL/MariaDB-Server?').format(e))

    def _recv(self, n):
        buf = bytearray()
        while len(buf) < n:
            c = self.sock.recv(n - len(buf))
            if not c:
                raise EOFError(tr("Verbindung vom Server beendet"))
            buf += c
        return bytes(buf)

    def _read_packet(self):
        out = b""
        while True:
            h = self._recv(4)
            ln = h[0] | (h[1] << 8) | (h[2] << 16)
            self.seq = (h[3] + 1) & 0xFF
            out += self._recv(ln)
            if ln < 0xFFFFFF:
                return out

    def _send_packet(self, payload):
        while True:
            chunk, payload = payload[:0xFFFFFF], payload[0xFFFFFF:]
            n = len(chunk)
            self.sock.sendall(bytes((n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF, self.seq)) + chunk)
            self.seq = (self.seq + 1) & 0xFF
            if n < 0xFFFFFF:
                return

    @staticmethod
    def _lenenc(b, off):
        c = b[off]
        if c < 0xFB:
            return c, off + 1
        if c == 0xFC:
            return int.from_bytes(b[off + 1:off + 3], "little"), off + 3
        if c == 0xFD:
            return int.from_bytes(b[off + 1:off + 4], "little"), off + 4
        if c == 0xFE:
            return int.from_bytes(b[off + 1:off + 9], "little"), off + 9
        return None, off + 1  # 0xFB = NULL

    @staticmethod
    def _raise_err(p):
        code = int.from_bytes(p[1:3], "little")
        msg = p[9:].decode("utf-8", errors="replace") if p[3:4] == b"#" else p[3:].decode("utf-8", errors="replace")
        hint = ""
        if code in (1045, 1044):
            hint = tr("\nBenutzername oder Passwort stimmen nicht (oder der Benutzer darf sich von diesem Computer nicht anmelden).")
        elif code == 1049:
            hint = tr("\nDie angegebene Datenbank existiert nicht.")
        elif code in (1142, 1143):
            hint = tr("\nDem Benutzer fehlt die Leseberechtigung für diese Tabelle.")
        elif code == 1146:
            hint = tr("\nDie Tabelle/View wurde nicht gefunden.")
        raise HMAError(tr('MySQL/MariaDB: {0}{1}').format(msg, hint))

    @staticmethod
    def _scramble(plugin, pw, salt):
        import hashlib
        if not pw:
            return b""
        pw = pw.encode("utf-8")
        if plugin == "mysql_native_password":
            s1 = hashlib.sha1(pw).digest()
            s3 = hashlib.sha1(salt[:20] + hashlib.sha1(s1).digest()).digest()
            return bytes(a ^ b for a, b in zip(s1, s3))
        if plugin == "caching_sha2_password":
            s1 = hashlib.sha256(pw).digest()
            s3 = hashlib.sha256(hashlib.sha256(s1).digest() + salt[:20]).digest()
            return bytes(a ^ b for a, b in zip(s1, s3))
        raise HMAError(tr("Der Server verlangt das Anmeldeverfahren '{0}', das HMA nicht unterstützt. Bitte einen Benutzer mit "
                          "'mysql_native_password' oder 'caching_sha2_password' verwenden.").format(plugin))

    def _handshake(self, user, password, database, sslmode):
        import ssl
        import struct
        p = self._read_packet()
        if p[0] == 0xFF:
            self._raise_err(p)
        if p[0] != 10:
            raise HMAError(tr("Unerwartete Antwort des Servers. Ist das wirklich ein MySQL/MariaDB-Server?"))
        off = p.index(b"\0", 1) + 1
        off += 4
        salt = p[off:off + 8]
        off += 9
        caps = struct.unpack_from("<H", p, off)[0]
        off += 5
        caps |= struct.unpack_from("<H", p, off)[0] << 16
        off += 2
        auth_len = p[off]
        off += 11
        salt += p[off:off + max(13, auth_len - 8) - 1]
        off += max(13, auth_len - 8)
        plugin = p[off:].split(b"\0")[0].decode("ascii", errors="ignore") or "mysql_native_password"
        cf = (self.CAP_LONG_PASSWORD | self.CAP_PROTOCOL_41 | self.CAP_SECURE_CONNECTION | self.CAP_PLUGIN_AUTH
              | self.CAP_PLUGIN_AUTH_LENENC | (self.CAP_CONNECT_WITH_DB if database else 0))
        tls = False
        if sslmode != "disable" and caps & self.CAP_SSL:
            cf |= self.CAP_SSL
            self._send_packet(struct.pack("<IIB", cf, 1 << 24, 45) + b"\0" * 23)
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            try:
                self.sock = ctx.wrap_socket(self.sock, server_hostname=self.host)
            except ssl.SSLError as e:
                if sslmode == "prefer":
                    raise _SslFallback()
                raise _net_error(e, "MySQL/MariaDB", self.host, self.port)
            tls = True
        elif sslmode == "require":
            raise HMAError(tr("Der Server bietet keine verschlüsselte Verbindung (SSL) an, die Einstellung verlangt sie "
                           "aber. Verschlüsselung auf 'bevorzugt' oder 'aus' stellen."))
        tok = self._scramble(plugin, password, salt)
        resp = (struct.pack("<IIB", cf, 1 << 24, 45) + b"\0" * 23 + user.encode("utf-8") + b"\0"
                + bytes([len(tok)]) + tok + ((database or "").encode("utf-8") + b"\0" if database else b"")
                + plugin.encode() + b"\0")
        self._send_packet(resp)
        while True:
            p = self._read_packet()
            if p[0] == 0x00:
                return
            if p[0] == 0xFF:
                self._raise_err(p)
            if p[0] == 0xFE:  # Wechsel des Anmeldeverfahrens
                plugin = p[1:].split(b"\0")[0].decode("ascii", errors="ignore")
                salt = p[1 + len(plugin) + 1:].rstrip(b"\0")
                self._send_packet(self._scramble(plugin, password, salt))
            elif p[0] == 0x01:  # AuthMoreData (caching_sha2_password)
                if p[1:2] == b"\x03":
                    continue  # schnelle Anmeldung ok, es folgt OK-Paket
                if p[1:2] == b"\x04":
                    if not tls:
                        raise HMAError(tr("Der Server verlangt für diese Anmeldung eine verschlüsselte Verbindung "
                                       "(caching_sha2_password), sie ist aber nicht möglich. SSL am Server aktivieren "
                                       "oder einen Benutzer mit 'mysql_native_password' verwenden."))
                    self._send_packet(password.encode("utf-8") + b"\0")
                else:
                    raise HMAError(tr("Unerwartete Antwort bei der Anmeldung."))
            else:
                raise HMAError(tr("Unerwartete Antwort bei der Anmeldung."))

    def _cmd(self, sql):
        self.seq = 0
        self._send_packet(b"\x03" + sql.encode("utf-8"))

    def _exec(self, sql):
        """Befehl ohne Ergebnismenge (BEGIN, ROLLBACK …)."""
        self._cmd(sql)
        p = self._read_packet()
        if p[0] == 0xFF:
            self._raise_err(p)

    @_guard_query("MySQL/MariaDB")
    def query(self, sql):
        self._exec("START TRANSACTION READ ONLY")
        try:
            self._cmd(sql)
            p = self._read_packet()
            if p[0] == 0xFF:
                self._raise_err(p)
            if p[0] == 0x00:
                return [], []
            ncol, _o = self._lenenc(p, 0)
            cols, types = [], []
            for _ in range(ncol):
                d = self._read_packet()
                o = 0
                vals = []
                for _i in range(6):  # catalog, schema, table, org_table, name, org_name
                    ln, o = self._lenenc(d, o)
                    vals.append(d[o:o + ln])
                    o += ln
                o += 1  # 0x0c
                o += 2 + 4  # charset, length
                cols.append(vals[4].decode("utf-8", errors="replace"))
                types.append(d[o])
            p = self._read_packet()  # EOF nach den Spalten
            rows = []
            while True:
                p = self._read_packet()
                if p[0] == 0xFE and len(p) < 9:
                    break
                if p[0] == 0xFF:
                    self._raise_err(p)
                o, vals = 0, []
                for t in types:
                    ln, o = self._lenenc(p, o)
                    if ln is None:
                        vals.append("")
                        continue
                    v = p[o:o + ln].decode("utf-8", errors="replace")
                    o += ln
                    if t in (0, 4, 5, 246) and v and v[-1:].isdigit():  # DECIMAL/FLOAT/DOUBLE
                        try:
                            from decimal import Decimal
                            v = format(Decimal(v).normalize(), "f")
                        except Exception:
                            pass
                    elif v.startswith("0000-00-00"):
                        v = ""
                    vals.append(v)
                rows.append(vals)
            return cols, rows
        finally:
            try:
                self._exec("ROLLBACK")
            except Exception:
                pass

    def close(self):
        try:
            if self.sock:
                try:
                    self.seq = 0
                    self._send_packet(b"\x01")
                except Exception:
                    pass
                self.sock.close()
        finally:
            self.sock = None


def _my_connect(db, password):
    host = (db.get("host") or "").strip()
    user = (db.get("user") or "").strip()
    if not host or not user:
        raise HMAError(tr("Bitte Server und Benutzer eintragen (Tab 'Datenbank')."))
    return MySqlConnection(host, _db_port(db, "mysql"), user, password or "", (db.get("dbname") or "").strip(),
                           (db.get("sslmode") or "prefer").strip())


def mysql_list_tables(db, password):
    conn = _my_connect(db, password)
    try:
        _c, rows = conn.query("SHOW FULL TABLES")
    finally:
        conn.close()
    return sorted(r[0] for r in rows)


def mysql_read_table(db, password, log=None):
    table = (db.get("table") or "").strip()
    if not table:
        raise HMAError(tr("Bitte eine Tabelle oder View wählen (Tab 'Datenbank', Knopf 'Tabellen laden')."))
    conn = _my_connect(db, password)
    try:
        if log:
            log(tr('MySQL/MariaDB: lese {0} …').format(table))
        cols, rows = conn.query("SELECT * FROM `" + table.replace("`", "``") + "`")
    finally:
        conn.close()
    return _server_table(cols, rows, tr('MySQL/MariaDB, Tabelle {0}').format(table))


# --- HTTP-Hilfen (FileMaker, Google Sheets) ---
def _http(url, method="GET", headers=None, data=None, timeout=30, verify=True):
    import ssl
    import urllib.error
    import urllib.request
    h = {"User-Agent": "HMA"}
    h.update(headers or {})
    req = urllib.request.Request(url, data=data, headers=h, method=method)
    ctx = None
    if url.lower().startswith("https"):
        ctx = _ssl_ctx()
        if not verify:
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
    try:
        with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
            return resp.status, resp.headers.get("Content-Type", ""), resp.read()
    except urllib.error.HTTPError as e:
        try:
            body = e.read()
        except Exception:
            body = b""
        return e.code, e.headers.get("Content-Type", "") if e.headers else "", body
    except urllib.error.URLError as e:
        r = getattr(e, "reason", e)
        if isinstance(r, ssl.SSLCertVerificationError):
            raise HMAError(tr("Das Zertifikat des Servers wird nicht als vertrauenswürdig erkannt (z. B. selbstsigniert). "
                           "Wenn du dem Server vertraust, 'Zertifikat nicht prüfen' aktivieren.") + (_cert_hint() if sys.platform == "darwin" else ""))
        raise HMAError(tr('Keine Verbindung möglich: {0}\nBitte Adresse, Internetverbindung, VPN, Proxy/Firewall prüfen.').format(r))
    except (OSError, ValueError) as e:
        raise HMAError(tr('Fehler bei der Abfrage: {0}').format(e))


# --- FileMaker Data API ---
def _fm_base(db):
    host = (db.get("host") or "").strip().rstrip("/")
    if not host:
        raise HMAError(tr("Bitte die Serveradresse eintragen (z. B. fm.beispiel.at)."))
    if not re.match(r"https?://", host, re.I):
        host = "https://" + host
    if host.lower().startswith("http://") and not re.match(r"http://(localhost|127\.0\.0\.1|\[::1\])(:|/|$)", host, re.I):
        raise HMAError(tr("Eine unverschlüsselte Verbindung (http://) ist nicht erlaubt, weil dabei das Passwort im Klartext "
                       "übertragen würde. Bitte die Adresse ohne 'http://' angeben (FileMaker verwendet https)."))
    port = (db.get("port") or "").strip()
    if port:
        if not port.isdigit():
            raise HMAError(tr("Ungültiger Port: '{0}'.").format(port))
        host += ":" + port
    dbn = (db.get("dbname") or "").strip()
    if not dbn:
        raise HMAError(tr("Bitte den Datenbanknamen eintragen (Name der FileMaker-Datei ohne .fmp12)."))
    import urllib.parse
    return f"{host}/fmi/data/vLatest/databases/{urllib.parse.quote(dbn, safe='')}"


def _fm_json(status, body):
    try:
        j = json.loads(body.decode("utf-8", errors="replace"))
    except ValueError:
        raise HMAError(tr('Der Server hat keine gültige FileMaker-Antwort geliefert (HTTP {0}). Stimmen Adresse und '
                          'Datenbankname, und ist die Data API am Server aktiviert?').format(status))
    msgs = j.get("messages") or [{}]
    code = str(msgs[0].get("code", "0"))
    return j, code, msgs[0].get("message", "")


_FM_HINTS = {"212": tr("Benutzername oder Passwort stimmen nicht."), "105": tr("Das Layout wurde nicht gefunden."),
             "802": tr("Die Datenbank konnte nicht geöffnet werden (Name prüfen)."),
             "954": tr("Die Data API ist für diese Datenbank/diesen Benutzer nicht aktiviert (Erweiterte Zugriffsrechte "
                    "'fmrest' im Rechtesatz des Benutzers)."),
             "952": tr("Die Sitzung ist ungültig geworden."), "401": tr("Keine Datensätze gefunden.")}


def _fm_check(code, msg):
    if code not in ("0", "401"):
        raise HMAError(tr('FileMaker (Fehler {0}): {1}').format(code, msg) + ("\n" + _FM_HINTS[code] if code in _FM_HINTS else ""))


class FmSession:
    def __init__(self, db, password):
        import base64
        self.base = _fm_base(db)
        self.verify = (db.get("tls_verify", "1") != "0")
        user = (db.get("user") or "").strip()
        if not user:
            raise HMAError(tr("Bitte den Benutzer eintragen."))
        auth = base64.b64encode((user + ":" + (password or "")).encode("utf-8")).decode()
        st, _ct, body = _http(self.base + "/sessions", "POST", {"Authorization": "Basic " + auth,
                                                                 "Content-Type": "application/json"}, b"{}",
                              verify=self.verify)
        j, code, msg = _fm_json(st, body)
        _fm_check(code, msg)
        self.token = j["response"]["token"]

    def get(self, path):
        st, _ct, body = _http(self.base + path, "GET", {"Authorization": "Bearer " + self.token}, verify=self.verify)
        j, code, msg = _fm_json(st, body)
        _fm_check(code, msg)
        return j, code

    def close(self):
        try:
            _http(self.base + "/sessions/" + self.token, "DELETE", {"Content-Type": "application/json"},
                  verify=self.verify, timeout=10)
        except Exception:
            pass


def fm_list_layouts(db, password):
    s = FmSession(db, password)
    try:
        j, _c = s.get("/layouts")
    finally:
        s.close()
    out = []

    def walk(items):
        for it in items or []:
            if it.get("isFolder"):
                walk(it.get("folderLayoutNames"))
            elif it.get("name"):
                out.append(it["name"])
    walk(j.get("response", {}).get("layouts"))
    return out


def _fm_value(v):
    if v is None:
        return ""
    if isinstance(v, float):
        return _num_text(v)
    return str(v)


def fm_read_layout(db, password, log=None):
    import urllib.parse
    layout = (db.get("table") or "").strip()
    if not layout:
        raise HMAError(tr("Bitte ein Layout wählen (Tab 'Datenbank', Knopf 'Layouts laden')."))
    s = FmSession(db, password)
    cols, recs, offset, page = [], [], 1, 500
    try:
        while True:
            if log:
                log(tr('FileMaker: lese {0}, ab Datensatz {1} …').format(layout, offset))
            j, code = s.get(f"/layouts/{urllib.parse.quote(layout, safe='')}/records?_offset={offset}&_limit={page}")
            data = (j.get("response") or {}).get("data") or []
            for rec in data:
                fd = rec.get("fieldData") or {}
                for k in fd:
                    if k not in cols:
                        cols.append(k)
                recs.append(fd)
            if code == "401" or len(data) < page:
                break
            offset += page
    finally:
        s.close()
    rows = [[_fm_value(fd.get(c)) for c in cols] for fd in recs]
    return _server_table(cols, rows, tr('FileMaker, Layout {0}').format(layout))


# --- Google Sheets (öffentlich per Link) ---
def gsheets_csv_url(link):
    link = (link or "").strip()
    if not link:
        raise HMAError(tr("Bitte den Link zur Google-Tabelle eintragen."))
    gid = re.search(r"[#&?]gid=(\d+)", link)
    m = re.search(r"/spreadsheets/d/e/([\w-]+)", link)
    if m:  # "Im Web veröffentlicht"
        u = f"https://docs.google.com/spreadsheets/d/e/{m.group(1)}/pub?output=csv"
        return u + (f"&gid={gid.group(1)}&single=true" if gid else "")
    m = re.search(r"/spreadsheets/d/([\w-]+)", link) or re.fullmatch(r"([\w-]{20,})", link)
    if not m:
        raise HMAError(tr("Aus dem Link konnte keine Google-Tabelle erkannt werden. Bitte die Adresse aus dem Browser "
                       "kopieren (…/spreadsheets/d/…)."))
    return f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=csv" + (f"&gid={gid.group(1)}" if gid else "")


def gsheets_read(db, settings=None, log=None):
    url = db.get("gs_url_override") or gsheets_csv_url(db.get("gs_url"))
    if log:
        log(tr("Google Sheets: lade Tabelle …"))
    st, ct, body = _http(url)
    head = body[:400].lstrip().lower()
    if st in (401, 403) or "text/html" in ct.lower() or head.startswith((b"<!doctype", b"<html")):
        raise HMAError(tr("Die Google-Tabelle ist nicht öffentlich lesbar. In Google Sheets 'Teilen' → Allgemeiner Zugriff "
                       "→ 'Jeder, der den Link hat' → 'Betrachter' einstellen (oder 'Datei → Im Web veröffentlichen'). "
                       "Private Tabellen kann HMA nicht abrufen; dort bitte als .xlsx/.csv herunterladen."))
    if st == 404:
        raise HMAError(tr("Die Google-Tabelle wurde nicht gefunden. Bitte den Link prüfen."))
    if st != 200:
        raise HMAError(tr('Google Sheets antwortet mit Fehler {0}.').format(st))
    st2 = dict(settings or DEFAULT_SETTINGS)
    st2["delimiter"] = ","
    st2["encoding"] = "auto"
    headers, rows, _i = parse_csv_bytes(body, st2, "Google Sheets")
    return headers, rows, "Google Sheets"


# --- Zentrale Abfrage für den Tab 'Datenbank' ---
def db_load_table(db, secret, schema, log=None):
    """Liest je nach Anbieter die Tabelle. Gibt (headers, rows, info) zurück."""
    prov = (db.get("provider") or "airtable")
    if prov == "airtable":
        return records_to_table(airtable_fetch(db, secret, log=log), db, schema)
    if prov == "postgres":
        return pg_read_table(db, secret, log)
    if prov == "mysql":
        return mysql_read_table(db, secret, log)
    if prov == "filemaker":
        return fm_read_layout(db, secret, log)
    if prov == "gsheets":
        return gsheets_read(db, schema.get("settings"), log)
    raise HMAError(tr('Unbekannte Datenquelle: {0}').format(prov))


def db_list_tables(db, secret):
    prov = db.get("provider")
    if prov == "postgres":
        return pg_list_tables(db, secret)
    if prov == "mysql":
        return mysql_list_tables(db, secret)
    if prov == "filemaker":
        return fm_list_layouts(db, secret)
    return []


def db_ready(db, secret):
    """Fehlende Pflichtangaben für den gewählten Anbieter (leer = vollständig)."""
    prov = db.get("provider") or "airtable"
    missing = []
    if prov == "airtable":
        if not (secret or "").strip():
            missing.append(tr("API-Schlüssel"))
        for k, n in (("base_id", "Base ID"), ("table_id", "Table ID")):
            if not (db.get(k) or "").strip():
                missing.append(n)
    elif prov in ("postgres", "mysql", "filemaker"):
        for k, n in (("host", "Server"), ("user", tr("Benutzer")), ("table", "Layout" if prov == "filemaker" else tr("Tabelle"))):
            if not (db.get(k) or "").strip():
                missing.append(n)
        if prov == "filemaker" and not (db.get("dbname") or "").strip():
            missing.append(tr("Datenbank"))
    elif prov == "gsheets":
        if not (db.get("gs_url") or "").strip():
            missing.append("Link")
    return missing


# ---------------------------------------------------------------------------
# Weitere Tabellenformate: LibreOffice Calc (.ods), altes Excel (.xls, BIFF5/8), dBase/FoxPro (.dbf)
# ---------------------------------------------------------------------------
def _num_text(f):
    if f != f:  # NaN
        return ""
    return str(int(f)) if f.is_integer() and abs(f) < 1e15 else format(f, ".15g")


_BUILTIN_DATE_FMT = set(range(14, 23)) | set(range(27, 37)) | set(range(45, 48)) | set(range(50, 59))


def _is_date_fmt(code):
    """Erkennt, ob ein Excel-Zahlenformat ein Datum/eine Uhrzeit darstellt."""
    c = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]|_.|\*.', "", code or "")
    c = c.split(";")[0].lower()
    return bool(re.search(r"[ymdhs]", c)) and "general" not in c


def _excel_date(f, mode1904=False):
    """Excel-Seriennummer -> ISO-Datum (mit Uhrzeit, wenn vorhanden)."""
    from datetime import datetime as _dt, timedelta as _td
    try:
        base = _dt(1904, 1, 1) if mode1904 else _dt(1899, 12, 30)
        if not mode1904 and f < 61:
            base = _dt(1899, 12, 31)  # 1900-Schaltjahr-Fehler von Excel
        d = base + _td(days=f)
        if f < 1 and f >= 0:
            return d.strftime("%H:%M:%S")
        return d.strftime("%Y-%m-%d") if abs(f - round(f)) < 1e-9 else d.strftime("%Y-%m-%d %H:%M:%S")
    except (OverflowError, ValueError):
        return _num_text(f)


def _xls_kind(path):
    """Erkennt den tatsächlichen Inhalt einer .xls-Datei: 'xls' (echt), 'xlsx' (falsch benannt), 'text' (CSV/HTML)."""
    try:
        head = Path(path).read_bytes()[:16]
    except OSError as e:
        raise HMAError(tr('Die Datei konnte nicht gelesen werden: {0}').format(e))
    if head[:2] == b"PK":
        return "xlsx"
    if head[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "xls"
    return "text"


# --- ODS ---
_ODS_NS = {"office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
           "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
           "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0"}


def _oq(ns, name):
    return "{%s}%s" % (_ODS_NS[ns], name)


_T_P, _T_S, _T_TAB, _T_BR = _oq("text", "p"), _oq("text", "s"), _oq("text", "tab"), _oq("text", "line-break")
_T_CELL, _T_COVERED, _T_ROW = _oq("table", "table-cell"), _oq("table", "covered-table-cell"), _oq("table", "table-row")
_A_VT, _A_VAL, _A_DATE = _oq("office", "value-type"), _oq("office", "value"), _oq("office", "date-value")
_A_TIME, _A_BOOL, _A_STR = _oq("office", "time-value"), _oq("office", "boolean-value"), _oq("office", "string-value")
_A_CREP, _A_RREP, _A_SC = _oq("table", "number-columns-repeated"), _oq("table", "number-rows-repeated"), _oq("text", "c")


def _ods_text(el):
    if not len(el):  # häufigster Fall: reiner Text ohne Unterelemente
        return el.text or ""
    parts = []

    def rec(e):
        if e.text:
            parts.append(e.text)
        for ch in e:
            if ch.tag == _T_S:
                parts.append(" " * int(ch.get(_A_SC, "1")))
            elif ch.tag == _T_TAB:
                parts.append("\t")
            elif ch.tag == _T_BR:
                parts.append("\n")
            else:
                rec(ch)
            if ch.tail:
                parts.append(ch.tail)
    rec(el)
    return "".join(parts)


def _ods_cell_value(c):
    vt = c.get(_A_VT)
    if vt is None:
        ps = c.findall(_T_P)
        return ps[0].text or "" if len(ps) == 1 and not len(ps[0]) else "\n".join(_ods_text(p) for p in ps)
    if vt in ("float", "percentage", "currency"):
        try:
            return _num_text(float(c.get(_A_VAL, "")))
        except ValueError:
            return c.get(_A_VAL, "")
    if vt == "date":
        return c.get(_A_DATE, "").replace("T", " ")
    if vt == "time":
        tv = c.get(_A_TIME, "")  # PT05H06M07S
        m = re.fullmatch(r"P(?:(\d+)D)?T(\d+)H(\d+)M(\d+)(?:\.\d+)?S", tv)
        return "%02d:%02d:%02d" % (int(m.group(2)) + 24 * int(m.group(1) or 0), int(m.group(3)), int(m.group(4))) if m else tv
    if vt == "boolean":
        return "TRUE" if c.get(_A_BOOL) == "true" else "FALSE"
    sv = c.get(_A_STR)
    if sv is not None:
        return sv
    return "\n".join(_ods_text(p) for p in c.findall(_T_P))


def _ods_sheets(path):
    try:
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            if "META-INF/manifest.xml" in names and b"encryption-data" in zf.read("META-INF/manifest.xml"):
                raise HMAError(tr("Die ODS-Datei ist passwortgeschützt und kann nicht gelesen werden."))
            root = ET.fromstring(zf.read("content.xml"))
    except HMAError:
        raise
    except Exception as e:
        raise HMAError(tr('Die ODS-Datei konnte nicht gelesen werden: {0}').format(e))
    ss = root.find(_oq("office", "body") + "/" + _oq("office", "spreadsheet"))
    if ss is None:
        raise HMAError(tr("Die Datei enthält keine Tabellen."))
    return [(t.get(_oq("table", "name")), t) for t in ss.findall(_oq("table", "table"))]


def ods_sheet_names(path):
    return [n for n, _t in _ods_sheets(path)]


def _ods_read(path, sheet):
    sheets = _ods_sheets(path)
    if not sheets:
        raise HMAError(tr("Die ODS-Datei enthält keine Tabellenblätter."))
    tbl = next((t for n, t in sheets if n == sheet), sheets[0][1])
    containers = {_oq("table", "table-header-rows"), _oq("table", "table-row-group"), _oq("table", "table-rows")}

    def walk(e):
        for ch in e:
            if ch.tag == _T_ROW:
                yield ch
            elif ch.tag in containers:
                yield from walk(ch)
    table, rowno, budget = [], 0, 20_000_000
    for row in walk(tbl):
        rep = int(row.get(_A_RREP, "1"))
        vals, gap = [], 0
        for c in row:
            tag = c.tag
            if tag != _T_CELL and tag != _T_COVERED:
                continue
            v = _ods_cell_value(c) if tag == _T_CELL else ""
            rc = c.get(_A_CREP)
            n = int(rc) if rc else 1
            if v == "":
                gap += n  # leere Zellen erst anhängen, wenn danach noch etwas kommt (Spalten bis zum Tabellenende)
            else:
                if gap:
                    vals.extend([""] * min(gap, 4096))
                    gap = 0
                vals.extend([v] * min(n, 100000))
        if not vals:  # leere (evtl. millionenfach wiederholte) Zeilen nicht ausbreiten
            rowno += rep
            continue
        for _ in range(min(rep, 10000)):
            rowno += 1
            budget -= len(vals) + 1
            if budget < 0:
                raise HMAError(tr("Die ODS-Tabelle ist ungewöhnlich groß (über 20 Millionen Zellen) und wird nicht gelesen. "
                               "Bitte als .csv oder .xlsx speichern."))
            table.append((rowno, list(vals)))
    return table


# --- XLS (OLE2-Container + BIFF) ---
_ERR_TXT = {0x00: "#NULL!", 0x07: "#DIV/0!", 0x0F: "#VALUE!", 0x17: "#REF!", 0x1D: "#NAME?", 0x24: "#NUM!",
            0x2A: "#N/A"}


def _xls_guard(fn):
    """Technische Lesefehler bei beschädigten .xls-Dateien in eine verständliche Meldung umwandeln."""
    import functools
    import struct as _st

    @functools.wraps(fn)
    def wrapper(*a, **k):
        try:
            return fn(*a, **k)
        except HMAError:
            raise
        except (_st.error, IndexError, ValueError, KeyError, OverflowError, UnicodeError, TypeError):
            raise HMAError(tr("Die Excel-Datei (.xls) ist beschädigt oder hat einen nicht lesbaren Aufbau. Bitte in Excel "
                           "oder LibreOffice öffnen und als .xlsx oder .csv speichern."))
    return wrapper


def _ole_stream(data, names):
    """Liest einen benannten Stream aus einem OLE2-Container (Compound File)."""
    import struct
    if data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        raise HMAError(tr("Das ist kein Excel-Dokument im Format .xls (OLE2). Sehr alte Excel-Versionen (vor 5.0) werden "
                       "nicht unterstützt – bitte in Excel als .xlsx oder .csv speichern."))
    ssz = 1 << struct.unpack_from("<H", data, 30)[0]
    mssz = 1 << struct.unpack_from("<H", data, 32)[0]
    n_fat, dir_start, _t, cutoff, mf_start, n_mf, dif_start, n_dif = struct.unpack_from("<8I", data, 44)
    END = 0xFFFFFFFA
    difat = list(struct.unpack_from("<109I", data, 76))
    s = dif_start
    for _ in range(n_dif):
        if s >= END:
            break
        off = (s + 1) * ssz
        vals = struct.unpack_from("<%dI" % (ssz // 4), data, off)
        difat.extend(vals[:-1])
        s = vals[-1]
    fat = []
    for s in [x for x in difat if x < END][:n_fat]:
        fat.extend(struct.unpack_from("<%dI" % (ssz // 4), data, (s + 1) * ssz))

    def chain(start, table):
        out, s, guard = [], start, 0
        while s < END and guard < len(table) + 10:
            out.append(s)
            s = table[s] if s < len(table) else END
            guard += 1
        return out

    def sectors(ch):
        return b"".join(data[(s + 1) * ssz:(s + 2) * ssz] for s in ch)
    dirs = sectors(chain(dir_start, fat))
    entries = []
    for i in range(len(dirs) // 128):
        e = dirs[i * 128:(i + 1) * 128]
        nlen = struct.unpack_from("<H", e, 64)[0]
        name = e[:max(0, nlen - 2)].decode("utf-16le", errors="replace")
        size = struct.unpack_from("<Q", e, 120)[0] if ssz == 4096 else struct.unpack_from("<I", e, 120)[0]
        entries.append((name, e[66], struct.unpack_from("<I", e, 116)[0], size))
    target = next((e for e in entries if e[1] == 2 and e[0] in names), None)
    if target is None:
        raise HMAError(tr("In der .xls-Datei wurde kein Tabellenstream gefunden (evtl. ältere Excel-Version oder "
                       "verschlüsselt). Bitte als .xlsx oder .csv speichern."))
    _n, _t2, start, size = target
    if size < cutoff:
        root = next(e for e in entries if e[1] == 5)
        ministream = sectors(chain(root[2], fat))[:root[3]]
        minifat = []
        for s in chain(mf_start, fat):
            minifat.extend(struct.unpack_from("<%dI" % (ssz // 4), data, (s + 1) * ssz))
        return b"".join(ministream[s * mssz:(s + 1) * mssz] for s in chain(start, minifat))[:size]
    return sectors(chain(start, fat))[:size]


class _SegReader:
    """Liest Strings über CONTINUE-Grenzen hinweg (Shared String Table)."""

    def __init__(self, segs):
        self.segs, self.i, self.p = segs, 0, 0

    def left(self):
        return len(self.segs[self.i]) - self.p if self.i < len(self.segs) else 0

    def _next(self):
        self.i += 1
        self.p = 0

    def read(self, n):
        out = self.segs[self.i][self.p:self.p + n]
        self.p += n
        return out

    def skip(self, n):
        while n > 0 and self.i < len(self.segs):
            if self.left() == 0:
                self._next()
                continue
            k = min(n, self.left())
            self.p += k
            n -= k

    def string(self):
        import struct
        if self.left() == 0:
            self._next()
        if self.left() < 3:
            return None
        cch, flags = struct.unpack("<HB", self.read(3))
        runs = struct.unpack("<H", self.read(2))[0] if flags & 8 else 0
        ext = struct.unpack("<I", self.read(4))[0] if flags & 4 else 0
        is16, remaining, out = flags & 1, cch, []
        while remaining > 0:
            if self.left() == 0:
                self._next()
                if self.i >= len(self.segs):
                    break
                is16 = self.read(1)[0] & 1
            width = 2 if is16 else 1
            take = min(remaining, self.left() // width)
            if take <= 0:
                break
            raw = self.read(take * width)
            out.append(raw.decode("utf-16le", errors="replace") if is16 else raw.decode("latin-1"))
            remaining -= take
        self.skip(4 * runs + ext)
        return "".join(out)


def _biff_iter(buf, pos=0):
    import struct
    n = len(buf)
    while pos + 4 <= n:
        rid, ln = struct.unpack_from("<HH", buf, pos)
        yield rid, buf[pos + 4:pos + 4 + ln]
        pos += 4 + ln


def _biff8_string(d, off):
    """XLUnicodeString (16-Bit-Länge) ab off; gibt (Text, Folgeposition) zurück."""
    import struct
    cch, flags = struct.unpack_from("<HB", d, off)
    off += 3
    runs = ext = 0
    if flags & 8:
        runs = struct.unpack_from("<H", d, off)[0]
        off += 2
    if flags & 4:
        ext = struct.unpack_from("<I", d, off)[0]
        off += 4
    if flags & 1:
        txt = d[off:off + 2 * cch].decode("utf-16le", errors="replace")
        off += 2 * cch
    else:
        txt = d[off:off + cch].decode("latin-1")
        off += cch
    return txt, off + 4 * runs + ext


def _rk_value(rk):
    import struct
    if rk & 2:
        v = rk >> 2
        if v & 0x20000000:
            v -= 0x40000000
        f = float(v)
    else:
        f = struct.unpack("<d", struct.pack("<Q", (rk & 0xFFFFFFFC) << 32))[0]
    return f / 100.0 if rk & 1 else f


def _xls_workbook(path):
    """Liest Blattliste und Shared Strings. Gibt (buf, biff8, codepage, sheets, sst) zurück."""
    import struct
    try:
        data = Path(path).read_bytes()
    except OSError as e:
        raise HMAError(tr('Die Datei konnte nicht gelesen werden: {0}').format(e))
    buf = _ole_stream(data, ("Workbook", "Book"))
    biff8, cp, sheets, sst, depth = True, "cp1252", [], [], 0
    fmts, xfs, date1904 = {}, [], False
    recs = list(_biff_iter(buf))
    k = 0
    while k < len(recs):
        rid, d = recs[k]
        k += 1
        if rid in (0x0809, 0x0009):
            ver = struct.unpack_from("<H", d, 0)[0]
            if ver < 0x0500:
                raise HMAError(tr("Diese Excel-Datei ist zu alt (vor Excel 5.0). Bitte als .xlsx oder .csv speichern."))
            biff8 = ver >= 0x0600
            depth += 1
        elif rid == 0x000A:
            break
        elif rid == 0x002F:
            raise HMAError(tr("Die Excel-Datei ist passwortgeschützt/verschlüsselt und kann nicht gelesen werden."))
        elif rid == 0x041E:
            fid = struct.unpack_from("<H", d, 0)[0]
            try:
                fmts[fid] = _biff8_string(d, 2)[0] if biff8 else d[3:3 + d[2]].decode(cp, errors="replace")
            except (struct.error, IndexError):
                pass
        elif rid in (0x00E0, 0x0243, 0x0443, 0x0043) and len(d) >= 4:
            xfs.append(struct.unpack_from("<H", d, 2)[0])
        elif rid == 0x0022:
            date1904 = bool(struct.unpack_from("<H", d, 0)[0])
        elif rid == 0x0042:
            c = struct.unpack_from("<H", d, 0)[0]
            cp = "cp%d" % c if c not in (10000, 65001) else ("mac_roman" if c == 10000 else "utf-8")
        elif rid == 0x0085:
            off, vis, typ = struct.unpack_from("<IBB", d, 0)
            cch = d[6]
            if biff8:
                is16 = d[7] & 1
                name = d[8:8 + cch * (2 if is16 else 1)].decode("utf-16le" if is16 else "latin-1", errors="replace")
            else:
                name = d[7:7 + cch].decode(cp, errors="replace")
            sheets.append((name, off, typ))
        elif rid == 0x00FC and biff8:
            segs = [d[8:]]
            while k < len(recs) and recs[k][0] == 0x003C:
                segs.append(recs[k][1])
                k += 1
            rd = _SegReader(segs)
            total = struct.unpack_from("<I", d, 4)[0]
            for _ in range(total):
                s = rd.string()
                if s is None:
                    break
                sst.append(s)
    datexf = {i for i, f in enumerate(xfs) if f in _BUILTIN_DATE_FMT or (f in fmts and _is_date_fmt(fmts[f]))}
    return buf, biff8, cp, [s for s in sheets if s[2] == 0], sst, (datexf, date1904)


@_xls_guard
def xls_sheet_names(path):
    return [s[0] for s in _xls_workbook(path)[3]]


@_xls_guard
def _xls_read(path, sheet):
    import struct
    buf, biff8, cp, sheets, sst, (datexf, date1904) = _xls_workbook(path)
    if not sheets:
        raise HMAError(tr("Die Excel-Datei enthält keine Tabellenblätter."))
    _name, offset, _t = next((s for s in sheets if s[0] == sheet), sheets[0])
    cells, pending, depth = {}, None, 0

    def put(r, c, v):
        cells[(r, c)] = v

    def num(r, c, xf, f):
        put(r, c, _excel_date(f, date1904) if xf in datexf else _num_text(f))

    def text8(d, off):  # Text eines LABEL/STRING-Records
        if biff8:
            return _biff8_string(d, off)[0]
        n = struct.unpack_from("<H", d, off)[0]
        return d[off + 2:off + 2 + n].decode(cp, errors="replace")
    for rid, d in _biff_iter(buf, offset):
        if rid in (0x0809, 0x0009):
            depth += 1
        elif rid == 0x000A:
            depth -= 1
            if depth <= 0:
                break
        elif depth != 1:
            continue
        elif rid == 0x00FD and biff8:
            r, c, _x, i = struct.unpack_from("<HHHI", d, 0)
            put(r, c, sst[i] if i < len(sst) else "")
        elif rid == 0x0203:
            r, c, xf = struct.unpack_from("<HHH", d, 0)
            num(r, c, xf, struct.unpack_from("<d", d, 6)[0])
        elif rid == 0x027E:
            r, c, xf, rk = struct.unpack_from("<HHHI", d, 0)
            num(r, c, xf, _rk_value(rk))
        elif rid == 0x00BD:
            r, c1 = struct.unpack_from("<HH", d, 0)
            last = struct.unpack_from("<H", d, len(d) - 2)[0]
            for j in range(last - c1 + 1):
                xf, rk = struct.unpack_from("<HI", d, 4 + 6 * j)
                num(r, c1 + j, xf, _rk_value(rk))
        elif rid in (0x0204, 0x00D6):
            r, c, _x = struct.unpack_from("<HHH", d, 0)
            put(r, c, text8(d, 6))
        elif rid == 0x0006:
            r, c, xf = struct.unpack_from("<HHH", d, 0)
            res = d[6:14]
            if res[6:8] == b"\xff\xff":
                t = res[0]
                if t == 0:
                    pending = (r, c)
                elif t == 1:
                    put(r, c, "TRUE" if res[2] else "FALSE")
                elif t == 2:
                    put(r, c, _ERR_TXT.get(res[2], "#ERR"))
                else:
                    put(r, c, "")
            else:
                num(r, c, xf, struct.unpack("<d", res)[0])
        elif rid == 0x0207 and pending:
            put(pending[0], pending[1], text8(d, 0))
            pending = None
        elif rid == 0x0205:
            r, c, _x = struct.unpack_from("<HHH", d, 0)
            put(r, c, (_ERR_TXT.get(d[6], "#ERR")) if d[7] else ("TRUE" if d[6] else "FALSE"))
    if not cells:
        return []
    by_row = {}
    for (r, c), v in cells.items():
        by_row.setdefault(r, {})[c] = v
    table = []
    for r in sorted(by_row):
        width = max(by_row[r]) + 1
        table.append((r + 1, [by_row[r].get(i, "") for i in range(width)]))
    return table


# --- DBF (dBase III/IV, FoxPro, Visual FoxPro) ---
_DBF_LANG = {0x01: "cp437", 0x02: "cp850", 0x03: "cp1252", 0x04: "mac_roman", 0x08: "cp865", 0x09: "cp437",
             0x0A: "cp850", 0x0B: "cp437", 0x0D: "cp437", 0x0E: "cp850", 0x0F: "cp437", 0x10: "cp850", 0x11: "cp437",
             0x12: "cp850", 0x13: "cp932", 0x14: "cp850", 0x15: "cp437", 0x16: "cp850", 0x17: "cp865", 0x18: "cp437",
             0x19: "cp437", 0x1A: "cp850", 0x1B: "cp437", 0x1C: "cp863", 0x1D: "cp850", 0x1F: "cp852", 0x22: "cp852",
             0x23: "cp852", 0x24: "cp860", 0x25: "cp850", 0x26: "cp866", 0x37: "cp850", 0x40: "cp852", 0x4D: "cp936",
             0x4E: "cp949", 0x4F: "cp950", 0x50: "cp874", 0x57: "cp1252", 0x58: "cp1252", 0x59: "cp1252",
             0x64: "cp852", 0x65: "cp866", 0x66: "cp865", 0x67: "cp861", 0x6A: "cp737", 0x6B: "cp857", 0x6C: "cp863",
             0x78: "cp950", 0x79: "cp949", 0x7A: "cp936", 0x7B: "cp932", 0x7C: "cp874", 0x86: "cp737", 0x87: "cp852",
             0x88: "cp857", 0xC8: "cp1250", 0xC9: "cp1251", 0xCA: "cp1254", 0xCB: "cp1253", 0xCC: "cp1257"}


def _sibling(path, exts):
    for e in exts:
        for cand in (path.with_suffix(e), path.with_suffix(e.upper())):
            if cand.exists():
                return cand
    return None


def _codec(name):
    import codecs
    n = (name or "").strip().lower().replace("_", "-")
    if n.isdigit():
        n = "cp" + n
    try:
        return codecs.lookup(n).name
    except LookupError:
        return None


def _dbf_memo_reader(path, vfp):
    mf = _sibling(path, (".fpt", ".dbt") if vfp else (".dbt", ".fpt"))
    if not mf:
        return None
    import struct
    raw = mf.read_bytes()
    is_fpt = mf.suffix.lower() == ".fpt"
    if is_fpt:
        bs = struct.unpack_from(">H", raw, 6)[0] or 64
    else:
        bs = struct.unpack_from("<H", raw, 20)[0] if len(raw) > 22 and raw[16:20] != b"\0\0\0\0" else 0
        bs = bs if bs >= 64 and bs % 64 == 0 else 512

    def read(block):
        pos = block * bs
        if block <= 0 or pos >= len(raw):
            return b""
        if is_fpt:
            _typ, ln = struct.unpack_from(">II", raw, pos)
            return raw[pos + 8:pos + 8 + ln]
        if raw[pos:pos + 4] == b"\xff\xff\x08\x00":  # dBase IV
            ln = struct.unpack_from("<I", raw, pos + 4)[0]
            return raw[pos + 8:pos + ln]
        end = raw.find(b"\x1a", pos)
        return raw[pos:end if end >= 0 else len(raw)]
    return read


def _dbf_read(path, encoding="auto"):
    """Liest eine DBF-Datei. Gibt (Spaltennamen, [(Zeilennummer, Werte)]) zurück."""
    import struct
    from datetime import datetime, timedelta
    path = Path(path)
    raw = path.read_bytes()
    if len(raw) < 33:
        raise HMAError(tr('Die DBF-Datei ist leer oder beschädigt: {0}').format(path.name))
    ver, nrec, hlen, rlen, ldid = raw[0], struct.unpack_from("<I", raw, 4)[0], struct.unpack_from("<H", raw, 8)[0], \
        struct.unpack_from("<H", raw, 10)[0], raw[29]
    vfp = ver in (0x30, 0x31, 0x32)
    fields, pos = [], 32
    while pos + 32 <= len(raw) and raw[pos] != 0x0D:
        d = raw[pos:pos + 32]
        ln, dec = d[16], d[17]
        if chr(d[11]) == "C" and dec and not vfp:  # Clipper/FoxPro 2: Zeichenfelder > 255 Byte
            ln, dec = ln + 256 * dec, 0
        fields.append((d[:11].split(b"\0")[0], chr(d[11]), ln, dec))
        pos += 32
    if not fields or rlen < 2:
        raise HMAError(tr('Die DBF-Datei enthält keine Spalten: {0}').format(path.name))
    start = hlen if hlen >= pos + 1 else pos + 1
    avail = (len(raw) - start) // rlen
    n = nrec if 0 < nrec <= avail else avail
    body = raw[start:start + n * rlen]
    enc = None
    if encoding not in ("auto", "", None):
        enc = _codec(encoding)
    if not enc:
        cpg = _sibling(path, (".cpg",))
        if cpg:
            enc = _codec(cpg.read_text(encoding="ascii", errors="ignore"))
    if not enc and ldid in _DBF_LANG:
        enc = _DBF_LANG[ldid]
    if not enc:
        try:
            body.decode("utf-8")
            enc = "utf-8"
        except UnicodeDecodeError:
            enc = "cp1252"
    memo = _dbf_memo_reader(path, vfp) if any(f[1] in "MGPB" and (f[1] == "M" or f[1] != "B") for f in fields) else None
    names, spec, off = [], [], 1
    for nm, typ, ln, dec in fields:
        if typ == "0" or nm.startswith(b"_NullFlags"):
            off += ln
            continue
        names.append(nm.decode(enc, errors="replace").strip())
        spec.append((typ, off, ln, dec))
        off += ln
    rows = []
    for i in range(n):
        rec = body[i * rlen:(i + 1) * rlen]
        if rec[:1] == b"*":  # als gelöscht markiert
            continue
        vals = []
        for typ, o, ln, dec in spec:
            b = rec[o:o + ln]
            v = ""
            try:
                if typ == "C":
                    v = b.decode(enc, errors="replace").rstrip(" \0")
                elif typ in "NF":
                    s = b.decode("ascii", errors="ignore").strip()
                    if s and not set(s) <= {"*", "."}:
                        f = float(s)
                        v = str(int(f)) if (dec == 0 or f.is_integer()) else format(f, ".15g")
                elif typ == "D":
                    s = b.decode("ascii", errors="ignore").strip()
                    v = f"{s[:4]}-{s[4:6]}-{s[6:8]}" if len(s) == 8 and s.isdigit() and s != "00000000" else s
                elif typ == "L":
                    c = b[:1].decode("ascii", errors="ignore").upper()
                    v = "TRUE" if c in "TY" and c else ("FALSE" if c in "FN" and c else "")
                elif typ in "I+" and ln == 4:
                    v = str(struct.unpack("<i", b)[0])
                elif typ == "2" and ln == 2:
                    v = str(struct.unpack("<h", b)[0])
                elif typ in "BO" and ln == 8:
                    v = _num_text(struct.unpack("<d", b)[0])
                elif typ == "Y" and ln == 8:
                    v = _num_text(struct.unpack("<q", b)[0] / 10000.0)
                elif typ in "T@" and ln == 8:
                    jd, ms = struct.unpack("<II", b)
                    if jd:
                        dt = datetime(1970, 1, 1) + timedelta(days=jd - 2440588, milliseconds=ms)
                        v = dt.strftime("%Y-%m-%d %H:%M:%S")
                elif typ in "MGP":
                    blk = 0
                    if ln == 4:
                        blk = struct.unpack("<I", b)[0]
                    else:
                        s = b.decode("ascii", errors="ignore").strip()
                        blk = int(s) if s.isdigit() else 0
                    if memo and blk and typ == "M":
                        v = memo(blk).decode(enc, errors="replace").rstrip("\0\x1a")
                else:
                    v = b.decode(enc, errors="replace").strip()
            except (ValueError, struct.error, OverflowError):
                v = ""
            vals.append(v)
        rows.append((i + 1, vals))
    return _unique_headers(names), rows


# ---------------------------------------------------------------------------
# Lokale Datenbank-Dateien (nur lesend): Access (.accdb/.mdb), LibreOffice Base (.odb), SQLite
# ---------------------------------------------------------------------------
DB_FILE_EXT = {".accdb": "Access", ".mdb": "Access", ".odb": "LibreOffice Base",
               ".sqlite": "SQLite", ".sqlite3": "SQLite", ".db": "SQLite", ".gpkg": "SQLite"}
SHEET_EXT = (".xlsx", ".xlsm", ".xls", ".ods") + tuple(DB_FILE_EXT)


def output_stem(path, sheet="", n_sheets=0):
    """Name der Ausgabedatei ohne Endung: <Datei>; bei Mappen/Datenbanken mit mehreren Blättern/Tabellen
    <Datei>_<Blatt>, damit die Matrizen verschiedener Blätter einander nicht überschreiben."""
    stem = Path(path).stem
    if sheet and n_sheets > 1:
        safe = re.sub(r'[<>:"/\\|?*\x00-\x1f\s]+', "_", sheet).strip("._")
        if safe:
            stem += "_" + safe
    return stem


def sheet_names(path):
    """Blätter einer Excel-Datei bzw. Tabellen einer Datenbank-Datei."""
    path = Path(path)
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        return xlsx_sheet_names(path)
    if ext == ".xls":
        kind = _xls_kind(path)
        return xlsx_sheet_names(path) if kind == "xlsx" else (xls_sheet_names(path) if kind == "xls" else [])
    if ext == ".ods":
        return ods_sheet_names(path)
    kind = DB_FILE_EXT.get(ext)
    if kind == "SQLite":
        return _sqlite_tables(path)
    if kind == "Access":
        return _access_tables(path)
    if kind == "LibreOffice Base":
        return list(_odb_load(path))
    return []


def db_file_table(path, table=""):
    """Liest eine Tabelle aus einer Datenbank-Datei. Gibt (Spaltennamen, [Zeilen als Listen], Tabellenname) zurück."""
    path = Path(path)
    kind = DB_FILE_EXT.get(path.suffix.lower())
    names = [table] if (kind == "Access" and table) else sheet_names(path)  # Access: nicht erneut auflisten (langsam)
    if not names:
        raise HMAError(tr('In {0} wurde keine Tabelle gefunden.').format(path.name))
    name = table if table in names else names[0]
    if kind == "SQLite":
        cols, rows = _sqlite_read(path, name)
    elif kind == "Access":
        cols, rows = _access_read(path, name)
    else:
        cols, rows = _odb_load(path)[name]
    return cols, rows, name


def _copy_to_temp(path):
    import shutil
    import tempfile
    import atexit
    d = Path(tempfile.mkdtemp(prefix="hma_"))
    atexit.register(shutil.rmtree, str(d), True)  # Datenkopie beim Beenden wieder löschen
    target = d / path.name
    shutil.copy2(path, target)
    return target


def _sqlite_connect(path):
    import sqlite3
    try:
        con = sqlite3.connect(Path(path).resolve().as_uri() + "?mode=ro", uri=True)
        con.execute("SELECT count(*) FROM sqlite_master").fetchone()
        return con
    except Exception:  # z. B. Datei gesperrt oder WAL: Kopie lesen
        return sqlite3.connect(str(_copy_to_temp(Path(path))))


def _sqlite_tables(path):
    try:
        con = _sqlite_connect(path)
        try:
            rows = con.execute("SELECT name FROM sqlite_master WHERE type IN ('table','view') "
                               "AND name NOT LIKE 'sqlite\\_%' ESCAPE '\\' AND name NOT LIKE 'gpkg\\_%' ESCAPE '\\' "
                               "AND name NOT LIKE 'rtree\\_%' ESCAPE '\\' ORDER BY name").fetchall()
        finally:
            con.close()
    except Exception as e:
        raise HMAError(tr('Die Datenbank konnte nicht gelesen werden ({0}): {1}').format(Path(path).name, e))
    return [r[0] for r in rows]


def _sqlite_read(path, table):
    try:
        con = _sqlite_connect(path)
        try:
            cur = con.execute('SELECT * FROM "' + table.replace('"', '""') + '"')
            cols = [d[0] for d in cur.description]
            data = cur.fetchall()
        finally:
            con.close()
    except Exception as e:
        raise HMAError(tr("Tabelle '{0}' konnte nicht gelesen werden: {1}").format(table, e))

    def txt(v):
        if v is None or isinstance(v, (bytes, bytearray)):
            return ""
        if isinstance(v, float):
            return str(int(v)) if v.is_integer() else format(v, ".15g")
        return str(v)
    return cols, [[txt(v) for v in r] for r in data]


# --- Access über PowerShell/OleDb (benötigt Windows und die Access-Datenbank-Engine, z. B. durch Office) ---
_ACCESS_PS = r'''param([string]$Db, [string]$Out, [string]$Table = "")
$ErrorActionPreference = 'Stop'
$conn = $null; $lastErr = ''
foreach ($p in @('Microsoft.ACE.OLEDB.16.0', 'Microsoft.ACE.OLEDB.12.0', 'Microsoft.Jet.OLEDB.4.0')) {
  try {
    $c = New-Object System.Data.OleDb.OleDbConnection("Provider=$p;Data Source=$Db;Mode=Read;Persist Security Info=False")
    $c.Open(); $conn = $c; break
  } catch { $lastErr = $_.Exception.Message }
}
if (-not $conn) { [Console]::Error.WriteLine("NOCONN: " + $lastErr); exit 3 }
$utf8 = New-Object System.Text.UTF8Encoding($false)
try {
  if ($Table -eq '') {
    $s = $conn.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($r in $s.Rows) {
      $t = [string]$r['TABLE_TYPE']; $n = [string]$r['TABLE_NAME']
      if (($t -eq 'TABLE' -or $t -eq 'VIEW') -and -not $n.StartsWith('MSys') -and -not $n.StartsWith('~')) { $names.Add($n) }
    }
    [IO.File]::WriteAllText($Out, (($names | Sort-Object) -join "`n"), $utf8)
  } else {
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = 'SELECT * FROM [' + $Table.Replace(']', ']]') + ']'
    $dt = New-Object System.Data.DataTable
    $ad = New-Object System.Data.OleDb.OleDbDataAdapter $cmd
    [void]$ad.Fill($dt)
    $inv = [Globalization.CultureInfo]::InvariantCulture
    $sb = New-Object System.Text.StringBuilder
    function Q($v) {
      if ($v -eq $null -or $v -is [DBNull] -or $v -is [byte[]]) { return '""' }
      if ($v -is [datetime]) { $t = $v.ToString('yyyy-MM-dd HH:mm:ss', $inv) }
      elseif ($v -is [IFormattable]) { $t = $v.ToString($null, $inv) }
      else { $t = [string]$v }
      return '"' + $t.Replace('"', '""') + '"'
    }
    $hdr = @(); foreach ($col in $dt.Columns) { $hdr += (Q $col.ColumnName) }
    [void]$sb.Append(($hdr -join ';') + "`r`n")
    foreach ($row in $dt.Rows) {
      $cells = @(); foreach ($v in $row.ItemArray) { $cells += (Q $v) }
      [void]$sb.Append(($cells -join ';') + "`r`n")
    }
    [IO.File]::WriteAllText($Out, $sb.ToString(), $utf8)
  }
} catch { [Console]::Error.WriteLine("ERR: " + $_.Exception.Message); exit 4 }
finally { $conn.Close() }
'''


def _access_run(path, table=""):
    import subprocess
    import tempfile
    path = Path(path)
    if os.name != "nt":
        raise HMAError(tr("Access-Dateien können nur unter Windows gelesen werden. Bitte die Tabelle in Access als "
                       "CSV exportieren."))
    import atexit
    import shutil
    tmp = Path(tempfile.mkdtemp(prefix="hma_"))
    atexit.register(shutil.rmtree, str(tmp), True)  # ausgelesene Tabelle beim Beenden wieder löschen
    ps1, out = tmp / "read.ps1", tmp / "out.txt"
    ps1.write_text(_ACCESS_PS, encoding="utf-8-sig")
    root = os.environ.get("SystemRoot", r"C:\Windows")
    exes = ["powershell.exe"] + [p for p in (root + r"\SysWOW64\WindowsPowerShell\v1.0\powershell.exe",
                                             root + r"\Sysnative\WindowsPowerShell\v1.0\powershell.exe")
                                 if os.path.exists(p)]
    last = ""
    for exe in exes:
        try:
            cmd = [exe, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(ps1),
                   "-Db", str(path), "-Out", str(out)] + (["-Table", table] if table else [])
            r = subprocess.run(cmd, capture_output=True, timeout=180, creationflags=0x08000000)
        except FileNotFoundError:
            continue
        except subprocess.TimeoutExpired:
            raise HMAError(tr("Das Lesen der Access-Datei hat zu lange gedauert (Zeitüberschreitung)."))
        err = (r.stderr or b"").decode("utf-8", errors="replace").strip()
        if r.returncode == 0 and out.exists():
            return out.read_text(encoding="utf-8-sig")
        last = err or tr('Exit-Code {0}').format(r.returncode)
        if r.returncode != 3:  # nur "Datenbank-Engine nicht gefunden" mit der anderen PowerShell-Variante (32/64 Bit) wiederholen
            break
    hint = ""
    if "NOCONN" in last:
        hint = (tr("\n\nDie Access-Datenbank-Engine (OLE DB) wurde nicht gefunden bzw. passt nicht zur Version von "
                "Office/Python (32 oder 64 Bit). Mit installiertem Microsoft Office ist sie normalerweise vorhanden; "
                "sonst 'Microsoft Access Database Engine 2016 Redistributable' installieren. Alternativ die Tabelle in "
                "Access als CSV exportieren (Externe Daten > Exportieren > Textdatei)."))
    raise HMAError(tr('Die Access-Datei konnte nicht gelesen werden ({0}):\n{1}{2}').format(path.name, last, hint))


def _access_tables(path):
    return [n for n in _access_run(path).splitlines() if n.strip()]


def _access_read(path, table):
    text = _access_run(path, table)
    reader = csv.reader(io.StringIO(text, newline=""), delimiter=";", quotechar='"')
    rows = [r for r in reader]
    if not rows:
        raise HMAError(tr("Tabelle '{0}' ist leer.").format(table))
    return rows[0], rows[1:]


# --- LibreOffice Base: eingebettete HSQLDB (Textskript in der .odb); nur Tabellen, keine gespeicherten Abfragen ---
_ODB_CACHE = {}
_SQL_ID = r'(?:"((?:[^"]|"")+)"|([A-Za-z_][\w$]*))'


def _sql_split_top(s):
    parts, depth, cur, q = [], 0, [], None
    for ch in s:
        if q:
            cur.append(ch)
            if ch == q:
                q = None
        elif ch in "'\"":
            q = ch
            cur.append(ch)
        elif ch == "(":
            depth += 1
            cur.append(ch)
        elif ch == ")":
            depth -= 1
            cur.append(ch)
        elif ch == "," and depth == 0:
            parts.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
    if cur:
        parts.append("".join(cur))
    return parts


def _sql_values(s):
    """Werteliste eines HSQLDB-INSERT: 'Text' (mit '' und \\uXXXX), NULL, Zahlen."""
    out, i, n = [], 0, len(s)
    while i < n:
        while i < n and s[i] in " ,":
            i += 1
        if i >= n:
            break
        if s[i] == "'":
            i += 1
            buf = []
            while i < n:
                if s[i] == "'":
                    if i + 1 < n and s[i + 1] == "'":
                        buf.append("'")
                        i += 2
                        continue
                    i += 1
                    break
                buf.append(s[i])
                i += 1
            txt = "".join(buf)
            out.append(re.sub(r"\\u([0-9a-fA-F]{4})", lambda m: chr(int(m.group(1), 16)), txt))
        else:
            j = i
            while j < n and s[j] != ",":
                j += 1
            tok = s[i:j].strip()
            i = j
            if tok.upper() == "NULL":
                out.append("")
            else:
                m = re.match(r"^[A-Za-z]+\s*'(.*)'$", tok)  # z. B. TIMESTAMP'2020-01-01 00:00:00'
                out.append(m.group(1) if m else tok)
    return out


def _odb_load(path):
    path = Path(path)
    st = path.stat()
    key = (str(path), st.st_mtime_ns, st.st_size)
    if key in _ODB_CACHE:
        return _ODB_CACHE[key]
    try:
        zf = zipfile.ZipFile(path)
    except Exception as e:
        raise HMAError(tr('Die Base-Datei konnte nicht geöffnet werden ({0}): {1}').format(path.name, e))
    with zf:
        names = zf.namelist()
        if "database/script" not in names:
            content = zf.read("content.xml").decode("utf-8", "ignore") if "content.xml" in names else ""
            if any("firebird" in n.lower() for n in names):
                raise HMAError(tr("Diese Base-Datei verwendet die eingebettete Firebird-Datenbank, die ohne Zusatzsoftware "
                               "nicht gelesen werden kann. Bitte die Tabelle in LibreOffice als CSV exportieren "
                               "(Tabelle öffnen, Daten kopieren in Calc, als CSV speichern) oder in Base unter "
                               "Extras > SQL die Daten als Textdatei exportieren."))
            if "connection-resource" in content:
                raise HMAError(tr("Diese Base-Datei ist nur eine Verknüpfung auf eine externe Datenbank. Bitte die "
                               "zugrunde liegende Datenbank-Datei (z. B. Access oder SQLite) direkt auswählen."))
            raise HMAError(tr("In dieser Base-Datei wurde keine lesbare eingebettete Datenbank gefunden."))
        script = zf.read("database/script").decode("utf-8", errors="replace")
        has_data = "database/data" in names and zf.getinfo("database/data").file_size > 0
    tables, order, cached = {}, [], set()
    create = re.compile(r'^CREATE\s+(?:(MEMORY|CACHED|TEXT)\s+)?TABLE\s+(?:PUBLIC\.)?' + _SQL_ID + r'\s*\((.*)\)\s*$', re.I)
    insert = re.compile(r'^INSERT\s+INTO\s+(?:PUBLIC\.)?' + _SQL_ID + r'\s+VALUES\s*\((.*)\)\s*$', re.I)
    for line in script.splitlines():
        line = line.strip()
        m = create.match(line)
        if m:
            name = (m.group(2) or "").replace('""', '"') or m.group(3)
            if (m.group(1) or "").upper() in ("CACHED", "TEXT"):
                cached.add(name)
            cols = []
            for d in _sql_split_top(m.group(4)):
                d = d.strip()
                if re.match(r"^(CONSTRAINT|PRIMARY\s+KEY|FOREIGN\s+KEY|UNIQUE|CHECK)\b", d, re.I):
                    continue
                mc = re.match(_SQL_ID, d)
                if mc:
                    cols.append((mc.group(1) or "").replace('""', '"') or mc.group(2))
            tables[name] = (cols, [])
            order.append(name)
            continue
        m = insert.match(line)
        if m:
            name = (m.group(1) or "").replace('""', '"') or m.group(2)
            if name in tables:
                tables[name][1].append(_sql_values(m.group(3)))
    if not tables:
        raise HMAError(tr("In dieser Base-Datei wurden keine Tabellen gefunden."))
    if cached and has_data and any(t in cached and not tables[t][1] for t in tables):
        raise HMAError(tr("Diese Base-Datei speichert die Daten in einem Format, das ohne Zusatzsoftware nicht gelesen "
                       "werden kann (CACHED-Tabellen). Bitte die Tabelle als CSV exportieren."))
    result = {}
    for name in order:
        cols, rows = tables[name]
        result[name] = (cols, [(r + [""] * len(cols))[:len(cols)] for r in rows])
    _ODB_CACHE.clear()
    _ODB_CACHE[key] = result
    return result


def parse_csv_bytes(raw, settings, name):
    """CSV/Text (Bytes) -> (headers, rows, info); rows = [(zeilennummer, [werte])]."""
    text = _decode(raw, settings.get("encoding", "auto"))
    delim = settings.get("delimiter", "auto")
    delim = {"tab": "\t"}.get(delim, delim)
    if delim == "auto":
        delim = _detect_delimiter(text)

    class _Dialect(csv.Dialect):  # RFC 4180
        delimiter = delim
        quotechar = '"'
        doublequote = True
        escapechar = None
        skipinitialspace = False
        lineterminator = "\n"
        quoting = csv.QUOTE_MINIMAL

    reader = csv.reader(io.StringIO(text, newline=""), dialect=_Dialect)
    try:
        headers = _unique_headers(next(reader))
    except StopIteration:
        raise HMAError(tr('Datei ist leer: {0}').format(name))
    rows = [(reader.line_num, row) for row in reader if row]
    return headers, rows, tr("Trennzeichen: ") + ("Tab" if delim == "\t" else delim)



def load_table(path, settings, sheet=""):
    """Liest CSV/TXT/XLSX sowie Tabellen aus Access-/Base-/SQLite-Dateien. Gibt (headers, rows, info) zurück; rows = [(zeilennummer, [werte])]."""
    path = Path(path)
    ext = path.suffix.lower()
    if ext == ".dbf":
        try:
            cols, data = _dbf_read(path, settings.get("encoding", "auto"))
        except HMAError:
            raise
        except Exception as e:
            raise HMAError(tr('DBF-Datei konnte nicht gelesen werden: {0}').format(e))
        rows = [(i, (r + [""] * len(cols))[:len(cols)]) for i, r in data if any(str(x).strip() for x in r)]
        if not rows:
            raise HMAError(tr('Tabelle ist leer: {0}').format(path.name))
        return cols, rows, "dBase (DBF)"
    if ext in DB_FILE_EXT:
        cols, data, tname = db_file_table(path, sheet)
        headers = _unique_headers([str(c).strip() for c in cols])
        rows = [(i, (list(r) + [""] * len(headers))[:len(headers)]) for i, r in enumerate(data, start=2)
                if any(str(x).strip() for x in r)]
        if not rows:
            raise HMAError(tr('Tabelle ist leer: {0}').format(tname))
        return headers, rows, tr('{0}, Tabelle {1}').format(DB_FILE_EXT[ext], tname)
    if ext == ".xls":
        kind = _xls_kind(path)
        if kind == "text":
            raw = path.read_bytes()
            if raw.lstrip()[:1] == b"<":
                raise HMAError(tr('{0} ist keine echte Excel-Datei, sondern eine Webseite (HTML) mit der Endung .xls. Bitte in Excel '
                                  'öffnen und als .xlsx oder .csv speichern.').format(path.name))
            return parse_csv_bytes(raw, settings, path.name)
    if ext in (".xlsx", ".xlsm", ".xls", ".ods"):
        reader = {".xlsx": _xlsx_read, ".xlsm": _xlsx_read, ".xls": _xls_read, ".ods": _ods_read}[ext]
        if ext == ".xls" and kind == "xlsx":
            reader = _xlsx_read
        try:
            table = reader(path, sheet)
        except HMAError:
            raise
        except Exception as e:
            raise HMAError(tr('Die Datei konnte nicht gelesen werden ({0}): {1}').format(path.name, e))
        table = [(n, r) for n, r in table if any(str(x).strip() for x in r)]
        if not table:
            raise HMAError(tr('Tabelle ist leer: {0}').format(path.name))
        headers = _unique_headers([str(h).strip() for h in table[0][1]])
        rows = []
        for n, r in table[1:]:
            r = (r + [""] * len(headers))[:len(headers)]
            rows.append((n, r))
        return headers, rows, {".ods": "LibreOffice Calc", ".xls": "Excel (.xls)"}.get(ext, "Excel")

    return parse_csv_bytes(path.read_bytes(), settings, path.name)

def used_columns(schema):
    m, st = schema["mapping"], schema["settings"]
    cols = [(k, m[k]) for k in MAP_KEYS if m.get(k) and m[k] != PREFIX_COL]
    if st.get("ref_key_col"):
        cols.append(("ref_key_col", st["ref_key_col"]))
    return cols


def check_schema(schema, headers):
    """Gibt (probleme_erforderlich, probleme_optional) zurück."""
    m = schema["mapping"]
    hs = set(headers)
    req, opt = [], []
    if not m.get("id1"):
        req.append(tr("ID: mindestens Teil 1 wählen"))
    if not m.get("above"):
        req.append(tr("'liegt über': Spalte wählen"))
    for k, col in used_columns(schema):
        if col not in hs:
            label = tr("Verweise bestehen aus") if k == "ref_key_col" else ROLE_LABELS.get(k, k)
            (req if k in ("id1", "above") else opt).append(tr("{0} → '{1}' nicht in der Tabelle").format(label, col))
    for group in DISTINCT_GROUPS:  # dieselbe Spalte für zwei Aufgaben ergibt unsinnige Matrizen
        seen = {}
        for k in group:
            col = m.get(k)
            if col:
                seen.setdefault(col, []).append(ROLE_LABELS[k])
        for col, roles in seen.items():
            if len(roles) > 1:
                req.append(tr("Die Spalte '{0}' ist mehrfach zugeordnet ({1}) – jede Aufgabe braucht ihre eigene Spalte").format(col, tr(' und ').join(roles)))
    return req, opt


def distinct_values(headers, rows, col, schema=None):
    if col == PREFIX_COL:
        st = (schema or default_schema())["settings"]
        try:
            pre = re.compile(st["id_prefix_regex"]) if st["id_prefix_regex"] else None
        except re.error:
            pre = None
        mp = (schema or default_schema())["mapping"]
        idxs = [headers.index(mp[k]) for k in ("id1", "id2", "id3", "id4") if mp.get(k) in headers]
        out = Counter()
        for _, r in rows:
            rid = "_".join(p for p in ((r[i] if i < len(r) else "").strip() for i in idxs) if p)
            mm = pre.match(rid) if pre else None
            out[mm.group(0).rstrip("_") if mm else ""] += 1
        return out
    if col not in headers:
        return Counter()
    i = headers.index(col)
    return Counter((r[i] if i < len(r) else "").strip() for _, r in rows)


def _norm(s):
    return re.sub(r"[^0-9a-zäöüß]", "", s.lower())


def suggest_mapping(headers):
    """Schlägt Spalten vor, in der Reihenfolge der Aliase: exakter Treffer, sonst kürzester Teiltreffer."""
    normed = [(h, _norm(h)) for h in headers]
    out = {}
    for key, aliases in ASSIGN_ALIASES.items():
        for a in aliases:
            na = _norm(a)
            hit = next((h for h, n in normed if n == na), None)
            if not hit and len(na) >= 4:
                cands = [(len(n), h) for h, n in normed if na in n]
                hit = min(cands)[1] if cands else None
            if hit:
                out[key] = hit
                break
    return out


# ---------------------------------------------------------------------------
# Hilfsfunktionen für den XML-Aufbau (aus den Originalskripten)
# ---------------------------------------------------------------------------
def indent(elem, level=0):
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        for child in elem:
            indent(child, level + 1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if not elem.tail or not elem.tail.strip():
            elem.tail = i


def make_hmcnode(node_id, name, description, node_type, open_flag=False, layer="0"):
    data = ET.Element("data", {"key": "d2"})
    attrs = {"bookmarked": "false", "description": description or "", "id": node_id, "index": "0",
             "layer": layer, "name": name or "", "type": node_type, "valid": "true", "x": "0.0", "y": "0.0"}
    if open_flag:
        attrs["open"] = "true"
    h = ET.Element("hmcnode", attrs)
    h.set("xmlns", "")
    data.append(h)
    return data


def make_hmcedge(edge_type):
    h = ET.Element("hmcedge", {"type": edge_type, "valid": "true"})
    h.set("xmlns", "")
    ET.SubElement(h, "source", {"x": "0.0", "y": "0.0"})
    ET.SubElement(h, "target", {"x": "0.0", "y": "0.0"})
    return h


def se_first(base_id):
    return base_id.split("_", 1)[0]


def se_num(base_id):
    try:
        return int(se_first(base_id))
    except Exception:
        return 10 ** 9


def se_label(base_id):
    """Erster ID-Teil für Merge-IDs: Zahl ohne führende Nullen, sonst unverändert."""
    p = se_first(base_id)
    try:
        return str(int(p))
    except Exception:
        return p


def se_suffix(base_id):
    parts = base_id.split("_", 1)
    return parts[1] if len(parts) == 2 else ""


class UnionFind:
    def __init__(self):
        self.parent = {}
        self.rank = {}

    def add(self, x):
        if x not in self.parent:
            self.parent[x] = x
            self.rank[x] = 0

    def find(self, x):
        p = self.parent[x]
        if p != x:
            self.parent[x] = self.find(p)
        return self.parent[x]

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra == rb:
            return
        if self.rank[ra] < self.rank[rb]:
            ra, rb = rb, ra
        self.parent[rb] = ra
        if self.rank[ra] == self.rank[rb]:
            self.rank[ra] += 1


def strongly_connected(adj):
    """Iterativer Tarjan: gibt alle starken Zusammenhangskomponenten zurück."""
    index, low, stack, on, out = {}, {}, [], set(), []
    counter = 0
    nodes = set(adj) | {w for ws in adj.values() for w in ws}
    for root in sorted(nodes):
        if root in index:
            continue
        work = [(root, 0)]
        while work:
            v, pi = work.pop()
            if pi == 0:
                index[v] = low[v] = counter
                counter += 1
                stack.append(v)
                on.add(v)
            nbrs = adj.get(v, [])
            descended = False
            for k in range(pi, len(nbrs)):
                w = nbrs[k]
                if w not in index:
                    work.append((v, k + 1))
                    work.append((w, 0))
                    descended = True
                    break
                elif w in on:
                    low[v] = min(low[v], index[w])
            if descended:
                continue
            if low[v] == index[v]:
                comp = []
                while True:
                    w = stack.pop()
                    on.discard(w)
                    comp.append(w)
                    if w == v:
                        break
                out.append(comp)
            if work:
                u = work[-1][0]
                low[u] = min(low[u], low[v])
    return out


def shortest_cycle(adj, comp):
    """Kürzester Kreis durch die kleinste SE der Komponente (als Liste, Start = Ende)."""
    members = set(comp)
    start = sorted(comp, key=lambda x: (se_num(x), x))[0]
    prev = {}
    queue = [start]
    seen = {start}
    while queue:
        v = queue.pop(0)
        for w in adj.get(v, []):
            if w not in members:
                continue
            if w == start:
                path = [v]
                while path[-1] != start:
                    path.append(prev[path[-1]])
                path.reverse()
                return path + [start]
            if w not in seen:
                seen.add(w)
                prev[w] = v
                queue.append(w)
    return [start]


def verify_hmcx(path):
    """Selbstprüfung der erzeugten .hmcx. Gibt eine Liste von Problemen zurück (leer = in Ordnung)."""
    problems = []
    try:
        with zipfile.ZipFile(path) as zf:
            if sorted(zf.namelist()) != ["matrix.xml", "project.xml"]:
                problems.append(tr("ZIP enthält nicht genau matrix.xml und project.xml"))
            root = ET.fromstring(zf.read("matrix.xml"))
            ET.fromstring(zf.read("project.xml"))
    except Exception as e:
        return [tr('Datei nicht lesbar: {0}').format(e)]
    ns = "{http://graphml.graphdrawing.org/xmlns/graphml}"
    ids = []
    for n in root.iter(ns + "node"):
        nid = n.get("id")
        ids.append(nid)
        h = n.find(f"./{ns}data/hmcnode")
        if h is None or h.get("id") != nid:
            problems.append(tr("hmcnode passt nicht zu node '{0}'").format(nid))
    if any(not i for i in ids):
        problems.append(tr("Node ohne ID"))
    dup = [k for k, v in Counter(ids).items() if v > 1]
    if dup:
        problems.append(tr('{0} doppelte Node-ID(s), z. B. {1}').format(len(dup), dup[0]))
    idset = set(ids)
    n_edges = 0
    for e in root.iter(ns + "edge"):
        n_edges += 1
        if e.get("source") not in idset or e.get("target") not in idset:
            problems.append(tr('Kante {0} verweist auf fehlenden Node').format(e.get('id')))
    for g in root.iter(ns + "graph"):
        if g.get("id") == "g1":
            if g.get("parse.nodes") != str(len(ids)) or g.get("parse.edges") != str(n_edges):
                problems.append(tr("Knoten-/Kantenzahl im Graph-Kopf stimmt nicht"))
    return problems


def ascii_xml(text):
    """Nicht-ASCII-Zeichen als numerische Zeichenreferenzen (&#223;) – unabhängig von der Kodierung lesbar."""
    return text.encode("ascii", "xmlcharrefreplace").decode("ascii")


# ---------------------------------------------------------------------------
# Kern: Tabelle → .hmcx
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Harris Matrix Editor (HME): JSON-Ausgabe samt Layout (Portierung von "Auto-Layout" des HME)
# ---------------------------------------------------------------------------
HME_PHASE_COLORS = ["#ef4444", "#f97316", "#22c55e", "#3b82f6", "#a855f7", "#eab308", "#14b8a6", "#ec4899"]
HME_OBJECT_COLORS = ["#10b981", "#8b5cf6", "#06b6d4", "#f43f5e", "#84cc16", "#f59e0b", "#6366f1", "#0ea5e9"]


_PUNCT_ORDER = " _-,;:!?.'\"()[]{}@*/\\&#%`^+<=>|~$"


def _label_key(s):
    """Sortierschlüssel wie JavaScript localeCompare (ICU): Leerzeichen/Satzzeichen < Ziffern < Buchstaben,
    Groß/Klein erst bei Gleichstand (klein vor groß), Akzente ignoriert."""
    prim = []
    for ch in s:
        for c in unicodedata.normalize("NFD", ch):
            if unicodedata.combining(c):
                continue
            if c in _PUNCT_ORDER:
                prim.append((0, _PUNCT_ORDER.index(c), ""))
            elif c.isdigit():
                prim.append((1, int(c) if c.isascii() else 0, ""))
            elif c.isalpha():
                prim.append((2, 0, c.casefold()))
            else:
                prim.append((0, len(_PUNCT_ORDER), c))
    tert = [(1 if ch.isupper() else 0) for ch in s]
    return prim, tert


# ----------------------------------------------------------------------------------------------------------------
# Third-party notice: HME - Harris Matrix Editor
#
# The function hme_layout() below re-implements the "Auto-Layout" procedure of the Harris Matrix Editor (HME,
# https://github.com/oeai-dac/HME-HarrisMatrixEditor), so that the positions written to the .json file are the
# ones HME itself calculates. HME is licensed under the MIT License:
#
#   MIT License
#
#   Copyright (c) 2026 oeai-dac
#
#   Permission is hereby granted, free of charge, to any person obtaining a copy
#   of this software and associated documentation files (the "Software"), to deal
#   in the Software without restriction, including without limitation the rights
#   to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
#   copies of the Software, and to permit persons to whom the Software is
#   furnished to do so, subject to the following conditions:
#
#   The above copyright notice and this permission notice shall be included in all
#   copies or substantial portions of the Software.
#
#   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#   IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#   FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#   AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#   LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
#   OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
#   SOFTWARE.
# ----------------------------------------------------------------------------------------------------------------
def hme_layout(nodes, edges, phases, objects, horizontal_gap=20, vertical_gap=50, phase_gap=70):
    """Berechnet x/y-Positionen genau nach dem Verfahren 'Auto-Layout' des HME (nodes: id, label, phase;
    edges: source (oben) / target (unten); phases: Liste in Reihenfolge von oben nach unten; objects: id, nodeIds).
    Gibt {node_id: (x, y)} zurück."""
    node_h, start_y = 28, 50
    object_gap = horizontal_gap * 3
    label = {n["id"]: n["label"] for n in nodes}
    parents, children = {}, {}
    for e in edges:
        children.setdefault(e["source"], []).append(e["target"])
        parents.setdefault(e["target"], []).append(e["source"])

    def width(nid):
        return max(50, len(re.sub(r"^SU\s*", "", label[nid], flags=re.I)) * 9 + 30)

    node_obj = {}
    for o in objects:
        for nid in o["nodeIds"]:
            node_obj.setdefault(nid, []).append(o["id"])
    obj_by_id = {o["id"]: o for o in objects}

    # 1. Ränge innerhalb jeder Phase (topologisch)
    phase_ids = [p["id"] for p in phases]
    phase_ranks = {}
    for pid in phase_ids:
        ids = [n["id"] for n in nodes if n["phase"] == pid]
        if not ids:
            continue
        idset = set(ids)
        indeg = {i: 0 for i in ids}
        adj = {i: [] for i in ids}
        for e in edges:
            if e["source"] in idset and e["target"] in idset:
                indeg[e["target"]] += 1
                adj[e["source"]].append(e["target"])
        rank = {i: 0 for i in ids}
        queue = [i for i in ids if indeg[i] == 0]
        qi = 0
        while qi < len(queue):
            cur = queue[qi]
            qi += 1
            for t in adj[cur]:
                rank[t] = max(rank[t], rank[cur] + 1)
                indeg[t] -= 1
                if indeg[t] == 0:
                    queue.append(t)
        phase_ranks[pid] = rank
    known = set(phase_ids)
    none_ids = [n["id"] for n in nodes if not n["phase"] or n["phase"] not in known]
    if none_ids:
        phase_ranks["__none__"] = {i: 0 for i in none_ids}
    all_ids = phase_ids + (["__none__"] if "__none__" in phase_ranks else [])

    # 2. Spalten pro Phase
    col_data = {}
    for pid in all_ids:
        ranks = phase_ranks.get(pid)
        if not ranks:
            continue
        max_rank = max([0] + list(ranks.values()))
        obj_ids, loose = [], []
        for nid in ranks:
            ol = node_obj.get(nid)
            if ol:
                if ol[0] not in obj_ids:
                    obj_ids.append(ol[0])
            else:
                loose.append(nid)
        loose.sort(key=lambda i: _label_key(label[i]))

        def obj_min(oid):
            ls = sorted(label[i] for i in obj_by_id[oid]["nodeIds"] if i in label)
            return ls[0] if ls else ""
        obj_ids.sort(key=lambda oid: _label_key(obj_min(oid)))
        columns = []
        for oid in obj_ids:
            by_rank = {}
            for nid in obj_by_id[oid]["nodeIds"]:
                if nid in ranks:
                    by_rank.setdefault(ranks[nid], []).append(nid)
            for lst in by_rank.values():
                lst.sort(key=lambda i: _label_key(label[i]))
            max_w = 0
            for r in range(max_rank + 1):
                lst = by_rank.get(r, [])
                rw = sum(width(i) for i in lst) + horizontal_gap * max(0, len(lst) - 1)
                max_w = max(max_w, rw)
            columns.append({"type": "object", "by_rank": by_rank, "max_w": max_w})
        for nid in loose:
            columns.append({"type": "loose", "by_rank": {ranks[nid]: [nid]}, "max_w": width(nid)})
        col_data[pid] = {"columns": columns, "max_rank": max_rank}

    # 3. Positionen aus gegebener Spaltenreihenfolge
    def compute(orders):
        pos = {}
        cur_y = start_y
        for idx, pid in enumerate(all_ids):
            data = col_data.get(pid)
            if not data:
                continue
            if idx > 0:
                cur_y += phase_gap
            cols = orders.get(pid)
            if not cols:
                continue

            def gap(i):
                nxt = cols[i + 1]
                return object_gap if (cols[i]["type"] == "object" or nxt["type"] == "object") else horizontal_gap
            total = sum(c["max_w"] for c in cols) + sum(gap(i) for i in range(len(cols) - 1))
            x = max(50, 400 - total / 2)
            xs = []
            for i, c in enumerate(cols):
                xs.append(x)
                x += c["max_w"] + (gap(i) if i < len(cols) - 1 else 0)
            y0 = cur_y
            for r in range(data["max_rank"] + 1):
                rank_y = y0 + r * (node_h + vertical_gap)
                for ci, c in enumerate(cols):
                    lst = c["by_rank"].get(r, [])
                    rw = sum(width(i) for i in lst) + horizontal_gap * max(0, len(lst) - 1)
                    nx = xs[ci] + (c["max_w"] - rw) / 2
                    for nid in lst:
                        pos[nid] = (nx, rank_y)
                        nx += width(nid) + horizontal_gap
                cur_y = rank_y + node_h
        return pos

    # 4. Lange Kanten (über 2+ Phasen)
    pidx = {n["id"]: (phase_ids.index(n["phase"]) if n["phase"] in known else -1) for n in nodes}
    long_nodes = set()
    for e in edges:
        a, b = pidx.get(e["source"], -1), pidx.get(e["target"], -1)
        if a >= 0 and b >= 0 and abs(b - a) >= 2:
            long_nodes.add(e["source"])
            long_nodes.add(e["target"])

    def col_ids(c):
        return [i for lst in c["by_rank"].values() for i in lst]

    # 5. Spalten nach Median der verbundenen Nodes sortieren
    def sort_by_median(columns, pos, direction):
        items = []
        for ci, c in enumerate(columns):
            ids = col_ids(c)
            conn = []
            for nid in ids:
                for o in (parents.get(nid, []) if direction == "down" else children.get(nid, [])):
                    if o in pos:
                        conn.append(pos[o][0] + width(o) / 2)
            median = ci * 1000
            if conn:
                conn.sort()
                mid = len(conn) // 2
                median = conn[mid] if len(conn) % 2 == 1 else (conn[mid - 1] + conn[mid]) / 2
            items.append((0 if conn else 1, median, ci, c))
        items.sort(key=lambda t: (t[0], t[1]))
        return [t[3] for t in items]

    # 6. Spalten mit langen Kanten nach außen
    def move_long_outer(columns, pos):
        if len(columns) <= 2:
            return columns
        xs_min, xs_max = float("inf"), float("-inf")
        for c in columns:
            for i in col_ids(c):
                if i in pos:
                    xs_min = min(xs_min, pos[i][0])
                    xs_max = max(xs_max, pos[i][0] + width(i))
        center = (xs_min + xs_max) / 2
        left, right, normal = [], [], []
        for c in columns:
            ids = col_ids(c)
            if any(i in long_nodes for i in ids):
                conn = []
                for nid in ids:
                    for o in children.get(nid, []) + parents.get(nid, []):
                        if o in pos:
                            conn.append(pos[o][0] + width(o) / 2)
                avg = sum(conn) / len(conn) if conn else center
                (left if avg < center else right).append((avg, c))
            else:
                normal.append(c)
        if not left and not right:
            return columns
        left.sort(key=lambda t: t[0])
        right.sort(key=lambda t: -t[0])
        return [t[1] for t in left] + normal + [t[1] for t in right]

    # 7. Mehrere Durchläufe (von oben nach unten und zurück)
    orders = {pid: list(col_data[pid]["columns"]) for pid in all_ids if pid in col_data}
    for _ in range(4):
        pos = compute(orders)
        for pid in all_ids:
            if pid in col_data:
                orders[pid] = sort_by_median(orders[pid], pos, "down")
                pos = compute(orders)
        for pid in reversed(all_ids):
            if pid in col_data:
                orders[pid] = sort_by_median(orders[pid], pos, "up")
                pos = compute(orders)
    final = compute(orders)
    for pid in all_ids:
        if pid in orders:
            orders[pid] = move_long_outer(orders[pid], final)
    return compute(orders)


def hme_depths(ids, above_edges):
    """Stratigraphische Tiefe jeder Unit (längster Weg von oben) und eine topologische Reihenfolge.
    above_edges: (oben, unten). Erwartet zyklenfreie Daten (Zyklen wurden vorher als Fehler gemeldet)."""
    parents = {i: [] for i in ids}
    children = {i: [] for i in ids}
    for s, t in above_edges:
        if s in parents and t in parents:
            parents[t].append(s)
            children[s].append(t)
    indeg = {i: len(parents[i]) for i in ids}
    depth = {i: 0 for i in ids}
    topo = [i for i in sorted(ids) if indeg[i] == 0]
    k = 0
    while k < len(topo):
        cur = topo[k]
        k += 1
        for t in children[cur]:
            depth[t] = max(depth[t], depth[cur] + 1)
            indeg[t] -= 1
            if indeg[t] == 0:
                topo.append(t)
    return depth, topo, parents, children


def hme_phase_order(periods, node_period, above_edges, depth=None):
    """Reihenfolge der Phasen (oben = jünger), aus den Beziehungen zwischen den Gruppen abgeleitet.
    Bei Widersprüchen (Gruppen überlappen sich) entscheidet die mittlere stratigraphische Tiefe.
    Gibt (Liste, Konflikt) zurück."""
    depth = depth or {}
    names = sorted(periods)
    mean = {}
    for p in names:
        ds = [depth[s] for s, q in node_period.items() if q == p and s in depth]
        mean[p] = sum(ds) / len(ds) if ds else 0
    out_e = {p: set() for p in names}
    indeg = {p: 0 for p in names}
    for s, t in above_edges:
        a, b = node_period.get(s), node_period.get(t)
        if a and b and a != b and b not in out_e[a]:
            out_e[a].add(b)
            indeg[b] += 1
    weight = Counter()
    for s, t in above_edges:
        a_, b_ = node_period.get(s), node_period.get(t)
        if a_ and b_ and a_ != b_:
            weight[(a_, b_)] += 1
    order, left, conflict = [], set(names), False
    while left:
        ready = sorted((p for p in left if indeg[p] == 0), key=lambda p: (mean[p], p))
        if not ready:  # Zyklus: die im Mittel höchste Gruppe zuerst
            conflict = True
            ready = [min(left, key=lambda p: (mean[p], p))]
        p = ready[0]
        order.append(p)
        left.discard(p)
        for q in out_e[p]:
            indeg[q] -= 1
    return order, conflict


def hme_assign_undated(order, node_period, depth_topo, parents, children):
    """Units ohne Periode werden stratigraphisch eingeordnet und zu möglichst wenigen Gruppen
    ('unknown-1', 'unknown-2', ...) zusammengefasst, die zwischen zwei datierten Perioden liegen.
    Gibt ({unit: slot}, [belegte Slots aufsteigend], Zahl widersprüchlicher Units) zurück;
    slot s = direkt nach der Periode mit Index s (-1 = vor der ersten)."""
    n_per = len(order)
    idx = {p: i for i, p in enumerate(order)}
    lo, hi = {}, {}
    for v in depth_topo:  # größter Periodenindex aller Vorgänger (oben liegende)
        m = -1
        for p in parents[v]:
            m = max(m, lo[p], idx.get(node_period.get(p), -1))
        lo[v] = m
    for v in reversed(depth_topo):  # kleinster Periodenindex aller Nachfolger (unten liegende)
        m = n_per
        for c in children[v]:
            m = min(m, hi[c], idx.get(node_period.get(c), n_per))
        hi[v] = m
    undated = [v for v in depth_topo if not node_period.get(v)]
    iv, contradict = {}, 0
    for v in undated:
        lft, rgt = lo[v], hi[v] - 1
        if lft > rgt:  # Widerspruch zwischen oben/unten liegenden Perioden
            contradict += 1
            rgt = lft
        iv[v] = (lft, rgt)
    chosen, last = [], -2
    for lft, rgt in sorted(set(iv.values()), key=lambda t: (t[1], t[0])):
        if lft > last:  # kleinste Zahl an Punkten, die jedes Intervall trifft
            chosen.append(rgt)
            last = rgt
    slot = {v: min(c for c in chosen if c >= iv[v][0]) for v in undated}
    return slot, sorted(set(slot.values())), contradict


def build_hme(se_nodes, periods, phases, above_edges, undated_groups=True, type_map=None):
    """HME-Datenstruktur aus den fertigen HMA-Daten. Übergeordnete Gruppe -> Phase, untergeordnete -> Objekt.
    Gibt (Daten, Konflikt, Info) zurück."""
    ids = {sid: str(i) for i, sid in enumerate(sorted(se_nodes, key=lambda x: (se_num(x), x)), start=1)}
    node_period = {sid: info["period_id"] for sid, info in se_nodes.items()}
    edges_pairs = [(s, t) for (s, t, _) in above_edges]
    depth, topo, parents, children = hme_depths(list(se_nodes), edges_pairs)
    order, conflict = hme_phase_order(periods, node_period, edges_pairs, depth)
    n_undated = sum(1 for v in se_nodes if not node_period.get(v))

    def plan(order_):
        """Phasenliste (oben -> unten) und Zuordnung jeder Unit zu einer Phase für eine Periodenreihenfolge."""
        plist, uphase = list(order_), dict(node_period)
        inf = {"undated": 0, "unknown_groups": 0, "contradict": 0}
        if undated_groups and n_undated:
            slot, used, contradict = hme_assign_undated(order_, node_period, topo, parents, children)
            used_names = set(order_)
            names = {}
            for k, s in enumerate(used, start=1):
                nm = f"unknown-{k}"
                while nm in used_names:
                    nm += "_"
                used_names.add(nm)
                names[s] = nm
            plist = ([names[-1]] if -1 in names else [])
            for i, p in enumerate(order_):
                plist.append(p)
                if i in names:
                    plist.append(names[i])
            for v, s in slot.items():
                uphase[v] = names[s]
            inf = {"undated": n_undated, "unknown_groups": len(used), "contradict": contradict}
        pos_ = {p: i for i, p in enumerate(plist)}
        back = sum(1 for s, t in edges_pairs
                   if uphase.get(s) in pos_ and uphase.get(t) in pos_ and pos_[uphase[s]] > pos_[uphase[t]])
        return plist, uphase, inf, back

    best = plan(order)
    if conflict and 1 < len(order) <= 7:
        # Zyklus zwischen den Gruppen: die Reihenfolge wählen, bei der am wenigsten Beziehungen "rückwärts" laufen
        from itertools import permutations
        rank = {p: i for i, p in enumerate(order)}
        best_key = (best[3], 0)
        for perm in permutations(order):
            cand = plan(perm)
            key = (cand[3], sum(abs(i - rank[p]) for i, p in enumerate(perm)))
            if key < best_key:
                best, best_key, order = cand, key, list(perm)
    phase_list, unit_phase, info, _back = best
    ph_id = {p: str(i) for i, p in enumerate(phase_list, start=1)}
    hme_phases = []
    for i, p in enumerate(phase_list):
        color = HME_PHASE_COLORS[order.index(p) % len(HME_PHASE_COLORS)] if p in periods else "#9ca3af"
        hme_phases.append({"id": ph_id[p], "name": p, "color": color})
    obj_names = sorted(phases)
    hme_objects = []
    for i, name in enumerate(obj_names):
        members = [ids[s] for s, inf in sorted(se_nodes.items(), key=lambda kv: (se_num(kv[0]), kv[0]))
                   if inf["phase_id"] == name]
        if members:
            hme_objects.append({"id": str(len(hme_objects) + 1), "name": name,
                                "color": HME_OBJECT_COLORS[i % len(HME_OBJECT_COLORS)], "nodeIds": members})
    nodes = []
    for sid, inf in sorted(se_nodes.items(), key=lambda kv: (se_num(kv[0]), kv[0])):
        desc = " – ".join(x for x in (inf["name"], (inf["description"] or "").replace("\r", "").replace("\n", " ")) if x)
        nodes.append({"id": ids[sid], "label": sid, "description": desc,
                      "type": "interface" if inf["node_type"] == "SURFACE" else (type_map or {}).get(inf.get("kat"), "deposit"),
                      "phase": ph_id.get(unit_phase.get(sid), ""), "x": 0, "y": 0})
    edges = [{"id": f"e{ids[s]}-{ids[t]}", "source": ids[s], "target": ids[t]} for (s, t, _) in above_edges]
    pos = hme_layout(nodes, edges, hme_phases, hme_objects)
    for n in nodes:
        x, y = pos.get(n["id"], (0, 0))
        n["x"], n["y"] = int(round(x)), int(round(y))
    data = {"version": "1.0", "exportDate": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z"),
            "nodes": nodes, "edges": edges, "phases": hme_phases, "objects": hme_objects}
    return data, conflict, info


def verify_hme(path):
    """Selbstprüfung der erzeugten HME-Datei (gleiche Struktur wie die Beispieldatei des HME)."""
    problems = []
    try:
        d = json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception as e:
        return [tr('JSON nicht lesbar: {0}').format(e)]
    for k in ("nodes", "edges", "phases", "objects"):
        if not isinstance(d.get(k), list):
            problems.append(tr("Liste '{0}' fehlt").format(k))
    if problems:
        return problems
    ids = [n.get("id") for n in d["nodes"]]
    if len(set(ids)) != len(ids):
        problems.append(tr("doppelte Node-IDs"))
    idset = set(ids)
    labels = [n.get("label") for n in d["nodes"]]
    if len(set(labels)) != len(labels):
        problems.append(tr("doppelte Labels"))
    for n in d["nodes"]:
        if n.get("type") not in ("layer", "deposit", "fill", "structure", "interface"):
            problems.append(tr('unbekannter Typ bei {0}').format(n.get('label')))
        if not isinstance(n.get("x"), (int, float)) or not isinstance(n.get("y"), (int, float)):
            problems.append(tr('Koordinaten fehlen bei {0}').format(n.get('label')))
    eids = set()
    for e in d["edges"]:
        if e.get("source") not in idset or e.get("target") not in idset or e["source"] == e["target"]:
            problems.append(tr('ungültige Kante {0}').format(e.get('id')))
        if e.get("id") in eids:
            problems.append(tr('doppelte Kanten-ID {0}').format(e.get('id')))
        eids.add(e.get("id"))
    pids = {p.get("id") for p in d["phases"]}
    for n in d["nodes"]:
        if n.get("phase") and n["phase"] not in pids:
            problems.append(tr('Phase von {0} unbekannt').format(n.get('label')))
    for o in d["objects"]:
        if any(i not in idset for i in o.get("nodeIds", [])):
            problems.append(tr('Objekt {0} verweist auf unbekannte Nodes').format(o.get('name')))
    return problems[:10]


# ---------------------------------------------------------------------------
# Voransicht: eigenes Layout (Perioden-Blöcke bzw. Schichten nach stratigraphischer Tiefe) und Szene für den Canvas
# ---------------------------------------------------------------------------
PV_NODE_COLORS = {"deposit": "#9ec5f0", "interface": "#8fd19e", "layer": "#f4d27a", "fill": "#d4b3e8",
                  "structure": "#f0a98a", "tug": "#f2b8b8"}
PV_NODE_NAMES = {"deposit": "Deposit", "interface": "Interface", "layer": "Layer", "fill": "Fill",
                 "structure": "Structure", "tug": "T / G / U"}
PV_MAX_LABEL = 24


def preview_graph(model):
    """Vereinheitlichte Sicht für die Voransicht aus dem Modell von convert(model_only=True):
    {"nodes": {id: {label, name, desc, type, outer, inner, block}}, "above": [(oben, unten)], "contemp": [(a, b)],
    "block_order": [Block, ...] (Perioden von oben nach unten), "block_label": {Block: Anzeigename}}.
    Ein Block ist eine Periode (HME: Phase); Units ohne Periode werden wie beim HME-Export zu 'unknown-N' gebündelt."""
    nodes, above, contemp = {}, [], []
    # Reihenfolge der Perioden und Zuordnung undatierter Units wie im HME-Export
    try:
        real_edges = [e for e in model["above_edges"] if e[0] in model["se_nodes"] and e[1] in model["se_nodes"]]
        data, _conf, _info = build_hme(model["se_nodes"], model["periods"], model["phases"], real_edges,
                                       undated_groups=(model["hme_undated"] if model["out_format"] == "hme" else True),
                                       type_map=model["hme_types"])
    except Exception:
        data = None
    block_of, order, label_of = {}, [], {}
    if not data and model["out_format"] != "hme":   # Notfall: Perioden alphabetisch, damit die Übergruppen erhalten bleiben
        order = sorted(model["periods"])
        label_of = {p: p for p in order}
    if data:
        ph = {p["id"]: p["name"] for p in data["phases"]}
        order = [p["name"] for p in data["phases"]]
        for n in data["nodes"]:
            block_of[n["label"]] = ph.get(n["phase"], "")
        for nm in order:
            if nm in model["periods"] or model["out_format"] == "hme":
                label_of[nm] = nm
            else:
                label_of[nm] = tr("ohne Periode ") + nm.rsplit("-", 1)[-1].rstrip("_")
    if model["out_format"] == "hme":
        lab = {n["id"]: n["label"] for n in data["nodes"]}
        obj_of = {}
        for o in data["objects"]:
            for nid in o["nodeIds"]:
                obj_of[nid] = o["name"]
        for n in data["nodes"]:
            nodes[n["label"]] = {"label": n["label"], "name": "", "desc": n["description"], "type": n["type"],
                                 "outer": ph.get(n["phase"], ""), "inner": obj_of.get(n["id"], ""),
                                 "block": ph.get(n["phase"], ""),
                                 "ident": model.get("idents", {}).get(n["label"], [])}
        above = [(lab[e["source"]], lab[e["target"]]) for e in data["edges"]]
        return {"nodes": nodes, "above": above, "contemp": [], "block_order": order, "block_label": label_of,
                "format": "hme"}
    for sid, inf in model["se_nodes"].items():
        nodes[sid] = {"label": sid, "name": inf["name"] or "", "desc": inf["description"] or "",
                      "type": "interface" if inf["node_type"] == "SURFACE" else "deposit",
                      "outer": inf["period_id"] or "", "inner": inf["phase_id"] or "",
                      "block": inf["period_id"] or block_of.get(sid, ""),
                      "ident": model.get("idents", {}).get(sid, [])}
    for k, (_ttype, tname, _layer) in model["tug_nodes"].items():
        nodes[k] = {"label": k, "name": tname, "desc": "", "type": "tug", "outer": "", "inner": "", "block": "@" + k}
    above = [(s, t) for (s, t, _) in model["above_edges"] if s in nodes and t in nodes]
    seen = set()
    for a, b, _ in model["zeitgleich_edges"]:
        key = tuple(sorted((a, b)))
        if a in nodes and b in nodes and key not in seen:
            seen.add(key)
            contemp.append((a, b))
    full_order = (["@T"] if "T" in model["tug_nodes"] else []) + order + \
                 [f"@{k}" for k in "UG" if k in model["tug_nodes"]]
    return {"nodes": nodes, "above": above, "contemp": contemp, "block_order": full_order if order or
            model["tug_nodes"] else [], "block_label": label_of, "format": "hmc"}


def pv_list(items, limit=12):
    items = list(items)
    if not items:
        return "–"
    s = ", ".join(items[:limit])
    return s + (f", … (+{len(items) - limit})" if len(items) > limit else "")


def pv_hover_text(n):
    """Stratigraphische Kurzinfo einer Unit (Hover in der Voransicht)."""
    return (tr('{0}\nliegt_unter: {1}\nliegt_ueber: {2}\nzeitgleich_mit: {3}\nident_mit: {4}').format(n['full'], pv_list(n['under']), pv_list(n['over']), pv_list(n['contemp']), pv_list(n['ident'])))


def pv_common_prefix(labels):
    """Gemeinsamer Anfang der Beschriftungen (bis einschließlich eines Trennzeichens), wenn ihn die meisten haben."""
    labels = [l for l in labels if l]
    if len(labels) < 3:
        return ""
    cnt = Counter()
    for l in labels:
        for m in re.finditer(r"[_\-. /]", l):
            cnt[l[:m.end()]] += 1
    cands = [p for p, c in cnt.items() if c >= 0.6 * len(labels)]
    return max(cands, key=len) if cands else ""


def pv_fit_label(label, width_px, char_px, prefix=""):
    """Kürzt eine Beschriftung, damit sie in die Breite passt: erst den gemeinsamen Anfang (prefix) durch '…'
    ersetzen, dann hinten mit '…' kürzen."""
    maxc = int((width_px - 8) / char_px)
    if maxc < 2:
        return ""
    if len(label) <= maxc:
        return label
    if prefix and label.startswith(prefix) and len(label) > len(prefix):
        label = "…" + label[len(prefix):]
        if len(label) <= maxc:
            return label
    return label[:max(1, maxc - 1)] + "…"


PV_ROW_BLOCK_GAP = 76   # Mindestabstand zwischen zwei Perioden-Blöcken
PV_MODES = {"auto": tr("Automatisch"), "blocks": tr("Perioden-Blöcke"), "depth": tr("Strikt nach Tiefe")}


class _PvBlock:
    """Baustein im Gruppenbaum: Unit (kind 'node') oder Gruppe (kind 'group', level 0 = übergeordnet, 1 = unter)."""
    __slots__ = ("kind", "level", "name", "id", "children", "lo", "hi", "ids", "lanes", "w", "left", "pad_x")

    def __init__(self, kind, level=0, name="", nid=None):
        self.kind, self.level, self.name, self.id = kind, level, name, nid
        self.children, self.lo, self.hi, self.ids = [], 0, 0, []
        self.lanes, self.w, self.left = [], 0.0, 0.0
        self.pad_x = {0: 14, 1: 7}.get(level, 0) if kind == "group" else 0


def _pv_make_rep(ids, pairs, edges=()):
    """Zeitgleiche Units zu Klassen zusammenfassen (Union-Find). Eine zeitgleich-Beziehung wird nur übernommen, wenn
    sie der Stratigraphie nicht widerspricht, d. h. keine der beiden Klassen liegt direkt oder über Dritte über der
    anderen; solche Paare werden übersprungen. Gibt ({unit: Klasse}, Zahl übersprungener Paare) zurück."""
    uf = {i: i for i in ids}

    def find(x):
        while uf[x] != x:
            uf[x] = uf[uf[x]]
            x = uf[x]
        return x
    succ = {i: set() for i in ids}  # Klasse -> Nachfolger (liegt darunter); Einträge werden über find() aufgelöst
    for s, t in edges:
        if s in uf and t in uf and s != t:
            succ[s].add(t)

    def reaches(a, b):
        stack, seen = [a], {a}
        while stack:
            for y in succ[stack.pop()]:
                y = find(y)
                if y == b:
                    return True
                if y not in seen:
                    seen.add(y)
                    stack.append(y)
        return False
    skipped = 0
    todo = sorted({tuple(sorted(p, key=lambda x: (se_num(x), x))) for p in pairs if p[0] in uf and p[1] in uf},
                  key=lambda p: (se_num(p[0]), p[0], se_num(p[1]), p[1]))
    for a, b in todo:
        ra, rb = find(a), find(b)
        if ra == rb:
            continue
        if reaches(ra, rb) or reaches(rb, ra):
            skipped += 1
            continue
        uf[rb] = ra
        succ[ra] |= succ[rb]
        succ[rb] = set()
    return {i: find(i) for i in ids}, skipped


def _pv_rank(ids, edges, contemp):
    """Zeile jeder Unit = längster Weg von oben; zeitgleiche Units auf gleicher Höhe, Quellen dicht an ihren
    Nachfolgern. Zyklen werden toleriert: Kanten, die gegen die gefundene Reihenfolge laufen, zählen nicht
    (werden später als Rückwärtskanten gezeichnet). Gibt ({unit: zeile ab 0}, Zahl Rückwärtskanten, Zahl
    zeitgleich-Paare, die der Stratigraphie widersprechen und deshalb nicht auf gleicher Höhe liegen) zurück."""
    rep, n_skipped = _pv_make_rep(ids, contemp, edges)  # widersprüchliche zeitgleich-Paare werden nicht angeglichen
    classes = sorted(set(rep.values()), key=lambda x: (se_num(x), x))
    succ = {c: set() for c in classes}
    pred = {c: set() for c in classes}
    for s, t in edges:
        a, b = rep[s], rep[t]
        if a != b:
            succ[a].add(b)
            pred[b].add(a)
    deg = {c: len(pred[c]) for c in classes}
    ready = [c for c in classes if deg[c] == 0]
    queued, order, qi = set(ready), [], 0
    while len(order) < len(classes):
        if qi >= len(ready):  # Zyklus: die Klasse mit den wenigsten offenen Vorgängern zuerst
            c = min((c for c in classes if c not in queued), key=lambda c: (deg[c], se_num(c), c))
            ready.append(c)
            queued.add(c)
        c = ready[qi]
        qi += 1
        order.append(c)
        for d in succ[c]:
            deg[d] -= 1
            if deg[d] == 0 and d not in queued:
                ready.append(d)
                queued.add(d)
    pos = {c: k for k, c in enumerate(order)}
    depth = {c: 0 for c in classes}
    n_back = 0
    for c in order:
        for d in succ[c]:
            if pos[d] > pos[c]:
                depth[d] = max(depth[d], depth[c] + 1)
            else:
                n_back += 1
    for c in reversed(order):
        fw = [d for d in succ[c] if pos[d] > pos[c]]
        if fw and not any(pos[p] < pos[c] for p in pred[c]):
            depth[c] = min(depth[d] for d in fw) - 1
    out = {i: depth[rep[i]] for i in ids}
    lo = min(out.values())
    return {i: r - lo for i, r in out.items()}, n_back, n_skipped


def layout_preview(g, mode="auto"):
    """Berechnet die Szene. mode 'blocks': Perioden (Phasen) liegen als Blöcke übereinander (wie im HME bzw. HMC),
    innerhalb eines Blocks gilt die lokale stratigraphische Tiefe. mode 'depth': alle Units strikt nach
    stratigraphischer Tiefe. Gruppen sind überschneidungsfreie Kästen, Kanten verlaufen nur waagerecht/senkrecht."""
    from bisect import bisect_right
    nodes, above, contemp = g["nodes"], g["above"], g["contemp"]
    ids = sorted(nodes, key=lambda x: (se_num(x), x))
    if not ids:
        return {"w": 200, "h": 100, "groups": [], "nodes": [], "edges": [], "types": [], "prefix": "",
                "mode": "depth", "n_back": 0}
    if mode == "auto" or mode not in ("blocks", "depth"):
        # HME stapelt Phasen als Blöcke übereinander; die HMC-Voransicht bleibt standardmäßig strikt nach Tiefe
        mode = "blocks" if g.get("block_order") and g.get("format") == "hme" else "depth"
    if mode == "blocks" and not g.get("block_order"):
        mode = "depth"

    # --- Zeilen (Ränge) ---
    block_starts = set()
    n_skip = 0
    if mode == "blocks":
        keys = list(g["block_order"])
        for i in ids:
            k = nodes[i].get("block")
            if k not in keys:
                keys.append(k)
        used = [k for k in keys if any(nodes[i].get("block") == k for i in ids)]
        rank, off = {}, 0
        for bk in used:
            bids = [i for i in ids if nodes[i].get("block") == bk]
            bset = set(bids)
            rk, _nb, nsk = _pv_rank(bids, [(s, t) for s, t in above if s in bset and t in bset],
                                    [(a, b) for a, b in contemp if a in bset and b in bset])
            n_skip += nsk
            if off:
                block_starts.add(off)
            for i in bids:
                rank[i] = off + rk[i]
            off += max(rk.values()) + 1
        n_skip += sum(1 for a, b in contemp if nodes[a].get("block") != nodes[b].get("block"))  # über Blöcke hinweg
        outer_of = {i: g.get("block_label", {}).get(nodes[i].get("block"), "") for i in ids}
    else:
        rank, _nb, n_skip = _pv_rank(ids, above, contemp)
        outer_of = {i: nodes[i]["outer"] for i in ids}
    n_layers = max(rank.values()) + 1
    n_back = sum(1 for s, t in above if rank[s] >= rank[t])

    width = {i: min(150, max(44, 6.6 * min(len(nodes[i]["label"]), 22) + 16)) for i in ids}
    parents = {i: [] for i in ids}
    children_of = {i: [] for i in ids}
    for s, t in above:
        parents[t].append(s)
        children_of[s].append(t)
    neigh = {i: parents[i] + children_of[i] for i in ids}

    # --- Gruppenbaum ---
    root = _PvBlock("group", level=-1)
    outer_b, inner_b = {}, {}
    for i in ids:
        o, n = outer_of[i], nodes[i]["inner"]
        leaf = _PvBlock("node", nid=i)
        leaf.lo = leaf.hi = rank[i]
        leaf.ids = [i]
        leaf.w = width[i]
        if o and o not in outer_b:
            outer_b[o] = _PvBlock("group", 0, o)
            root.children.append(outer_b[o])
        if o and n:
            if (o, n) not in inner_b:
                inner_b[(o, n)] = _PvBlock("group", 1, n)
                outer_b[o].children.append(inner_b[(o, n)])
            inner_b[(o, n)].children.append(leaf)
        elif o:
            outer_b[o].children.append(leaf)
        elif n:
            if (None, n) not in inner_b:
                inner_b[(None, n)] = _PvBlock("group", 1, n)
                root.children.append(inner_b[(None, n)])
            inner_b[(None, n)].children.append(leaf)
        else:
            root.children.append(leaf)
    all_groups = []

    def measure(b):
        if b.kind == "node":
            return
        for c in b.children:
            measure(c)
        b.lo = min(c.lo for c in b.children)
        b.hi = max(c.hi for c in b.children)
        b.ids = [i for c in b.children for i in c.ids]
        if b.level >= 0:
            all_groups.append(b)
    measure(root)

    def assign_lanes(b):
        """Kinder in Spuren: Kinder in derselben Spur haben getrennte Zeilenbereiche (also nie nebeneinander)."""
        if b.kind == "node":
            return
        lanes, last_hi, where = [], [], {}
        for c in sorted(b.children, key=lambda c: (c.lo, -c.hi, c.name or "", c.id or "")):
            cand = [k for k in range(len(lanes)) if last_hi[k] < c.lo]
            pick = None
            if cand:
                hits = {k: 0 for k in cand}
                for i in c.ids:
                    for p in parents[i]:
                        if where.get(p) in hits:
                            hits[where[p]] += 1
                pick = max(cand, key=lambda k: (hits[k], last_hi[k], -k))
            if pick is None:
                lanes.append([])
                last_hi.append(-1)
                pick = len(lanes) - 1
            lanes[pick].append(c)
            last_hi[pick] = c.hi
            for i in c.ids:
                where[i] = pick
        b.lanes = lanes
        for c in b.children:
            assign_lanes(c)
    assign_lanes(root)

    gap = 12

    def size(b):
        if b.kind == "node":
            return b.w
        lane_w = [max(size(c) for c in lane) for lane in b.lanes]
        b.w = 2 * b.pad_x + sum(lane_w) + gap * (len(lane_w) - 1)
        return b.w
    size(root)

    xc_of = {}
    block_off = {}

    def place(b, left):
        b.left = left
        if b.kind == "node":
            xc_of[b.id] = left + b.w / 2
            return
        x0 = left + b.pad_x
        for lane in b.lanes:
            lw = max(c.w for c in lane)
            for c in lane:
                cl = x0 + (lw - c.w) / 2
                if b is root and mode == "blocks":  # Blöcke zu ihren Nachbarn hin verschieben (im Rahmen der Breite)
                    cl = min(max(cl + block_off.get(id(c), 0.0), x0), x0 + lw - c.w)
                place(c, cl)
            x0 += lw + gap

    def reorder(b):
        """Spuren nach dem Schwerpunkt ihrer Nachbarn (außerhalb der Spur) sortieren."""
        changed = False
        if b.kind == "group" and len(b.lanes) > 1:
            keys = []
            for k, lane in enumerate(b.lanes):
                mine = {i for c in lane for i in c.ids}
                pos = [xc_of[i] for i in mine]
                cur = sum(pos) / len(pos)
                nb = [xc_of[n] for i in mine for n in neigh[i] if n not in mine]
                keys.append((sum(nb) / len(nb) if nb else cur, k))
            order = sorted(range(len(keys)), key=lambda k: keys[k])
            if order != list(range(len(order))):
                b.lanes = [b.lanes[k] for k in order]
                changed = True
        for c in b.children:
            changed = reorder(c) or changed
        return changed

    place(root, 0.0)
    for _ in range(10):
        if not reorder(root):
            break
        place(root, 0.0)
    if mode == "blocks" and len(root.children) > 1:
        for _ in range(8):
            for c in root.children:
                mine = set(c.ids)
                d = [xc_of[n] - xc_of[i] for i in c.ids for n in neigh[i] if n not in mine]
                if d:
                    block_off[id(c)] = block_off.get(id(c), 0.0) + 0.5 * (sum(d) / len(d))
            place(root, 0.0)
            reorder(root)
            place(root, 0.0)

    margin = 30
    shift = margin + 6
    x = {i: xc_of[i] + shift for i in ids}

    # --- Kanten (nur waagerecht/senkrecht): Korridore für lange Kanten ---
    layer_iv = [[] for _ in range(n_layers)]
    for i in ids:
        layer_iv[rank[i]].append((x[i] - width[i] / 2, x[i] + width[i] / 2))
    lefts, rights = [], []
    for r in range(n_layers):
        layer_iv[r].sort()
        lefts.append([a for a, _ in layer_iv[r]])
        rights.append([b for _, b in layer_iv[r]])
    PAD = 5.0

    def blocked(xv, r0, r1):
        for r in range(r0, r1 + 1):
            k = bisect_right(lefts[r], xv + PAD) - 1
            if k >= 0 and rights[r][k] >= xv - PAD:
                return True
        return False

    corr = {}

    def conflict(xv, ru, rl, s, t):
        b0 = int(xv // 6)
        for bk in (b0 - 1, b0, b0 + 1):
            for (cx, a, b, cs, ct) in corr.get(bk, ()):
                if abs(cx - xv) < 6 and a < rl and ru < b and cs != s and ct != t:
                    return True
        return False

    # Geometrie je Kante in "oben -> unten"-Richtung (u = obere, l = untere Unit); Rückwärtskanten (rank[s] > rank[t])
    # werden später umgedreht, Kanten zwischen gleich hohen Units laufen unten herum.
    geo = []
    for (s, t) in above:
        if rank[s] < rank[t]:
            geo.append((s, t, False))
        elif rank[s] > rank[t]:
            geo.append((t, s, True))
        else:
            geo.append((s, t, None))
    xcor = {}
    for k in sorted(range(len(above)), key=lambda k: -abs(rank[above[k][1]] - rank[above[k][0]])):
        u, lo_, rev = geo[k]
        if rev is None:
            continue
        ru, rl = rank[u], rank[lo_]
        if rl - ru <= 1:
            continue
        xs, xt = x[u], x[lo_]
        pick = None
        tried = [xs, xt]
        for d in range(6, 600, 6):
            tried += [xs - d, xs + d, xt - d, xt + d]
        order = sorted(range(len(tried)), key=lambda q: (abs(xs - tried[q]) + abs(xt - tried[q]), q))
        for q in order:
            cx = tried[q]
            if not blocked(cx, ru + 1, rl - 1) and not conflict(cx, ru, rl, u, lo_):
                pick = cx
                break
        if pick is None:
            pick = xs
        xcor[k] = pick
        corr.setdefault(int(pick // 6), []).append((pick, ru, rl, u, lo_))

    segs = {}
    for k, (u, lo_, rev) in enumerate(geo):
        ru, rl = rank[u], rank[lo_]
        xs, xt = x[u], x[lo_]
        if rev is None:                      # gleiche Zeile: unten herum
            segs.setdefault(ru, []).append([min(xs, xt), max(xs, xt), {("s", u), ("s", lo_)}, k, "S", 0])
        elif rl - ru <= 1:
            if xs != xt:
                segs.setdefault(ru, []).append([min(xs, xt), max(xs, xt), {("s", u), ("t", lo_)}, k, "M", 0])
        else:
            xc = xcor[k]
            if xc != xs:
                segs.setdefault(ru, []).append([min(xs, xc), max(xs, xc), {("s", u)}, k, "A", 0])
            if xc != xt:
                segs.setdefault(rl - 1, []).append([min(xc, xt), max(xc, xt), {("t", lo_)}, k, "B", 0])
    gap_tracks = [0] * max(1, n_layers)
    track_of = {}
    for g_, lst in segs.items():
        lst.sort(key=lambda sg: (sg[0], sg[1]))
        tracks = []
        for sg in lst:
            placed = False
            for ti, tr in enumerate(tracks):
                if all(sg[0] > o[1] + 4 or sg[1] < o[0] - 4 or (sg[2] & o[2]) for o in tr):
                    tr.append(sg)
                    track_of[(sg[3], sg[4])] = ti
                    placed = True
                    break
            if not placed:
                tracks.append([sg])
                track_of[(sg[3], sg[4])] = len(tracks) - 1
        gap_tracks[g_] = len(tracks)

    node_h, step, lead = 22, 6, 11
    gap_h = []
    for r in range(max(1, n_layers)):
        h_ = max(44, 2 * lead + step * max(0, gap_tracks[r] - 1))
        if (r + 1) in block_starts:
            h_ = max(h_, PV_ROW_BLOCK_GAP)
        gap_h.append(h_)
    ytop = [margin + 22]
    for r in range(n_layers - 1):
        ytop.append(ytop[-1] + node_h + gap_h[r])

    def track_y(gap_idx, ti):
        return ytop[gap_idx] + node_h + lead + step * ti

    edges = []
    for k, (s, t) in enumerate(above):
        u, lo_, rev = geo[k]
        ru, rl = rank[u], rank[lo_]
        xs, xt = x[u], x[lo_]
        pts = [(xs, ytop[ru] + node_h)]
        if rev is None:
            ya = track_y(ru, track_of[(k, "S")])
            pts += [(xs, ya), (xt, ya), (xt, ytop[rl] + node_h)]
        else:
            if rl - ru <= 1:
                if xs != xt:
                    ya = track_y(ru, track_of[(k, "M")])
                    pts += [(xs, ya), (xt, ya)]
            else:
                xc = xcor[k]
                if xc != xs:
                    ya = track_y(ru, track_of[(k, "A")])
                    pts += [(xs, ya), (xc, ya)]
                if xc != xt:
                    yb = track_y(rl - 1, track_of[(k, "B")])
                    pts += [(xc, yb), (xt, yb)]
            pts.append((xt, ytop[rl]))
        clean = [pts[0]]
        for p in pts[1:]:
            if p != clean[-1]:
                clean.append(p)
        if rev:
            clean.reverse()
        edges.append({"s": s, "t": t, "pts": clean, "kind": "above", "back": rev is not None})
    for a, b in contemp:
        if x[a] > x[b]:
            a, b = b, a
        if rank[a] == rank[b]:
            yy = ytop[rank[a]] + node_h / 2
            pts = [(x[a] + width[a] / 2, yy), (x[b] - width[b] / 2, yy)]
        else:
            pts = [(x[a], ytop[rank[a]] + node_h / 2), (x[b], ytop[rank[b]] + node_h / 2)]
        edges.append({"s": a, "t": b, "pts": pts, "kind": "contemp", "back": False, "clash": rank[a] != rank[b]})

    scene_nodes = []
    for i in ids:
        nd = nodes[i]
        scene_nodes.append({"id": i, "x": x[i] - width[i] / 2, "y": ytop[rank[i]], "w": width[i], "h": node_h,
                            "label": nd["label"], "type": nd["type"], "full": nd["label"], "name": nd["name"],
                            "desc": nd["desc"], "outer": nd["outer"], "inner": nd["inner"],
                            "over": sorted(children_of[i], key=lambda v: (se_num(v), v)),
                            "under": sorted(parents[i], key=lambda v: (se_num(v), v)),
                            "contemp": sorted({b for a, b in contemp if a == i} | {a for a, b in contemp if b == i},
                                              key=lambda v: (se_num(v), v)),
                            "ident": list(nd.get("ident") or [])})
    groups = []
    for b in all_groups:
        ptop, pbot = (20, 10) if b.level == 0 else (15, 6)
        y0, y1 = ytop[b.lo] - ptop, ytop[b.hi] + node_h + pbot
        groups.append({"x": b.left + shift, "y": y0, "w": b.w, "h": y1 - y0, "label": b.name, "level": b.level})
    if mode == "blocks":  # Perioden als durchgehende Bänder gleicher Breite
        l0 = [g_ for g_ in groups if g_["level"] == 0]
        if l0:
            x0_ = min(g_["x"] for g_ in l0)
            x1_ = max(g_["x"] + g_["w"] for g_ in l0)
            for g_ in l0:
                g_["x"], g_["w"] = x0_, x1_ - x0_
    groups.sort(key=lambda g_: g_["level"])
    tot_w = max([n["x"] + n["w"] for n in scene_nodes] + [g_["x"] + g_["w"] for g_ in groups]
                + [p[0] for e in edges for p in e["pts"]]) + margin
    tot_h = max([n["y"] + n["h"] for n in scene_nodes] + [g_["y"] + g_["h"] for g_ in groups]
                + [p[1] for e in edges for p in e["pts"]]) + margin
    used_types = [t for t in PV_NODE_NAMES if any(n["type"] == t for n in scene_nodes)]
    return {"w": tot_w, "h": tot_h, "groups": groups, "nodes": scene_nodes, "edges": edges, "types": used_types,
            "prefix": pv_common_prefix([n["label"] for n in scene_nodes]), "mode": mode, "n_back": n_back, "n_contemp_conflict": n_skip}


_LIST_QUOTES = "\"'`\u201c\u201d\u201e\u201a\u2018\u2019\u00ab\u00bb"


def convert(table_file, schema, out_file=None, sheet="", log=print, template_text=None, table=None, out_format=None, model_only=False):
    """
    Erzeugt <out_file> (.hmcx). Standard: <Tabellenname>.hmcx neben der Tabelle.
    Gibt ein Dict mit Pfad, Kennzahlen und dem Protokolltext zurück.
    """
    table_file = Path(table_file)
    schema = normalize_schema(schema)
    m, st = schema["mapping"], schema["settings"]
    out_format = out_format or st.get("out_format", "hmc")
    ext = ".json" if out_format == "hme" else ".hmcx"
    out_file = Path(out_file) if out_file else table_file.with_suffix(ext)
    project_name = out_file.stem
    out_file.parent.mkdir(parents=True, exist_ok=True)

    warnings = []

    def log_warn(code, message):
        warnings.append((code, message))

    try:
        prefix_re = re.compile(st["id_prefix_regex"]) if st["id_prefix_regex"] else None
        split_re = re.compile(st["list_split_regex"] or r"[;,]")
    except re.error as e:
        raise HMAError(tr('Ungültiger regulärer Ausdruck in den erweiterten Einstellungen: {0}').format(e))
    funde_suffix = st["descr_suffix"] if st["descr_suffix_on"] == "1" else ""
    kat_interface = set(schema["kategorie_interface"])
    outer_ignore = set(schema["group_outer_ignore"])
    inner_ignore = set(schema["group_inner_ignore"])
    filter_values = set(schema["filter_values"])
    errors = []  # (Titel, [Details], Hinweis) – bei Fehlern wird KEINE Datei erzeugt

    def strip_prefix(id_str):
        if not id_str:
            return ""
        return prefix_re.sub("", id_str) if prefix_re else id_str

    def parse_list(cell_text):
        if not cell_text:
            return []
        cell_text = str(cell_text).strip()
        if not cell_text or cell_text.lower() == "nan":
            return []
        # JSON-artige Liste ["A","B"] -> A,B; Anführungszeichen um einzelne IDs werden entfernt
        if cell_text[0] == "[" and cell_text[-1] == "]":
            cell_text = cell_text[1:-1]
        items = (p.strip().strip(_LIST_QUOTES).strip() for p in split_re.split(cell_text))
        return [strip_prefix(p) for p in items if p]

    if table is not None:  # bereits geladene Tabelle (z. B. aus der Datenbank-Abfrage)
        headers, rows, fmt_info = table
    else:
        headers, rows, fmt_info = load_table(table_file, st, sheet)
    req, opt = check_schema(schema, headers)
    if req or opt:
        raise HMAError(tr('Zuordnungsfehler in {0}: ').format(table_file.name) + "; ".join(req + opt))
    idx = {k: headers.index(c) for k, c in used_columns(schema)}

    tug_vals = {}

    def cell(row, key):
        i = idx.get(key)
        if i is None or i >= len(row):
            return ""
        return (row[i] or "").strip()

    def compose_id(row):
        parts = [cell(row, k) for k in ("id1", "id2", "id3", "id4") if k in idx]
        return "_".join(p for p in parts if p)

    # --- Zeilenfilter ---
    n_rows_total = len(rows)
    if m["filter_col"] and filter_values:
        kept = []
        for line_num, row in rows:
            if m["filter_col"] == PREFIX_COL:
                mm = prefix_re.match(compose_id(row)) if prefix_re else None
                val = mm.group(0).rstrip("_") if mm else ""
            else:
                val = cell(row, "filter_col")
            if val in filter_values:
                kept.append((line_num, row))
        rows = kept
        if not rows:
            raise HMAError(tr("Der Zeilenfilter lässt keine einzige Zeile übrig. Bitte die Filterwerte prüfen."))

    # --- Typischer Anwenderfehler: mehrere Projekte/Grabungen in einer Tabelle ---
    if prefix_re and not (m["filter_col"] and filter_values):
        pref = Counter()
        for _, row in rows:
            mm = prefix_re.match(compose_id(row))
            if mm:
                pref[mm.group(0).rstrip("_")] += 1
        if len(pref) > 1:
            errors.append((
                tr('Die Tabelle enthält {0} verschiedene Projekte/Grabungen (ID-Präfixe)').format(len(pref)),
                [tr('{0}: {1} Zeilen').format(p, n) for p, n in pref.most_common(12)],
                tr("Eine Matrix gilt immer für EINE Grabung. SE-Nummern wiederholen sich von Grabung zu Grabung; "
                "werden sie vermischt, entstehen doppelte IDs, falsche Verweise und Kreise. Lösung: im Tab "
                "'Werte' beim Zeilenfilter 'Projekt-Präfix der ID' wählen und EIN Projekt anhaken.")))

    # --- Template ---
    tpl_text = template_text if template_text is not None else TEMPLATE_XML
    root = ET.fromstring(tpl_text.encode("utf-8"))
    g1 = None
    for g in root.iter():
        if g.tag.split("}", 1)[-1] == "graph" and g.get("id") == "g1":
            g1 = g
            break
    if g1 is None:
        raise HMAError(tr("<graph id='g1'> im Template nicht gefunden."))
    for c in list(g1):
        g1.remove(c)

    node_type_conflicts = []

    # --- PASS 1: Knoten ---
    se_base = {}
    ref_index = {}
    dup_ids = {}
    se_line = {}
    ref_dups = []
    for line_num, row in rows:
        if len(row) > len(headers):
            preview = "; ".join(row[len(headers):][:3])
            raise HMAError(
                tr('Parsing-Fehler: Zeile {0} hat zusätzliche Felder (Spaltenverschiebung). Ursache sehr '
                   'wahrscheinlich: Anführungszeichen im Export oder falsches Trennzeichen. Auszug: {1}').format(line_num, preview))
        raw_id = compose_id(row)
        if not raw_id:
            continue
        base_id = strip_prefix(raw_id)
        if base_id in se_base:
            dup_ids.setdefault(base_id, []).append(line_num)

        name = cell(row, "name")
        descr = cell(row, "descr")
        kat = cell(row, "kategorie")
        node_type = "SURFACE" if (kat and kat in kat_interface) else "DEPOSIT"

        outer = cell(row, "group_outer")
        period_id = None if (not outer or outer in outer_ignore) else outer
        inner = cell(row, "group_inner") or None
        if inner in inner_ignore:
            inner = None
        phase_id = inner

        se_line[base_id] = line_num
        tug_vals[base_id] = raw_id
        se_base[base_id] = {"name": name, "description": descr, "has_funde": bool(descr),
                            "node_type": node_type, "period_id": period_id, "phase_id": phase_id, "kat": kat}
        if st["ref_key_col"]:
            key = strip_prefix(cell(row, "ref_key_col"))
            if key:
                if key in ref_index and ref_index[key] != base_id:
                    ref_dups.append(tr("Verweis-Wert '{0}' gehört zu mehreren SEs ({1}, {2})").format(key, ref_index[key], base_id))
                else:
                    ref_index.setdefault(key, base_id)

    if dup_ids:
        errors.append((
            tr('Doppelte IDs: {0} ID(s) kommen mehrfach vor').format(len(dup_ids)),
            [tr("'{0}' in den Zeilen {1}").format(k, ', '.join(str(x) for x in sorted(set(v + [se_line.get(k, 0)]))[:4]))
             for k, v in list(dup_ids.items())[:12]],
            tr("Häufigste Ursache: die Tabelle enthält mehrere Grabungen/Projekte, und der Projekt-Präfix wurde von "
            "der ID entfernt. Lösung: im Tab 'Werte' den Zeilenfilter auf EIN Projekt setzen (oder den "
            "ID-Präfix im Tab 'Erweitert' leeren / ein weiteres ID-Feld zuordnen).")))
    if ref_dups:
        errors.append((
            tr('Mehrdeutige Verweise: {0} Verweis-Wert(e) gehören zu mehreren SEs').format(len(ref_dups)),
            ref_dups[:12],
            tr("Die Spalte, auf die die Beziehungen verweisen, ist nicht eindeutig (z. B. SE-Nummern aus mehreren "
            "Grabungen). Lösung: Zeilenfilter auf EIN Projekt setzen.")))

    def resolve(token):
        if st["ref_key_col"]:
            return ref_index.get(token)
        return token if token in se_base else None

    # --- PASS 2: Beziehungen ---
    above_raw_base = []
    zeitgleich_raw_dirs_base = set()
    ident_raw_dirs_base = set()
    unresolved = Counter()
    for line_num, row in rows:
        raw_id = compose_id(row)
        if not raw_id:
            continue
        src_base = strip_prefix(raw_id)
        if src_base not in se_base:
            continue
        for t in parse_list(cell(row, "above")):
            r_ = resolve(t)
            if r_:
                above_raw_base.append((src_base, r_))
            else:
                unresolved[tr("liegt über")] += 1
        for t in parse_list(cell(row, "below")):
            r_ = resolve(t)
            if r_:
                above_raw_base.append((r_, src_base))
            else:
                unresolved[tr("liegt unter")] += 1
        for t in parse_list(cell(row, "contemp")):
            r_ = resolve(t)
            if r_:
                zeitgleich_raw_dirs_base.add((src_base, r_))
            else:
                unresolved[tr("zeitgleich")] += 1
        for t in parse_list(cell(row, "ident")):
            r_ = resolve(t)
            if r_ == src_base:
                log_warn("IDENT_SELF_REFERENCE", tr('{0} ident {1} (Selbstreferenz) – ignoriert').format(src_base, t))
                continue
            if not r_:
                log_warn("IDENT_TARGET_MISSING", tr('{0} ident {1} – Ziel-SE existiert nicht – ignoriert').format(src_base, t))
                continue
            ident_raw_dirs_base.add((src_base, r_))

    # --- Ident-Gruppen ---
    uf = UnionFind()
    for bid in se_base:
        uf.add(bid)
    for (a, b) in ident_raw_dirs_base:
        uf.union(a, b)
    for (a, b) in sorted(ident_raw_dirs_base):
        if (b, a) not in ident_raw_dirs_base:
            log_warn("IDENT_NOT_SYMMETRIC", tr('{0} ident {1}, aber {2} nennt {3} nicht (bitte in der Datenbank prüfen)').format(a, b, b, a))

    groups = {}
    for bid in se_base:
        groups.setdefault(uf.find(bid), []).append(bid)

    # --- Merge-IDs ---
    base_to_final = {}
    final_nodes = {}
    ident_members = {}

    def build_merged_id(members_sorted):
        sfx = se_suffix(members_sorted[0])
        ids_part = "=".join(se_label(x) for x in members_sorted)
        merged_base = f"{ids_part}_{sfx}" if sfx else ids_part
        if funde_suffix and any(se_base[x]["has_funde"] for x in members_sorted):
            return merged_base + funde_suffix
        return merged_base

    for members in groups.values():
        members_sorted = sorted(members, key=lambda x: (se_num(x), x))

        if len(members_sorted) == 1:
            mm = members_sorted[0]
            final_id = mm + funde_suffix if (funde_suffix and se_base[mm]["has_funde"]) else mm
            base_to_final[mm] = final_id
            final_nodes[final_id] = {k: se_base[mm][k] for k in ("name", "description", "node_type",
                                                                  "period_id", "phase_id", "kat")}
            continue

        smallest = members_sorted[0]
        if len(members_sorted) > 8:
            log_warn("IDENT_LARGE_GROUP",
                     tr("{0} SEs werden als identisch zusammengeführt (ab {1}) – sehr ungewöhnlich, bitte die Spalte 'ident mit' prüfen").format(len(members_sorted), members_sorted[0]))
        suffixes = {se_suffix(x) for x in members_sorted}
        if len(suffixes) > 1:
            log_warn("IDENT_SUFFIX_MISMATCH",
                     tr('{0} – Suffixe {1}; verwende Suffix von kleinster ID ({2})').format(', '.join(members_sorted), sorted(suffixes), smallest))

        node_types = {se_base[x]["node_type"] for x in members_sorted}
        winner_type = se_base[smallest]["node_type"]
        merged_final_id = build_merged_id(members_sorted)

        if len(node_types) > 1:
            node_type_conflicts.append({
                "merged_node_id": merged_final_id, "members_sorted": members_sorted, "winner_type": winner_type,
                "types_by_se": {x: se_base[x]["node_type"] for x in members_sorted}})

        period_vals = [se_base[x]["period_id"] for x in members_sorted]
        phase_vals = [se_base[x]["phase_id"] for x in members_sorted]
        if len(set(period_vals)) > 1:
            log_warn("IDENT_GROUP_PERIOD_MISMATCH",
                     tr('Ident-Gruppe {0} hat unterschiedliche Perioden: {1} – kleinste ID gewinnt ({2}: {3})').format(members_sorted, sorted(set(period_vals), key=lambda x: str(x)), smallest, se_base[smallest]['period_id']))
        if len(set(phase_vals)) > 1:
            log_warn("IDENT_GROUP_PHASE_MISMATCH",
                     tr('Ident-Gruppe {0} hat unterschiedliche Phasen: {1} – kleinste ID gewinnt ({2}: {3})').format(members_sorted, sorted(set(phase_vals), key=lambda x: str(x)), smallest, se_base[smallest]['phase_id']))

        for x in members_sorted:
            base_to_final[x] = merged_final_id
        ident_members[merged_final_id] = list(members_sorted)

        parts = []
        for x in members_sorted:
            if (se_base[x]["description"] or "").strip() == "":
                continue
            parts.extend([tr('[SE {0}]').format(x), se_base[x]["description"], ""])
        final_nodes[merged_final_id] = {
            "name": se_base[smallest]["name"], "description": "\n".join(parts).rstrip(),
            "node_type": winner_type, "period_id": se_base[smallest]["period_id"],
            "phase_id": se_base[smallest]["phase_id"], "kat": se_base[smallest]["kat"]}

    # --- Perioden/Phasen ---
    se_nodes = final_nodes
    periods = {}
    phases = {}
    for final_id, info in se_nodes.items():
        period_id, phase_id = info["period_id"], info["phase_id"]
        if period_id:
            periods.setdefault(period_id, {"name": period_id, "phase_ids": set(), "se_ids": set()})
            periods[period_id]["se_ids"].add(final_id)
        if phase_id:
            phases.setdefault(phase_id, {"name": phase_id, "period_id": None, "se_ids": set()})
            phases[phase_id]["se_ids"].add(final_id)
            if period_id:
                phases[phase_id]["period_id"] = period_id
                periods.setdefault(period_id, {"name": period_id, "phase_ids": set(), "se_ids": set()})
                periods[period_id]["phase_ids"].add(phase_id)

    # --- Kanten ---
    above_set = set()
    for (s_base, t_base) in above_raw_base:
        s, t = base_to_final.get(s_base), base_to_final.get(t_base)
        if not s or not t or s == t:
            continue
        above_set.add((s, t))
    above_edges = [(s, t, "ABOVE") for (s, t) in sorted(above_set)]

    zeitgleich_dirs_final = set()
    for (a_base, b_base) in zeitgleich_raw_dirs_base:
        a, b = base_to_final.get(a_base), base_to_final.get(b_base)
        if not a or not b or a == b:
            continue
        zeitgleich_dirs_final.add((a, b))

    zeitgleich_edges = []
    emitted_pairs = set()
    for (a, b) in sorted(zeitgleich_dirs_final):
        pair = frozenset((a, b))
        if pair in emitted_pairs:
            continue
        if (b, a) not in zeitgleich_dirs_final:
            log_warn("CONTEMPORARY_NOT_SYMMETRIC", tr('{0} zeitgleich {1}, aber {2} nicht zeitgleich {3}').format(a, b, b, a))
        src, tgt = sorted((a, b))
        zeitgleich_edges.append((src, tgt, "CONTEMPORARY"))
        emitted_pairs.add(pair)

    # --- Optional: T (Top surface), G (Interface to Geology), U (Unexcavated) ---
    TUG_DEF = {"T": ("TOP_SURFACE", "Top surface", "0"), "G": ("GEOLOGY", "Interface to Geology", "2"),
               "U": ("UNEXCAVATED", "Unexcavated", "1")}
    tug_nodes = {}
    tug_sel = {k: set(schema[f"tug_{k.lower()}_values"]) for k in "TGU"}
    for k in "TGU":
        if st[f"tug_{k.lower()}_on"] == "1":
            if out_format == "hme":
                log_warn("HME_NO_TUG", tr('{0} ({1}) gibt es im HME nicht – wird für die HME-Ausgabe weggelassen.').format(TUG_DEF[k][1], k))
            else:
                tug_nodes[k] = TUG_DEF[k]
    if tug_nodes:
        for k in tug_nodes:
            if k in se_nodes or k in se_base:
                raise HMAError(tr("Eine SE hat die ID '{0}' – das kollidiert mit dem Node '{1}' ({2}). Bitte T/G/U deaktivieren oder "
                                  'die SE-ID ändern.').format(k, TUG_DEF[k][1], k))
        extra = set()
        counts_tug = {k: 0 for k in "TGU"}
        for base, raw in tug_vals.items():
            fin = base_to_final.get(base)
            if not fin:
                continue
            for k in "TGU":
                if k in tug_nodes and tug_sel[k] and raw in tug_sel[k]:
                    counts_tug[k] += 1
                    extra.add(("T", fin) if k == "T" else (fin, k))
        for k in "TGU":
            if k in tug_nodes and tug_sel[k] and not counts_tug[k]:
                log_warn("TUG_NO_MATCH", tr('{0} ({1}): Die gewählten SE-IDs kommen in der Tabelle nicht vor.').format(TUG_DEF[k][1], k))
        # Standardkette T → U → G, wenn keine eigenen Werte gewählt sind
        if "T" in tug_nodes and "U" in tug_nodes and not tug_sel["T"]:
            extra.add(("T", "U"))
        if "U" in tug_nodes and "G" in tug_nodes and not tug_sel["G"]:
            extra.add(("U", "G"))
        if "T" in tug_nodes and "G" in tug_nodes and "U" not in tug_nodes and not tug_sel["T"] and not tug_sel["G"]:
            extra.add(("T", "G"))
        above_edges = sorted(set(above_edges) | {(a, b, "ABOVE") for (a, b) in extra})

    if out_format == "hme" and zeitgleich_edges:
        log_warn("HME_NO_CONTEMPORARY", tr("{0} 'zeitgleich'-Beziehung(en) gibt es im HME nicht – sie werden für die HME-Ausgabe weggelassen.").format(len(zeitgleich_edges)))
    edges = above_edges + zeitgleich_edges

    # --- Logikprüfungen: es wird nur eine Datei erzeugt, die HMC auch öffnen kann ---
    # Gleicher Text in übergeordneter/untergeordneter Gruppe (oder wie eine SE-ID) ist erlaubt:
    # intern bekommen die Gruppen eindeutige IDs, der angezeigte Name bleibt gleich.
    used_ids = set(se_nodes) | set(tug_nodes)
    pid_emit, phid_emit = {}, {}

    def uniq(base):
        cand, n = base, 2
        while cand in used_ids:
            cand = f"{base}_{n}"
            n += 1
        used_ids.add(cand)
        return cand
    for pid in sorted(periods):
        pid_emit[pid] = uniq(pid)
    for phid in sorted(phases):
        phid_emit[phid] = uniq(phid)

    adj = {}
    for (s, t, _) in above_edges:
        adj.setdefault(s, []).append(t)
    cyc = [c for c in strongly_connected(adj) if len(c) > 1]
    if cyc:
        cyc.sort(key=len, reverse=True)
        details = []
        for c in cyc[:6]:
            path = shortest_cycle(adj, c)
            shown = ", ".join(sorted(c, key=lambda x: (se_num(x), x))[:8]) + (" …" if len(c) > 8 else "")
            details.append(tr('{0} SEs im Kreis, z. B.: {1}   [betroffen: {2}]').format(len(c), ' → '.join(path), shown))
        errors.append((
            tr('Zyklische Stratigraphie: {0} Kreis-Gruppe(n), insgesamt {1} SEs').format(len(cyc), sum(len(c) for c in cyc)),
            details,
            tr("A liegt über B, B über C und C wieder über A ist stratigraphisch unmöglich; HMC kann so eine Matrix "
            "nicht öffnen. Ursachen: Fehler in den Daten, falsch zugeordnete Spalten, oder mehrere Grabungen "
            "vermischt (Zeilenfilter!).")))

    if errors:
        msg = [tr('Es wurde KEINE {0}-Datei erzeugt, weil die Daten logische Fehler enthalten, mit denen {1} die Datei '
                  'nicht öffnen könnte:').format(ext, 'HME' if out_format == 'hme' else 'HMC'), ""]
        for i, (title, details, hint) in enumerate(errors, start=1):
            msg.append(f"{i}. {title}")
            msg.extend("     - " + d for d in details)
            if len(details) >= 12:
                msg.append("     - …")
            msg.append(tr('   Hinweis: {0}').format(hint))
            msg.append("")
        raise HMAError("\n".join(msg).rstrip())

    # Widersprüche zeitgleich/über (Warnung): A über B, aber A zeitgleich B
    cparent = {}

    def cfind(x):
        cparent.setdefault(x, x)
        while cparent[x] != x:
            cparent[x] = cparent[cparent[x]]
            x = cparent[x]
        return x

    for (a, b, _) in zeitgleich_edges:
        cparent[cfind(a)] = cfind(b)
    for (s, t, _) in above_edges:
        if cfind(s) == cfind(t):
            log_warn("STRAT_CONTEMPORARY_CONFLICT", tr('{0} liegt über {1}, ist aber (direkt oder über Dritte) auch zeitgleich – Widerspruch in den Daten').format(s, t))
    if len(se_nodes) + len(periods) + len(phases) > 2000:
        log_warn("HMC_LARGE", tr('{0} Nodes – sehr große Matrix, HMC kann langsam werden. Besser einzeln nach Projekt/Zeitraum filtern.').format(len(se_nodes)))

    if model_only:  # nur für die Voransicht: nichts schreiben
        return {"model": {"out_format": out_format, "se_nodes": se_nodes, "periods": periods, "phases": phases,
                          "above_edges": above_edges, "zeitgleich_edges": zeitgleich_edges, "tug_nodes": tug_nodes,
                          "hme_types": schema.get("hme_types") or {}, "idents": ident_members,
                          "hme_undated": st.get("hme_undated", "1") == "1"},
                "warnings": warnings}

    hme_conflict = False
    if out_format == "hme":
        hme_data, hme_conflict, hme_info = build_hme(se_nodes, periods, phases, above_edges,
                                                     undated_groups=st.get("hme_undated", "1") == "1",
                                                     type_map=schema.get("hme_types") or {})
        tc = Counter(n["type"] for n in hme_data["nodes"])
        log(tr("  HME-Node-Typen: ") + ", ".join(f"{k} {tc[k]}" for k in HME_TYPES if tc.get(k)))
        if hme_info["undated"]:
            log(tr("  HME: {0} Units ohne übergeordnete Gruppe -> {1} Phase(n) 'unknown-N' (je nach stratigraphischer "
                   'Lage zwischen den datierten Gruppen)').format(hme_info['undated'], hme_info['unknown_groups']))
            if hme_info["contradict"]:
                log_warn("HME_UNKNOWN_CONTRADICTION", tr('{0} Unit(s) ohne Gruppe liegen zwischen Gruppen, die sich stratigraphisch widersprechen; die '
                                                         "Zuordnung zu 'unknown-N' ist dort nur ungefähr.").format(hme_info['contradict']))
        elif st.get("hme_undated", "1") != "1" and any(not i["period_id"] for i in se_nodes.values()):
            log_warn("HME_UNDATED_BAND", tr("Units ohne übergeordnete Gruppe stehen im HME in einem eigenen Band am "
                                          "Ende. Option 'Undatierte Units in unknown-Phasen einordnen' aktivieren."))
        if hme_conflict:
            log_warn("HME_PHASE_ORDER", tr("Die übergeordneten Gruppen (HME: Phasen) überlappen sich stratigraphisch "
                                         "(Zyklus zwischen den Gruppen); die Reihenfolge der Phasen folgt der mittleren "
                                         "stratigraphischen Tiefe. Bitte im HME unter 'Validate' prüfen."))
        tmp_file = out_file.with_name(out_file.name + ".tmp")
        try:
            tmp_file.write_text(json.dumps(hme_data, indent=2, ensure_ascii=False), encoding="utf-8")
        except (PermissionError, OSError) as e:
            raise HMAError(tr('Die Datei kann nicht in diesem Ordner gespeichert werden:\n  {0}\nGrund: {1}').format(out_file.parent, e))
        problems = verify_hme(tmp_file)
        if problems:
            tmp_file.unlink()
            raise HMAError(tr("Die erzeugte Datei hat die Selbstprüfung nicht bestanden und wurde verworfen:\n  - ")
                            + "\n  - ".join(problems))
        tmp_file.replace(out_file)
        edge_counter = len(hme_data["edges"]) + 1
    else:
        # --- Graph bauen ---
        graph_id_counter = 2
        graph_containers = {("root", None): g1}
        for pid in sorted(periods):
            pnode = ET.SubElement(g1, "node", {"id": pid_emit[pid]})
            pgraph = ET.SubElement(pnode, "graph", {"edgedefault": "directed", "id": f"g{graph_id_counter}"})
            graph_id_counter += 1
            pnode.append(make_hmcnode(pid_emit[pid], pid, "", "PERIOD_GROUP", open_flag=True))
            graph_containers[("period", pid)] = pgraph

        for phid, ph in sorted(phases.items()):
            parent = graph_containers.get(("period", ph["period_id"]), g1)
            phnode = ET.SubElement(parent, "node", {"id": phid_emit[phid]})
            phgraph = ET.SubElement(phnode, "graph", {"edgedefault": "directed", "id": f"g{graph_id_counter}"})
            graph_id_counter += 1
            phnode.append(make_hmcnode(phid_emit[phid], phid, "", "PHASE_GROUP", open_flag=True))
            graph_containers[("phase", phid)] = phgraph

        for se_id, info in sorted(se_nodes.items()):
            key = ("root", None)
            if info["phase_id"] and ("phase", info["phase_id"]) in graph_containers:
                key = ("phase", info["phase_id"])
            elif info["period_id"] and ("period", info["period_id"]) in graph_containers:
                key = ("period", info["period_id"])
            node = ET.SubElement(graph_containers[key], "node", {"id": se_id})
            node.append(make_hmcnode(se_id, info["name"], info["description"], info["node_type"]))

        for k, (ttype, tname, tlayer) in sorted(tug_nodes.items()):
            node = ET.SubElement(g1, "node", {"id": k})
            node.append(make_hmcnode(k, tname, "", ttype, layer=tlayer))

        edge_counter = 1
        for (src, tgt, etype) in edges:
            e = ET.SubElement(g1, "edge", {"id": f"e{edge_counter}", "source": src, "target": tgt})
            d = ET.SubElement(e, "data", {"key": "d2"})
            d.append(make_hmcedge(etype))
            edge_counter += 1

        g1.set("parse.nodes", str(len(se_nodes) + len(periods) + len(phases) + len(tug_nodes)))
        g1.set("parse.edges", str(edge_counter - 1))

        # --- XML zusammensetzen ---
        indent(g1)
        new_g1_xml = ET.tostring(g1, encoding="unicode")
        try:
            graph_start = tpl_text.index('<graph edgedefault="directed" id="g1"')
            data_pos = tpl_text.index('<data key="d3"')
        except ValueError:
            raise HMAError(tr("Template hat nicht die erwartete Struktur (graph g1 / data d3)."))
        close_pos = tpl_text.rfind("</graph>", 0, data_pos)
        graph_end = close_pos + len("</graph>")
        matrix_xml = tpl_text[:graph_start] + new_g1_xml + tpl_text[graph_end:]

        project_xml = ('<?xml version="1.0" encoding="UTF-8" ?>\n'
                       f'<ProjectProperties Name="{escape(project_name, {chr(34): "&quot;"})}" '
                       'Description="" ExcavationSite=""></ProjectProperties>\n')

        # Alles als reines ASCII (Zeichenreferenzen): Umlaute/ß/Sonderzeichen bleiben unabhängig davon korrekt,
        # wie HMC die Dateien einliest.
        tmp_file = out_file.with_name(out_file.name + ".tmp")
        try:
            with zipfile.ZipFile(tmp_file, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("matrix.xml", ascii_xml(matrix_xml))
                zf.writestr("project.xml", ascii_xml(project_xml))
        except (PermissionError, OSError) as e:
            try:
                tmp_file.unlink()
            except Exception:
                pass
            raise HMAError(
                tr('Die Datei kann nicht in diesem Ordner gespeichert werden:\n'
                   '  {0}\n'
                   'Grund: {1}\n'
                   "Meist fehlt die Schreibberechtigung (z. B. Ordner unter 'C:\\Program Files') oder die Datei ist in "
                   "einem anderen Programm (HMC) geöffnet. Bitte oben bei 'Speicherort (.hmcx)' einen Ordner wählen, in "
                   'dem du schreiben darfst (z. B. Dokumente oder Desktop).').format(out_file.parent, e))
        problems = verify_hmcx(tmp_file)
        if problems:
            tmp_file.unlink()
            raise HMAError(tr("Die erzeugte Datei hat die Selbstprüfung nicht bestanden und wurde verworfen:\n  - ")
                            + "\n  - ".join(problems[:10]))
        tmp_file.replace(out_file)

    # --- Protokoll ---
    warn_map = {}
    for code, msg in warnings:
        warn_map.setdefault(code, []).append(msg)

    def section(lines, title, codes, empty_text=tr("keine Fehler")):
        lines.append("=" * 60)
        lines.append(title)
        lines.append("=" * 60)
        lines.append("")
        if not any(warn_map.get(c) for c in codes):
            lines.append(empty_text)
            lines.append("")
            return
        for c in codes:
            msgs = warn_map.get(c, [])
            if msgs:
                lines.append(c)
                lines.append(tr('{0} Fehler gefunden').format(len(msgs)))
                lines.extend(msgs)
                lines.append("")

    rep = [tr('Projekt: {0}').format(project_name), tr('Tabelle: {0}').format(table_file.name)
           + (f" ({tr('Tabelle') if table_file.suffix.lower() in DB_FILE_EXT else tr('Blatt')}: {sheet or tr('erste')})"
              if table_file.suffix.lower() in SHEET_EXT else ""),
           tr('Zeitpunkt: {0}').format(datetime.now().strftime('%Y-%m-%d %H:%M:%S')), tr('Erzeugt mit: HMA v{0}').format(APP_VERSION),
           tr("Spaltenzuordnung:")]
    id_cols = [m[k] for k in ("id1", "id2", "id3", "id4") if m[k]]
    rep.append(tr("  ID: ") + " + ".join(id_cols) + (tr(" (mit _ verbunden)") if len(id_cols) > 1 else ""))
    labels = [("name", "Name"), ("descr", tr("Funde/Beschreibung")), ("kategorie", tr("Kategorie")),
              ("group_outer", tr("Übergeordnete Gruppe")), ("group_inner", tr("Untergeordnete Gruppe")),
              ("above", tr("liegt über")), ("below", tr("liegt unter")),
              ("contemp", tr("zeitgleich mit")), ("ident", tr("ident mit"))]
    for k, lab in labels:
        if m[k]:
            rep.append(f"  {lab}: {m[k]}")
    if st["ref_key_col"]:
        rep.append(tr("  Beziehungen verweisen auf: Spalte '{0}'").format(st['ref_key_col']))
    if kat_interface:
        rep.append(tr("  Interface-Werte (Kategorie): ") + ", ".join(sorted(kat_interface)))
    if outer_ignore:
        rep.append(tr("  Ignorierte Werte (übergeordnete Gruppe): ") + ", ".join(sorted(outer_ignore)))
    for k in "TGU":
        if k in tug_nodes:
            sel = sorted(tug_sel[k])
            rep.append(tr('  {0} ({1}): aktiv').format(TUG_DEF[k][1], k)
                       + (tr(', SEs: {0}').format(', '.join(sel)) if sel else tr(" (Standardverbindung)")))
    if inner_ignore:
        rep.append(tr("  Ignorierte Werte (untergeordnete Gruppe): ") + ", ".join(sorted(inner_ignore)))
    if m["filter_col"] and filter_values:
        fname = tr("Projekt-Präfix der ID") if m["filter_col"] == PREFIX_COL else m["filter_col"]
        rep.append(tr('  Zeilenfilter: {0} = {1}  ({2} von {3} Zeilen)').format(fname, ', '.join(sorted(filter_values)), len(rows), n_rows_total))
    rep.append(tr('Ergebnis: {0} Nodes, {1} übergeordnete Gruppen, {2} untergeordnete Gruppen, {3} Kanten  ({4})').format(len(se_nodes) + len(tug_nodes), len(periods), len(phases), edge_counter - 1, fmt_info))
    if unresolved:
        rep.append(tr("Hinweis: Verweise auf SEs, die nicht in der Tabelle stehen, wurden übersprungen (")
                   + ", ".join(f"{k}: {v}" for k, v in unresolved.items())
                   + tr(") – normal bei Teilauszügen."))
    rep.append("")

    lines = []
    lines.append("=" * 60)
    lines.append(tr("1. Node-Type Fehler (DEPOSIT / SURFACE)"))
    lines.append("=" * 60)
    lines.append("")
    if not node_type_conflicts:
        lines += [tr("keine Fehler"), ""]
    else:
        lines += [tr('{0} Konflikt(e) gefunden (für jeden Konflikt wurde der Typ der numerisch kleinsten SE übernommen).').format(len(node_type_conflicts)), ""]
        for i, c in enumerate(node_type_conflicts, start=1):
            lines += [tr('Konflikt {0}').format(i), tr('Merged Node: {0}').format(c['merged_node_id']),
                      tr('Ident-Gruppe (Base-IDs): {0}').format(', '.join(c['members_sorted'])),
                      tr('Übernommener Typ (kleinste SE): {0}').format(c['winner_type']), tr("Typen je SE:")]
            lines += [f"- {x}: {c['types_by_se'].get(x, '')}" for x in c["members_sorted"]]
            lines.append("")
    section(lines, tr("2. Gleichsetzung (ident)"),
            ["IDENT_SELF_REFERENCE", "IDENT_TARGET_MISSING", "IDENT_NOT_SYMMETRIC", "IDENT_SUFFIX_MISMATCH",
             "IDENT_GROUP_PERIOD_MISMATCH", "IDENT_GROUP_PHASE_MISMATCH"])
    section(lines, tr("3. Zeitgleich-Setzung"), ["CONTEMPORARY_NOT_SYMMETRIC"])
    section(lines, tr("4. Gruppen (übergeordnet / untergeordnet)"), ["PHASE_OBJECT_DOUBLE_ASSIGNMENT"])
    section(lines, tr("5. Stratigraphie und Größe"), ["STRAT_CONTEMPORARY_CONFLICT", "IDENT_LARGE_GROUP", "HMC_LARGE"])
    report = "\n".join(rep + lines).rstrip() + "\n"

    log(tr('  {0} Nodes, {1} + {2} Gruppen, {3} Kanten').format(len(se_nodes), len(periods), len(phases), edge_counter - 1))
    return {"hmcx": out_file, "format": out_format, "project": project_name, "nodes": len(se_nodes), "periods": len(periods),
            "phases": len(phases), "edges": edge_counter - 1, "warnings": len(warnings),
            "conflicts": len(node_type_conflicts), "report": report}


# ---------------------------------------------------------------------------
# Oberfläche
# ---------------------------------------------------------------------------
ID_KEYS = ["id1", "id2", "id3", "id4"]
NONE_LABEL = tr("(nicht verwenden)")
NO_FILTER = tr("(kein Filter)")
PREFIX_LABEL = tr("(Projekt-Präfix der ID, z. B. XTESTX)")
REF_DEFAULT = tr("(zusammengesetzte ID)")
C_OK, C_BAD, C_MUTED = "#1b7a2b", "#b00020", "#777777"


def run_gui(config_path=None):
    import tkinter as tk
    from tkinter import filedialog, messagebox, simpledialog, ttk

    class Tooltip:
        def __init__(self, widget, text):
            self.w, self.text, self.tip = widget, text, None
            widget.bind("<Enter>", self.show)
            widget.bind("<Leave>", self.hide)

        def show(self, _e=None):
            if self.tip or not self.text:
                return
            self.tip = tk.Toplevel(self.w)
            self.tip.wm_overrideredirect(True)
            self.tip.wm_geometry(f"+{self.w.winfo_rootx() + 20}+{self.w.winfo_rooty() + self.w.winfo_height() + 4}")
            tk.Label(self.tip, text=self.text, justify="left", background="#ffffe0", relief="solid",
                     borderwidth=1, wraplength=440, padx=6, pady=3).pack()

        def hide(self, _e=None):
            if self.tip:
                self.tip.destroy()
                self.tip = None

    class SearchCombobox(ttk.Combobox):
        """Combobox mit Textsuche: Tippen filtert die Liste (z. B. 'liegt'), Enter/Pfeil-runter öffnet sie.
        Der Wert wird erst übernommen, wenn er genau einem Eintrag entspricht (sonst springt er zurück)."""

        def __init__(self, parent, variable, **kw):
            self._ext = variable
            self._txt = tk.StringVar(value=variable.get())
            self._all = []
            kw["textvariable"] = self._txt
            kw["state"] = "normal"
            super().__init__(parent, **kw)
            variable.trace_add("write", lambda *a: self._sync_in())
            self.bind("<KeyRelease>", self._key)
            self.bind("<<ComboboxSelected>>", self._selected, add="+")
            self.bind("<FocusOut>", lambda e: self.after(250, self._commit), add="+")
            self.bind("<Return>", self._enter, add="+")
            self.bind("<ButtonPress-1>", self._hook_list, add="+")
            self._picked = None
            self._hook_list()

        def _hook_list(self, _e=None):
            # Tk wählt per Listen-Index aus der vollen Werteliste; bei gefilterter Liste wäre das die falsche Spalte.
            # Daher den angeklickten Text direkt aus der Liste lesen.
            try:
                lb = f"{self}.popdown.f.l"
                if not getattr(self, "_hooked", False):
                    self.tk.call("ttk::combobox::PopdownWindow", str(self))
                    lbw = self.nametowidget(lb)
                    lbw.bind("<ButtonRelease-1>", lambda e, w=lbw: self._pick_from(w, e), add="+")
                    lbw.bind("<Return>", lambda e, w=lbw: self._pick_from(w, e, key=True), add="+")
                    self._hooked = True
            except Exception:
                pass

        def _pick_from(self, w, e, key=False):
            try:
                if key:
                    sel = w.curselection()
                    idx = sel[0] if sel else None
                else:
                    idx = w.nearest(e.y)
                self._picked = w.get(idx) if idx is not None else None
            except Exception:
                self._picked = None


        def set_values(self, values):
            self._all = list(values)
            self.configure(values=self._all)

        def _sync_in(self):
            if self._txt.get() != self._ext.get():
                self._txt.set(self._ext.get())
            self.configure(values=self._all)

        def _matches(self):
            t = self._txt.get().strip().lower()
            return [v for v in self._all if t in v.lower()] if t else list(self._all)

        def _key(self, e):
            if e.keysym in ("Up", "Down", "Return", "Tab", "Escape", "Left", "Right", "Home", "End",
                            "Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R"):
                return
            self.configure(values=self._matches())

        def _selected(self, _e=None):
            if self._picked:
                self._txt.set(self._picked)
                self._picked = None
            self._ext.set(self._txt.get())
            self.configure(values=self._all)

        def _popup_open(self):
            try:
                return bool(int(self.tk.call("winfo", "ismapped", f"{self}.popdown")))
            except Exception:
                return False

        def _commit(self, _e=None):
            if self._popup_open():
                self.after(300, self._commit)  # Liste offen: später erneut prüfen, Klick-Auswahl nicht überschreiben
                return
            text = self._txt.get().strip()
            if not text:
                self._ext.set("")
            else:
                exact = next((v for v in self._all if v.lower() == text.lower()), None)
                if exact is not None:
                    self._ext.set(exact)
                    self._txt.set(exact)
                else:
                    self._txt.set(self._ext.get())
            self.configure(values=self._all)

        def _enter(self, _e=None):
            m = self._matches()
            if len(m) == 1:
                self._txt.set(m[0])
                self._ext.set(m[0])
                self.configure(values=self._all)
            else:
                self.configure(values=m)
                try:
                    self.tk.call("ttk::combobox::Post", str(self))
                except Exception:
                    pass

    AUTO_LABEL = tr("(automatisch)")
    EX_LEAD, EX_ETC = tr("Bsp.: "), "..., etc."

    class MultiPick(ttk.Frame):
        """Durchsuchbare Mehrfachauswahl: Suchfeld + scrollbare Liste, Klick schaltet einen Eintrag ein/aus."""

        def __init__(self, parent, selected, on_change=None, height=6, width=46):
            super().__init__(parent)
            self.sel, self.on_change = selected, on_change
            self.items, self.shown = [], []
            self.q = tk.StringVar()
            top = ttk.Frame(self)
            top.pack(fill="x")
            ttk.Label(top, text=tr("Suchen:")).pack(side="left")
            self.entry = ttk.Entry(top, textvariable=self.q, width=24)
            self.entry.pack(side="left", padx=4)
            self.btn_clear = ttk.Button(top, text=tr("Auswahl leeren"), command=self.clear)
            self.btn_clear.pack(side="left", padx=4)
            self.lbl = ttk.Label(top, text="", foreground=C_MUTED)
            self.lbl.pack(side="left", padx=6)
            body = ttk.Frame(self)
            body.pack(fill="x", pady=(3, 0))
            self.lb = tk.Listbox(body, height=height, width=width, exportselection=False, activestyle="none",
                                 selectmode="browse")
            sb = ttk.Scrollbar(body, orient="vertical", command=self.lb.yview)
            self.lb.configure(yscrollcommand=sb.set)
            self.lb.pack(side="left", fill="x", expand=True)
            sb.pack(side="left", fill="y")
            self.q.trace_add("write", lambda *a: self.refresh())
            self.lb.bind("<ButtonRelease-1>", self._click)
            self.lb.bind("<MouseWheel>", self._wheel)
            self.lb.bind("<Button-4>", self._wheel)
            self.lb.bind("<Button-5>", self._wheel)

        def _wheel(self, e):
            if sys.platform == "darwin":
                self.lb.yview_scroll(wheel_units(e), "units")
                return "break"
            d = -1 if (e.delta > 0 if e.delta else e.num == 4) else 1
            self.lb.yview_scroll(d * 2, "units")
            return "break"

        def set_items(self, items):
            self.items = list(items)
            self.refresh()

        def set_enabled(self, on):
            st = "normal" if on else "disabled"
            for w in (self.entry, self.btn_clear):
                w.configure(state=st)
            self.lb.configure(state=st)

        def refresh(self):
            top = self.lb.yview()[0]
            t = self.q.get().strip().lower()
            allv = self.items + sorted(v for v in self.sel if v not in set(self.items))
            self.shown = [v for v in allv if t in v.lower()] if t else allv
            self.lb.delete(0, "end")
            for v in self.shown:
                self.lb.insert("end", ("✓  " if v in self.sel else "     ") + v)
            self.lb.yview_moveto(top)
            n = len(self.sel)
            self.lbl.configure(text=tr("keine Auswahl") if not n else tr('{0} ausgewählt').format(n) if n > 3 else
                               tr("ausgewählt: ") + ", ".join(sorted(self.sel)))

        def _click(self, e):
            try:
                i = self.lb.nearest(e.y)
                v = self.shown[i]
            except Exception:
                return
            (self.sel.discard if v in self.sel else self.sel.add)(v)
            self.refresh()
            if self.on_change:
                self.on_change()

        def clear(self):
            self.sel.clear()
            self.refresh()
            if self.on_change:
                self.on_change()

    class ScrollFrame(ttk.Frame):
        def __init__(self, parent):
            super().__init__(parent)
            self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
            vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
            self.canvas.configure(yscrollcommand=vsb.set)
            self.canvas.pack(side="left", fill="both", expand=True)
            self._need = False
            self.inner = ttk.Frame(self.canvas)
            win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

            def update_bar(_e=None):
                # Scrollbalken nur anzeigen, wenn der Inhalt höher ist als das sichtbare Fenster
                try:
                    ch = self.canvas.winfo_height()
                    if ch <= 1 or not self.winfo_ismapped():
                        return
                    need_h = self.inner.winfo_reqheight()
                    need = need_h > ch
                    self.canvas.configure(scrollregion=(0, 0, self.canvas.winfo_width(), max(need_h, ch)))
                    if need and not self._need:
                        vsb.pack(side="right", fill="y")
                    elif not need:
                        if self._need:
                            vsb.pack_forget()
                        self.canvas.yview_moveto(0)
                    self._need = need
                except Exception:
                    pass

            def poll():
                update_bar()
                self.after(300, poll)

            self.inner.bind("<Configure>", update_bar)
            self.canvas.bind("<Configure>", lambda e: (self.canvas.itemconfigure(win, width=e.width), update_bar()))
            self.after(300, poll)

            def wheel(e):
                if not self._need:
                    return
                self.canvas.yview_scroll(wheel_units(e), "units")

            self.canvas.bind("<Enter>", lambda e: (self.canvas.bind_all("<MouseWheel>", wheel),
                                                  self.canvas.bind_all("<Button-4>", wheel),
                                                  self.canvas.bind_all("<Button-5>", wheel)))
            self.canvas.bind("<Leave>", lambda e: (self.canvas.unbind_all("<MouseWheel>"),
                                                  self.canvas.unbind_all("<Button-4>"),
                                                  self.canvas.unbind_all("<Button-5>")))

    class App(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title(f"{APP_NAME} v{APP_VERSION}")
            self.minsize(900, 640)
            self.store = ConfigStore(config_path or pick_config_path())
            self.headers, self.rows = [], []
            self.files = []
            self.sheets = []
            self._stamp = None
            self._busy = False
            self._last_hmcx = None
            self._vl_key = None
            self._report = ""
            try:
                self._icon64 = tk.PhotoImage(data=ICON_64_B64)
                self.iconphoto(True, self._icon64)
            except Exception:
                self._icon64 = None

            last = self.store.data.get("last", {})
            names = sorted(self.store.data["schemas"])
            start_name = last.get("schema") if last.get("schema") in self.store.data["schemas"] else \
                (names[0] if names else "")
            sc = normalize_schema(last) if last.get("mapping") else \
                (self.store.data["schemas"][start_name] if start_name else normalize_schema({}))

            self.v_schema = tk.StringVar(value=start_name)
            self.v_file = tk.StringVar(value="")
            self.v_out = tk.StringVar(value="")
            self.v_sheet = tk.StringVar(value="")
            self.map_vars = {k: tk.StringVar(value=sc["mapping"][k]) for k in MAP_KEYS}
            self.set_vars = {k: tk.StringVar(value=sc["settings"][k]) for k in DEFAULT_SETTINGS}
            self.v_suffix_on = tk.BooleanVar(value=sc["settings"]["descr_suffix_on"] == "1")
            self.v_ref = tk.StringVar(value=sc["settings"]["ref_key_col"])
            self.v_filter_disp = tk.StringVar(value="")
            self.kat_interface = set(sc["kategorie_interface"])
            self.hme_types = dict(sc["hme_types"])
            self.outer_ignore = set(sc["group_outer_ignore"])
            self.inner_ignore = set(sc["group_inner_ignore"])
            self.tug_sel = {k: set(sc[f"tug_{k}_values"]) for k in "tgu"}
            self.filter_values = set(sc["filter_values"])
            self.v_tug_on = {k: tk.BooleanVar(value=sc["settings"][f"tug_{k}_on"] == "1") for k in "tgu"}
            self.v_source = tk.StringVar(value=sc["source"])
            self.db_vars = {k: tk.StringVar(value=sc["db"][k]) for k in DEFAULT_DB if k not in ("provider", "overwrite", "resolve_col",
                                                                                              "as_text")}
            self.db_vars["resolve_col"] = tk.StringVar(value=sc["db"]["resolve_col"] or AUTO_LABEL)
            self.v_overwrite = tk.BooleanVar(value=sc["db"]["overwrite"] == "1")
            self.v_astext = tk.BooleanVar(value=sc["db"]["as_text"] == "1")
            self.v_dburl = tk.StringVar(value="")
            self.v_provider = tk.StringVar(value=DB_PROVIDER_NAME.get(sc["db"]["provider"], "Airtable"))
            self._ssl_cbs, self.cb_dbtable, self.prov_frames = [], {}, {}
            self.v_secret = {k: tk.StringVar(value="") for k, _n in DB_PROVIDERS}
            self.v_savekeys = {k: tk.BooleanVar(value=False) for k, _n in DB_PROVIDERS}
            self.v_apikey, self.v_savekey = self.v_secret["airtable"], self.v_savekeys["airtable"]
            for k, _n in DB_PROVIDERS:
                try:
                    key = self.store.get_secret(k)
                except Exception:
                    key = ""
                if key:
                    self.v_secret[k].set(key)
                    self.v_savekeys[k].set(True)

            self._build()
            self.on_source(select=False)
            self._apply_geometry()
            try:  # zuletzt verwendete Tabelle wieder laden, falls sie noch existiert
                lf = [f for f in (last.get("files") or []) if Path(f).exists()]
                if last.get("out_dir") and Path(last["out_dir"]).is_dir():
                    self.v_out.set(last["out_dir"])
                if lf and self.v_source.get() == "file":
                    self.v_sheet.set(last.get("sheet", ""))
                    self.set_files(lf)
            except Exception:
                pass

            for v in list(self.map_vars.values()) + [self.v_suffix_on, self.v_ref]:
                v.trace_add("write", lambda *a: self.after_idle(self.refresh))
            self.v_sheet.trace_add("write", lambda *a: self.after_idle(self.on_sheet_changed))
            for k in ("delimiter", "encoding"):
                self.set_vars[k].trace_add("write", lambda *a: self.after_idle(self.reload_table))
            self.set_vars["id_prefix_regex"].trace_add("write", lambda *a: self.after_idle(self.refresh))
            self.v_filter_disp.trace_add("write", lambda *a: self.after_idle(self.on_filter_col))

            self.refresh()
            self.protocol("WM_DELETE_WINDOW", self.on_close)

        def _apply_geometry(self):
            """Fenster zentriert öffnen, so groß, dass die Tabs 'Datenbank' und 'Spalten' ohne Scrollen passen
            (begrenzt auf den sichtbaren Bildschirm)."""
            try:
                self.update_idletasks()
                sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
                kids = self.nb.winfo_children()
                header = self.nb.winfo_reqheight() - max(k.winfo_reqheight() for k in kids)
                chrome = self.winfo_reqheight() - self.nb.winfo_reqheight()
                content_h = max(self.sf_cols.inner.winfo_reqheight(), self.tab_db.winfo_reqheight())
                content_w = max(self.sf_cols.inner.winfo_reqwidth(), self.tab_db.winfo_reqwidth())
                w = max(900, content_w + 60)
                h = max(640, chrome + header + content_h + 24)
                w, h = min(w, sw - 20), min(h, sh - 80)
                self.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h - 40) // 2)}")
            except Exception:
                self.geometry("980x720+60+40")
            try:
                self.state("normal")
                self.lift()
                self.focus_force()
            except Exception:
                pass

        # ---------- Aufbau ----------
        def _build(self):
            pad = {"padx": 6, "pady": 2}
            f1 = ttk.LabelFrame(self, text=tr("Quelle und Speicherort"))
            f1.pack(fill="x", padx=10, pady=(10, 4))
            f1.columnconfigure(1, weight=1)
            rs = ttk.Frame(f1)
            rs.grid(row=1, column=0, sticky="w", padx=6, pady=3)
            ttk.Radiobutton(rs, text=tr("Datenbank / Online"), variable=self.v_source, value="db",
                            command=self.on_source).pack(side="left")
            ttk.Radiobutton(rs, text=tr("Datei (Tabelle / Datenbank):"), variable=self.v_source, value="file",
                            command=self.on_source).pack(side="left", padx=(12, 0))
            self.ent_file = ttk.Entry(f1, textvariable=self.v_file, state="readonly")
            self.ent_file.grid(row=1, column=1, sticky="ew", **pad)
            sh = ttk.Frame(f1)
            sh.grid(row=1, column=2, **pad)
            ttk.Label(sh, text=tr("Blatt / Tabelle:")).pack(side="left")
            self.cb_sheet = ttk.Combobox(sh, textvariable=self.v_sheet, state="disabled", width=18)
            self.cb_sheet.pack(side="left", padx=(4, 0))
            self.btn_pick = ttk.Button(f1, text=tr("Datei wählen…"), command=self.pick_files)
            self.btn_pick.grid(row=1, column=3, **pad)
            self.lbl_fileinfo = ttk.Label(f1, text="", foreground=C_MUTED)
            self.lbl_fileinfo.grid(row=2, column=1, sticky="w", padx=6)
            ttk.Label(f1, text=tr("Speicherort:")).grid(row=3, column=0, sticky="w", **pad)
            ttk.Entry(f1, textvariable=self.v_out).grid(row=3, column=1, columnspan=2, sticky="ew", **pad)
            ttk.Button(f1, text=tr("Ordner wählen…"), command=self.pick_out).grid(row=3, column=3, **pad)
            self.lbl_outhint = ttk.Label(f1, text=tr("leer = neben der Tabelle"), foreground=C_MUTED)
            self.lbl_outhint.grid(row=4, column=1, sticky="w", padx=6)

            f2 = ttk.LabelFrame(self, text=tr("Mapping-Schema"))
            f2.pack(fill="x", padx=10, pady=4)
            f2.columnconfigure(1, weight=1)
            ttk.Label(f2, text=tr("Schema:")).grid(row=0, column=0, sticky="w", **pad)
            self.cb_schema = ttk.Combobox(f2, textvariable=self.v_schema, state="readonly")
            self.cb_schema.grid(row=0, column=1, sticky="ew", **pad)
            self.cb_schema.bind("<<ComboboxSelected>>", lambda e: self.load_schema(self.v_schema.get()))
            bf = ttk.Frame(f2)
            bf.grid(row=0, column=2, sticky="e")
            for txt, cmd in ((tr("Speichern"), self.save_schema), (tr("Speichern unter…"), self.save_schema_as),
                             (tr("Löschen"), self.delete_schema), (tr("Export…"), self.export_schema),
                             (tr("Import…"), self.import_schema)):
                ttk.Button(bf, text=txt, command=cmd).pack(side="left", padx=2, pady=2)
            ttk.Label(f2, text=tr('Gespeichert in: {0}  ·  der zuletzt verwendete Stand wird beim Beenden automatisch gemerkt').format(self.store.path), foreground=C_MUTED).grid(
                row=1, column=0, columnspan=2, sticky="w", padx=6)
            self._lang_picker(f2).grid(row=1, column=2, sticky="e", padx=6)

            try:
                st = ttk.Style(self)
                from tkinter import font as tkfont
                base = tkfont.nametofont("TkDefaultFont")
                tab_font = tkfont.Font(family=base.actual("family"), size=base.actual("size") + 1, weight="bold")
                self._tab_font = tab_font
                st.configure("TNotebook.Tab", font=tab_font, padding=(12, 5))
            except Exception:
                pass
            bottom = ttk.Frame(self)  # zuerst packen, damit die Buttonleiste immer am unteren Rand sichtbar bleibt
            bottom.pack(side="bottom", fill="x", padx=10, pady=(2, 10))
            self.nb = ttk.Notebook(self)
            self.nb.pack(fill="both", expand=True, padx=10, pady=4)
            self._tab_db()
            self._tab_columns()
            self._tab_values()
            self._tab_preview()
            self._tab_graph()
            self._tab_advanced()
            self._tab_log()

            self.btn_run = ttk.Button(bottom, text=tr("Matrix erzeugen"), command=self.run)
            self.btn_run.pack(side="left")
            self.btn_open_file = ttk.Button(bottom, text=tr("Datei öffnen"), command=self.open_file, state="disabled")
            self.btn_open_file.pack(side="left", padx=(10, 4))
            self.btn_open_dir = ttk.Button(bottom, text=tr("Ordner öffnen"), command=self.open_dir, state="disabled")
            self.btn_open_dir.pack(side="left", padx=4)
            self.status = ttk.Label(bottom, text="", foreground=C_MUTED)
            self.status.pack(side="left", padx=12)
            fmt = ttk.Frame(bottom)
            fmt.pack(side="right")
            ttk.Label(fmt, text=tr("Ausgabe:")).pack(side="left", padx=(0, 6))
            ttk.Radiobutton(fmt, text=tr("HMC – Harris Matrix Composer"), variable=self.set_vars["out_format"],
                            value="hmc").pack(side="left")
            ttk.Radiobutton(fmt, text=tr("HME – Harris Matrix Editor"), variable=self.set_vars["out_format"],
                            value="hme").pack(side="left", padx=(10, 0))
            self.set_vars["out_format"].trace_add("write", lambda *a: self.after_idle(self._on_format))
            self._apply_format_state()

        def _lang_picker(self, parent):
            """Kippschalter EN | GER; die Wahl wird gespeichert und gilt ab dem nächsten Start."""
            try:
                bg = ttk.Style(self).lookup("TFrame", "background") or "#f0f0f0"
            except Exception:
                bg = "#f0f0f0"
            fr = ttk.Frame(parent)
            self.v_lang = tk.StringVar(value=UI_LANG)
            w, h = 46, 22
            lb_en = ttk.Label(fr, text="EN", cursor="hand2")
            lb_de = ttk.Label(fr, text="GER", cursor="hand2")
            sw = tk.Canvas(fr, width=w, height=h, bg=bg, highlightthickness=0, cursor="hand2")
            lb_en.pack(side="left", padx=(0, 5))
            sw.pack(side="left")
            lb_de.pack(side="left", padx=(5, 0))

            def draw():
                de = self.v_lang.get() == "de"
                sw.delete("all")
                col = "#0067c0"  # Blau wie die Kontrollkästchen
                sw.create_oval(1, 1, h - 1, h - 1, fill=col, outline=col)
                sw.create_oval(w - h + 1, 1, w - 1, h - 1, fill=col, outline=col)
                sw.create_rectangle(h // 2, 1, w - h // 2, h - 1, fill=col, outline=col)
                x0 = w - h + 4 if de else 4
                sw.create_oval(x0, 4, x0 + h - 8, h - 4, fill="white", outline="white")
                lb_en.configure(font=("TkDefaultFont", 9, "normal" if de else "bold"))
                lb_de.configure(font=("TkDefaultFont", 9, "bold" if de else "normal"))

            def choose(code):
                if code != self.v_lang.get():
                    self.v_lang.set(code)
                    draw()
                    self.on_lang(code)

            sw.bind("<Button-1>", lambda e: choose("en" if self.v_lang.get() == "de" else "de"))
            lb_en.bind("<Button-1>", lambda e: choose("en"))
            lb_de.bind("<Button-1>", lambda e: choose("de"))
            draw()
            return fr

        def on_lang(self, code):
            self.store.data.setdefault("ui", {})["language"] = code
            self.store.save()
            messagebox.showinfo(APP_SHORT, "Die Sprache wird beim nächsten Start von HMA übernommen.\n"
                                           "The language will be applied the next time HMA is started.")

        def _row(self, parent, r, label, key, tip, required=False, target=""):
            """Zeile mit Auswahlfeld in Grid-Zeile r; darunter (Zeile r+1) die Beispielwerte der Spalte."""
            lbl = ttk.Label(parent, text=label + (" *" if required else ""))
            lbl.grid(row=r, column=0, sticky="w", padx=6, pady=(2, 0))
            Tooltip(lbl, tip)
            self.rowlbls[key] = lbl
            cb = SearchCombobox(parent, self.map_vars[key], width=34)
            cb.grid(row=r, column=1, sticky="w", padx=6, pady=(2, 0))
            stat = ttk.Label(parent, text="", width=24)
            stat.grid(row=r, column=2, sticky="w", padx=4, pady=(2, 0))
            if target:
                tl = ttk.Label(parent, text="→ " + target, foreground=C_MUTED)
                tl.grid(row=r, column=3, sticky="w", padx=4, pady=(2, 0))
                self.tlbls[key] = (tl, target)
            ex = ttk.Label(parent, text="", foreground=C_MUTED, font=("TkDefaultFont", 8))
            ex.grid(row=r + 1, column=1, columnspan=3, sticky="w", padx=8, pady=(0, 2))
            self.cbs[key] = cb
            self.stats[key] = stat
            self.exlbls[key] = ex
            return cb

        SSL_LABELS = {"prefer": tr("bevorzugt (empfohlen)"), "require": tr("erforderlich"), "disable": tr("aus"),
                      "verify-full": tr("erforderlich, Zertifikat prüfen")}

        def _prov_key(self):
            name = self.v_provider.get()
            return next((k for k, v in DB_PROVIDERS if v == name), "airtable")

        def _tab_db(self):
            tab = ttk.Frame(self.nb)
            self.tab_db = tab
            self.nb.add(tab, text=tr("Datenbank"))
            pad = {"padx": 6, "pady": 3}
            top = ttk.LabelFrame(tab, text=tr("Datenquelle"))
            top.pack(fill="x", padx=8, pady=(8, 2))
            ttk.Label(top, text=tr("Anbieter:")).pack(side="left", **pad)
            cb = ttk.Combobox(top, textvariable=self.v_provider, state="readonly", width=26,
                              values=[n for _k, n in DB_PROVIDERS])
            cb.pack(side="left", **pad)
            cb.bind("<<ComboboxSelected>>", lambda _e: self._prov_changed())
            ttk.Label(top, text=tr("Alle Anbieter werden nur gelesen – HMA verändert nie etwas in der Datenbank."),
                      foreground=C_MUTED).pack(side="left", padx=10)
            self.db_holder = ttk.Frame(tab)
            self.db_holder.pack(fill="x")
            for key, _n in DB_PROVIDERS:
                fr = ttk.LabelFrame(self.db_holder, text=self._prov_title(key))
                fr.columnconfigure(1, weight=1)
                self.prov_frames[key] = fr
                getattr(self, "_prov_" + key)(fr)
            common = ttk.LabelFrame(tab, text=tr("Ausgabe"))
            common.pack(fill="x", padx=8, pady=4)
            common.columnconfigure(1, weight=1)
            r = 0
            ttk.Label(common, text=tr("Dateiname (ohne Endung):")).grid(row=r, column=0, sticky="w", **pad)
            ttk.Entry(common, textvariable=self.db_vars["out_name"]).grid(row=r, column=1, sticky="ew", **pad)
            ttk.Label(common, text=tr("Ordner: siehe 'Speicherort' oben (leer = Dokumente)"),
                      foreground=C_MUTED).grid(row=r, column=2, sticky="w", padx=6)
            r += 1
            ttk.Checkbutton(common, text=tr("Vorhandene Datei ohne Nachfrage überschreiben"),
                            variable=self.v_overwrite).grid(row=r, column=1, columnspan=2, sticky="w", **pad)
            r += 1
            bf = ttk.Frame(common)
            bf.grid(row=r, column=0, columnspan=3, sticky="w", padx=6, pady=(8, 6))
            ttk.Button(bf, text=tr("Verbindung testen / Daten abrufen"), command=self.db_fetch).pack(side="left")
            self.lbl_dbstatus = ttk.Label(bf, text="", foreground=C_MUTED)
            self.lbl_dbstatus.pack(side="left", padx=10)
            ttk.Label(tab, text=tr("Beim Klick auf 'Matrix erzeugen' werden die Daten immer frisch aus der Datenquelle "
                                "abgerufen. 'Daten abrufen' lädt sie vorab, damit du im Tab 'Spalten' die "
                                "Zuordnung treffen kannst."), foreground=C_MUTED, wraplength=820,
                      justify="left").pack(anchor="w", padx=14, pady=(0, 8))
            self._show_provider()

        @staticmethod
        def _prov_title(key):
            return {"airtable": tr("Airtable-Verbindung"), "postgres": tr("PostgreSQL-Server"), "mysql": tr("MySQL-/MariaDB-Server"),
                    "filemaker": tr("FileMaker-Server (Data API)"), "gsheets": "Google Sheets"}[key]

        def _show_provider(self):
            cur = self._prov_key()
            for k, fr in self.prov_frames.items():
                if k == cur:
                    fr.pack(fill="x", padx=8, pady=4)
                else:
                    fr.pack_forget()

        def _prov_changed(self):
            self.db_vars["table"].set("")
            self._show_provider()
            self.lbl_dbstatus.configure(text="")
            if self.v_source.get() == "db":
                self.on_source(select=False)

        def _secret_rows(self, box, r, prov, label, hint, extra_note=""):
            pad = {"padx": 6, "pady": 3}
            ttk.Label(box, text=label + ":").grid(row=r, column=0, sticky="w", **pad)
            ttk.Entry(box, textvariable=self.v_secret[prov], show="•").grid(row=r, column=1, sticky="ew", **pad)
            ttk.Label(box, text=hint, foreground=C_MUTED).grid(row=r, column=2, sticky="w", padx=6)
            r += 1
            cbk = ttk.Checkbutton(box, text=tr("Speichern: verschlüsselt auf diesem Computer (nur für dieses "
                                            "Windows-Benutzerkonto lesbar, nicht im Schema)"),
                                  variable=self.v_savekeys[prov])
            cbk.grid(row=r, column=1, columnspan=2, sticky="w", **pad)
            if not secret_available():
                self.v_savekeys[prov].set(False)
                cbk.configure(state="disabled", text=tr("Speichern ist nur unter Windows möglich (wird für diese "
                                                     "Sitzung gemerkt)"))
            return r + 1

        def _entry_rows(self, box, r, rows):
            pad = {"padx": 6, "pady": 3}
            for key, label, hint in rows:
                ttk.Label(box, text=label + ":").grid(row=r, column=0, sticky="w", **pad)
                ttk.Entry(box, textvariable=self.db_vars[key]).grid(row=r, column=1, sticky="ew", **pad)
                ttk.Label(box, text=hint, foreground=C_MUTED).grid(row=r, column=2, sticky="w", padx=6)
                r += 1
            return r

        def _table_row(self, box, r, label, btn_text):
            pad = {"padx": 6, "pady": 3}
            ttk.Label(box, text=label + ":").grid(row=r, column=0, sticky="w", **pad)
            cb = ttk.Combobox(box, textvariable=self.db_vars["table"], width=40)
            cb.grid(row=r, column=1, sticky="w", **pad)
            self.cb_dbtable[box] = cb
            ttk.Button(box, text=btn_text, command=lambda: self.db_list(cb)).grid(row=r, column=2, sticky="w", padx=6)
            return r + 1

        def _ssl_row(self, box, r, with_verify):
            pad = {"padx": 6, "pady": 3}
            ttk.Label(box, text=tr("Verschlüsselung:")).grid(row=r, column=0, sticky="w", **pad)
            keys = ["prefer", "require", "disable"] + (["verify-full"] if with_verify else [])
            cb = ttk.Combobox(box, state="readonly", width=34, values=[self.SSL_LABELS[k] for k in keys])
            cb.grid(row=r, column=1, sticky="w", **pad)
            cb.set(self.SSL_LABELS.get(self.db_vars["sslmode"].get(), self.SSL_LABELS["prefer"]))

            def sel(_e=None, cb=cb, keys=keys):
                self.db_vars["sslmode"].set(next((k for k in keys if self.SSL_LABELS[k] == cb.get()), "prefer"))
            cb.bind("<<ComboboxSelected>>", sel)
            self._ssl_cbs.append((cb, keys))
            ttk.Label(box, text=tr("SSL/TLS; 'bevorzugt' verschlüsselt, wenn der Server es anbietet"),
                      foreground=C_MUTED).grid(row=r, column=2, sticky="w", padx=6)
            return r + 1

        def _prov_airtable(self, box):
            pad = {"padx": 6, "pady": 3}
            r = 0
            ttk.Label(box, text=tr("Airtable-Link einfügen:")).grid(row=r, column=0, sticky="w", **pad)
            ttk.Entry(box, textvariable=self.v_dburl).grid(row=r, column=1, sticky="ew", **pad)
            ttk.Button(box, text=tr("Übernehmen"), command=self.db_parse_url).grid(row=r, column=2, **pad)
            r += 1
            ttk.Label(box, text=tr("Adresse aus dem Browser kopieren (…/appXXX/tblXXX/viwXXX): Base, Tabelle und View "
                                "werden erkannt."), foreground=C_MUTED).grid(row=r, column=1, sticky="w", padx=6)
            r += 1
            r = self._entry_rows(box, r, (
                ("base_id", "Base ID", tr("beginnt mit 'app'")),
                ("table_id", "Table ID", tr("beginnt mit 'tbl' (oder Tabellenname)")),
                ("view_id", tr("View ID"), tr("optional, beginnt mit 'viw'; leer = alle Datensätze der Tabelle")),
                ("fields", tr("Felder"), tr("optional, Namen kommagetrennt wie in Airtable; leer = alle Felder"))))
            r = self._secret_rows(box, r, "airtable", tr("API-Schlüssel"), "Personal Access Token")
            ttk.Label(box, text=tr("Token erstellen: airtable.com/create/tokens, Berechtigung 'data.records:read', "
                                "Zugriff auf die Base geben. Der Token wird nur einmal angezeigt."),
                      foreground=C_MUTED).grid(row=r, column=1, columnspan=2, sticky="w", padx=6)
            r += 1
            ttk.Separator(box).grid(row=r, column=0, columnspan=3, sticky="ew", pady=6)
            r += 1
            ttk.Checkbutton(box, text=tr("Verknüpfungen als Text abrufen (Primärfeld der verknüpften Tabelle statt "
                                      "recXXXX)"), variable=self.v_astext).grid(row=r, column=1, columnspan=2,
                                                                                sticky="w", **pad)
            r += 1
            ttk.Label(box, text=tr("Alle Werte kommen dann so an, wie Airtable sie anzeigt. Andere Spalten der "
                                "verknüpften Tabelle: Lookup-Feld in Airtable anlegen. Ist das Häkchen aus, "
                                "werden Verknüpfungen nur über die Auflösungs-Spalte (unten) ersetzt."),
                      foreground=C_MUTED, wraplength=760, justify="left").grid(row=r, column=1, columnspan=2,
                                                                               sticky="w", padx=6)
            r += 1
            ttk.Label(box, text=tr("Verknüpfungen auflösen mit:")).grid(row=r, column=0, sticky="w", **pad)
            self.cb_resolve = SearchCombobox(box, self.db_vars["resolve_col"], width=34)
            self.cb_resolve.grid(row=r, column=1, sticky="w", **pad)
            self.cb_resolve.set_values([AUTO_LABEL])
            r += 1
            ttk.Label(box, text=tr("Airtable liefert bei verknüpften Feldern (z. B. 'liegt über') interne Kennungen "
                                "(recXXXX…). HMA ersetzt sie durch den Wert dieser Spalte. Automatisch = die im Tab "
                                "'Spalten' zugeordnete ID (bzw. 'Verweise bestehen aus'). Das funktioniert bei "
                                "Verknüpfungen innerhalb derselben Tabelle."), foreground=C_MUTED,
                      wraplength=760, justify="left").grid(row=r, column=1, columnspan=2, sticky="w", padx=6)

        def _prov_sql(self, box, prov, with_verify):
            r = self._entry_rows(box, 0, (
                ("host", "Server", tr("Name oder IP-Adresse, z. B. db.beispiel.at")),
                ("port", tr("Port"), tr("leer = Standard (") + DB_DEFAULT_PORT[prov] + ")"),
                ("dbname", tr("Datenbank"), tr("Name der Datenbank")),
                ("user", tr("Benutzer"), tr("am besten ein Benutzer mit reiner Leseberechtigung"))))
            r = self._secret_rows(box, r, prov, tr("Passwort"), "")
            r = self._ssl_row(box, r, with_verify)
            r = self._table_row(box, r, tr("Tabelle / View"), tr("Tabellen laden"))
            ttk.Label(box, text=tr("HMA liest ausschließlich (schreibgeschützte Sitzung) und liest die gewählte Tabelle "
                                "bzw. View vollständig. Für Auswertungen über mehrere Tabellen am besten in der "
                                "Datenbank eine View anlegen und diese wählen."), foreground=C_MUTED,
                      wraplength=760, justify="left").grid(row=r, column=1, columnspan=2, sticky="w", padx=6, pady=(0, 4))

        def _prov_postgres(self, box):
            self._prov_sql(box, "postgres", True)

        def _prov_mysql(self, box):
            self._prov_sql(box, "mysql", False)
            ttk.Label(box, text=tr("Hinweis: MySQL/MariaDB-Anbindung bisher nur gegen einen Nachbau getestet, nicht gegen "
                                "einen echten Server."), foreground=C_BAD).grid(
                row=9, column=1, columnspan=2, sticky="w", padx=6, pady=(0, 4))

        def _prov_filemaker(self, box):
            pad = {"padx": 6, "pady": 3}
            r = self._entry_rows(box, 0, (
                ("host", "Server", tr("z. B. fm.beispiel.at (FileMaker Server mit aktivierter Data API)")),
                ("port", tr("Port"), tr("leer = Standard (443)")),
                ("dbname", tr("Datenbank"), tr("Name der FileMaker-Datei, ohne .fmp12")),
                ("user", tr("Benutzer"), tr("Konto mit dem Zugriffsrecht 'fmrest'"))))
            r = self._secret_rows(box, r, "filemaker", tr("Passwort"), "")
            r = self._table_row(box, r, "Layout", tr("Layouts laden"))
            self.v_fm_noverify = tk.BooleanVar(value=self.db_vars["tls_verify"].get() == "0")

            def sync(*_a):
                self.db_vars["tls_verify"].set("0" if self.v_fm_noverify.get() else "1")
            ttk.Checkbutton(box, text=tr("Zertifikat des Servers nicht prüfen (bei selbstsigniertem Zertifikat)"),
                            variable=self.v_fm_noverify, command=sync).grid(row=r, column=1, columnspan=2, sticky="w", **pad)
            r += 1
            ttk.Label(box, text=tr("Gelesen werden die Felder des gewählten Layouts (alle Datensätze). Die Data API muss am "
                                "Server aktiviert sein. Lokale FileMaker-Dateien (.fmp12) können nicht direkt gelesen "
                                "werden – dort die Tabelle als CSV/Excel exportieren oder die Datei über FileMaker "
                                "Server freigeben. Hinweis: bisher nur gegen einen Nachbau der Schnittstelle getestet."),
                      foreground=C_MUTED, wraplength=760, justify="left").grid(row=r, column=1, columnspan=2,
                                                                               sticky="w", padx=6, pady=(0, 4))

        def _prov_gsheets(self, box):
            r = self._entry_rows(box, 0, (("gs_url", "Link", tr("Adresse der Tabelle aus dem Browser kopieren")),))
            ttk.Label(box, text=tr("Die Tabelle muss für 'Jeden mit dem Link' lesbar sein (Teilen → Allgemeiner Zugriff → "
                                "Betrachter) oder im Web veröffentlicht sein. Ist die Adresse mit '#gid=…' kopiert, "
                                "wird genau dieses Blatt gelesen, sonst das erste. Private Tabellen kann HMA nicht "
                                "abrufen – dort bitte als .xlsx oder .csv herunterladen."), foreground=C_MUTED, wraplength=760,
                      justify="left").grid(row=r, column=1, columnspan=2, sticky="w", padx=6, pady=(0, 4))

        def db_list(self, combo):
            prov = self._prov_key()
            schema = normalize_schema(self.gather_schema())
            self.lbl_dbstatus.configure(text=tr("Lese Liste …"))
            self.update_idletasks()
            try:
                names = db_list_tables(schema["db"], self.v_secret[prov].get())
            except HMAError as e:
                self.lbl_dbstatus.configure(text=tr("Fehler"))
                messagebox.showerror(tr("Datenbank"), str(e))
                return
            except Exception:
                self.lbl_dbstatus.configure(text=tr("Fehler"))
                messagebox.showerror(tr("Datenbank"), traceback.format_exc())
                return
            combo.configure(values=names)
            self.persist_key()
            self.lbl_dbstatus.configure(text=tr('{0} gefunden – bitte aus der Liste wählen').format(len(names)))
            if names and not self.db_vars["table"].get().strip():
                self.db_vars["table"].set(names[0])

        def _tab_columns(self):
            sf = ScrollFrame(self.nb)
            self.sf_cols = sf
            self.nb.add(sf, text=tr("Spalten"))
            t = sf.inner
            self.cbs, self.stats, self.exlbls, self.rowlbls, self.tlbls = {}, {}, {}, {}, {}
            bar = ttk.Frame(t)
            bar.pack(fill="x", padx=6, pady=(6, 2))
            ttk.Button(bar, text=tr("Spalten automatisch erkennen"), command=self.auto_suggest).pack(side="left")
            ttk.Button(bar, text=tr("Zuordnung leeren"), command=self.clear_mapping).pack(side="left", padx=6)
            self.lbl_csvinfo = ttk.Label(bar, text=tr("Noch keine Datei geladen."), foreground=C_MUTED)
            self.lbl_csvinfo.pack(side="left", padx=8)
            self.lbl_multi = ttk.Label(t, text="", foreground=C_BAD)
            self.lbl_multi.pack(anchor="w", padx=8)
            ttk.Label(t, text=tr("Tipps: In die Felder tippen filtert die Spaltenliste (z. B. \"liegt\"), dann Enter oder Pfeil-runter. "
                           "Mauszeiger über eine Bezeichnung = Erklärung."),
                      foreground=C_MUTED).pack(anchor="w", padx=8)

            fk = ttk.LabelFrame(t, text=tr("Knoten"))
            fk.pack(fill="x", padx=6, pady=4)
            lbl = ttk.Label(fk, text="ID *")
            lbl.grid(row=0, column=0, sticky="w", padx=6, pady=2)
            Tooltip(lbl, tr("Eindeutige ID der Stratigraphischen Einheit. Bis zu vier Spalten werden mit _ "
                         "verbunden (z. B. SE-Nr + Typ + Datierung → 125_Vf_MA). Teil 1 ist Pflicht; Teil 2 "
                         "wird erst nach Teil 1 freigeschaltet usw."))
            idf = ttk.Frame(fk)
            idf.grid(row=0, column=1, columnspan=3, sticky="w", padx=6, pady=2)
            for i, key in enumerate(ID_KEYS):
                cb = SearchCombobox(idf, self.map_vars[key], width=16)
                cb.pack(side="left")
                self.cbs[key] = cb
                if i < 3:
                    ttk.Label(idf, text="_").pack(side="left", padx=2)
            self.lbl_idex = ttk.Label(fk, text="", foreground=C_MUTED)
            self.lbl_idex.grid(row=1, column=1, columnspan=3, sticky="w", padx=6)
            self._row(fk, 2, "Name", "name", tr("Anzeigename des Knotens, z. B. der SE-Typ."), target=tr("Node-Name"))
            self._row(fk, 4, tr("Funde / Beschreibung"), "descr",
                      tr("Der Inhalt wird zur Beschreibung (description) des Knotens."))
            sfx = ttk.Frame(fk)
            sfx.grid(row=4, column=3, sticky="w", padx=4)
            ttk.Checkbutton(sfx, text=tr("Suffix an ID, wenn gefüllt:"), variable=self.v_suffix_on).pack(side="left")
            self.ent_suffix = ttk.Entry(sfx, textvariable=self.set_vars["descr_suffix"], width=10)
            self.ent_suffix.pack(side="left", padx=4)
            self._row(fk, 6, tr("Kategorie"), "kategorie",
                      tr("Welche Werte Interface (SURFACE) sind, wählst du im Tab 'Werte'. Alle anderen → DEPOSIT."),
                      target=tr("Node-Type"))

            fg = ttk.LabelFrame(t, text=tr("Gruppierung"))
            fg.pack(fill="x", padx=6, pady=4)
            self._row(fg, 0, tr("Übergeordnete Gruppe"), "group_outer",
                      tr("Äußere Gruppe, z. B. Periode oder Phase. Einzelne Werte (z. B. 'unklar') "
                      "kannst du im Tab 'Werte' ignorieren."), target="PERIOD_GROUP")
            self._row(fg, 2, tr("Untergeordnete Gruppe"), "group_inner",
                      tr("Innere Gruppe, z. B. Phase oder Befund; liegt in der übergeordneten Gruppe."),
                      target="PHASE_GROUP")

            fb = ttk.LabelFrame(t, text=tr("Stratigraphische Beziehungen"))
            fb.pack(fill="x", padx=6, pady=4)
            self._row(fb, 0, tr("liegt über"), "above",
                      tr("Mehrere Verweise in einer Zelle, getrennt durch , oder ;."), required=True,
                      target=tr("ABOVE-Kante"))
            self._row(fb, 2, tr("liegt unter"), "below",
                      tr("Optional, umgekehrte Richtung: 'liegt unter X' ⇒ X liegt über diese SE."),
                      target=tr("ABOVE-Kante"))
            self._row(fb, 4, tr("zeitgleich mit"), "contemp", "", target=tr("CONTEMPORARY-Kante"))
            self._row(fb, 6, tr("ident mit"), "ident",
                      tr("Identische SEs werden zu einem Node zusammengeführt (z. B. 3=45_Ps). Nicht dieselbe Spalte "
                      "wie 'zeitgleich mit' verwenden!"), target=tr("Node-Merge"))
            lbl = ttk.Label(fb, text=tr("Verweise bestehen aus"))
            lbl.grid(row=8, column=0, sticky="w", padx=6, pady=(6, 2))
            Tooltip(lbl, tr("Standard: die oben zusammengesetzte ID. Enthalten die Beziehungs-Spalten nur z. B. die "
                         "SE-Nummer, hier diese Spalte wählen (muss innerhalb der Tabelle eindeutig sein)."))
            self.cb_ref = SearchCombobox(fb, self.v_ref, width=34)
            self.cb_ref.grid(row=8, column=1, sticky="w", padx=6, pady=(6, 2))
            ttk.Label(t, text=tr("* erforderlich"), foreground=C_MUTED).pack(anchor="w", padx=8, pady=(0, 8))

        def _tab_values(self):
            sf = ScrollFrame(self.nb)
            self.nb.add(sf, text=tr("Werte"))
            t = sf.inner
            self.fr_kat = ttk.LabelFrame(t, text=tr("Kategorie: welche Werte sind ein Interface (SURFACE)?"))
            self.fr_kat.pack(fill="x", padx=6, pady=6)
            self.fr_outer = ttk.LabelFrame(t, text=tr("Übergeordnete Gruppe: welche Werte werden ignoriert?"))
            self.fr_outer.pack(fill="x", padx=6, pady=6)
            self.fr_inner = ttk.LabelFrame(t, text=tr("Untergeordnete Gruppe: welche Werte werden ignoriert?"))
            self.fr_inner.pack(fill="x", padx=6, pady=6)
            ff = ttk.LabelFrame(t, text=tr("Zeilenfilter: nur bestimmte Zeilen verwenden"))
            ff.pack(fill="x", padx=6, pady=6)
            top = ttk.Frame(ff)
            top.pack(fill="x", padx=6, pady=4)
            ttk.Label(top, text=tr("Filtern nach Spalte:")).pack(side="left")
            self.cb_filter = SearchCombobox(top, self.v_filter_disp, width=40)
            self.cb_filter.pack(side="left", padx=6)
            ttk.Label(ff, text=tr("Enthält die Tabelle mehrere Grabungen/Projekte, genau EIN Projekt anhaken "
                               "(sonst sind IDs und Verweise nicht eindeutig). Nichts angehakt = alle Zeilen."),
                      foreground=C_MUTED, wraplength=820, justify="left").pack(anchor="w", padx=8)
            self.fr_filter = ttk.Frame(ff)
            self.fr_filter.pack(fill="x")

            ft = ttk.LabelFrame(t, text=tr("Optional: T, G und U einfügen (Top surface, Interface to Geology, Unexcavated)"))
            ft.pack(fill="x", padx=6, pady=6)
            self.fr_tug = ft
            ttk.Label(ft, text=tr("Standardmäßig aus. Ohne Auswahl von Werten gilt: T über U und U über G."),
                      foreground=C_MUTED).pack(anchor="w", padx=8, pady=(4, 2))
            self.tug_picks = {}
            spec = {"t": (tr("Top surface (T) einfügen"), tr("Direkt unter T liegen diese SEs (ID wie in der Knoten-Gruppe):")),
                    "g": (tr("Interface to Geology (G) einfügen"), tr("Direkt über G liegen diese SEs:")),
                    "u": (tr("Unexcavated (U) einfügen"), tr("Direkt über U liegen diese SEs:"))}
            for k in "tgu":
                box = ttk.Frame(ft)
                box.pack(fill="x", padx=8, pady=4)
                ttk.Checkbutton(box, text=spec[k][0], variable=self.v_tug_on[k],
                                command=self._tug_state).pack(anchor="w")
                ttk.Label(box, text=spec[k][1], foreground=C_MUTED).pack(anchor="w", padx=22)
                mp = MultiPick(box, self.tug_sel[k])
                mp.pack(anchor="w", padx=22, fill="x")
                self.tug_picks[k] = mp

        def _tab_preview(self):
            tab = ttk.Frame(self.nb)
            self.nb.add(tab, text=tr("Tabelle"))
            ttk.Label(tab, text=tr("Erste Zeilen der zugeordneten Spalten (die Matrix siehst du im Tab 'Voransicht')"), foreground=C_MUTED).pack(
                anchor="w", padx=8, pady=(6, 2))
            frame = ttk.Frame(tab)
            frame.pack(fill="both", expand=True, padx=8, pady=4)
            self.tree = ttk.Treeview(frame, show="headings")
            sy = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
            sx = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
            self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
            self.tree.grid(row=0, column=0, sticky="nsew")
            sy.grid(row=0, column=1, sticky="ns")
            sx.grid(row=1, column=0, sticky="ew")
            frame.rowconfigure(0, weight=1)
            frame.columnconfigure(0, weight=1)

        def _tab_graph(self):
            tab = ttk.Frame(self.nb)
            self.tab_graph = tab
            self.nb.add(tab, text=tr("Voransicht"))
            self._gr_scene, self._gr_sig, self._gr_zoom, self._gr_sel = None, None, 1.0, None
            self._gr_idx, self._gr_press, self._gr_list, self._gr_tip = {}, (0, 0), [], None
            self._gr_hover, self._gr_fonts, self._gr_note = None, {}, ""
            bar = ttk.Frame(tab)
            bar.pack(fill="x", padx=6, pady=(6, 2))
            btn = ttk.Button(bar, text=tr("Aktualisieren"), command=self.graph_refresh)
            btn.pack(side="left")
            Tooltip(btn, tr("Liest die Daten neu ein: bei 'Datenbank / Online' wird die Datenquelle neu abgerufen, bei einer Tabelle "
                         "wird die Datei neu gelesen. Danach wird die Voransicht neu berechnet."))
            ttk.Button(bar, text=tr("Alles anzeigen"), command=self.graph_fit).pack(side="left", padx=(8, 2))
            ttk.Button(bar, text="100 %", command=lambda: self.graph_zoom_to(1.0)).pack(side="left", padx=2)
            ttk.Button(bar, text="\u2212", width=3, command=lambda: self._gr_zoom_step(1 / 1.25)).pack(side="left", padx=(8, 1))
            ttk.Button(bar, text="+", width=3, command=lambda: self._gr_zoom_step(1.25)).pack(side="left", padx=1)
            ttk.Label(bar, text=tr("Anordnung:")).pack(side="left", padx=(12, 2))
            self.v_gr_mode = tk.StringVar(value=PV_MODES["auto"])
            cbm = ttk.Combobox(bar, textvariable=self.v_gr_mode, values=list(PV_MODES.values()), state="readonly",
                               width=16)
            cbm.pack(side="left")
            cbm.bind("<<ComboboxSelected>>", lambda e: self.graph_update(force=True))
            Tooltip(cbm, tr("Perioden-Blöcke: Die Perioden (HME: Phasen) liegen als Blöcke übereinander, innerhalb eines "
                         "Blocks zählt die stratigraphische Tiefe.\nStrikt nach Tiefe: alle Units nach ihrer Tiefe in der "
                         "gesamten Matrix.\nAutomatisch: bei HME Blöcke (wie der HME-Auto-Layout), bei HMC strikt nach Tiefe.\n"
                         "Die Voransicht dient nur der Kontrolle der Daten; die Anordnung im Zielprogramm kann abweichen."))
            self.v_gr_groups = tk.BooleanVar(value=True)
            self.v_gr_contemp = tk.BooleanVar(value=True)
            ttk.Checkbutton(bar, text=tr("Gruppen"), variable=self.v_gr_groups, command=self._gr_draw).pack(
                side="left", padx=(12, 2))
            ttk.Checkbutton(bar, text=tr("zeitgleich"), variable=self.v_gr_contemp, command=self._gr_draw).pack(
                side="left", padx=2)
            self.lbl_graph = ttk.Label(bar, text="", foreground=C_MUTED)
            self.lbl_graph.pack(side="left", padx=10)

            body = ttk.Frame(tab)
            body.pack(fill="both", expand=True, padx=6, pady=(2, 6))
            panel = ttk.Frame(body, width=310)
            panel.pack(side="right", fill="y", padx=(6, 0))
            panel.pack_propagate(False)
            frame = ttk.Frame(body)
            frame.pack(side="left", fill="both", expand=True)
            self.cv = tk.Canvas(frame, bg="white", highlightthickness=1, highlightbackground="#cccccc",
                                closeenough=3.0)
            sy = ttk.Scrollbar(frame, orient="vertical", command=self.cv.yview)
            sx = ttk.Scrollbar(frame, orient="horizontal", command=self.cv.xview)
            self.cv.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
            self.cv.grid(row=0, column=0, sticky="nsew")
            sy.grid(row=0, column=1, sticky="ns")
            sx.grid(row=1, column=0, sticky="ew")
            frame.rowconfigure(0, weight=1)
            frame.columnconfigure(0, weight=1)

            # Bedienfeld rechts: ID-Liste + Karteikarte
            ttk.Label(panel, text=tr("Units (Klick, Pfeiltasten oder Mausrad)")).pack(anchor="w")
            self.v_gr_q = tk.StringVar(value="")
            ent = ttk.Entry(panel, textvariable=self.v_gr_q)
            ent.pack(fill="x", pady=(2, 3))
            self.v_gr_q.trace_add("write", lambda *a: self._gr_populate_list())
            lf = ttk.Frame(panel)
            lf.pack(fill="x")
            self.lb = tk.Listbox(lf, height=11, exportselection=False, activestyle="none", selectmode="browse")
            lsb = ttk.Scrollbar(lf, orient="vertical", command=self.lb.yview)
            self.lb.configure(yscrollcommand=lsb.set)
            self.lb.pack(side="left", fill="x", expand=True)
            lsb.pack(side="left", fill="y")
            self.lb.bind("<<ListboxSelect>>", self._gr_list_select)
            card = ttk.LabelFrame(panel, text=tr("Karteikarte"))
            card.pack(fill="both", expand=True, pady=(8, 0))
            cf = ttk.Frame(card)
            cf.pack(fill="both", expand=True, padx=4, pady=4)
            self.card = tk.Text(cf, wrap="word", height=14, width=34, state="disabled", bg="#f2f2f2", relief="flat",
                                font=("TkDefaultFont", 9))
            csb = ttk.Scrollbar(cf, orient="vertical", command=self.card.yview)
            self.card.configure(yscrollcommand=csb.set)
            self.card.pack(side="left", fill="both", expand=True)
            csb.pack(side="left", fill="y")
            self.card.tag_configure("id", font=("TkDefaultFont", 11, "bold"), foreground="#1f4e79", spacing3=7)
            self.card.tag_configure("key", foreground=C_MUTED)
            self.card.tag_configure("val", spacing3=7)
            ttk.Label(panel, text=(tr("Scrollen = Verschieben · ⌘ + Scrollen = Zoom · Ziehen = Verschieben · Hover = Beziehungen")
                                   if sys.platform == "darwin" else
                                   tr("Mausrad = Zoom · Ziehen = Verschieben · Hover = Beziehungen")),
                      foreground=C_MUTED, wraplength=300, justify="left").pack(anchor="w", pady=(4, 0))

            self.cv.bind("<MouseWheel>", self._gr_wheel)
            self.cv.bind("<Button-4>", self._gr_wheel)
            self.cv.bind("<Button-5>", self._gr_wheel)
            self.cv.bind("<ButtonPress-1>", self._gr_press_ev)
            self.cv.bind("<B1-Motion>", lambda e: (self._gr_tip_hide(), self.cv.scan_dragto(e.x, e.y, gain=1)))
            self.cv.tag_bind("node", "<ButtonRelease-1>", self._gr_click)
            self.cv.tag_bind("node", "<Enter>", self._gr_node_enter)
            self.cv.tag_bind("node", "<Motion>", self._gr_tip_move)
            self.cv.tag_bind("node", "<Leave>", lambda e: self._gr_tip_hide())
            self.cv.tag_bind("edge", "<Enter>", self._gr_edge_enter)
            self.cv.tag_bind("edge", "<Motion>", self._gr_tip_move)
            self.cv.tag_bind("edge", "<Leave>", self._gr_edge_leave)
            self.nb.bind("<<NotebookTabChanged>>", lambda e: self._on_tab_changed(), add="+")

        def _on_tab_changed(self):
            try:
                if self.nb.select() == str(self.tab_graph):
                    if self._files_changed():  # Datei (z. B. Access-Datenbank) wurde zwischenzeitlich gespeichert
                        self._gr_note = tr(" · Datei hat sich geändert, neu eingelesen ") + datetime.now().strftime("%H:%M:%S")
                        self.reload_table(first_load=False)
                        self.graph_update(force=True)
                    else:
                        self.graph_update()
            except Exception:
                pass

        def _gr_message(self, text, color=C_MUTED):
            self._gr_scene = None
            self._gr_populate_list()
            self._gr_card(None)
            self.cv.delete("all")
            self.cv.create_text(20, 20, anchor="nw", text=text, fill=color, width=700, font=("TkDefaultFont", 10))
            self.cv.configure(scrollregion=(0, 0, 800, 400))

        def _gr_signature(self):
            try:
                return json.dumps([self.gather_schema(), id(self.rows), len(self.rows), len(self.headers),
                                   self.v_gr_mode.get()],
                                  sort_keys=True, default=str)
            except Exception:
                return None

        def graph_refresh(self):
            """'Aktualisieren': Daten neu holen (Datenbank per API bzw. Datei neu lesen) und Voransicht neu berechnen."""
            self._gr_note = ""
            try:
                if self.v_source.get() == "db":
                    self.lbl_graph.configure(text=tr("rufe Daten aus der Datenbank ab …"))
                    self.update_idletasks()
                    self.db_fetch()
                    self._gr_note = tr(" · Datenbank neu abgerufen ") + datetime.now().strftime("%H:%M:%S")
                elif self.files and all(Path(f).exists() for f in self.files):
                    self.reload_table(first_load=False)
                    self._gr_note = tr(" · Datei neu gelesen ") + datetime.now().strftime("%H:%M:%S")
            except Exception:
                self._gr_note = ""
            self.graph_update(force=True)

        def graph_update(self, force=False):
            if not self.headers or not self.rows:
                self._gr_sig = None
                self._gr_message(tr("Noch keine Tabelle geladen. Bitte im Tab 'Spalten' bzw. 'Datenbank' eine Tabelle laden."))
                return
            schema = normalize_schema(self.gather_schema())
            req, opt = check_schema(schema, self.headers)
            if req or opt:
                self._gr_sig = None
                self._gr_message(tr("Die Spaltenzuordnung ist noch nicht vollständig:\n\n") + "\n".join(req + opt))
                return
            sig = self._gr_signature()
            if not force and sig is not None and sig == self._gr_sig:
                return
            self.lbl_graph.configure(text=tr("berechne …"))
            self.update_idletasks()
            try:
                res = convert(Path("Voransicht.csv"), schema, table=(self.headers, self.rows, ""), model_only=True,
                              log=lambda *a, **k: None)
                if len(res["model"]["se_nodes"]) > 3000:
                    raise HMAError(tr("Zu viele Units für die Voransicht ({0}). Bitte im Tab 'Werte' einen Zeilenfilter auf ein Projekt setzen.").format(len(res['model']['se_nodes'])))
                mode = next((k for k, v in PV_MODES.items() if v == self.v_gr_mode.get()), "auto")
                scene = layout_preview(preview_graph(res["model"]), mode)
            except HMAError as e:
                self._gr_sig = None
                self.lbl_graph.configure(text="")
                self._gr_message(tr("Voransicht nicht möglich:\n\n") + str(e), C_BAD)
                return
            except Exception:
                self._gr_sig = None
                self.lbl_graph.configure(text="")
                self._gr_message(tr("Unerwarteter Fehler in der Voransicht:\n\n") + traceback.format_exc(), C_BAD)
                return
            self._gr_scene, self._gr_sig, self._gr_sel = scene, sig, None
            self._gr_idx = {n["id"]: i for i, n in enumerate(scene["nodes"])}
            fmt = res["model"]["out_format"].upper()
            back = (tr(' · {0} Beziehung(en) gegen die Reihenfolge (orange)').format(scene['n_back']) if scene.get("n_back") else "")
            if scene.get("n_contemp_conflict"):
                back += tr(' · {0} zeitgleich-Paar(e) widersprechen der Stratigraphie').format(scene["n_contemp_conflict"])
            self.lbl_graph.configure(text=tr('{0} Units · {1} Beziehungen · {2} · Ausgabe {3}{4}{5}').format(len(scene['nodes']), sum(1 for e in scene['edges'] if e['kind'] == 'above'), PV_MODES[scene['mode']], fmt, back, self._gr_note))
            self._gr_populate_list()
            self._gr_card(None)
            self.graph_fit()

        def graph_fit(self):
            sc = self._gr_scene
            if not sc:
                return
            self.update_idletasks()
            cw = max(200, self.cv.winfo_width())
            ch = max(200, self.cv.winfo_height())
            self.graph_zoom_to(min(1.0, cw / sc["w"], ch / sc["h"]))

        def graph_zoom_to(self, z):
            if not self._gr_scene:
                return
            self._gr_zoom = max(0.08, min(3.0, z))
            self._gr_draw()
            self.cv.xview_moveto(0)
            self.cv.yview_moveto(0)

        def _gr_zoom_at(self, factor, px, py):
            """Zoomt um den Faktor, der Punkt (px, py) im Fenster bleibt dabei an seiner Stelle."""
            sc = self._gr_scene
            if not sc:
                return
            z = max(0.08, min(3.0, self._gr_zoom * factor))
            f = z / self._gr_zoom
            cx, cy = self.cv.canvasx(px), self.cv.canvasy(py)
            self._gr_zoom = z
            self._gr_draw()
            self.cv.xview_moveto(max(0.0, (cx * f - px) / max(1.0, sc["w"] * z)))
            self.cv.yview_moveto(max(0.0, (cy * f - py) / max(1.0, sc["h"] * z)))

        def _gr_zoom_step(self, factor):
            self._gr_zoom_at(factor, self.cv.winfo_width() / 2, self.cv.winfo_height() / 2)

        def _gr_pan(self, dx, dy):
            """Verschiebt den Ausschnitt um dx/dy Pixel."""
            sc = self._gr_scene
            if not sc:
                return
            w, h = max(1.0, sc["w"] * self._gr_zoom), max(1.0, sc["h"] * self._gr_zoom)
            if dx:
                self.cv.xview_moveto(max(0.0, self.cv.xview()[0] + dx / w))
            if dy:
                self.cv.yview_moveto(max(0.0, self.cv.yview()[0] + dy / h))

        def _gr_wheel(self, e):
            if not self._gr_scene:
                return
            self._gr_tip_hide()
            state = getattr(e, "state", 0) or 0
            if sys.platform == "darwin":
                # Mac: Trackpad (Zwei-Finger-Scrollen) und Mausrad verschieben das Bild, mit Strg/⌘ wird gezoomt.
                d = getattr(e, "delta", 0) or 0
                if state & 0xC:
                    self._gr_zoom_at(1.03 ** max(-10, min(10, d)), e.x, e.y)
                else:
                    step = max(-160, min(160, -d * 8))
                    if state & 0x1:  # Shift = waagerecht (so meldet Tk das seitliche Wischen)
                        self._gr_pan(step, 0)
                    else:
                        self._gr_pan(0, step)
                return "break"
            up = (e.delta > 0) if e.delta else (e.num == 4)
            self._gr_zoom_at(1.15 if up else 1 / 1.15, e.x, e.y)
            return "break"

        def _gr_press_ev(self, e):
            self._gr_press = (e.x, e.y)
            self.cv.scan_mark(e.x, e.y)

        def _gr_font(self, size):
            f = self._gr_fonts.get(size)
            if f is None:
                try:
                    from tkinter import font as tkfont
                    f = tkfont.Font(family="TkDefaultFont", size=-size)
                    f.measure("x")
                except Exception:
                    f = False
                self._gr_fonts[size] = f
            return f

        def _gr_fit_text(self, label, px, size):
            """Beschriftung so kürzen ('…'), dass sie in die Node passt."""
            f = self._gr_font(size)
            pre = (self._gr_scene or {}).get("prefix", "")
            if not f:
                return pv_fit_label(label, px, size * 0.62, pre)
            if f.measure(label) <= px - 6:
                return label
            if pre and label.startswith(pre) and len(label) > len(pre):  # gemeinsamen Anfang durch '…' ersetzen
                label = "…" + label[len(pre):]
                if f.measure(label) <= px - 6:
                    return label
            lo, hi, best = 1, len(label) - 1, 0
            while lo <= hi:
                mid = (lo + hi) // 2
                if f.measure(label[:mid] + "…") <= px - 6:
                    best, lo = mid, mid + 1
                else:
                    hi = mid - 1
            return label[:best] + "…" if best else ""

        def _gr_draw(self):
            sc, z = self._gr_scene, self._gr_zoom
            cv = self.cv
            self._gr_tip_hide()
            cv.delete("all")
            self._gr_hover = None
            if not sc:
                return
            cv.configure(scrollregion=(0, 0, sc["w"] * z, sc["h"] * z))
            fs = max(7, int(round(11 * z)))
            if self.v_gr_groups.get():
                for g in sc["groups"]:
                    cv.create_rectangle(g["x"] * z, g["y"] * z, (g["x"] + g["w"]) * z, (g["y"] + g["h"]) * z,
                                        fill="#f1efe4" if g["level"] == 0 else "#e3ecf7",
                                        outline="#b9b49a" if g["level"] == 0 else "#9db4cf")
                    if z >= 0.45:
                        cv.create_text((g["x"] + 5) * z, (g["y"] + 3) * z, anchor="nw", text=g["label"],
                                       fill="#777777", font=("TkDefaultFont", max(6, fs - 3)))
            for k, e in enumerate(sc["edges"]):
                flat = [c * z for p in e["pts"] for c in p]
                tags = ("edge", e["kind"], f"x{k}", f"e{self._gr_idx[e['s']]}", f"e{self._gr_idx[e['t']]}")
                if e["kind"] == "contemp":
                    if not self.v_gr_contemp.get():
                        continue
                    cv.create_line(*flat, fill="#9b1c1c" if e.get("clash") else "#d04a4a", dash=(4, 3),
                                   width=2 if e.get("clash") else 1, tags=tags)
                else:
                    cv.create_line(*flat, fill="#d98a1f" if e.get("back") else "#666666", arrow="last",
                                   arrowshape=(7, 8, 3), tags=tags)
            for i, n in enumerate(sc["nodes"]):
                cv.create_rectangle(n["x"] * z, n["y"] * z, (n["x"] + n["w"]) * z, (n["y"] + n["h"]) * z,
                                    fill=PV_NODE_COLORS.get(n["type"], "#cccccc"), outline="#4a6a99",
                                    tags=("node", f"n{i}", "rect"))
                txt = self._gr_fit_text(n["label"], n["w"] * z, fs) if z >= 0.3 else ""
                if txt:
                    cv.create_text((n["x"] + n["w"] / 2) * z, (n["y"] + n["h"] / 2) * z, text=txt,
                                   font=("TkDefaultFont", -fs), tags=("node", f"n{i}"))
            lx = 10
            for t in sc["types"]:
                cv.create_rectangle(lx, 6, lx + 14, 18, fill=PV_NODE_COLORS[t], outline="#4a6a99", tags="legend")
                cv.create_text(lx + 19, 12, anchor="w", text=PV_NODE_NAMES[t], font=("TkDefaultFont", 8),
                               tags="legend")
                lx += 30 + 7 * len(PV_NODE_NAMES[t])
            if self._gr_sel is not None:
                self._gr_highlight(self._gr_sel)

        # ----- Bedienfeld: Liste und Karteikarte -----
        def _gr_populate_list(self):
            self.lb.delete(0, "end")
            q = self.v_gr_q.get().strip().lower()
            self._gr_list = [n["id"] for n in (self._gr_scene["nodes"] if self._gr_scene else [])
                             if not q or q in n["id"].lower()]
            for i in self._gr_list:
                self.lb.insert("end", i)
            if self._gr_sel is not None and self._gr_scene:
                self._gr_list_mark(self._gr_scene["nodes"][self._gr_sel]["id"])

        def _gr_list_mark(self, nid):
            self.lb.selection_clear(0, "end")
            if nid in self._gr_list:
                pos = self._gr_list.index(nid)
                self.lb.selection_set(pos)
                self.lb.see(pos)

        def _gr_list_select(self, _e=None):
            sel = self.lb.curselection()
            if not sel or not self._gr_scene:
                return
            idx = self._gr_idx.get(self._gr_list[sel[0]])
            if idx is not None:
                self._gr_select(idx, from_list=True)

        def _gr_card(self, n):
            self.card.configure(state="normal")
            self.card.delete("1.0", "end")
            if n is None:
                self.card.insert("end", tr("Eine Unit in der Liste oder in der Matrix auswählen."), "key")
            else:
                def line(key, val):
                    self.card.insert("end", key + "\n", "key")
                    self.card.insert("end", (val or "–") + "\n", "val")
                self.card.insert("end", n["full"] + "\n", "id")
                line(tr("Typ"), PV_NODE_NAMES.get(n["type"], n["type"]))
                line("Name", n["name"])
                line(tr("Übergeordnete Gruppe"), n["outer"])
                line(tr("Untergeordnete Gruppe"), n["inner"])
                line(tr("liegt über"), ", ".join(n["over"]))
                line(tr("liegt unter"), ", ".join(n["under"]))
                line(tr("zeitgleich mit"), ", ".join(n["contemp"]))
                line(tr("ident mit"), ", ".join(n["ident"]))
                line(tr("Beschreibung"), (n["desc"] or "").strip())
            self.card.configure(state="disabled")
            self.card.yview_moveto(0)

        def _gr_highlight(self, idx):
            cv = self.cv
            cv.itemconfigure(f"n{idx}&&rect", width=3, outline="#d9480f")
            cv.itemconfigure(f"e{idx}&&above", fill="#1f6fd6", width=2)
            cv.itemconfigure(f"e{idx}&&contemp", width=2)
            try:
                cv.tag_raise(f"e{idx}&&above", "edge")
            except Exception:
                pass

        def _gr_select(self, idx, from_list=False, scroll=True):
            sc = self._gr_scene
            if not sc:
                return
            if self._gr_sel is not None:
                p = self._gr_sel
                self.cv.itemconfigure(f"n{p}&&rect", width=1, outline="#4a6a99")
                self.cv.itemconfigure(f"e{p}&&above", fill="#666666", width=1)
                self.cv.itemconfigure(f"e{p}&&contemp", width=1)
            self._gr_sel = idx
            self._gr_highlight(idx)
            n = sc["nodes"][idx]
            self._gr_card(n)
            if not from_list:
                if n["id"] not in self._gr_list and self.v_gr_q.get():
                    self.v_gr_q.set("")
                self._gr_list_mark(n["id"])
            if scroll:
                z = self._gr_zoom
                cw, ch = max(100, self.cv.winfo_width()), max(100, self.cv.winfo_height())
                self.cv.xview_moveto(max(0.0, ((n["x"] + n["w"] / 2) * z - cw / 2) / max(1.0, sc["w"] * z)))
                self.cv.yview_moveto(max(0.0, ((n["y"] + n["h"] / 2) * z - ch / 2) / max(1.0, sc["h"] * z)))

        def _gr_click(self, e):
            if abs(e.x - self._gr_press[0]) > 4 or abs(e.y - self._gr_press[1]) > 4:
                return  # war ein Ziehen
            idx = next((int(t[1:]) for t in self.cv.gettags("current") if t[0] == "n" and t[1:].isdigit()), None)
            if idx is not None:
                self._gr_select(idx, scroll=False)

        # ----- Hover -----
        def _gr_tip_show(self, text, e):
            self._gr_tip_hide()
            tw = tk.Toplevel(self)
            tw.wm_overrideredirect(True)
            tw.wm_geometry(f"+{e.x_root + 14}+{e.y_root + 14}")
            tk.Label(tw, text=text, justify="left", background="#ffffe0", relief="solid", borderwidth=1,
                     padx=6, pady=3, font=("TkDefaultFont", 9)).pack()
            self._gr_tip = tw

        def _gr_tip_move(self, e):
            if self._gr_tip:
                self._gr_tip.wm_geometry(f"+{e.x_root + 14}+{e.y_root + 14}")

        def _gr_tip_hide(self):
            if self._gr_tip:
                try:
                    self._gr_tip.destroy()
                except Exception:
                    pass
                self._gr_tip = None

        def _gr_node_enter(self, e):
            sc = self._gr_scene
            idx = next((int(t[1:]) for t in self.cv.gettags("current") if t[0] == "n" and t[1:].isdigit()), None)
            if sc and idx is not None:
                self._gr_tip_show(pv_hover_text(sc["nodes"][idx]), e)

        def _gr_edge_enter(self, e):
            sc = self._gr_scene
            tags = self.cv.gettags("current")
            k = next((int(t[1:]) for t in tags if t[0] == "x" and t[1:].isdigit()), None)
            if not sc or k is None:
                return
            ed = sc["edges"][k]
            if ed["kind"] == "contemp":
                text = tr('{0}\nzeitgleich mit\n{1}').format(ed['s'], ed['t'])
            else:
                text = tr('{0}\nliegt über\n{1}').format(ed['s'], ed['t'])
            cur = self.cv.find_withtag("current")
            if cur:
                item = cur[0]
                self._gr_hover = (item, self.cv.itemcget(item, "fill"), self.cv.itemcget(item, "width"))
                self.cv.itemconfigure(item, fill="#d9480f", width=3)
            self._gr_tip_show(text, e)

        def _gr_edge_leave(self, _e=None):
            self._gr_tip_hide()
            if self._gr_hover:
                item, fill, width = self._gr_hover
                try:
                    self.cv.itemconfigure(item, fill=fill, width=width)
                except Exception:
                    pass
                self._gr_hover = None

        def _tab_advanced(self):
            tab = ttk.Frame(self.nb)
            self.nb.add(tab, text=tr("Erweitert"))
            tab.columnconfigure(1, weight=1)
            rows = [
                ("delimiter", tr("CSV-Trennzeichen:"), ["auto", ",", ";", "tab", "|"], tr("auto = aus der Kopfzeile erkennen")),
                ("encoding", tr("Zeichenkodierung:"), ["auto", "utf-8-sig", "utf-8", "cp1252", "latin-1", "utf-16"],
                 tr("auto = UTF-8, sonst Windows-1252 (Umlaute/ß/Sonderzeichen)")),
                ("id_prefix_regex", tr("ID-Präfix entfernen (Regex):"), None,
                 tr("Leer = IDs unverändert übernehmen. Beispiele siehe Tabelle unten.")),
                ("list_split_regex", tr("Trenner zwischen Verweisen (Regex):"), None,
                 tr("Standard [;,]  – mehrere IDs in einer Beziehungszelle")),
            ]
            for r, (key, label, values, hint) in enumerate(rows):
                ttk.Label(tab, text=label).grid(row=r * 2, column=0, sticky="w", padx=10, pady=(10, 0))
                if values:
                    w = ttk.Combobox(tab, textvariable=self.set_vars[key], values=values, state="readonly", width=22)
                else:
                    w = ttk.Entry(tab, textvariable=self.set_vars[key], width=30)
                w.grid(row=r * 2, column=1, sticky="w", padx=6, pady=(10, 0))
                ttk.Label(tab, text=hint, foreground=C_MUTED).grid(row=r * 2 + 1, column=1, sticky="w", padx=6)
            cb = ttk.Checkbutton(tab, text=tr("HME: Units ohne übergeordnete Gruppe in Phasen 'unknown-N' einordnen "
                                           "(nach stratigraphischer Lage)"),
                                 variable=self.set_vars["hme_undated"], onvalue="1", offvalue="0")
            self.cb_undated = cb
            cb.grid(row=len(rows) * 2, column=0, columnspan=2, sticky="w", padx=10, pady=(14, 0))
            ex = ttk.LabelFrame(tab, text=tr("Beispiele für 'ID-Präfix entfernen (Regex)'"))
            ex.grid(row=len(rows) * 2 + 1, column=0, columnspan=2, sticky="w", padx=10, pady=(16, 0))
            table = ((tr("Regex"), tr("Entfernt"), tr("Beispiel")),
                     (r"^\d{6}_", tr("genau 6 Ziffern + _"), "123456_17 → 17"),
                     (r"^\d+_", tr("beliebig viele Ziffern + _"), "2026_17 → 17"),
                     ("^[A-Z]+_", tr("Großbuchstaben + _"), "XTESTX_17 → 17"),
                     (r"^\w+?_", tr("alles bis zum ersten _"), "123456_17 → 17"))
            for i, row in enumerate(table):
                for j, text in enumerate(row):
                    kw = {"font": ("Consolas", 10)} if (j == 0 and i) else ({"font": ("TkDefaultFont", 9, "bold")} if not i else {})
                    ttk.Label(ex, text=text, **kw).grid(row=i, column=j, sticky="w", padx=(10, 24), pady=2)

        def _tab_log(self):
            tab = ttk.Frame(self.nb)
            self.tab_log = tab
            self.nb.add(tab, text=tr("Protokoll"))
            bar = ttk.Frame(tab)
            bar.pack(fill="x", padx=8, pady=(6, 0))
            ttk.Button(bar, text=tr("Protokoll speichern…"), command=self.save_report).pack(side="left")
            ttk.Button(bar, text=tr("Kopieren"), command=self.copy_report).pack(side="left", padx=6)
            ttk.Button(bar, text=tr("Leeren"), command=self.clear_log).pack(side="left")
            inner = ttk.Frame(tab)
            inner.pack(fill="both", expand=True, padx=8, pady=6)
            self.txt = tk.Text(inner, wrap="word", state="disabled", bg="white",
                               font=("Consolas", 10) if sys.platform.startswith("win") else ("Courier", 10))
            sc = ttk.Scrollbar(inner, orient="vertical", command=self.txt.yview)
            self.txt.configure(yscrollcommand=sc.set)
            self.txt.pack(side="left", fill="both", expand=True)
            sc.pack(side="right", fill="y")
            self.txt.tag_configure("head", foreground="#1f4e79", font=("TkDefaultFont", 10, "bold"))
            self.txt.tag_configure("code", foreground="#b86a1d", font=("TkDefaultFont", 10, "bold"))
            self.txt.tag_configure("ok", foreground=C_OK)
            self.txt.tag_configure("err", foreground=C_BAD, font=("TkDefaultFont", 10, "bold"))

        # ---------- Zustand ----------
        def gather_schema(self):
            settings = {k: v.get() for k, v in self.set_vars.items()}
            settings["descr_suffix_on"] = "1" if self.v_suffix_on.get() else "0"
            settings["ref_key_col"] = self.v_ref.get() if self.v_ref.get() not in (NONE_LABEL, REF_DEFAULT) else ""
            mapping = {k: (v.get() if v.get() != NONE_LABEL else "") for k, v in self.map_vars.items()}
            for k in "tgu":
                settings[f"tug_{k}_on"] = "1" if self.v_tug_on[k].get() else "0"
            return {"mapping": mapping, "settings": settings,
                    "tug_t_values": sorted(self.tug_sel["t"]), "tug_g_values": sorted(self.tug_sel["g"]),
                    "tug_u_values": sorted(self.tug_sel["u"]),
                    "files": [str(p) for p in self.files], "sheet": self.v_sheet.get(),
                    "out_dir": self.v_out.get().strip(), "source": self.v_source.get(),
                    "db": {"provider": self._prov_key(), **{k: v.get().strip() for k, v in self.db_vars.items()},
                           "resolve_col": ("" if self.db_vars["resolve_col"].get() in (AUTO_LABEL, NONE_LABEL)
                                           else self.db_vars["resolve_col"].get()),
                           "overwrite": "1" if self.v_overwrite.get() else "0",
                           "as_text": "1" if self.v_astext.get() else "0"},
                    "kategorie_interface": sorted(self.kat_interface),
                    "hme_types": dict(self.hme_types),
                    "group_outer_ignore": sorted(self.outer_ignore),
                    "group_inner_ignore": sorted(self.inner_ignore),
                    "filter_values": sorted(self.filter_values)}

        def apply_schema(self, sc):
            sc = normalize_schema(sc)
            for k in MAP_KEYS:
                self.map_vars[k].set(sc["mapping"][k])
            for k in DEFAULT_SETTINGS:
                if k in self.set_vars:
                    self.set_vars[k].set(sc["settings"][k])
            self.v_suffix_on.set(sc["settings"]["descr_suffix_on"] == "1")
            self.v_ref.set(sc["settings"]["ref_key_col"])
            self.kat_interface = set(sc["kategorie_interface"])
            self.hme_types = dict(sc["hme_types"])
            self.outer_ignore = set(sc["group_outer_ignore"])
            self.inner_ignore = set(sc["group_inner_ignore"])
            self.tug_sel = {k: set(sc[f"tug_{k}_values"]) for k in "tgu"}
            for k in "tgu":
                self.v_tug_on[k].set(sc["settings"][f"tug_{k}_on"] == "1")
            self.filter_values = set(sc["filter_values"])
            self.v_source.set(sc["source"])
            for k, v in self.db_vars.items():
                v.set(sc["db"][k] or (AUTO_LABEL if k == "resolve_col" else ""))
            self.v_overwrite.set(sc["db"]["overwrite"] == "1")
            self.v_astext.set(sc["db"]["as_text"] == "1")
            self.v_provider.set(DB_PROVIDER_NAME.get(sc["db"]["provider"], "Airtable"))
            if self.prov_frames:
                self._show_provider()
                self.v_fm_noverify.set(sc["db"]["tls_verify"] == "0")
                for cb, keys in self._ssl_cbs:
                    cb.set(self.SSL_LABELS[sc["db"]["sslmode"] if sc["db"]["sslmode"] in keys else "prefer"])
            self._vl_key = None

        def persist_last(self):
            last = self.gather_schema()
            last.update({"schema": self.v_schema.get()})
            self.store.data["last"] = last
            self.persist_key()
            return self.store.save()

        def on_close(self):
            self.persist_last()
            self.destroy()

        # ---------- Dateien ----------
        def pick_files(self):
            start = str(Path(self.files[0]).parent) if self.files else str(Path.home())
            files = filedialog.askopenfilenames(
                title=tr("Tabelle(n) wählen"), initialdir=start,
                filetypes=[(tr("Tabellen und Datenbanken"), "*.csv *.txt *.tsv *.xlsx *.xlsm *.xls *.ods *.dbf *.accdb *.mdb "
                                                        "*.odb *.sqlite *.sqlite3 *.db *.gpkg"),
                           (tr("CSV / Text"), "*.csv *.txt *.tsv"), ("Excel", "*.xlsx *.xlsm *.xls"),
                           ("LibreOffice Calc", "*.ods"), (tr("dBase"), "*.dbf"),
                           ("Access", "*.accdb *.mdb"), ("LibreOffice Base", "*.odb"),
                           ("SQLite", "*.sqlite *.sqlite3 *.db *.gpkg"), (tr("Alle Dateien"), "*.*")])
            if files:
                self.v_sheet.set("")
                self.set_files(list(files))

        def pick_out(self):
            d = filedialog.askdirectory(title=tr("Speicherort wählen"))
            if d:
                self.v_out.set(d)

        def set_files(self, files):
            self.files = [Path(f) for f in files]
            self.v_file.set(", ".join(p.name for p in self.files))
            first = self.files[0]
            if first.suffix.lower() in SHEET_EXT:
                is_db = first.suffix.lower() in DB_FILE_EXT
                try:
                    if is_db:
                        self.lbl_fileinfo.configure(text=tr("Datenbank wird gelesen …"))
                        self.update_idletasks()
                    self.sheets = sheet_names(first)
                except Exception as e:
                    self.sheets = []
                    messagebox.showerror(tr("Datenbank-Datei") if is_db else tr("Excel-Datei"),
                                         tr('Tabellen bzw. Blätter konnten nicht gelesen werden:\n{0}').format(e))
                self.cb_sheet.configure(values=self.sheets, state="readonly")
                if self.v_sheet.get() not in self.sheets:
                    self.v_sheet.set(self.sheets[0] if self.sheets else "")
            else:
                self.sheets = []
                self.cb_sheet.configure(values=[], state="disabled")
                self.v_sheet.set("")
            self.reload_table(first_load=True)

        def on_sheet_changed(self):
            if self.files and self.sheets:
                self.reload_table(first_load=True)

        def on_source(self, select=True):
            db = self.v_source.get() == "db"
            try:
                self.nb.tab(self.tab_db, state="normal" if db else "hidden")
            except Exception:
                pass
            self.btn_pick.configure(state="disabled" if db else "normal")
            if db:
                self.cb_sheet.configure(state="disabled")
                self.lbl_fileinfo.configure(text=tr('Datenbank ({0}) – die Daten werden beim Erzeugen abgerufen.').format(self.v_provider.get()))
                self.lbl_outhint.configure(text=tr("leer = Ordner Dokumente"))
                if select:
                    self.nb.select(self.tab_db)
            else:
                self.lbl_outhint.configure(text=tr("leer = neben der Tabelle"))
                self.lbl_fileinfo.configure(text="")
                if self.sheets:
                    self.cb_sheet.configure(state="readonly")
                if select:
                    self.nb.select(self.sf_cols)
                if self.files:
                    self.reload_table(first_load=False)

        def db_parse_url(self):
            found = parse_airtable_url(self.v_dburl.get())
            if not found:
                messagebox.showinfo(tr("Airtable-Link"), tr("In dem Text wurde keine Base ID (app…) gefunden."))
                return
            for k, v in found.items():
                self.db_vars[k].set(v)
            self.lbl_dbstatus.configure(text=tr("Übernommen: ") + ", ".join(k.replace("_id", "").capitalize() for k in found))

        def _db_table(self, log=None):
            schema = normalize_schema(self.gather_schema())
            return db_load_table(schema["db"], self.v_secret[schema["db"]["provider"]].get(), schema, log=log)

        def db_fetch(self):
            self.lbl_dbstatus.configure(text=tr("Rufe Daten ab …"))
            self.update_idletasks()
            try:
                table = self._db_table()
            except HMAError as e:
                self.lbl_dbstatus.configure(text=tr("Fehler"))
                messagebox.showerror(tr("Datenbank"), str(e))
                return
            except Exception:
                self.lbl_dbstatus.configure(text=tr("Fehler"))
                messagebox.showerror(tr("Datenbank"), traceback.format_exc())
                return
            self._set_table(table)
            self.persist_key()
            self.lbl_dbstatus.configure(text=tr('OK – {0} Datensätze, {1} Spalten').format(len(table[1]), len(table[0])))

        def _set_table(self, table):
            self.headers, self.rows, info = table
            hs = set(self.headers)
            first = not any(self.map_vars[k].get() in hs for k in MAP_KEYS)
            self.cb_resolve.set_values([AUTO_LABEL] + self.headers)
            self._table_ready(info, first_load=first)

        def persist_key(self):
            for k, _n in DB_PROVIDERS:
                try:
                    key = self.v_secret[k].get().strip()
                    if self.v_savekeys[k].get() and key:
                        if not self.store.set_secret(k, key):
                            self.v_savekeys[k].set(False)
                    else:
                        self.store.set_secret(k, None)
                except Exception:
                    pass

        def reload_table(self, first_load=False):
            if not self.files:
                return
            try:
                if self.files[0].suffix.lower() in DB_FILE_EXT:
                    self.lbl_fileinfo.configure(text=tr("Datenbank wird gelesen …"))
                    self.update_idletasks()
                self.headers, self.rows, info = load_table(self.files[0], self.gather_schema()["settings"],
                                                           self.v_sheet.get())
                self._stamp = self._file_stamp()
            except Exception as e:
                self.headers, self.rows = [], []
                messagebox.showerror(tr("Tabelle konnte nicht gelesen werden"), str(e))
                self.refresh()
                return
            self._table_ready(info, first_load)

        def _file_stamp(self):
            out = []
            for f in self.files:
                try:
                    st = Path(f).stat()
                    out.append((str(f), st.st_mtime_ns, st.st_size))
                except OSError:
                    out.append((str(f), 0, 0))
            return out

        def _files_changed(self):
            return (self.v_source.get() == "file" and bool(self.files) and getattr(self, "_stamp", None) is not None
                    and self._file_stamp() != self._stamp)

        def _table_ready(self, info, first_load=False):
            self.lbl_fileinfo.configure(text=info)
            for key, cb in self.cbs.items():
                cb.set_values(self.headers if key in ("id1", "above") else [NONE_LABEL] + self.headers)
            self.cb_ref.set_values([REF_DEFAULT] + self.headers)
            self.cb_filter.set_values([NO_FILTER, PREFIX_LABEL] + self.headers)
            hs = set(self.headers)
            sug = suggest_mapping(self.headers)
            changed = 0
            for key in MAP_KEYS:
                cur = self.map_vars[key].get()
                if key in sug and cur not in hs and (first_load or not cur):
                    self.map_vars[key].set(sug[key])
                    changed += 1
            self._vl_key = None
            txt = tr('{0} Zeilen, {1} Spalten').format(len(self.rows), len(self.headers))
            if changed:
                txt += tr(' – {0} Zuordnung(en) automatisch erkannt').format(changed)
            if self.v_source.get() == "file" and len(self.files) > 1:
                txt += tr(' – {0} Dateien (gleiche Zuordnung für alle)').format(len(self.files))
            self.lbl_csvinfo.configure(text=txt)
            self.refresh()

        # ---------- Schemata ----------
        def refresh_schema_list(self):
            self.cb_schema.configure(values=sorted(self.store.data["schemas"]))

        def load_schema(self, name):
            sc = self.store.data["schemas"].get(name)
            if sc is None:
                return
            self.apply_schema(sc)
            self.on_source(select=False)
            files = [f for f in sc.get("files", []) if Path(f).exists()] if sc.get("source") == "file" else []
            if sc.get("source") == "db":
                pass
            elif files:  # gemerkte Tabelle des Schemas direkt wieder laden
                self.v_sheet.set(sc.get("sheet", ""))
                self.set_files(files)
            else:
                self.reload_table(first_load=False)
            if sc.get("out_dir") and Path(sc["out_dir"]).is_dir():
                self.v_out.set(sc["out_dir"])

        def _write_config(self, msg):
            ok = self.store.save()
            self.set_status(msg if ok else tr("Konnte die Konfigurationsdatei nicht schreiben (Schreibschutz?)."), ok)

        def save_schema(self):
            name = self.v_schema.get()
            if not name:
                return self.save_schema_as()
            self.store.data["schemas"][name] = normalize_schema(self.gather_schema())
            self._write_config(tr("Schema '{0}' gespeichert.").format(name))

        def save_schema_as(self):
            name = simpledialog.askstring(tr("Mapping-Schema speichern unter"),
                                          tr("Name des Schemas (z. B. Name der Datenbank):"), parent=self)
            if not name or not name.strip():
                return
            name = name.strip()
            if name in self.store.data["schemas"] and not messagebox.askyesno(
                    tr("Schema überschreiben"), tr("Schema '{0}' existiert bereits. Überschreiben?").format(name)):
                return
            self.v_schema.set(name)
            self.store.data["schemas"][name] = normalize_schema(self.gather_schema())
            self.refresh_schema_list()
            self._write_config(tr("Schema '{0}' gespeichert.").format(name))

        def delete_schema(self):
            name = self.v_schema.get()
            if not name or name not in self.store.data["schemas"]:
                return
            if messagebox.askyesno(tr("Schema löschen"), tr("Schema '{0}' wirklich löschen?").format(name)):
                del self.store.data["schemas"][name]
                self.store.save()
                rest = sorted(self.store.data["schemas"])
                self.v_schema.set(rest[0] if rest else "")
                self.refresh_schema_list()

        def export_schema(self):
            f = filedialog.asksaveasfilename(title=tr("Schema exportieren"), defaultextension=".json",
                                             initialfile=f"{self.v_schema.get() or 'schema'}.json",
                                             filetypes=[(tr("Mapping-Schema"), "*.json")])
            if f:
                Path(f).write_text(json.dumps({"name": self.v_schema.get(),
                                               "schema": {**self.gather_schema(), "files": [], "sheet": "",
                                                          "out_dir": ""}},
                                              ensure_ascii=False, indent=2), encoding="utf-8")
                self.set_status(tr('Schema exportiert: {0}').format(f), True)

        def import_schema(self):
            f = filedialog.askopenfilename(title=tr("Schema importieren"), filetypes=[(tr("Mapping-Schema"), "*.json")])
            if not f:
                return
            try:
                raw = json.loads(Path(f).read_text(encoding="utf-8"))
                sc = normalize_schema(raw.get("schema", raw))
                name = raw.get("name") or Path(f).stem
            except Exception as e:
                messagebox.showerror(tr("Import"), tr('Datei konnte nicht gelesen werden:\n{0}').format(e))
                return
            name = simpledialog.askstring(tr("Schema importieren"), tr("Name des Schemas:"), initialvalue=name,
                                          parent=self) or name
            self.store.data["schemas"][name] = sc
            self.store.save()
            self.refresh_schema_list()
            self.v_schema.set(name)
            self.load_schema(name)

        # ---------- Zuordnung ----------
        def auto_suggest(self):
            if not self.headers:
                messagebox.showinfo(tr("Automatisch erkennen"), tr("Bitte zuerst eine Datei wählen."))
                return
            sug = suggest_mapping(self.headers)
            for key in MAP_KEYS:
                if key != "filter_col":
                    self.map_vars[key].set(sug.get(key, ""))
            self.set_status(tr('{0} Spalten erkannt – bitte prüfen.').format(len(sug)), True)

        def clear_mapping(self):
            for key in MAP_KEYS:
                self.map_vars[key].set("")
            self.v_ref.set("")
            self.filter_values = set()

        def _on_format(self):
            self._vl_key = None
            self.refresh()
            self._on_tab_changed()

        def _tug_state(self):
            hme = self.set_vars["out_format"].get() == "hme"
            for k in "tgu":
                self.tug_picks[k].set_enabled(self.v_tug_on[k].get() and not hme)

        def _apply_format_state(self):
            """HME gewählt: Felder, die der HME nicht verarbeitet, ausgrauen (HMC: alles wieder aktiv)."""
            hme = self.set_vars["out_format"].get() == "hme"
            fg = getattr(self, "_lbl_fg", None) or ttk.Style().lookup("TLabel", "foreground") or "black"
            self._lbl_fg = fg
            self.cbs["contemp"].configure(state="disabled" if hme else "normal")
            self.rowlbls["contemp"].configure(foreground=C_MUTED if hme else fg)
            if hme:
                self.stats["contemp"].configure(text=tr("im HME nicht verwendet"), foreground=C_MUTED)
                self.exlbls["contemp"].configure(text="")
            targets = {"kategorie": "Node-Type (Layer, Deposit, Fill, Structure, Interface)", "group_outer": "Phase",
                       "group_inner": "Object", "above": "Relation", "below": "Relation"}
            for key, (lbl, orig) in self.tlbls.items():
                lbl.configure(text="→ " + (targets.get(key, orig) if hme else orig))
            self.fr_kat.configure(text=tr("Kategorie: jedem Wert einen HME-Node-Typ zuordnen") if hme else
                                  tr("Kategorie: welche Werte sind ein Interface (SURFACE)?"))
            self.fr_tug.configure(text=tr("T, G und U einfügen – im HME nicht verwendet") if hme else
                                  tr("Optional: T, G und U einfügen (Top surface, Interface to Geology, Unexcavated)"))

            def walk(w):
                for ch in w.winfo_children():
                    if isinstance(ch, ttk.Checkbutton):
                        ch.state(["disabled"] if hme else ["!disabled"])
                    elif isinstance(ch, ttk.Frame) and not isinstance(ch, MultiPick):
                        walk(ch)
            walk(self.fr_tug)
            self._tug_state()
            self.cb_undated.state(["!disabled"] if hme else ["disabled"])

        def _all_ids(self):
            idxs = [self.headers.index(self.map_vars[k].get()) for k in ID_KEYS
                    if self.map_vars[k].get() in self.headers]
            fc = self.map_vars["filter_col"].get()
            fi = self.headers.index(fc) if fc in self.headers else None
            try:
                pre = re.compile(self.set_vars["id_prefix_regex"].get()) if self.set_vars["id_prefix_regex"].get() else None
            except re.error:
                pre = None
            seen, out = set(), []
            for _, r in self.rows:
                rid = "_".join(p for p in ((r[i] if i < len(r) else "").strip() for i in idxs) if p)
                if self.filter_values and fc:  # Zeilenfilter berücksichtigen
                    if fc == PREFIX_COL:
                        mm = pre.match(rid) if pre else None
                        val = mm.group(0).rstrip("_") if mm else ""
                    else:
                        val = (r[fi] if fi is not None and fi < len(r) else "").strip()
                    if val not in self.filter_values:
                        continue
                if rid and rid not in seen:
                    seen.add(rid)
                    out.append(rid)
            return out

        def rebuild_tug(self):
            ids = self._all_ids() if (self.headers and self.rows) else []
            for k in "tgu":
                self.tug_picks[k].sel = self.tug_sel[k]
                self.tug_picks[k].set_items(ids)
            self._tug_state()

        def on_filter_col(self):
            d = self.v_filter_disp.get()
            new = "" if d in ("", NO_FILTER) else PREFIX_COL if d == PREFIX_LABEL else d
            if new != self.map_vars["filter_col"].get():
                self.filter_values = set()
                self.map_vars["filter_col"].set(new)

        def refresh(self):
            hs = set(self.headers)
            for k in MAP_KEYS:
                if self.map_vars[k].get() == NONE_LABEL:
                    self.map_vars[k].set("")
            if self.v_ref.get() in (NONE_LABEL, REF_DEFAULT):
                self.v_ref.set("")
            fc = self.map_vars["filter_col"].get()
            self.v_filter_disp.set(NO_FILTER if not fc else PREFIX_LABEL if fc == PREFIX_COL else fc)
            # ID-Teile: Teil n+1 nur wählbar, wenn Teil n gewählt
            prev = True
            for key in ID_KEYS:
                if not prev and self.map_vars[key].get():
                    self.map_vars[key].set("")
                self.cbs[key].configure(state="normal" if prev else "disabled")
                prev = bool(self.map_vars[key].get())
            # doppelt vergebene Spalten
            dup = set()
            for group in DISTINCT_GROUPS:
                cnt = Counter(self.map_vars[k].get() for k in group if self.map_vars[k].get())
                dup |= {k for k in group if cnt.get(self.map_vars[k].get(), 0) > 1}
            for key, stat in self.stats.items():
                cur = self.map_vars[key].get()
                required = key == "above"
                if key in dup:
                    stat.configure(text=tr("✗ Spalte doppelt vergeben"), foreground=C_BAD)
                elif not cur:
                    stat.configure(text=tr("● erforderlich") if required else "optional",
                                   foreground=C_BAD if required else C_MUTED)
                elif not self.headers:
                    stat.configure(text=tr("(keine Datei geladen)"), foreground=C_MUTED)
                elif cur in hs:
                    stat.configure(text=tr("✓ gefunden"), foreground=C_OK)
                else:
                    stat.configure(text=tr("✗ nicht in dieser Tabelle"), foreground=C_BAD)
            self.ent_suffix.configure(state="normal" if self.v_suffix_on.get() else "disabled")
            cache_key = (id(self.rows), len(self.rows), tuple(self.headers))
            if getattr(self, "_ex_key", None) != cache_key:
                self._ex_key, self._ex_cache = cache_key, {}
            for key, lab in self.exlbls.items():
                cur = self.map_vars[key].get()
                if cur not in hs or key in dup:
                    lab.configure(text="")
                    continue
                if cur not in self._ex_cache:
                    self._ex_cache[cur] = example_values(self.headers, self.rows, cur, 40, EX_LEAD, EX_ETC)
                lab.configure(text=self._ex_cache[cur])
            ex, color = "", C_MUTED
            if any(k in dup for k in ID_KEYS):
                ex, color = tr("✗ Dieselbe Spalte ist mehrfach als ID-Teil gewählt."), C_BAD
            elif self.rows and self.map_vars["id1"].get() in hs:
                parts = []
                for key in ID_KEYS:
                    c = self.map_vars[key].get()
                    if c in hs:
                        i = self.headers.index(c)
                        v = (self.rows[0][1][i] if i < len(self.rows[0][1]) else "").strip()
                        if v:
                            parts.append(v)
                ex = tr("Beispiel (erste Zeile): ") + "_".join(parts)
            elif not self.map_vars["id1"].get():
                ex = tr("Mindestens Teil 1 wählen; danach wird Teil 2 freigeschaltet usw.")
            self.lbl_idex.configure(text=ex, foreground=color)
            # Hinweis auf den häufigen Fehler "mehrere Projekte in einer Tabelle"
            warn = ""
            if self.rows and self.map_vars["id1"].get() in hs and not (fc and self.filter_values):
                try:
                    pc = distinct_values(self.headers, self.rows, PREFIX_COL,
                                         normalize_schema(self.gather_schema()))
                    pc = {k: v for k, v in pc.items() if k}
                except Exception:
                    pc = {}
                if len(pc) > 1:
                    warn = (tr("⚠ Die Tabelle enthält {0} Projekte (ID-Präfixe). Bitte im Tab 'Werte' einen Zeilenfilter auf EIN "
                               'Projekt setzen.').format(len(pc)))
            self.lbl_multi.configure(text=warn)
            self.refresh_schema_list()
            self.refresh_preview()
            key = (self.map_vars["kategorie"].get(), self.map_vars["group_outer"].get(), self.map_vars["group_inner"].get(),
                   fc,
                   tuple(self.map_vars[k].get() for k in ID_KEYS), self.set_vars["id_prefix_regex"].get(),
                   tuple(self.headers), len(self.rows), self.set_vars["out_format"].get())
            self._apply_format_state()
            if key != self._vl_key:
                self._vl_key = key
                self.rebuild_value_lists()

        def _value_list(self, frame, col, colname, selected, empty_text, label_checked, render=None):
            for w in frame.winfo_children():
                w.destroy()
            if not col:
                ttk.Label(frame, text=empty_text, foreground=C_MUTED).pack(anchor="w", padx=8, pady=6)
                return
            counts = distinct_values(self.headers, self.rows, col, normalize_schema(self.gather_schema())) \
                if self.headers else Counter()
            vals = sorted({v for v in counts if v} | selected, key=lambda s: (-counts.get(s, 0), s))
            if self.headers and self.rows and not vals:
                ttk.Label(frame, text=tr("Spalte '{0}': Es befinden sich keine Werte in dieser Spalte.").format(colname),
                          foreground=C_MUTED).pack(anchor="w", padx=8, pady=6)
                return
            if len(vals) > 200:
                ttk.Label(frame, text=tr("Die Spalte '{0}' hat {1} verschiedene Werte – dafür ungeeignet.").format(colname, len(vals)), foreground=C_MUTED).pack(anchor="w", padx=8, pady=6)
                return
            ttk.Label(frame, text=tr("Spalte '{0}' – {1}").format(colname, label_checked), foreground=C_MUTED).pack(
                anchor="w", padx=8, pady=(4, 2))
            if render:
                render(frame, vals, counts)
                if not counts:
                    ttk.Label(frame, text=tr("Noch keine Tabelle geladen – gespeicherte Werte werden angezeigt."),
                              foreground=C_MUTED).pack(anchor="w", padx=8, pady=4)
                return
            for v in vals:
                var = tk.BooleanVar(value=v in selected)

                def toggle(v=v, var=var):
                    (selected.add if var.get() else selected.discard)(v)
                    self.after_idle(self.rebuild_tug)

                n = counts.get(v)
                text = tr('{0}   ({1} Zeilen)').format(v, n) if n else tr('{0}   (nicht in dieser Tabelle)').format(v)
                ttk.Checkbutton(frame, text=text, variable=var, command=toggle).pack(anchor="w", padx=12)
            if not counts:
                ttk.Label(frame, text=tr("Noch keine Tabelle geladen – gespeicherte Werte werden angezeigt."),
                          foreground=C_MUTED).pack(anchor="w", padx=8, pady=4)

        def _render_types(self, frame, vals, counts):
            """HME: pro Kategorie-Wert eine Auswahl des Node-Typs (Layer, Deposit, Fill, Structure, Interface)."""
            labels = {"deposit": tr("Deposit (Standard)"), "interface": "Interface", "fill": "Fill",
                      "structure": "Structure", "layer": "Layer"}
            rev = {v: k for k, v in labels.items()}
            grid = ttk.Frame(frame)
            grid.pack(anchor="w", padx=12, pady=(0, 4))
            for i, v in enumerate(vals):
                n = counts.get(v)
                ttk.Label(grid, text=tr('{0}   ({1} Zeilen)').format(v, n) if n else tr('{0}   (nicht in dieser Tabelle)').format(v)).grid(
                    row=i, column=0, sticky="w", padx=(0, 16), pady=1)
                cur = "interface" if v in self.kat_interface else self.hme_types.get(v, "deposit")
                var = tk.StringVar(value=labels[cur])
                cb = ttk.Combobox(grid, textvariable=var, values=list(labels.values()), state="readonly", width=18)
                cb.grid(row=i, column=1, sticky="w", pady=1)

                def picked(_e=None, v=v, var=var):
                    t = rev.get(var.get(), "deposit")
                    self.kat_interface.discard(v)
                    self.hme_types.pop(v, None)
                    if t == "interface":
                        self.kat_interface.add(v)
                    elif t != "deposit":
                        self.hme_types[v] = t
                    self.after_idle(self.rebuild_tug)
                cb.bind("<<ComboboxSelected>>", picked)

        def rebuild_value_lists(self):
            kc, oc, fc = (self.map_vars[k].get() for k in ("kategorie", "group_outer", "filter_col"))
            if self.set_vars["out_format"].get() == "hme":
                self._value_list(self.fr_kat, kc, kc, set(self.kat_interface) | set(self.hme_types),
                                 tr("Bitte im Tab 'Spalten' eine Kategorie-Spalte wählen. Ohne Auswahl sind alle "
                                 "Units Deposit."),
                                 tr("jedem Wert einen HME-Node-Typ zuordnen (nicht zugeordnet = Deposit)"),
                                 render=self._render_types)
            else:
                self._value_list(self.fr_kat, kc, kc, self.kat_interface,
                                 tr("Bitte im Tab 'Spalten' eine Kategorie-Spalte wählen. Ohne Auswahl sind alle SEs "
                                 "DEPOSIT."), tr("angehakt = Interface (SURFACE), sonst Deposit"))
            self._value_list(self.fr_outer, oc, oc, self.outer_ignore,
                             tr("Bitte im Tab 'Spalten' eine übergeordnete Gruppe wählen."),
                             tr("angehakt = ignorieren (z. B. 'unklar'), die SE landet dann in keiner Gruppe"))
            ic = self.map_vars["group_inner"].get()
            self._value_list(self.fr_inner, ic, ic, self.inner_ignore,
                             tr("Bitte im Tab 'Spalten' eine untergeordnete Gruppe wählen."),
                             tr("angehakt = ignorieren, die SE landet dann in keiner untergeordneten Gruppe"))
            self.rebuild_tug()
            self._value_list(self.fr_filter, fc, tr("Projekt-Präfix der ID") if fc == PREFIX_COL else fc,
                             self.filter_values, tr("Kein Filter gewählt – alle Zeilen werden verwendet."),
                             tr("angehakt = diese Zeilen verwenden"))

        def refresh_preview(self):
            self.tree.delete(*self.tree.get_children())
            cols = []
            for key in MAP_KEYS + ["ref"]:
                c = self.v_ref.get() if key == "ref" else self.map_vars[key].get()
                if c and c in self.headers and c not in cols:
                    cols.append(c)
            self.tree.configure(columns=cols)
            for c in cols:
                self.tree.heading(c, text=c)
                self.tree.column(c, width=150, stretch=False)
            idxs = [self.headers.index(c) for c in cols]
            for _, r in self.rows[:60]:
                self.tree.insert("", "end", values=[(r[i] if i < len(r) else "")[:70] for i in idxs])

        # ---------- Protokoll ----------
        def set_status(self, text, ok=True):
            self.status.configure(text=text, foreground=C_OK if ok else C_BAD)

        def log_line(self, line):
            self.txt.configure(state="normal")
            tag = None
            if re.fullmatch(r"[A-Z][A-Z_]{4,}", line):
                tag = "code"
            elif line.startswith("=") or re.match(r"^\d\. ", line):
                tag = "head"
            elif line.strip() == tr("keine Fehler"):
                tag = "ok"
            elif line.startswith(tr("FEHLER")) or line.startswith(tr("Es wurde KEINE")):
                tag = "err"
            if tag:
                self.txt.insert("end", line + "\n", tag)
            else:
                self.txt.insert("end", line + "\n")
            self.txt.configure(state="disabled")

        def log_text(self, text):
            for line in text.splitlines():
                self.log_line(line)
            self.txt.see("end")

        def clear_log(self):
            self.txt.configure(state="normal")
            self.txt.delete("1.0", "end")
            self.txt.configure(state="disabled")
            self._report = ""

        def save_report(self):
            if not self._report:
                messagebox.showinfo(tr("Protokoll"), tr("Noch kein Protokoll vorhanden."))
                return
            f = filedialog.asksaveasfilename(title=tr("Protokoll speichern"), defaultextension=".txt",
                                             initialfile=tr('Protokoll_{0}.txt').format(self.files[0].stem if self.files else 'HMA'),
                                             filetypes=[(tr("Textdatei"), "*.txt")])
            if f:
                Path(f).write_text(self._report, encoding="utf-8")

        def copy_report(self):
            if self._report:
                self.clipboard_clear()
                self.clipboard_append(self._report)
                self.set_status(tr("Protokoll in die Zwischenablage kopiert."), True)

        # ---------- Öffnen ----------
        def _open(self, path):
            try:
                if sys.platform.startswith("win"):
                    os.startfile(str(path))  # noqa
                else:
                    import subprocess
                    subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(path)])
            except Exception as e:
                messagebox.showerror(tr("Öffnen"), str(e))

        def open_file(self):
            if self._last_hmcx:
                self._open(self._last_hmcx)

        def open_dir(self):
            if self._last_hmcx:
                self._open(Path(self._last_hmcx).parent)

        # ---------- Ausführen ----------
        def reopen_in_hmc(self, state, target):
            try:
                os.startfile(str(target))
                self.log_line(tr("HMC: Datei wird neu geöffnet …"))
                self._hmc_poll(target.name, state, 0)
            except Exception as e:
                self.log_line(tr('HMC: Neu öffnen nicht möglich ({0}). Bitte die Datei selbst öffnen.').format(e))

        def _hmc_poll(self, name, state, n):
            hwnd = hmc_find_window(name)
            if hwnd:
                hmc_set_state(hwnd, state)
                if n >= 0:  # nach kurzer Pause noch einmal, falls HMC beim Laden die Position ändert
                    self.after(1500, lambda: (hmc_set_state(hmc_find_window(name) or hwnd, state)))
                return
            if n < 80:
                self.after(500, lambda: self._hmc_poll(name, state, n + 1))

        def run(self):
            if self._busy:
                return
            dbmode = self.v_source.get() == "db"
            table = None
            if dbmode:
                prov = self._prov_key()
                miss = db_ready(normalize_schema(self.gather_schema())["db"], self.v_secret[prov].get())
                if miss:
                    messagebox.showwarning(tr("Matrix erzeugen"), tr("Bitte im Tab 'Datenbank' ausfüllen: ") + ", ".join(miss) + ".")
                    self.nb.select(self.tab_db)
                    return
                self.lbl_dbstatus.configure(text=tr("Rufe Daten ab …"))
                self.update_idletasks()
                try:
                    table = self._db_table()
                except HMAError as e:
                    self.lbl_dbstatus.configure(text=tr("Fehler"))
                    self.nb.select(self.tab_db)
                    messagebox.showerror(tr("Datenbank"), str(e))
                    return
                except Exception:
                    self.lbl_dbstatus.configure(text=tr("Fehler"))
                    messagebox.showerror(tr("Datenbank"), traceback.format_exc())
                    return
                self._set_table(table)
                self.lbl_dbstatus.configure(text=tr('OK – {0} Datensätze').format(len(table[1])))
            else:
                if not self.files:
                    messagebox.showwarning(tr("Matrix erzeugen"), tr("Bitte zuerst eine Tabelle wählen."))
                    return
                if not all(Path(f).exists() for f in self.files):
                    messagebox.showwarning(tr("Matrix erzeugen"), tr("Die gewählte Tabelle existiert nicht mehr am "
                                           "gemerkten Ort.\nBitte die Tabelle neu wählen."))
                    return
                self.reload_table(first_load=False)  # Tabelle neu einlesen (z. B. nach neuem Datenbank-Export)
            schema = normalize_schema(self.gather_schema())
            req, opt = check_schema(schema, self.headers) if self.headers else ([tr("Keine Daten geladen")], [])
            if req or opt:
                messagebox.showwarning(tr("Zuordnung prüfen"), "\n".join(req + opt))
                return
            self.persist_last()
            self._busy = True
            self.btn_run.configure(state="disabled")
            self.nb.select(self.tab_log)
            self.clear_log()
            done, failed, report_all = [], [], []
            out_dir = Path(self.v_out.get().strip()) if self.v_out.get().strip() else None
            fmt = self.set_vars["out_format"].get()
            ext = ".json" if fmt == "hme" else ".hmcx"
            if dbmode:
                name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", self.db_vars["out_name"].get().strip()) or tr("Datenbank_Export")
                if out_dir is None:
                    docs = Path.home() / "Documents"
                    out_dir = docs if docs.is_dir() else Path.home()
                jobs = [(Path(name + ".csv"), out_dir / (name + ext))]
            else:
                jobs = [(f, (out_dir or f.parent) / (output_stem(f, self.v_sheet.get(), len(self.sheets)) + ext))
                        for f in self.files]
            for f, target in jobs:
                hmc_hwnd = hmc_state = None
                hmc_closed = False
                if target.exists() and os.name == "nt" and fmt == "hmc":
                    hmc_hwnd = hmc_find_window(target.name)
                if target.exists():
                    skip_q = dbmode and self.v_overwrite.get()
                    if hmc_hwnd and not skip_q:
                        ans = messagebox.askyesnocancel(
                            tr("Datei existiert bereits"),
                            tr('{0} existiert bereits in\n'
                               '{1}\n'
                               'und ist gerade in Harris Matrix Composer geöffnet.\n'
                               '\n'
                               'Ja = überschreiben; HMC-Fenster wird OHNE Speichern geschlossen und die neue Datei an derselben '
                               'Stelle wieder geöffnet\n'
                               'Nein = unter neuem Namen speichern\n'
                               'Abbrechen = diese Datei überspringen').format(target.name, target.parent))
                    elif not skip_q:
                        ans = messagebox.askyesnocancel(
                            tr("Datei existiert bereits"),
                            tr('{0} existiert bereits in\n'
                               '{1}\n'
                               '\n'
                               'Ja = überschreiben\n'
                               'Nein = unter neuem Namen speichern\n'
                               'Abbrechen = diese Datei überspringen').format(target.name, target.parent))
                    else:
                        ans = True
                    if ans is None:
                        continue
                    if ans is False:
                        hmc_hwnd = None
                        n = 2
                        while (target.parent / f"{f.stem}_{n}{ext}").exists():
                            n += 1
                        target = target.parent / f"{f.stem}_{n}{ext}"
                    elif hmc_hwnd:
                        hmc_state = hmc_get_state(hmc_hwnd)
                        self.log_line(tr("HMC wird geschlossen (ohne Speichern) …"))
                        self.update_idletasks()
                        if hmc_close(hmc_hwnd):
                            hmc_closed = True
                        else:
                            self.log_line(tr("HMC konnte nicht automatisch geschlossen werden. Die Datei wird trotzdem "
                                          "erzeugt; bitte in HMC danach schließen und neu öffnen."))
                self.log_line(tr("Verarbeite: ") + (tr('Datenbank ({0})').format(self.v_provider.get()) if dbmode else f.name))
                try:
                    res = convert(f, schema, target, sheet=self.v_sheet.get(), log=self.log_line, table=table, out_format=fmt)
                    done.append(res)
                    report_all.append(res["report"])
                    self.log_text(res["report"])
                    self.log_line(tr('→ gespeichert: {0}').format(res['hmcx']))
                    self.log_line("")
                except HMAError as e:
                    failed.append(f.name)
                    self.log_line(tr("FEHLER bei ") + f.name)
                    self.log_text(str(e))
                    report_all.append(tr('FEHLER bei {0}\n{1}').format(f.name, e))
                except Exception:
                    failed.append(f.name)
                    tb = traceback.format_exc()
                    self.log_line(tr("FEHLER (unerwartet):"))
                    self.log_text(tb)
                    report_all.append(tb)
                if hmc_closed and target.exists():
                    self.reopen_in_hmc(state=hmc_state, target=target)
            self._report = "\n".join(report_all)
            self._busy = False
            self.btn_run.configure(state="normal")
            if done:
                self._last_hmcx = done[-1]["hmcx"]
                self.btn_open_file.configure(state="normal")
                self.btn_open_dir.configure(state="normal")
            total_w = sum(r["warnings"] + r["conflicts"] for r in done)
            if failed:
                self.set_status(tr('Keine Datei erzeugt für: {0} – Details im Protokoll').format(', '.join(failed)), False)
                messagebox.showerror(tr("Matrix erzeugen"), tr("Es wurde keine Datei erzeugt, weil die Daten logische "
                                                        "Fehler enthalten.\n\nDetails und Lösungsvorschläge stehen "
                                                        "im Tab 'Protokoll'."))
            elif done:
                self.set_status(tr('✓ {0} von {1} Datei(en) verarbeitet').format(len(done), len(jobs))
                                + (tr(' – {0} Warnung(en), siehe Protokoll').format(total_w) if total_w else tr(" – keine Warnungen")),
                                True)

    App().mainloop()


# ---------------------------------------------------------------------------
# Kommandozeile
# ---------------------------------------------------------------------------
def run_cli(argv):
    import argparse
    ap = argparse.ArgumentParser(description=f"{APP_NAME} v{APP_VERSION}")
    ap.add_argument("--cli", action="store_true")
    ap.add_argument("--lang", choices=list(LANGS), default=None,
                    help=tr("Sprache der Meldungen (sonst: Umgebungsvariable HMA_LANG, Konfiguration, Systemsprache)"))
    ap.add_argument("--config", default=None)
    ap.add_argument("--schema", help=tr("Name eines gespeicherten Mapping-Schemas (sonst: zuletzt verwendeter Stand)"))
    ap.add_argument("--sheet", default="")
    ap.add_argument("--out", default=None, help=tr("Speicherordner für die Ausgabedatei"))
    ap.add_argument("--format", choices=["hmc", "hme"], default=None,
                    help=tr("Ausgabe: hmc = Harris Matrix Composer (.hmcx), hme = Harris Matrix Editor (.json)"))
    ap.add_argument("files", nargs="+")
    a = ap.parse_args(argv)
    store = ConfigStore(a.config or pick_config_path())
    if a.schema:
        if a.schema not in store.data["schemas"]:
            print(tr("Schema nicht gefunden. Vorhanden:"), ", ".join(sorted(store.data["schemas"])))
            return 2
        schema = store.data["schemas"][a.schema]
    else:
        schema = normalize_schema(store.data.get("last", {}))
    schema = normalize_schema(schema)
    if a.format:
        schema["settings"]["out_format"] = a.format
    ext = ".json" if schema["settings"]["out_format"] == "hme" else ".hmcx"
    rc = 0
    for f in a.files:
        print(tr('Verarbeite: {0}').format(f))
        try:
            n_sheets = len(sheet_names(Path(f))) if Path(f).suffix.lower() in SHEET_EXT else 0
        except Exception:
            n_sheets = 0
        out = (Path(a.out) if a.out else Path(f).parent) / (output_stem(f, a.sheet, n_sheets) + ext)
        try:
            res = convert(f, schema, out, sheet=a.sheet)
            print(res["report"])
        except HMAError as e:
            print(tr("FEHLER:"), e)
            rc = 1
    return rc


if __name__ == "__main__":
    if "--cli" in sys.argv:
        sys.exit(run_cli(sys.argv[1:]))
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("HMA.Adapter.1")
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    try:
        run_gui()
    except ImportError as e:
        print(tr("Die Oberfläche benötigt tkinter (gehört zur Standard-Python-Installation von python.org)."), e)
        print(tr("Alternativ: python {0} --cli --help").format(Path(__file__).name))
        sys.exit(1)
