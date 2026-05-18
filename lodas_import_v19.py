"""
LODAS Importskript  lodas_import_v19
====================================
Liest Zeitnachweis-Excel-Dateien aus einem Monatsordner und erzeugt
eine LODAS-ASCII-Importdatei.

Neu in v18 (Mai 2026):
  - LA 100 wird 1:1 mit den IST-Stunden aus dem Zeitnachweis gebucht.
    Die SOLL-Deckelung (min(IST, SOLL)) wird entfernt.
    Hintergrund: DATEV zieht AZK+-Stunden (LA 122) vom Abrechnungsbetrag
    ab. Damit das korrekt funktioniert, muessen die vollen IST-Stunden
    zunaechst ueber LA 100 eingebucht sein. DATEV rechnet dann:
    LA 100 (IST) - LA 122 (AZK+) = Nettoauszahlung (SOLL-Stunden).
    Eine Deckelung in LA 100 wuerde zu einer Unterzahlung fuehren.

Neu in v17 (Mai 2026):
  - AZK-Buchung (LA 122/124) = reine Ueberstunden aus normalen Arbeitstagen
    (IST - SOLL). FZA (LA 123) und Garantiezeit (LA 210) sind eigenstaendige
    Buchungen und werden nicht aus der AZK-Berechnung abgezogen oder addiert.
    Das Leitsystem verrechnet diese intern selbst.

Neu in v16 (Mai 2026):
  - Negative Stunden im Zeitnachweis werden generell ignoriert.
    Ersetzt den fragilen String-Vergleich auf '-SOLL Feiertag'.
    SOLL-Gegenbuchungen des Leitsystems haben immer negative Stunden-
    werte und sind keine echten Arbeitsstunden. Robuster gegen
    Schreibvarianten (Leerzeichen, fehlender Kommentar etc.).

Neu in v15 (Mai 2026):
  - Stunden-Rundung beim Einlesen: float-Werte aus der Stunden-Spalte des
    Zeitnachweis-Sheets werden auf 2 Nachkommastellen gerundet (round(..., 2)).
    Verhindert Gleitkomma-Akkumulation bei Werten wie 5.3333... (5h 20min).
    Damit stimmt die AZK-Eigenberechnung auf 2 Dezimalstellen mit dem
    Leitsystem-Saldo ueberein.

Neu in v14 (Mai 2026):
  - Datumsfilter: Zeilen im Zeitnachweis-Sheet ausserhalb des Abrechnungsmonats
    werden ignoriert (z.B. Unterschriftszeile mit Folgedatum). Verhindert
    fehlerhafte Einbeziehung von 0h-Zeilen mit Datum ausserhalb des Monats.
  - AZK-Netto-Berechnung: LA 122/124 wird jetzt korrekt aus drei Komponenten
    ermittelt:
      azk_netto = Ueberstunden_normale_Tage - FZA-Abzug + AZK-Kommentar-Buchungen
    FZA-Stunden (LA 123) werden als AZK-Abbau abgezogen.
    AZK-Kommentar-Buchungen (z.B. Garantiezeit LA 210, negativ = AZK+) werden
    addiert. Damit stimmt die Eigenberechnung mit dem Leitsystem-Saldo ueberein.

Neu in v13 (Mai 2026):
  - Garantiezeit (und alle AZK-Lohnarten mit Kommentar-Schluessel) werden
    jetzt mit ihrer eigenen la_nr aus stammdaten.xlsx gebucht, nicht mehr
    hardcodiert als LA 122. Garantiezeit -> LA 210 (negativ, AZK+).
    LA 122 bleibt ausschliesslich der normalen AZK-Eigenberechnung vorbehalten.

Neu in v12 (Mai 2026):
  - AZK-Berechnung komplett umgestellt: Statt Saldo-Differenz aus dem
    Leitsystem werden Ueberstunden jetzt selbst berechnet:
    AZK = Summe(IST - SOLL) ueber alle normalen Arbeitstage (ohne Kommentar).
    Positiv (Ueberstunden) -> LA 122 mit negativem Wert.
    Negativ (Minderstunden) -> LA 124 mit positivem Wert.
  - Leitsystem-Saldo wird weiterhin gelesen, dient aber nur noch als
    Kontrollgroesse im Report (Hinweis: Eigen vs. Leitsystem-Saldo).
  - Garantiezeit mit AZK=ja: wird zusaetzlich als eigene LA 122-Buchung
    erzeugt (ueber verarbeite_kommentar_block, unveraendert aus v11).

Neu in v11 (Mai 2026):
  - Garantiezeit (LA 202) wird als AZK-Buchung behandelt: Kommentar
    "Garantiezeit" wird erkannt, Stunden summiert und als LA 122 mit
    negativem Wert gebucht (AZK+). Voraussetzung: LA 202 hat AZK=ja
    in stammdaten.xlsx. AZK-Lohnarten mit Kommentar-Schluessel werden
    jetzt im schluessel_index aufgenommen und gesondert verarbeitet.

Neu in v10 (Mai 2026):
  - Kommentar "-SOLL Feiertag" wird komplett ignoriert (SOLL-Gegenbuchung
    aus Leitsystem). Zeilen mit diesem Kommentar erzeugen keine Buchung
    und fliessen nicht in LA 100.
  - Garantiezeit (LA 202) wird nicht mehr ausbezahlt, sondern ins AZK
    geschoben. Die Garantiezeit-Stunden werden als LA 122 mit negativem
    Wert gebucht (identisch zur AZK+-Logik). Die Steuerung erfolgt ueber
    das AZK-Flag in stammdaten.xlsx – im Skript wird Garantiezeit nicht
    gesondert behandelt.

Neu in v9 (Mai 2026):
  - Null-Buchungen werden nicht mehr in die Importdatei geschrieben.
    Bewegungsdaten mit Wert 0.00 werden beim Schreiben uebersprungen.

Neu in v8 (Mai 2026):
  - LA 100 SOLL-Deckelung: IST-Stunden werden pro Tag auf den SOLL-Wert
    aus stammdaten.xlsx gedeckelt. Ueberstunden fliessen nicht mehr in
    LA 100, sondern werden ueber die Saldo-Differenz (LA 122/124) abgebildet.
    Formel: LA100_Tag = min(IST_Tag, SOLL)
    Report-Hinweis wenn Deckelung greift (Anzahl Tage + Gesamtdifferenz).

Neu in v7 (Mai 2026):
  - Krank bei Lohnempfaengern: Keine Fehlzeit-Buchung (103) mehr.
    Nur LA 320 (Bewegungsdaten) wird gebucht.
  - Krank bei Gehaltsempfaengern: Weder Fehlzeit 103 noch LA 320.
    Krank wird fuer Gehaltsempfaenger komplett ignoriert.

Neu in v6 (Mai 2026):
  - AZK-Rundungshinweis: Wenn die gebuchte AZK-Stundenzahl (LA 122/124)
    von der Saldo-Differenz des Leitsystems abweicht (jede Abweichung > 0),
    wird ein Hinweis im Report ausgegeben.

Neu in v5 (Mai 2026):
  - Unbezahlte Fehlzeit: Kein Abzug mehr von LA 100.
    Es wird ausschliesslich eine Fehlzeit-Buchung (Schluessel 22) erzeugt.
    Die Stunden der unbezahlten Fehlzeit stehen im Zeitnachweis ohnehin
    nicht als Arbeitsstunden – ein rechnerischer Abzug war daher falsch.
    Entfernt: unbezahlt_schluessel-Set, unbezahlt_tage-Zaehlung,
    Abzugs-Berechnung in verarbeite_kommentar_block() und
    verarbeite_zeitnachweis().

Neu in v4 (Mai 2026):
  - AZK-Buchung (LA 122/124) wird jetzt aus der Saldo-Differenz der
    Uebersicht-Tabelle im Zeitnachweis-Sheet ermittelt (v2-Logik),
    statt aus der Auszahlungsgrenze. Die Auszahlungsgrenze steuert
    die AZK-Buchung nicht mehr.

Neu in v3 (Mai 2026):
  - LA 303 (Urlaub, Tagebasis) wird jetzt auch fuer Lohnempfaenger
    aus dem Abwesenheiten-Sheet gebucht (nicht nur fuer Gehaltsempfaenger).
    Urlaubstage werden gezaehlt und als Bewegungsdaten mit bs_nr=71 erzeugt.

Aenderungen aus v4.5:
  - Gehaltsempfaenger-Erkennung ueber Spalte "Bemerkung" = "Gehalt" in stammdaten.xlsx

Neu in v4.4:
  - Pfadeingabe entfaellt: Zeitnachweis-Ordner wird automatisch aus dem
    Abrechnungszeitraum (MMYYYY) zusammengesetzt:
    C:\\Users\\eikeh\\01_ImportAPI\\Lohnexport\\MMYYYY

Neu in v4.3:
  - Gehaltsempfaenger-Logik (Backlog #10)
    Identifikation: Bemerkung = "Gehalt" in stammdaten.xlsx
    Gehaltsempfaenger erhalten: Fehlzeit Krank (103), Fehlzeit Kind krank (200),
    Bewegungsdaten Urlaub LA 303 (Tagebasis) – alle anderen Buchungen entfallen.

Neu in v4.2:
  - Vorschuss/Abschlag-Verarbeitung (Backlog #15)
    Kommentar "Abschlag <Betrag>" im Zeitnachweis wird erkannt.
    Erzeugt: Vorschusstabelle (u_lod_bwd_buchung_vorschuss) + BWD LA 9070 bs_nr 3 (negativ)

Neu in v4.1 (Refactoring & Konsolidierung):
  - clean() und clean_val() zu einer Funktion zusammengefasst (none_if_empty)
  - standardregeln-Liste wird nur noch einmal aufgebaut
  - Schleife ueber kommentar_tage + abwesenheiten_tage zusammengefasst
  - Spaltennamen-Normalisierung als eigene Funktion norm_col()
  - lese_abwesenheiten: Datumsiteration mit timedelta statt manuell
  - SEP / SEP2 als Klassenattribute in ReportCollector
  - Codestruktur in klar benannte Bloecke gegliedert

Neu in v4.0 (uebernommen aus lodas_maerz_v1.x):
  - LA 301 und LA 123 werden mit SOLL-Stunden gebucht statt IST
  - Direkt-Buchungen fliessen in Auszahlungssumme ein (wenn Auszahlung=ja)
  - LA 202: Direktwert aus Zeitnachweis statt Tage x IST
  - soll_stunden_pro_tag als Parameter in verarbeite_kommentar_block()
  - lese_saldo_differenz(): Leitsystem-Endsaldo aus Zeitnachweis-Sheet
  - fza_stunden (LA 123) werden im ReportCollector gesammelt
  - lodas_sync_MMYYYY.txt: Sync-Report Leitsystem-Saldo vs. AZK-Buchung

Neu in v3.2:
  - End-Report: Zusammenfassung pro Mitarbeiter (LA100, AZK, Fehlzeiten)
  - Report wird in Konsole ausgegeben UND als lodas_report_MMYYYY.txt gespeichert

Neu in v3.1:
  - Name wird aus Zeile 1 der xlsx-Datei gelesen (Spalte C: "Nachname, Vorname")
    statt aus dem Dateinamen.

Neu in v3.0:
  - Abwesenheiten-Sheet wird zusaetzlich eingelesen.
  - Doppelbuchungen werden verhindert (Schluessel-Ebene).

Aufruf:
  python lodas_import_v12.py
"""

VERSION = "lodas_import_v19"

# ===========================================================================
# BLOCK 1 – Imports & Logging
# ===========================================================================
import os
import re
import logging
from collections import defaultdict
from datetime import datetime, timedelta

import pandas as pd

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger(__name__)

# ===========================================================================
# BLOCK 2 – Konfiguration & Konstanten
# ===========================================================================
STAMMDATEN_PFAD = r"C:\Users\eikeh\01_ImportAPI\stammdaten.xlsx"

# ===========================================================================
# BLOCK 3 – Hilfsfunktionen
# ===========================================================================
def safe_float(val, default=0.0):
    try:
        f = float(val)
        return default if pd.isna(f) else f
    except (TypeError, ValueError):
        return default

def clean(val, none_if_empty=False):
    """Bereinigt einen Wert (strip, Zeilenumbrueche entfernen).
    none_if_empty=True  -> gibt None zurueck wenn leer/nan
    none_if_empty=False -> gibt ''  zurueck wenn leer/nan
    """
    s = str(val or "").strip().replace("\r", "").replace("\n", "")
    empty = None if none_if_empty else ""
    return empty if s in ("", "nan") else s

def norm_col(col):
    """Normalisiert einen Spaltennamen (Whitespace zusammenfassen, strip)."""
    return re.sub(r"[\r\n\s]+", " ", str(col)).strip()

def datum_aus_zelle(zelle):
    if pd.isna(zelle):
        return None
    s = str(zelle).strip()
    m = re.match(r"(\d{2}\.\d{2}\.\d{4})", s)
    if m:
        return m.group(1)
    try:
        return pd.Timestamp(zelle).strftime("%d.%m.%Y")
    except Exception:
        return None

def bwd_satz(pnr, zeitraum, wert, bs_nr, la_nr):
    return {
        "pnr":                 pnr,
        "abrechnungszeitraum": zeitraum,
        "wert":                wert,
        "bs_nr":               bs_nr,
        "la_nr":               la_nr,
    }

def parse_mmyyyy(mmyyyy_str):
    mmyyyy_str = mmyyyy_str.strip()
    if not re.match(r"^\d{6}$", mmyyyy_str):
        raise ValueError(f"Ungaeltiges Format '{mmyyyy_str}' - erwartet: MMYYYY (z. B. 032026)")
    mm = mmyyyy_str[:2]
    yyyy = mmyyyy_str[2:]
    if not (1 <= int(mm) <= 12):
        raise ValueError(f"Ungueltiger Monat '{mm}'")
    return f"01.{mm}.{yyyy}"

def ask(prompt, default=None):
    if default:
        val = input(f"{prompt} [{default}]: ").strip()
        return val if val else default
    return input(f"{prompt}: ").strip()

# ===========================================================================
# BLOCK 4 – ReportCollector
# ===========================================================================
class ReportCollector:
    SEP  = "=" * 70
    SEP2 = "-" * 70

    def __init__(self):
        self.mitarbeiter   = []
        self.fehler        = []
        self.uebersprungen = []

    def add_ma(self, name, pnr, la100_h, azk_la, azk_h, fehlzeiten_anzahl,
               saldo_differenz=None, fza_stunden=0.0):
        self.mitarbeiter.append({
            "name":              name,
            "pnr":               pnr,
            "la100_h":           la100_h,
            "azk_la":            azk_la,
            "azk_h":             azk_h,
            "fehlzeiten_anzahl": fehlzeiten_anzahl,
            "saldo_differenz":   saldo_differenz,
            "fza_stunden":       fza_stunden,
        })

    def hinweis(self, ma, text):
        self.fehler.append({"ebene": "HINWEIS", "ma": ma, "text": text})

    def fehler_add(self, ma, text):
        self.fehler.append({"ebene": "FEHLER", "ma": ma, "text": text})

    def skip(self, dateiname, grund):
        self.uebersprungen.append(dateiname)
        self.fehler.append({"ebene": "FEHLER", "ma": dateiname, "text": grund})

    def render(self, az_eingabe, az_anzeige, gesamt_bwd, gesamt_fz, encoding="ansi"):
        col_name=22; col_pnr=5; col_la100=10; col_azk=14; col_fz=12; col_hw=8
        zeilen = []
        zeilen.append(self.SEP)
        zeilen.append(f"  LODAS Import Report  [{VERSION}]  -  {az_eingabe}  ({az_anzeige})")
        zeilen.append(self.SEP)
        zeilen.append("")
        zeilen.append(
            f"  {'Mitarbeiter':<{col_name}} {'PNr':<{col_pnr}} "
            f"{'LA 100':>{col_la100}} {'AZK-Buchung':>{col_azk}} "
            f"{'Fehlzeiten':>{col_fz}} {'Hinweise':>{col_hw}}"
        )
        zeilen.append("  " + self.SEP2)
        for ma in self.mitarbeiter:
            hinweis_count = sum(1 for e in self.fehler if e["ma"] == ma["name"])
            fehler_count  = sum(1 for e in self.fehler if e["ma"] == ma["name"] and e["ebene"] == "FEHLER")
            la100_str = f"{ma['la100_h']:.2f}h" if ma["la100_h"] is not None else "-"
            if ma["azk_la"] == "122":   azk_str = f"LA122 -{ma['azk_h']:.2f}h"
            elif ma["azk_la"] == "124": azk_str = f"LA124 +{ma['azk_h']:.2f}h"
            else:                       azk_str = "-"
            fz_str = f"{ma['fehlzeiten_anzahl']}x Fehlzeit" if ma["fehlzeiten_anzahl"] > 0 else "-"
            if fehler_count > 0:    hw_str = f"{hinweis_count}x E"
            elif hinweis_count > 0: hw_str = f"{hinweis_count}x !"
            else:                   hw_str = "OK"
            zeilen.append(
                f"  {ma['name']:<{col_name}} {ma['pnr']:<{col_pnr}} "
                f"{la100_str:>{col_la100}} {azk_str:>{col_azk}} "
                f"{fz_str:>{col_fz}} {hw_str:>{col_hw}}"
            )
        if self.uebersprungen:
            zeilen.append("  " + self.SEP2)
            for _ in self.uebersprungen:
                zeilen.append(
                    f"  {'[uebersprungen]':<{col_name}} {'':>{col_pnr}} "
                    f"{'':>{col_la100}} {'':>{col_azk}} {'':>{col_fz}} {'FEHLER':>{col_hw}}"
                )
        zeilen.append("")
        zeilen.append(self.SEP)
        zeilen.append(f"  Verarbeitete Mitarbeiter:  {len(self.mitarbeiter)}")
        zeilen.append(f"  Uebersprungene Dateien:    {len(self.uebersprungen)}")
        zeilen.append(f"  Erzeugte Bewegungsdaten:   {gesamt_bwd}")
        zeilen.append(f"  Erzeugte Fehlzeiten:       {gesamt_fz}")
        zeilen.append(self.SEP)
        if self.fehler:
            zeilen.append("")
            zeilen.append("  HINWEISE & FEHLER")
            zeilen.append("  " + self.SEP2)
            for e in self.fehler:
                symbol = "!!" if e["ebene"] == "FEHLER" else " !"
                zeilen.append(f"  [{symbol}] [{e['ma']}]  {e['text']}")
            zeilen.append(self.SEP)
        else:
            zeilen.append("")
            zeilen.append("  Keine Hinweise oder Fehler - Verarbeitung fehlerfrei.")
            zeilen.append(self.SEP)
        return zeilen

    def render_sync(self, az_eingabe, az_anzeige):
        col_pnr=6; col_name=28; col_sal=10; col_azk=10; col_nach=11
        zeilen = []
        zeilen.append(self.SEP)
        zeilen.append(f"  LODAS Sync-Report  [{VERSION}]  -  {az_eingabe}  ({az_anzeige})")
        zeilen.append(self.SEP)
        zeilen.append(
            f"  {'PNr':<{col_pnr}} {'Name':<{col_name}} "
            f"{'Saldo LS':>{col_sal}} {'AZK Netto':>{col_azk}} "
            f"{'Nachbuchung':>{col_nach}}"
        )
        zeilen.append("  " + self.SEP2)
        for ma in self.mitarbeiter:
            saldo = ma["saldo_differenz"]
            if ma["azk_la"] == "122":
                azk_netto = -ma["azk_h"]
            elif ma["azk_la"] == "124":
                azk_netto = ma["azk_h"]
            else:
                azk_netto = 0.0
            azk_str = f"{azk_netto:+.2f}"
            if saldo is not None:
                saldo_str       = f"{saldo:+.2f}"
                nachbuchung_str = f"{round(saldo - azk_netto, 2):+.2f}"
            else:
                saldo_str       = "n/a"
                nachbuchung_str = "n/a"
            zeilen.append(
                f"  {ma['pnr']:<{col_pnr}} {ma['name']:<{col_name}} "
                f"{saldo_str:>{col_sal}} {azk_str:>{col_azk}} "
                f"{nachbuchung_str:>{col_nach}}"
            )
        zeilen.append(self.SEP)
        return zeilen


report = ReportCollector()

# ===========================================================================
# BLOCK 5 – Stammdaten einlesen
# ===========================================================================
print("=" * 60)
print(f"  LODAS Importskript  {VERSION}")
print("=" * 60)
print(f"  Stammdaten: {STAMMDATEN_PFAD}")
print("=" * 60)

while True:
    az_eingabe = ask("Abrechnungszeitraum (Format MMYYYY, z. B. 032026)")
    try:
        abrechnungszeitraum_global = parse_mmyyyy(az_eingabe)
        log.info(f"Abrechnungszeitraum: {abrechnungszeitraum_global}")
        break
    except ValueError as e:
        print(f"  Fehler: {e} - bitte erneut eingeben.")

zeitnachweis_ordner = rf"C:\Users\eikeh\01_ImportAPI\Lohnexport\{az_eingabe}"
log.info(f"Zeitnachweis-Ordner: {zeitnachweis_ordner}")

def read_sheet_find_header(pfad, sheet, schluessel_spalte):
    raw = pd.read_excel(pfad, sheet_name=sheet, header=None)
    for i, row in raw.iterrows():
        if any(schluessel_spalte in str(v) for v in row.values):
            raw.columns = [norm_col(c) for c in raw.iloc[i]]
            df = raw.iloc[i + 1:].reset_index(drop=True)
            return df
    raise ValueError(f"Spalte mit '{schluessel_spalte}' nicht im Sheet '{sheet}' gefunden.")

log.info("Lese Stammdaten...")

if not os.path.exists(STAMMDATEN_PFAD):
    log.error(f"stammdaten.xlsx nicht gefunden: {STAMMDATEN_PFAD}")
    exit(1)

raw_mandant = pd.read_excel(STAMMDATEN_PFAD, sheet_name="Mandant", header=None)
mandant = {}
header_found = False
for i, row in raw_mandant.iterrows():
    vals = [str(v).strip() for v in row.values]
    if "Feld" in vals and not header_found:
        header_found = True
        continue
    if header_found:
        k = clean(row.iloc[0])
        v = clean(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
        if k:
            mandant[k] = v

df_mitarbeiter = read_sheet_find_header(STAMMDATEN_PFAD, "Mitarbeiter", "Name (wie in Zeitnachweis)")
df_mitarbeiter = df_mitarbeiter.dropna(subset=["Name (wie in Zeitnachweis)"])

mitarbeiter_map = {}
for _, row in df_mitarbeiter.iterrows():
    name = clean(row["Name (wie in Zeitnachweis)"])
    if not name:
        continue
    pnr = clean(row.get("Personalnummer", "")) or ""
    if not re.match(r"^\d{1,3}$", pnr):
        log.warning(f"Personalnummer '{pnr}' fuer '{name}' ignoriert")
        continue
    azg_raw    = clean(row.get("Auszahlungsgrenze", "") or "")
    bemerkung  = clean(row.get("Bemerkung", "") or "")
    ist_gehalt = bemerkung.lower() == "gehalt"
    mitarbeiter_map[name] = {
        "pnr":               pnr,
        "soll":              safe_float(row.get("Tages Arbeitszeit SOLL")),
        "ist":               safe_float(row.get("Tages Arbeitszeit IST")),
        "auszahlungsgrenze": safe_float(azg_raw),  # bei Gehalt ohnehin ignoriert
        "ist_gehalt":        ist_gehalt,
    }

df_lohnarten = read_sheet_find_header(STAMMDATEN_PFAD, "Lohnarten", "Bezeichnung")
df_lohnarten = df_lohnarten[
    df_lohnarten.apply(
        lambda r: bool(clean(r.get("Lohnart-Nr."), none_if_empty=True) or
                       clean(r.get("Fehlzeiten- Schluessel"), none_if_empty=True)),
        axis=1
    )
]

lohnarten_regeln = []
for _, row in df_lohnarten.iterrows():
    ziel = clean(row.get("Ziel", "") or "")
    if not ziel:
        continue
    lohnarten_regeln.append({
        "bezeichnung":   clean(row.get("Bezeichnung", "") or ""),
        "schluessel":    clean(row.get("Kommentar- Schluesselwort"), none_if_empty=True),
        "la_nr":         clean(row.get("Lohnart-Nr."), none_if_empty=True),
        "einheit":       clean(row.get("Einheit", "") or ""),
        "ziel":          ziel,
        "fehlzeit_key":  clean(row.get("Fehlzeiten- Schluessel"), none_if_empty=True),
        "bs_nr":         clean(row.get("Buchungs- schluessel (bs_nr)"), none_if_empty=True),
        "azk":           clean(row.get("AZK", "") or "").lower() == "ja",
        "auszahlung":    clean(row.get("Auszahlung", "") or "").lower() == "ja",
        "direkt":        clean(row.get("Direkt", "") or "").lower() == "ja",
    })

schluessel_index = {}
for regel in lohnarten_regeln:
    k = regel["schluessel"]
    if k:
        schluessel_index.setdefault(k, []).append(regel)

la122 = next((r for r in lohnarten_regeln if r["la_nr"] == "122"), None)
la124 = next((r for r in lohnarten_regeln if r["la_nr"] == "124"), None)

if not la122:
    log.warning("LA 122 nicht gefunden - AZK-Ueberschuss-Buchung deaktiviert")
if not la124:
    log.warning("LA 124 nicht gefunden - AZK-Unterschuss-Buchung deaktiviert")

log.info(f"  {len(mitarbeiter_map)} Mitarbeiter geladen")
log.info(f"  {len(lohnarten_regeln)} Lohnarten-Regeln geladen")

# ===========================================================================
# BLOCK 6 – Zeitnachweis verarbeiten
# ===========================================================================
def name_aus_zeitnachweis(datei_pfad):
    try:
        df_raw = pd.read_excel(datei_pfad, sheet_name="Zeitnachweis", header=None)
        name_raw = str(df_raw.iloc[0, 2]).strip()
        teile = name_raw.split(",", 1)
        if len(teile) == 2:
            return f"{teile[0].strip()}, {teile[1].strip()}"
    except Exception as e:
        log.error(f"  Name aus Zeitnachweis-Sheet konnte nicht gelesen werden: {e}")
    return None

def lese_azk_saldo_differenz(datei_pfad, mitarbeiter_name):
    """
    Liest die Saldo-Differenz aus der Uebersicht-Tabelle im Zeitnachweis-Sheet.
    Anker: Zeile mit Label 'Uebersicht' (positionsunabhaengig).
    Danach: Zeile mit Label 'Saldo' – erste zwei numerische Werte = Vormonat / aktueller Monat.
    Differenz = aktueller Monat - Vormonat.
    Positiv → AZK gestiegen → LA 122 mit negativem Wert buchen.
    Negativ → AZK gesunken  → LA 124 mit positivem Wert buchen.
    """
    try:
        df_raw = pd.read_excel(datei_pfad, sheet_name="Zeitnachweis", header=None)
    except Exception as e:
        log.warning(f"  Saldo-Differenz konnte nicht gelesen werden: {e}")
        return None

    in_uebersicht   = False
    saldo_differenz = None

    for i, row in df_raw.iterrows():
        label = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ""

        if not in_uebersicht:
            if label in ("uebersicht", "ubersicht", "\u00fcbersicht"):
                in_uebersicht = True
            continue

        if "saldo" not in label:
            continue

        werte = [
            float(cell) for cell in row.iloc[1:]
            if pd.notna(cell)
            and str(cell).strip() not in ("", "None")
            and isinstance(cell, (int, float))
        ]

        if len(werte) >= 2:
            saldo_vormonat  = werte[0]
            saldo_aktuell   = werte[1]
            saldo_differenz = round(saldo_aktuell - saldo_vormonat, 2)
            break

    if saldo_differenz is None:
        report.hinweis(mitarbeiter_name,
            "Saldo-Differenz nicht gefunden \u2013 keine AZK-Buchung erzeugt.")
        return None

    if saldo_differenz > 0:
        log.info(f"  AZK gestiegen: +{saldo_differenz}h \u2192 LA 122 mit -{saldo_differenz}h")
    elif saldo_differenz < 0:
        log.info(f"  AZK gesunken: {saldo_differenz}h \u2192 LA 124 mit +{abs(saldo_differenz)}h")
    else:
        log.info("  Saldo unveraendert \u2013 keine AZK-Buchung.")

    return saldo_differenz

def lese_abwesenheiten(datei_pfad):
    eintraege = []
    try:
        xl = pd.ExcelFile(datei_pfad)
        if "Abwesenheiten" not in xl.sheet_names:
            return eintraege
        raw = pd.read_excel(datei_pfad, sheet_name="Abwesenheiten", header=None)
        header_idx = None
        for i, row in raw.iterrows():
            if any("Beginn" in str(v) for v in row.values):
                header_idx = i
                break
        if header_idx is None:
            return eintraege
        raw.columns = [norm_col(c) for c in raw.iloc[header_idx]]
        df = raw.iloc[header_idx + 1:].reset_index(drop=True)
        for _, row in df.iterrows():
            beginn = datum_aus_zelle(row.get("Beginn"))
            ende   = datum_aus_zelle(row.get("Ende"))
            art    = clean(row.get("Art") or "")
            if not beginn or not art:
                continue
            dt_von = datetime.strptime(beginn, "%d.%m.%Y")
            dt_bis = datetime.strptime(ende, "%d.%m.%Y") if ende else dt_von
            dt = dt_von
            while dt <= dt_bis:
                eintraege.append({"datum": dt.strftime("%d.%m.%Y"), "stunden": 0.0, "kommentar": art})
                dt += timedelta(days=1)
    except Exception as e:
        log.warning(f"  Abwesenheiten-Sheet konnte nicht gelesen werden: {e}")
    return eintraege

def verarbeite_kommentar_block(schluessel, zeilen_sorted, pnr,
                                ist_stunden_pro_tag, soll_stunden_pro_tag,
                                ergebnis, auszahlungsstunden, direkt_buchungen):
    regeln = schluessel_index.get(schluessel, [])
    stunden_gezaehlt = False

    for regel in regeln:
        if regel["ziel"] == "Fehlzeiten":
            # Krank (103) wird fuer Lohnempfaenger nicht als Fehlzeit gebucht
            if regel["fehlzeit_key"] == "103":
                continue
            zeitraeume = []
            start = prev = None
            for z in zeilen_sorted:
                dt = datetime.strptime(z["datum"], "%d.%m.%Y")
                if start is None:
                    start = prev = dt
                else:
                    if (dt - prev).days > 3:
                        zeitraeume.append((start, prev))
                        start = dt
                    prev = dt
            if start:
                zeitraeume.append((start, prev))
            for von_dt, bis_dt in zeitraeume:
                ergebnis["fehlzeiten"].append({
                    "pnr":          pnr,
                    "datum_von":    von_dt.strftime("%d.%m.%Y"),
                    "datum_bis":    bis_dt.strftime("%d.%m.%Y"),
                    "fehlzeit_key": regel["fehlzeit_key"],
                })

        elif regel["ziel"] == "Bewegungsdaten":
            # AZK-Lohnart mit Kommentar-Schluessel (z.B. Garantiezeit -> AZK)
            if regel["azk"]:
                gesamt_azk_kommentar = sum(z["stunden"] for z in zeilen_sorted)
                if gesamt_azk_kommentar != 0 and regel["la_nr"]:
                    ergebnis["bewegungsdaten"].append(
                        bwd_satz(pnr, abrechnungszeitraum_global,
                                 f"-{gesamt_azk_kommentar:.2f}",
                                 regel["bs_nr"] or "1", regel["la_nr"]))
                    log.info(f"  AZK-Buchung aus {schluessel}: LA {regel['la_nr']} -{gesamt_azk_kommentar:.2f}h")
                continue
            if regel["direkt"]:
                gesamt_direkt = sum(z["stunden"] for z in zeilen_sorted)
                direkt_buchungen.append(
                    bwd_satz(pnr, abrechnungszeitraum_global, f"{gesamt_direkt:.2f}",
                             regel["bs_nr"] or "1", regel["la_nr"] or ""))
                log.info(f"  Direkt-Buchung LA {regel['la_nr']}: {gesamt_direkt:.2f}h")
                if regel["auszahlung"] and not stunden_gezaehlt:
                    auszahlungsstunden += gesamt_direkt
                    stunden_gezaehlt = True
                    log.info(f"  -> LA {regel['la_nr']} Direktwert fliesst in Auszahlungssumme: +{gesamt_direkt:.2f}h")
                continue

            if regel["einheit"] == "Stunden":
                if regel["la_nr"] == "123":
                    gesamt = round(len(zeilen_sorted) * soll_stunden_pro_tag, 2)
                    log.info(f"  LA 123: SOLL-Basis ({soll_stunden_pro_tag}h/Tag)")
                else:
                    gesamt = sum(z["stunden"] for z in zeilen_sorted)
                    log.info(f"  LA {regel['la_nr']}: 1:1 aus Zeitnachweis: {gesamt:.2f}h")
                ergebnis["bewegungsdaten"].append(
                    bwd_satz(pnr, abrechnungszeitraum_global, f"{gesamt:.2f}",
                             regel["bs_nr"] or "1", regel["la_nr"] or ""))
                if regel["auszahlung"] and not stunden_gezaehlt:
                    auszahlungsstunden += gesamt
                    stunden_gezaehlt = True

            elif regel["einheit"] == "Tage":
                anzahl = len(zeilen_sorted)
                ergebnis["bewegungsdaten"].append(
                    bwd_satz(pnr, abrechnungszeitraum_global, str(anzahl),
                             regel["bs_nr"] or "14", regel["la_nr"] or ""))
                if regel["auszahlung"] and not stunden_gezaehlt:
                    auszahlungsstunden += anzahl * ist_stunden_pro_tag
                    stunden_gezaehlt = True

    return auszahlungsstunden

def verarbeite_zeitnachweis(datei_pfad, datei_name):
    ergebnis = {"fehlzeiten": [], "bewegungsdaten": [], "vorschuesse": []}

    name = name_aus_zeitnachweis(datei_pfad)
    if name is None:
        report.skip(datei_name, "Name konnte nicht aus Zeitnachweis-Sheet gelesen werden")
        return ergebnis
    if name not in mitarbeiter_map:
        report.skip(datei_name, f"Mitarbeiter '{name}' nicht in stammdaten.xlsx")
        return ergebnis

    ma                   = mitarbeiter_map[name]
    pnr                  = ma["pnr"]
    soll_stunden_pro_tag = ma["soll"]
    ist_stunden_pro_tag  = ma["ist"]
    auszahlungsgrenze    = ma["auszahlungsgrenze"]
    ist_gehalt           = ma["ist_gehalt"]
    if ist_gehalt:
        log.info(f"  -> Gehaltsempfaenger: nur Fehlzeiten + Urlaub LA 303")

    saldo_differenz = lese_azk_saldo_differenz(datei_pfad, name)

    log.info(f"  Verarbeite: {datei_name} -> PNr {pnr} ({name})")

    df = pd.read_excel(datei_pfad, sheet_name="Zeitnachweis", header=None)
    header_idx = None
    for i, row in df.iterrows():
        if str(row.iloc[0]).strip() == "Datum":
            header_idx = i
            break
    if header_idx is None:
        report.fehler_add(name, "Kein 'Datum'-Header im Zeitnachweis-Sheet gefunden")
        return ergebnis

    df.columns = df.iloc[header_idx]
    df = df.iloc[header_idx + 1:].reset_index(drop=True)

    # Abrechnungsmonat aus globalem Zeitraum ableiten (01.MM.YYYY -> MM, YYYY)
    az_dt     = datetime.strptime(abrechnungszeitraum_global, "%d.%m.%Y")
    az_monat  = az_dt.month
    az_jahr   = az_dt.year

    daten_zeilen = []
    for _, row in df.iterrows():
        datum = datum_aus_zelle(row.get("Datum"))
        if datum is None:
            continue
        # Zeilen außerhalb des Abrechnungsmonats ignorieren
        try:
            zeilen_dt = datetime.strptime(datum, "%d.%m.%Y")
            if zeilen_dt.month != az_monat or zeilen_dt.year != az_jahr:
                log.info(f"  Zeile außerhalb Abrechnungsmonat ignoriert: {datum}")
                continue
        except ValueError:
            continue
        stunden_raw = row.get("Stunden")
        try:
            stunden = 0.0 if pd.isna(stunden_raw) else round(float(stunden_raw), 2)
        except (TypeError, ValueError):
            stunden = 0.0
        kommentar = clean(row.get("Kommentar") or "")
        daten_zeilen.append({"datum": datum, "stunden": stunden, "kommentar": kommentar})

    # Abschlag-Erkennung (gilt fuer alle MA-Typen)
    schluessel_aus_zeitnachweis = set()
    kommentar_tage = defaultdict(list)
    normale_tage   = []

    for zeile in daten_zeilen:
        k = zeile["kommentar"]
        # Abschlag-Erkennung: Kommentar beginnt mit "Abschlag"
        # -SOLL Feiertag: SOLL-Gegenbuchung aus Leitsystem -> komplett ignorieren
        if zeile["stunden"] < 0:
            log.info(f"  Negative Stunden ignoriert: {zeile['stunden']:.2f}h am {zeile['datum']} ('{k}')")
            continue
        if k and k.lower().startswith("abschlag"):
            teile = k.split(None, 1)
            if len(teile) == 2:
                betrag = safe_float(teile[1].replace(",", "."))
                if betrag > 0:
                    ergebnis["vorschuesse"].append({
                        "pnr":                 pnr,
                        "abrechnungszeitraum": abrechnungszeitraum_global,
                        "betrag":              betrag,
                    })
                    log.info(f"  Abschlag erkannt: {betrag:.2f} EUR am {zeile['datum']}")
                else:
                    report.hinweis(name, f"Abschlag am {zeile['datum']} - Betrag nicht erkannt: '{k}'")
            else:
                report.hinweis(name, f"Abschlag am {zeile['datum']} - kein Betrag gefunden: '{k}'")
        elif not ist_gehalt:
            # Stundenempfaenger: Kommentare normal verarbeiten
            if k and k in schluessel_index:
                kommentar_tage[k].append(zeile)
                schluessel_aus_zeitnachweis.add(k)
            elif k and k not in schluessel_index:
                report.hinweis(name, f"Unbekannter Kommentar '{k}' am {zeile['datum']} - nicht gebucht")
            else:
                normale_tage.append(zeile)
        # Gehaltsempfaenger: alle Zeitnachweis-Kommentare (ausser Abschlag) ignorieren

    abwesenheiten      = lese_abwesenheiten(datei_pfad)
    abwesenheiten_tage = defaultdict(list)

    if ist_gehalt:
        # Gehaltsempfaenger: nur Krank, Kind krank (Fehlzeit) und Urlaub LA 303 aus Abwesenheiten
        gehalt_schluessel = {"Kind krank", "Urlaub"}
        la303 = next((r for r in lohnarten_regeln if r["la_nr"] == "303"), None)
        for zeile in abwesenheiten:
            k = zeile["kommentar"]
            if k not in gehalt_schluessel:
                continue
            abwesenheiten_tage[k].append(zeile)

        for schluessel, zeilen in abwesenheiten_tage.items():
            zeilen_sorted = sorted(zeilen, key=lambda x: datetime.strptime(x["datum"], "%d.%m.%Y"))
            regeln = schluessel_index.get(schluessel, [])
            for regel in regeln:
                if regel["ziel"] == "Fehlzeiten":
                    # Fehlzeit-Clustering (identisch zu Stundenempfaenger)
                    zeitraeume = []
                    start = prev = None
                    for z in zeilen_sorted:
                        dt = datetime.strptime(z["datum"], "%d.%m.%Y")
                        if start is None:
                            start = prev = dt
                        else:
                            if (dt - prev).days > 1:
                                zeitraeume.append((start, prev))
                                start = dt
                            prev = dt
                    if start:
                        zeitraeume.append((start, prev))
                    for von_dt, bis_dt in zeitraeume:
                        ergebnis["fehlzeiten"].append({
                            "pnr":          pnr,
                            "datum_von":    von_dt.strftime("%d.%m.%Y"),
                            "datum_bis":    bis_dt.strftime("%d.%m.%Y"),
                            "fehlzeit_key": regel["fehlzeit_key"],
                        })
            # Urlaub LA 303: Tagebasis
            if schluessel == "Urlaub" and la303:
                anzahl = len(zeilen_sorted)
                ergebnis["bewegungsdaten"].append(
                    bwd_satz(pnr, abrechnungszeitraum_global, str(anzahl),
                             la303["bs_nr"] or "71", "303"))
                log.info(f"  Gehalt Urlaub LA 303: {anzahl} Tage")

        report.add_ma(
            name              = name,
            pnr               = pnr,
            la100_h           = None,
            azk_la            = None,
            azk_h             = 0.0,
            fehlzeiten_anzahl = len(ergebnis["fehlzeiten"]),
            saldo_differenz   = saldo_differenz,
            fza_stunden       = 0.0,
        )

    else:
        # Stundenempfaenger: normaler Pfad
        for zeile in abwesenheiten:
            k = zeile["kommentar"]
            if k not in schluessel_index:
                report.hinweis(name, f"Abwesenheiten-Sheet: Unbekannter Schluessel '{k}' am {zeile['datum']} - nicht gebucht")
                continue
            if k in schluessel_aus_zeitnachweis:
                continue
            abwesenheiten_tage[k].append(zeile)

        auszahlungsstunden = 0.0
        direkt_buchungen   = []

        for quelle in (kommentar_tage, abwesenheiten_tage):
            for schluessel, zeilen in quelle.items():
                zeilen_sorted = sorted(zeilen, key=lambda x: datetime.strptime(x["datum"], "%d.%m.%Y"))
                auszahlungsstunden = verarbeite_kommentar_block(
                    schluessel, zeilen_sorted, pnr,
                    ist_stunden_pro_tag, soll_stunden_pro_tag,
                    ergebnis, auszahlungsstunden, direkt_buchungen)

        standardregeln = [r for r in lohnarten_regeln
                          if r["schluessel"] is None
                          and r["ziel"] == "Bewegungsdaten"
                          and r["einheit"] == "Stunden"
                          and not r["azk"]
                          and not r["direkt"]]

        gesamt_stunden_la100 = 0.0
        if normale_tage:
            # LA 100: 1:1 IST-Stunden aus Zeitnachweis (keine SOLL-Deckelung)
            # DATEV bucht LA 100 (IST) ein und zieht LA 122 (AZK+) wieder ab
            # -> Nettoauszahlung = SOLL-Stunden. Deckelung wuerde zu Unterzahlung fuehren.
            gesamt_stunden_la100 = round(sum(z["stunden"] for z in normale_tage), 2)
            log.info(f"  LA 100: {gesamt_stunden_la100:.2f}h (1:1 IST, keine SOLL-Deckelung)")
            if any(r["auszahlung"] for r in standardregeln):
                auszahlungsstunden += gesamt_stunden_la100
            else:
                report.hinweis(name, "Keine Standardregel (Stundenlohn) mit Auszahlung=ja gefunden")

        # AZK-Buchung aus Eigenberechnung (v14): IST - SOLL pro normalem Arbeitstag
        # abzueglich FZA-Abzug (LA 123) und zuzueglich AZK-Kommentar-Buchungen (z.B. LA 210)
        ueberstunden = 0.0
        if normale_tage:
            sum_ist      = round(sum(z["stunden"] for z in normale_tage), 2)
            sum_soll     = round(len(normale_tage) * soll_stunden_pro_tag, 2)
            ueberstunden = round(sum_ist - sum_soll, 2)
            log.info(f"  Ueberstunden-Berechnung: {len(normale_tage)} Tage, "
                     f"IST {sum_ist:.2f}h - "
                     f"SOLL {sum_soll:.2f}h = {ueberstunden:+.2f}h")

        # AZK-Buchung: reine Ueberstunden aus normalen Arbeitstagen (IST - SOLL)
        # FZA (LA 123) und Garantiezeit (LA 210) sind eigenstaendige Buchungen
        # und veraendern den AZK im Leitsystem intern – das Skript rechnet das nicht nach.
        azk_la = None; azk_h = 0.0
        if ueberstunden > 0:
            if la122:
                ergebnis["bewegungsdaten"].append(
                    bwd_satz(pnr, abrechnungszeitraum_global, f"-{ueberstunden:.2f}", la122["bs_nr"] or "1", "122"))
                azk_la = "122"; azk_h = ueberstunden
                log.info(f"  AZK-Buchung: LA 122 -{ueberstunden:.2f}h (Ueberstunden IST-SOLL)")
        elif ueberstunden < 0:
            if la124:
                ergebnis["bewegungsdaten"].append(
                    bwd_satz(pnr, abrechnungszeitraum_global, f"{abs(ueberstunden):.2f}", la124["bs_nr"] or "1", "124"))
                azk_la = "124"; azk_h = abs(ueberstunden)
                log.info(f"  AZK-Buchung: LA 124 +{abs(ueberstunden):.2f}h (Minderstunden IST-SOLL)")
        else:
            log.info("  -> Kein AZK-Ausgleich notwendig (Ueberstunden = 0)")

        # Kontroll-Hinweis: Eigenberechnung vs. Leitsystem-Saldo
        if saldo_differenz is not None:
            report.hinweis(name, f"AZK-Kontrolle: Ueberstunden {ueberstunden:+.2f}h, Leitsystem-Saldo {saldo_differenz:+.2f}h")

        if normale_tage:
            la100_netto = round(gesamt_stunden_la100, 2)
            for regel in standardregeln:
                ergebnis["bewegungsdaten"].append(
                    bwd_satz(pnr, abrechnungszeitraum_global, f"{la100_netto:.2f}",
                             regel["bs_nr"] or "1", regel["la_nr"] or ""))
        else:
            la100_netto = 0.0

        ergebnis["bewegungsdaten"].extend(direkt_buchungen)

        # LA 303 wird bereits ueber verarbeite_kommentar_block() erzeugt
        # (Schluessel "Urlaub", Einheit "Tage" -> bs_nr 71)

        fza_stunden = sum(
            safe_float(bwd["wert"]) for bwd in ergebnis["bewegungsdaten"] if bwd["la_nr"] == "123"
        )

        report.add_ma(
            name              = name,
            pnr               = pnr,
            la100_h           = la100_netto if normale_tage else None,
            azk_la            = azk_la,
            azk_h             = azk_h,
            fehlzeiten_anzahl = len(ergebnis["fehlzeiten"]),
            saldo_differenz   = saldo_differenz,
            fza_stunden       = fza_stunden,
        )

    return ergebnis

# ===========================================================================
# BLOCK 7 – Hauptverarbeitung
# ===========================================================================
alle_fehlzeiten     = []
alle_bewegungsdaten = []
alle_vorschuesse    = []

xlsx_dateien = [f for f in os.listdir(zeitnachweis_ordner) if f.lower().endswith(".xlsx")]
if not xlsx_dateien:
    log.error("Keine .xlsx-Dateien im angegebenen Ordner gefunden.")
    exit(1)

log.info(f"Gefunden: {len(xlsx_dateien)} Zeitnachweis-Datei(en)")

for datei in sorted(xlsx_dateien):
    pfad    = os.path.join(zeitnachweis_ordner, datei)
    ergebnis = verarbeite_zeitnachweis(pfad, datei)
    alle_fehlzeiten.extend(ergebnis["fehlzeiten"])
    alle_bewegungsdaten.extend(ergebnis["bewegungsdaten"])
    alle_vorschuesse.extend(ergebnis["vorschuesse"])

# ===========================================================================
# BLOCK 8 – Output schreiben
# ===========================================================================
berater_nr   = mandant.get("Berater-Nr.", "")
mandant_nr   = mandant.get("Mandanten-Nr.", "")
trennzeichen = mandant.get("Feldtrennzeichen", ";")
begrenzer    = mandant.get("Stringbegrenzer", '"')
datumsformat = mandant.get("Datumsformat", "TT.MM.JJJJ")
zahlenkomma  = mandant.get("Zahlenkomma", ".")
encoding     = mandant.get("Encoding", "ansi")

def schreibe_datei(pfad, zeilen, encoding):
    inhalt = "\r\n".join(z.replace("\r", "").replace("\n", "") for z in zeilen)
    with open(pfad, "w", encoding=encoding, errors="replace", newline="") as f:
        f.write(inhalt)

# Import-Datei
import_zeilen = []
import_zeilen.append("[Allgemein]")
import_zeilen.append("Ziel=LODAS")
import_zeilen.append(f"BeraterNr={berater_nr}")
import_zeilen.append(f"MandantenNr={mandant_nr}")
import_zeilen.append(f"Feldtrennzeichen={trennzeichen}")
import_zeilen.append(f"Stringbegrenzer={begrenzer}")
import_zeilen.append(f"Datumsformat={datumsformat}")
import_zeilen.append(f"Zahlenkomma={zahlenkomma}")
import_zeilen.append(f"StammdatenGueltigAb={abrechnungszeitraum_global}")
import_zeilen.append("")
import_zeilen.append("[Satzbeschreibung]")
import_zeilen.append("1;u_lod_psd_fehlzeiten;pnr#psd;datum_von_ttmmjjjj#psd;datum_bis_ttmmjjjj#psd;grund_fehlzeiten#psd;")
import_zeilen.append("2;u_lod_bwd_buchung_standard;pnr#bwd;abrechnung_zeitraum#bwd;bs_wert_butab#bwd;bs_nr#bwd;la_eigene#bwd;")
import_zeilen.append("3;u_lod_bwd_buchung_vorschuss;pnr#bwd;abrechnung_zeitraum#bwd;nba_vorschuss#bwd;bemerkung#bwd;")
import_zeilen.append("")
import_zeilen.append("[Bewegungsdaten]")
for fz in alle_fehlzeiten:
    import_zeilen.append(f"1;{fz['pnr']};{fz['datum_von']};{fz['datum_bis']};{fz['fehlzeit_key']};")
for bwd in alle_bewegungsdaten:
    if safe_float(bwd["wert"]) == 0.0:
        continue
    import_zeilen.append(f"2;{bwd['pnr']};{bwd['abrechnungszeitraum']};{bwd['wert']};{bwd['bs_nr']};{bwd['la_nr']};")
for vs in alle_vorschuesse:
    bemerkung = f"Abschlag {az_eingabe}"
    import_zeilen.append(f"3;{vs['pnr']};{vs['abrechnungszeitraum']};{vs['betrag']:.2f};{bemerkung};")
    import_zeilen.append(f"2;{vs['pnr']};{vs['abrechnungszeitraum']};-{vs['betrag']:.2f};3;9070;")

ausgabe_pfad = f"lodas_import_{az_eingabe}.txt"
schreibe_datei(ausgabe_pfad, import_zeilen, encoding)

# Report
report_zeilen = report.render(
    az_eingabe = az_eingabe,
    az_anzeige = abrechnungszeitraum_global,
    gesamt_bwd = len(alle_bewegungsdaten),
    gesamt_fz  = len(alle_fehlzeiten),
    encoding   = encoding,
)
report_pfad = f"lodas_report_{az_eingabe}.txt"
print()
for z in report_zeilen:
    print(z)
schreibe_datei(report_pfad, report_zeilen, encoding)

# Sync-Report (deaktiviert – Backlog: Vorzeichen-Logik pruefen)
# sync_zeilen = report.render_sync(
#     az_eingabe = az_eingabe,
#     az_anzeige = abrechnungszeitraum_global,
# )
# sync_pfad = f"lodas_sync_{az_eingabe}.txt"
# for z in sync_zeilen:
#     print(z)
# schreibe_datei(sync_pfad, sync_zeilen, encoding)

print()
print(f"  Import-Datei:  {ausgabe_pfad}")
print(f"  Report-Datei:  {report_pfad}")
# print(f"  Sync-Datei:    {sync_pfad}")  # deaktiviert
print("=" * 70)
