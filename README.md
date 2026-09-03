# budget-watcher

Beobachtet einen Ordner, extrahiert Buchungen aus Bankauszugs-PDFs/CSVs (per lokaler KI) und schreibt sie in eine Excel-Datei zur Budgetplanung. Zuordnung in Kategorien läuft über einfache, frei erweiterbare Regeln – alles Unzuordenbare landet in „Sonstiges“ und kann per Korrektur in Excel nachträglich „gelernt“ werden.

## Aufbau

```
budget-watcher/
├── watcher.py            # das Wächter-Skript
├── eingabe/              # ← Bankauszugs-PDFs oder -CSVs hierhin legen
├── verarbeitet/          # erfolgreich verarbeitete Dateien landen hier
├── fehler/                # Dateien, die nicht verarbeitet werden konnten
├── ausgabe/
│   ├── budget.xlsx       # Ergebnis (Buchungen, Übersicht, Monatsvergleich, Abos)
│   └── backups/          # automatische Sicherungen vor jedem Schreibvorgang
└── config/
    ├── ai.yaml           # lokale KI: Endpoint, Modell, Auto-Start, Fallback
    └── rules/            # eine Markdown-Datei pro Kategorie (z. B. rewe → Lebensmittel)
        ├── 00-settings.md              # default_category
        ├── 01-gelernt/                 # vom Wächter automatisch ergänzt (siehe unten)
        ├── 10-lebensmittel.md
        ├── 20-miete-nebenkosten.md
        ├── 30-mobilitaet.md
        ├── 40-versicherungen-gesundheit.md
        ├── 50-streaming-online.md
        ├── 60-handy-internet.md
        └── 70-einnahmen.md
```

## Einrichtung (einmalig)

```bash
cd budget-watcher
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Lokale KI (empfohlen)

Der Wächter nutzt jeden OpenAI-kompatiblen **lokalen** Endpoint – typisch Ollama oder LM Studio.

**Ollama:**
```bash
brew install ollama
ollama pull qwen2.5:3b        # oder qwen2.5:7b für mehr Genauigkeit
```
`ollama serve` muss laufen – oder der Wächter startet es automatisch (siehe `auto_start` in `config/ai.yaml`).

**LM Studio:** Server starten und in `config/ai.yaml` die `base_url` auf `http://localhost:1234/v1` sowie das geladene Modell unter `model` setzen.

**Ohne KI** verarbeitet der eingebaute Fallback-Parser PDFs im üblichen deutschen Auszugsformat
(`TT.MM.JJJJ  EMPFÄNGER  BESCHREIBUNG  -Betrag`). Er ist robust, aber weniger flexibel als die KI.

**CSV-Import**: Alternativ zu PDFs akzeptiert `eingabe/` auch `.csv`-Exporte der Bank. Spalten
für Datum, Empfänger, Beschreibung und Betrag werden anhand gängiger deutscher/englischer
Kopfzeilen-Namen automatisch erkannt (z. B. „Buchungstag“, „Beguenstigter/Zahlungspflichtiger“,
„Verwendungszweck“, „Betrag“). Werden Spalten nicht erkannt, landet die CSV mit genauer
Fehlermeldung in `fehler/`. Sowohl `;` als auch `,` als Trennzeichen werden erkannt, ebenso
UTF-8- und CP1252-Kodierung (typisch bei älteren Sparkassen-/DKB-Exporten).

## Nutzung

```bash
source .venv/bin/activate
python watcher.py          # läuft dauerhaft und beobachtet eingabe/
python watcher.py --once   # verarbeitet aktuelle Dateien einmalig und beendet
```

**Im Hintergrund starten** (kein offenes Terminal nötig):

- Doppelklick auf `Watcher starten.app` (macOS) bzw. `Watcher starten.vbs` (Windows) – startet `watcher.py` unsichtbar im Hintergrund, kein Konsolenfenster bleibt offen.
- Alternativ per Terminal: `python start.py` (nutzt automatisch `.venv`, schreibt PID nach `watcher.pid` und Log nach `watcher.log`).
- Erneuter Start erkennt einen bereits laufenden Wächter und tut nichts.
- Der Hintergrundprozess läuft unabhängig vom startenden Fenster weiter (eigene Session) und endet nur durch `kill $(cat watcher.pid)` bzw. `taskkill /PID <pid> /F` oder einen Neustart des Rechners.

Prozessablauf je Datei:
1. Wächter meldet neue `.pdf`/`.csv` in `eingabe/` (wartet, bis die Datei vollständig kopiert ist)
2. Bei PDF: lokale KI wird bei Bedarf hochgefahren (`auto_start`), extrahiert Datum/Empfänger/Beschreibung/Betrag als JSON (sonst Fallback-Parser). Bei CSV: Spalten werden direkt gelesen.
3. `budget.xlsx` wird vorher nach `ausgabe/backups/` gesichert (die letzten 20 Sicherungen bleiben erhalten)
4. Manuelle Kategorie-Korrekturen aus vorigen Läufen werden erkannt und als neue Regeln unter `config/rules/01-gelernt/` gespeichert (siehe unten)
5. Regeln aus `config/rules/` ordnen jede Buchung einer Kategorie zu
6. Bereits vorhandene Buchungen (gleiches Datum, gleicher Empfänger, gleiche Beschreibung, gleicher Betrag) werden übersprungen, um Duplikate durch überlappende Auszüge zu vermeiden
7. Neue Zeilen werden in `ausgabe/budget.xlsx` angehängt; „Übersicht“, „Monatsvergleich“ und „Abos“ werden neu berechnet
8. Datei wird nach `verarbeitet/` verschoben (bei Problemen nach `fehler/`)

**Die vier Blätter in `budget.xlsx`:**
- **Buchungen** – alle Rohdaten, eine Zeile pro Buchung
- **Übersicht** – Summe je Kategorie (nur Ausgaben)
- **Monatsvergleich** – Ausgaben je Kategorie und Monat nebeneinander (Pivot-Tabelle)
- **Abos** – Empfänger mit demselben Betrag in mindestens zwei verschiedenen Monaten (erkannte Abos/wiederkehrende Zahlungen)

## Kategorien erweitern

Jede Kategorie ist eine eigene Markdown-Datei in `config/rules/`, z. B. `config/rules/10-lebensmittel.md`:

```markdown
# Lebensmittel

- rewe          „Rewe“ ist damit immer Lebensmittel
- aldi
- lidl
```

**Neue Kategorie** = neue Datei in `config/rules/` anlegen, z. B. `config/rules/45-hobbys.md`:

```markdown
# Hobbys

- kletterhalle
- fischen
```

Regeln:
- Erste Überschrift (`# Name`) = Kategoriename, jede Bullet-Zeile (`- keyword`) darunter = ein Keyword. Freier Text zwischen den Bullets (Notizen, Erklärungen) wird beim Einlesen ignoriert.
- Der Zahlen-Präfix im Dateinamen bestimmt die **Priorität** (aufsteigend, alphabetische Sortierung): die zuerst passende Kategorie gewinnt. Lücken (10, 20, 30, …) lassen Platz, um neue Kategorien dazwischenzuschieben, ohne andere Dateien umzubenennen.
- Keywords werden **kleingeschrieben als komplette Wörter** in Empfänger + Beschreibung gesucht (kein falsch-Matches wie „tk“ in „Marktkirche“).
- Keine passende Regel → `default_category` aus `config/rules/00-settings.md` (Standard: `Sonstiges`).
- `config/rules/` wird **rekursiv** durchsucht – Unterordner zum weiteren Gruppieren sind erlaubt, z. B. `config/rules/fixkosten/20-miete-nebenkosten.md`.
- Regeln werden **bei jeder Datei neu eingelesen** – neue/geänderte Dateien wirken sofort, ohne den Wächter neu zu starten oder `watcher.py` anzupassen.

## Automatisches Lernen aus Korrekturen

Ordnet eine Buchung „Sonstiges“ zu und korrigierst du das später manuell in `budget.xlsx`
(Spalte „Kategorie“ auf eine bestehende Kategorie ändern), merkt sich der Wächter das beim
nächsten Lauf: das erste Wort des Empfängers wird als neues Keyword unter
`config/rules/01-gelernt/<kategorie>.md` gespeichert. Künftige Buchungen desselben Empfängers
(auch mit leicht anderem Namen, z. B. „Kletterhalle Berlin“ statt „Kletterhalle München“) werden
dann automatisch richtig zugeordnet – ganz ohne manuelles Anlegen einer Regel.

Der Ordner `01-gelernt/` sortiert alphabetisch **vor** allen anderen Kategorien und hat damit
höchste Priorität. Nur Korrekturen auf eine **bestehende** Kategorie werden übernommen (Tippfehler
oder frei erfundene Kategorienamen in Excel werden ignoriert). Die generierten Dateien können
jederzeit von Hand nachbearbeitet oder gelöscht werden wie jede andere Regel-Datei auch.

## Grenzen / Hinweise

- **Gescannte Bild-PDFs** (Fotokopien ohne Textlayer) können nicht gelesen werden – solche PDFs landen in `fehler/`.
- **CSV-Spaltenerkennung** ist heuristisch (gängige deutsche/englische Kopfzeilen); exotische Bank-Exportformate (z. B. DKB mit mehrzeiliger Kopf-Präambel) werden nicht automatisch entschachtelt und landen ggf. in `fehler/`.
- Die „Übersicht“ und der „Monatsvergleich“ summieren nur **Ausgaben** (negative Beträge); Einnahmen sind im Blatt „Buchungen“ sichtbar.
- Beträge: Ausgaben negativ, Einnahmen positiv.
- **Duplikaterkennung** vergleicht Datum, Empfänger, Beschreibung und Betrag exakt – leicht unterschiedlich formulierte Duplikate (z. B. abweichender Verwendungszweck) werden nicht erkannt.
- **Backups** liegen unversioniert in `ausgabe/backups/`; die letzten 20 werden behalten, ältere automatisch gelöscht.
- Projekt bewusst **nicht versioniert** (außerhalb jeder Git-Repo).
