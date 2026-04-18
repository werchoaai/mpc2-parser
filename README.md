# MPC² Corrosion Ray — DL-EPR Parser & Analyzer

Automatische Auswertung von IPS PgU Touch `.ASC` Messdaten: Peak-Erkennung (Ja, Jr), Ladungsintegration (Qa, Qr), DOS-Berechnung (Jr/Ja, Qr/Qa) und direktes Befüllen der MPC² Auswertungs-Workbooks.

Built by **werchota.ai** for **MPC² GmbH** — Green Startupmark consulting project (Linear: [MAL-10](https://linear.app/malcolm-werchota-2026/issue/MAL-10)).

---

## Validierung

Getestet gegen Manuels manuelle Auswertung von Projekt 2 (8 ASC-Dateien):

| Metric | Mean abs. error vs. manuell | Sheets mit 0.00% Fehler |
|--------|---|---|
| **Ja** | <0.01% | 6 / 6 valide |
| **Jr** | <0.01% | 6 / 6 valide |
| **Qa** | 0.05% | 6 / 6 |
| **Qr** | 0.9% | 6 / 6 |

_(Die 2 "Ausreißer" auf Sheet 404/405 sind auf einen Daten-Label-Swap in Manuels Workbook zurückzuführen — die ASC-Dateien 0404 und 0405 sind in seiner Auswertung vertauscht. Parser-Output stimmt exakt mit der korrekten Zuordnung überein.)_

---

## Setup

```bash
cd ~/Documents/mpc2-parser
python3 -m venv .venv
source .venv/bin/activate
pip install numpy pandas openpyxl scipy matplotlib streamlit
```

---

## Die vier Output-Varianten

Wir haben bewusst **vier verschiedene Workflows** gebaut, damit MPC² entscheiden kann, welcher am besten in den Alltag passt.

### Variante 1 — Drop-in Ersatz für Manuels aktuelles Auswertungs-Workbook

**Was sie tut:** Liest einen Ordner voller ASCs und erzeugt eine `Auswertung_<Projekt>.xlsx` genau wie die, die Manuel heute per Hand pflegt. Ein Sheet pro Messung, Spalten A-I mit Rohdaten + Formeln, Summary-Block bei L-P mit Ja/Jr/Qa/Qr/DOS.

**Wann nutzen:** Zero-Change-Workflow. Manuel öffnet das File und findet alles genau wie gewohnt — nur komplett ausgefüllt.

```bash
PYTHONPATH=src python -m mpc2_parser.cli variant1 \
  "/path/to/Messdateien/Projekt X" \
  out/Auswertung_ProjektX.xlsx
```

**Pro:** Kein Umdenken, passt in den aktuellen Prozess.
**Con:** Zwischenschritt — die Werte müssen dann immer noch ins Messübersicht-Master übertragen werden.

---

### Variante 2 — Direkt ins Messübersicht-Master schreiben

**Was sie tut:** Parst die ASCs und hängt für jede Messung eine neue Zeile im "Corrosion Ray"-Sheet der `Messübersicht_Elektrochemische Messungen.xlsx` an. Füllt automatisch: Messung, Datum, Uhrzeit, Material, Probenbezeichnung, Messfläche (aus HAD), Temperatur, Ruhepotential, Ja, Jr, Jr/Ja, Qa, Qr, Qr/Qa, Prüfer, Dateiname.

**Wann nutzen:** Wenn man vom ASC direkt ins Master-Laborbuch will — maximale Zeitersparnis.

```bash
PYTHONPATH=src python -m mpc2_parser.cli variant2 \
  "/path/to/Messdateien/Projekt X" \
  "/path/to/Messübersicht_Elektrochemische Messungen.xlsx" \
  --output out/Messuebersicht_updated.xlsx
```

Die Original-Datei wird **nicht** überschrieben — es wird eine Kopie erzeugt. Bereits vorhandene Messung-IDs werden per Default übersprungen (oder mit `--overwrite` aktualisiert).

**Pro:** Spring vom ASC direkt ins Laborbuch. Weniger Kopier-Fehler.
**Con:** Kein Rohkurven-Archiv im Master — die detaillierte Kurvenansicht fehlt hier.

---

### Variante 3 — Web-App "Paste & Copy"

**Was sie tut:** Streamlit-App im Browser. Drag-and-drop eine oder mehrere ASC-Dateien (optional HAD), und sofort sieht man:
- Große Kacheln mit Ja, Jr, Qa, Qr, Jr/Ja, Qr/Qa, Ruhepotential
- Die beiden Kurven (Potenzial vs. Zeit und J vs. Zeit)
- **Kopier-Button** mit tab-separierten Werten zum direkten Einfügen ins Excel
- Manual Override des Split-Punktes per Slider, falls die Auto-Erkennung daneben liegt
- Toggle zwischen Vertex- und Midpoint-Methode

**Wann nutzen:** Einzelmessung, schnell checken, Werte in eine beliebige Tabelle kopieren. Mobile-freundlich — läuft auch am Handy.

```bash
cd ~/Documents/mpc2-parser
source .venv/bin/activate
streamlit run webapp/app.py
```

**Pro:** Interaktive Validierung, visuelle Kontrolle, läuft überall (kein Excel nötig).
**Con:** Einzelverarbeitung, nicht für Batch-Jobs geeignet.

---

### Variante 4 — Kombiniertes Workbook (alles in einem)

**Was sie tut:** Erzeugt **ein** Excel mit:
1. Einem "Übersicht"-Sheet (eine Zeile pro Messung, ready to copy in Messübersicht)
2. Einem Detail-Sheet pro Messung (wie Variante 1)

**Wann nutzen:** Wenn man sowohl den Überblick als auch die Details in einer Datei haben möchte. Manuel kann die Übersicht-Tabelle direkt in die Messübersicht kopieren und hat trotzdem die Kurven zur Hand.

```bash
PYTHONPATH=src python -m mpc2_parser.cli variant4 \
  "/path/to/Messdateien/Projekt X" \
  out/Combined_ProjektX.xlsx
```

**Pro:** Best of both worlds — Übersicht + Detail.
**Con:** Größere Datei, etwas mehr Komplexität.

---

## Split-Detection: Zwei Methoden + Manual Override

Das Herzstück der Auswertung ist die Trennung zwischen **Aktivierungs-Sweep** (Ja) und **Reaktivierungs-Sweep** (Jr). Manuel hat das bisher manuell gemacht ("in Zeile 1076 trenne ich"). Wir bieten zwei automatische Methoden plus ein manuelles Override.

### Methode A — Potenzial-Vertex (physikalisch)

Findet den Punkt, wo das Potenzial sein Maximum erreicht (das ist per Definition die Sweep-Umkehr).

- **Pro:** Physikalisch die korrekte Definition der "Double Loop".
- **Pro:** Funktioniert bei beliebigen Messzeiten und -raten.
- **Pro:** Unabhängig von der HAD-Datei (robust auch bei fehlenden Metadaten).
- **Con:** Kann bei sehr verrauschten Signalen um ein paar Zeilen schwanken (mit Glättung entschärft).

### Methode B — Mittelpunkt (heuristisch)

Teilt bei N/2 (Anzahl Punkte geteilt durch zwei).

- **Pro:** Simpel, deterministisch, in einem Satz erklärt.
- **Pro:** Schnell, keine Signalverarbeitung nötig.
- **Pro:** Entspricht Manuels bisheriger Praxis (Split ≈ 1060-1088 für ~2000 Zeilen).
- **Con:** Nimmt an, dass Vorwärts- und Rückwärts-Sweep gleich lang dauern (meist wahr).
- **Con:** Benutzt nicht die tatsächliche Kurve, nur die Anzahl der Punkte.

### Manual Override

In der Web-App per Slider. In der CLI per `--split-override=<index>` (für Einzel-ASC über den `json`-Befehl). Für Batch-Jobs via CSV-Konfigurationsdatei erweiterbar.

---

## Qr-Truncation (wichtig!)

DL-EPR-Konvention: Die Qr-Integration endet, sobald die Rückwärtsrampe wieder das Startpotenzial erreicht. Danach folgt der Post-Sweep-Bereich (Recovery / über-kathodisch), der **nicht** in Qr zählen soll.

Manuel macht das manuell ("SUM(I1060:I1950)" — Endzeile handverlesen). Wir haben das automatisiert (`find_reverse_endpoint` in `analysis.py`). Resultat: Qr-Fehler vs. Manuel von ~18% auf <3% reduziert.

---

## Package-Struktur

```
mpc2-parser/
├── src/mpc2_parser/
│   ├── parser.py          # ASC + HAD + Filename parser
│   ├── analysis.py        # Split-Detection, Ja/Jr/Qa/Qr, Truncation
│   ├── core.py            # process_measurement() orchestration
│   ├── cli.py             # Command-line interface
│   └── outputs/
│       ├── variant1_auswertung.py
│       ├── variant2_messuebersicht.py
│       └── variant4_combined.py
├── webapp/
│   └── app.py             # Variante 3 (Streamlit)
├── tests/
│   └── validate_projekt2.py   # Vergleich gegen Manuels manuelle Auswertung
└── out/                   # Demo-Outputs für Projekt 2
```

---

## Bekannte Data-Quality-Hinweise

- **Projekt 2 Auswertung 404/405 Swap:** In Manuels `Auswertung_ON2025-0003_...xlsx` scheinen die Sheets für ASC 0404 und 0405 vertauscht zu sein. Unsere Auswertung auf den Original-ASCs ist konsistent; die Manuel-Zuordnung ist es nicht.

---

## Nächste Schritte

1. ~~ASC-Parser mit Peak-Erkennung~~ ✓ (MAL-10)
2. ~~Excel-Template-Filler~~ ✓ (Variant 1 + 4)
3. ~~Messübersicht-Direct-Append~~ ✓ (Variant 2)
4. ~~Web-App~~ ✓ (Variant 3)
5. Review-Session mit Manuel und Simone — welche Variante(n) sollen Prio haben?
6. Ergebnisse aus Projekt 1 (~80 ASCs) batchen und Qualität auf größerem Datensatz validieren
7. Temperaturkompensations-Algorithmus drauf aufsetzen (MAL-13)
