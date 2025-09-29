# Stundenanalyse-Tool

![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![Status](https://img.shields.io/badge/status-POC-orange)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)

## Ziel

Das Tool analysiert Mitarbeiter-Einsatzzeiten und berechnet die Personalstunden je Uhrzeit. Es dient als Grundlage zur Bewertung der Personalproduktivität – z. B. in der Gastronomie

## Funktionen

- Einlesen eines Excel-Exports (in dieser Version konkretes Format erwartet)
- Zerlegung der Arbeitszeit je Stunde
- Rundung auf zwei Nachkommastellen
- Gruppierung nach Datum und Stunde
- Automatische Formatierung (Zentrierung, Spaltenbreite)
- Export als Excel-Datei mit neuem Dateinamen
- GUI zur Datei-Auswahl (kein Code-Kontakt nötig)

## Aufbau des Exports

- Verwendet das Tabellenblatt „Alle Mitarbeiter“
- Header ab Zeile 7
- Relevante Spalten: Tag, Startzeit, Endzeit, Dauer netto (dezimal)
- Beendet die Verarbeitung, sobald in Spalte A der Begriff „Summe:“ auftaucht

## Nutzung

1. Python installieren
2. Abhängigkeiten installieren:

```bash
pip install pandas openpyxl
```

3. Tool starten:

```bash
python staff_hours_analyzer_final.py
```

## Testdaten

Testdateien findest du im Ordner `example_data/` (optional).

## Geplant

- Konfigurierbarer Spaltenimport (für generische Exporte)
- Produktivitätskennzahlen (Umsatz je Stunde / Mitarbeiter)
- Monats-/Wochenauswertungen mit Medianvergleichen
- Portable `.exe` für einfache Nutzung durch Nicht-Tech-User


