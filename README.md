# Excel Data Merger & Sorter

Dieses Python-Skript ermöglicht es, Daten aus einer Quell-Excel-Datei in eine bestehende Ziel-Excel-Datei zu übertragen. Dabei werden Formatierungen, Formeln und Datentypen (insbesondere Datumsformate) beibehalten und der gesamte Bestand kann nach einer frei wählbaren Spalte sortiert werden.

---

### Voraussetzungen

- **Python 3.10+** muss installiert sein.
- Die Ziel-Excel-Datei muss mindestens eine Zeile mit Daten enthalten (Zeile 2), die als Vorlage für Formeln und Formatierungen dient.

---

### Installation

1. Entpacke den Ordner.
2. Öffne ein Terminal / eine Eingabeaufforderung in diesem Ordner.
3. Installiere die benötigten Bibliotheken:
   ```bash
   pip install -r requirements.txt
   ```

---

### Benutzung

1. Skript starten:
   `python script.py`

Bash

python script.py
Dateipfade angeben: Gib den Pfad zur Quell-Datei und zur Ziel-Datei an (z. B. daten.xlsx).

Spaltenzuordnung (Mapping):

Das Skript geht nacheinander alle Spalten der Ziel-Datei durch.

Gib für jede Ziel-Spalte den exakten Namen der Spalte aus der Quell-Datei ein.

Lässt du ein Feld leer, bleibt die Spalte für die neuen Daten leer (oder behält ihre Formel).

Hinweis: Spalten, die in der Ziel-Excel bereits Formeln enthalten, werden automatisch erkannt und übersprungen. Die Formeln werden für alle neuen Zeilen intelligent "nach unten gezogen".

Sortierung:

Gib den Namen der Spalte aus der Ziel-Datei an, nach welcher der gesamte Bestand (alte und neue Daten) sortiert werden soll.

Das Skript erkennt Datumsangaben, Zahlen und Texte und sortiert diese entsprechend.

⚠️ Wichtige Hinweise
Schreibschutz: Die Ziel-Excel-Datei muss geschlossen sein, während das Skript läuft. Andernfalls bricht das Programm mit einem PermissionError ab.

Formel-Logik: Das Skript nutzt einen Translator für Excel-Formeln. Relative Bezüge (z. B. =A2*1.19) werden für jede neue Zeile korrekt angepasst (z. B. =A50*1.19).

Datentypen: Datumsangaben werden für die Sortierung temporär in Excel-Seriennummern umgewandelt und beim Schreiben wieder in echte Datums-Objekte zurückgeführt, um Formatierungsfehler zu vermeiden.

Backup: Da das Skript die Zieldatei direkt überschreibt, ist es ratsam, vor der Ausführung eine Kopie der Zieldatei anzulegen.
