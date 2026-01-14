# Mitutoyo Data Converter

**Sprachen:** [Čeština](README.md) | [English](README_ENG.md) | [Deutsch](README_DE.md)

Eine moderne Desktop-Anwendung zum Konvertieren von Messdaten des Mitutoyo SJ-412 von .txt nach Excel unter Windows.

Erstellt mit Python, dem modernen GUI-Framework **customtkinter** (Dark/Light Mode) und **Pandas** für die Datenverarbeitung.

## 📋 Funktionen

- ✅ **Import von TXT-Dateien** vom Mitutoyo SJ-412
- ✅ **Automatische Verarbeitung** der Messdaten
- ✅ **Export nach Excel** (.xlsx)
- ✅ **Unterstützung vieler Rauheitsparameter** (Ra, Rz, Rq, Rp, Rv, usw.)
- ✅ **Modernes GUI** mit elegantem Design
- ✅ **Batch-Verarbeitung** mehrerer Dateien
- ✅ **Intuitive Bedienung**
- ✅ **Dark/Light Mode** und moderne Tabs (Tabview)

## 🖥️ Systemanforderungen

- **Betriebssystem:** Windows 10/11
- **Python:** 3.8 oder neuer
- **RAM:** mindestens 4GB
- **Speicherplatz:** 100MB für die App + Platz für Daten

## 🚀 Installation

### Option 1: Aus dem Quellcode starten

1. **Repository klonen:**

```bash
git clone https://github.com/Sumixon/mitutoyo-converter.git
cd mitutoyo-converter
```

2. **Virtuelle Umgebung erstellen:**

```bash
python -m venv venv
venv\Scripts\activate
```

3. **Abhängigkeiten installieren:**

```bash
pip install -r requirements.txt
```

4. **App starten:**

```bash
python main.pyw
```

### Option 2: Standalone-EXE erstellen

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=icon.ico main.pyw
```

Die EXE befindet sich anschließend im Ordner `dist/`.

## 🚀 Schnellstart

1. **App starten** – `python main.pyw`
   - Die UI-Sprache schalten Sie über die Flaggen oben im Fenster.
2. **Dateien importieren** – „📂 Dateien importieren“
3. TXT-Dateien vom Mitutoyo SJ-412 auswählen
4. Daten in der Tabelle prüfen
5. **Export nach Excel** – „📊 Nach Excel exportieren“
6. Speicherort auswählen

## 🔧 Technische Details

- **Framework:** customtkinter + ttk Treeview
- **Datenverarbeitung:** Pandas
- **Excel-Export:** OpenPyXL
- **UI-Übersetzungen:** `locales/translations.json` (CS/EN/DE)

## 🏳️ Flaggen (Quelle)

Die Flaggen sind statische Dateien (kein automatischer Download). Lege PNG-Dateien in `img/flags/` (empfohlen) oder `img/` ab.

- CZ: https://commons.wikimedia.org/wiki/File:Flag_of_the_Czech_Republic.svg
- EN (UK): https://commons.wikimedia.org/wiki/File:Flag_of_the_United_Kingdom.svg
- DE: https://commons.wikimedia.org/wiki/File:Flag_of_Germany.svg

## 📄 Lizenz

Distributed under the MIT License. See `LICENSE` for more information.

## 👨‍💻 Autor

**Roman Denev (Sumixon)**

- GitHub: [@Sumixon](https://github.com/Sumixon)
- Email: romna.denev@gmail.com
