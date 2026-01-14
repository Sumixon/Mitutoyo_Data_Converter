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

## 📊 Unterstützte Parameter

| Parameter | Einheit | Beschreibung                         |
| --------- | ------- | ------------------------------------ |
| Ra        | μm      | Arithmetische Mittenrauheit          |
| Rz        | μm      | Mittlere Rautiefe                    |
| Rq        | μm      | Quadratische Mittenrauheit (RMS)     |
| Rp        | μm      | Maximale Profilspitzenhöhe           |
| Rv        | μm      | Maximale Profiltal-Tiefe             |
| Rsk       | μm      | Profilschiefenwert                   |
| Rku       | μm      | Profilwölbungswert                   |
| Rc        | μm      | Mittlere Höhe der Profilelemente     |
| RPc       | /cm     | Anzahl der Profilelemente pro cm     |
| RSm       | μm      | Mittlerer Abstand der Profilelemente |
| RDq       | μm      | Mittlere quadratische Steigung       |
| Rmr       | %       | Traganteil (Materialanteil)          |
| Rdc       | μm      | Profilhöhe                           |
| Rt        | μm      | Gesamthöhe des Profils               |
| Rz1max    | μm      | Maximale Rautiefe                    |
| Rk        | μm      | Kernrauheitstiefe                    |
| Rpk       | μm      | Reduzierte Spitzenhöhe               |
| Rvk       | μm      | Reduzierte Riefentiefe               |
| Mr1       | %       | Materialanteil 1                     |
| Mr2       | %       | Materialanteil 2                     |
| A1        | -       | Fläche oberhalb des Kerns            |
| A2        | -       | Fläche unterhalb des Kerns           |

## 🔧 Technische Details

- **Framework:** customtkinter (moderner Wrapper um Tkinter) + ttk Treeview
- **Datenverarbeitung:** Pandas für Datenmanipulation
- **Excel-Export:** OpenPyXL zur Erstellung von .xlsx Dateien
- **GUI-Style:** Modernes Flat-Design, Karten (CTkFrame), Tabs (CTkTabview), Dark/Light Mode
- **Dateihandling:** UTF-8-Encoding mit Unterstützung für Fehlerzustände
- **Architektur:** Objektorientiertes Design mit modularer Struktur
- **UI-Übersetzungen:** `locales/translations.json` (CS/EN/DE)

## 📋 Format der Eingabedateien

Die App erwartet TXT-Dateien im Mitutoyo SJ-412 Format mit folgender Struktur:

```
//Header
Date;2025-01-01;
Time;10:30:15;

//CalcResult
Ra;1.234;μm
Rz;5.678;μm
Rq;1.456;μm
...

//Condition-A
Cutoff;0.8;mm
Speed;0.5;mm/s
...
```

## 🐛 Problemlösung

### Häufige Probleme:

**App startet nicht:**

- Prüfen Sie, ob Python 3.8+ installiert ist
- Prüfen Sie die Installation aller Abhängigkeiten: `pip install -r requirements.txt`

**Fehler beim Lesen der TXT-Datei:**

- Prüfen Sie, ob die Datei im korrekten Mitutoyo SJ-412 Format ist
- Prüfen Sie die Kodierung der Datei (sollte UTF-8 sein)

**Export nach Excel funktioniert nicht:**

- Prüfen Sie die Schreibrechte im Zielordner
- Stellen Sie sicher, dass die Ziel-Excel-Datei nicht geöffnet ist

**Langsame Verarbeitung:**

- Bei vielen Dateien ggf. in kleineren Batches verarbeiten
- Prüfen Sie den verfügbaren Arbeitsspeicher (RAM)

## 📄 Lizenz

Distributed under the MIT License. See `LICENSE` for more information.

## 👨‍💻 Autor

**Roman Denev (Sumixon)**

- GitHub: [@Sumixon](https://github.com/Sumixon)
- Email: roman.denev@gmail.com

## 📈 Changelog

### v2.1.0 (2026-01-08)

- ✅ Migration der GUI auf customtkinter (modernes Design, Dark/Light Mode)
- ✅ Neues Layout mit Karten (CTkFrame) und CTkTabview
- ✅ Verbesserte, gut lesbare Treeview-Tabelle im Dark Mode

### v2.0.0 (2025-01-01)

- ✅ Komplett überarbeitetes modernes UI
- ✅ Verbesserter TXT-Parser mit besserem Error Handling
- ✅ Erweiterte Unterstützung für Rauheitsparameter
- ✅ Optimierte Verarbeitung großer Dateien
- ✅ Tabs hinzugefügt für bessere Organisation

### v1.0.0 (2024-12-01)

- ✅ Erste Version der Anwendung
- ✅ Grundlegende Import/Export-Funktionalität
- ✅ Tkinter GUI mit einfachem Design

## 🔗 Nützliche Links

- [Mitutoyo SJ-412 Manual](https://mitutoyo.com/)
- [Python Documentation](https://docs.python.org/3/)
- [Pandas Documentation](https://pandas.pydata.org/docs/)
- [Tkinter Tutorial](https://docs.python.org/3/library/tkinter.html)

---

**Erstellt für präzise Messung der Oberflächenrauheit**
