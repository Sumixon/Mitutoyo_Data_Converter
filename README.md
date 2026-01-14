# Mitutoyo Data Converter

**Jazyky:** [Čeština](README.md) | [English](README_ENG.md) | [Deutsch](README_DE.md)

Moderní desktop aplikace pro převod měřicích dat z přístroje Mitutoyo SJ-412 z formátu .txt do Excel pro Windows.

Postavena na Pythonu, moderním GUI frameworku **customtkinter** (dark/light režim) a **Pandas** pro práci s daty.

## 📋 Funkce

- ✅ **Import TXT souborů** z měřicího přístroje Mitutoyo SJ-412
- ✅ **Automatické zpracování** měřicích dat
- ✅ **Export do Excel** formátu (.xlsx)
- ✅ **Podpora všech parametrů drsnosti** (Ra, Rz, Rq, Rp, Rv, atd.)
- ✅ **Moderní GUI** s elegantním designem
- ✅ **Batch processing** - zpracování více souborů najednou
- ✅ **Intuitivní uživatelské rozhraní**
- ✅ **Světlý / tmavý režim** a moderní karty (Tabview)

## 🖥️ Systémové požadavky

- **Operační systém:** Windows 10/11
- **Python:** 3.8 nebo novější
- **RAM:** Minimálně 4GB
- **Místo na disku:** 100MB pro aplikaci + místo pro data

## 🚀 Instalace

### Možnost 1: Spuštění ze zdrojového kódu

1. **Klonování repozitáře:**

```bash
git clone https://github.com/Sumixon/mitutoyo-converter.git
cd mitutoyo-converter
```

2. **Vytvoření virtuálního prostředí:**

```bash
python -m venv venv
venv\Scripts\activate
```

3. **Instalace závislostí:**

```bash
pip install -r requirements.txt
```

4. **Spuštění aplikace:**

```bash
python main.pyw
```

### Možnost 2: Vytvoření standalone EXE

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=icon.ico main.pyw
```

Výsledný EXE soubor najdete ve složce `dist/`.

## 🚀 Rychlý start

1. **Spusťte aplikaci** - `python main.pyw`
   - Jazyk UI přepnete vlajkami v horní části okna.
2. **Importujte soubory** - klikněte na "📂 Importovat soubory"
3. **Vyberte TXT soubory** z měřicího přístroje Mitutoyo SJ-412
4. **Zkontrolujte data** v tabulce
5. **Exportujte do Excel** - klikněte na "📊 Exportovat do Excel"
6. **Uložte soubor** na požadované místo

## 📊 Podporované parametry

| Parametr | Jednotka | Popis                            |
| -------- | -------- | -------------------------------- |
| Ra       | μm       | Průměrná aritmetická drsnost     |
| Rz       | μm       | Průměrná výška drsnosti          |
| Rq       | μm       | Průměrná kvadratická drsnost     |
| Rp       | μm       | Maximální výška výstupku         |
| Rv       | μm       | Maximální hloubka prohlubně      |
| Rsk      | μm       | Šikmost profilu                  |
| Rku      | μm       | Špičatost profilu                |
| Rc       | μm       | Průměrná výška elementu          |
| RPc      | /cm      | Počet elementů na cm             |
| RSm      | μm       | Průměrná vzdálenost elementů     |
| RDq      | μm       | Střední kvadratický sklon        |
| Rmr      | %        | Relativní délka nesoucí křivky   |
| Rdc      | μm       | Výška profilu                    |
| Rt       | μm       | Celková výška profilu            |
| Rz1max   | μm       | Maximální výška drsnosti         |
| Rk       | μm       | Hloubka jádra drsnosti           |
| Rpk      | μm       | Redukovaná výška výstupků        |
| Rvk      | μm       | Redukovaná hloubka prohlubní     |
| Mr1      | %        | Relativní délka nesoucí křivky 1 |
| Mr2      | %        | Relativní délka nesoucí křivky 2 |
| A1       | -        | Plocha nad jádrem                |
| A2       | -        | Plocha pod jádrem                |

## 🔧 Technické detaily

- **Framework:** customtkinter (moderní wrapper nad Tkinter) + ttk Treeview
- **Data processing:** Pandas pro manipulaci s daty
- **Excel export:** OpenPyXL pro vytváření .xlsx souborů
- **GUI Style:** Modern flat design, karty (CTkFrame), záložky (CTkTabview), dark/light režim
- **File handling:** UTF-8 encoding s podporou chybových stavů
- **Architektura:** Objektově orientovaný design s modulární strukturou
- **Překlady UI:** `locales/translations.json` (CS/EN/DE)

## 🏳️ Vlajky (zdroj)

Vlajky jsou v projektu staticky (žádné automatické stahování). Vložte PNG soubory do `img/flags/` (preferované) nebo do `img/`.

- CZ: https://commons.wikimedia.org/wiki/File:Flag_of_the_Czech_Republic.svg
- EN (UK): https://commons.wikimedia.org/wiki/File:Flag_of_the_United_Kingdom.svg
- DE: https://commons.wikimedia.org/wiki/File:Flag_of_Germany.svg

## 📋 Formát vstupních souborů

Aplikace očekává TXT soubory ve formátu Mitutoyo SJ-412 s následující strukturou:

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

## 🐛 Řešení problémů

### Časté problémy:

**Aplikace se nespustí:**

- Zkontrolujte instalaci Python 3.8+
- Ověřte instalaci všech závislostí: `pip install -r requirements.txt`

**Chyba při čtení TXT souboru:**

- Zkontrolujte, že soubor je ve správném formátu Mitutoyo SJ-412
- Ověřte kódování souboru (mělo by být UTF-8)

**Export do Excel nefunguje:**

- Zkontrolujte oprávnění k zápisu do cílové složky
- Ujistěte se, že cílový Excel soubor není otevřený

**Pomalé zpracování:**

- Pro velké množství souborů zvažte zpracování po menších dávkách
- Zkontrolujte dostupnou RAM

## 🤝 Přispívání

1. Fork repozitáře
2. Vytvořte feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit změny (`git commit -m 'Add some AmazingFeature'`)
4. Push do branch (`git push origin feature/AmazingFeature`)
5. Otevřete Pull Request

### Coding standards:

- Používejte Python PEP 8
- Přidejte dokumentaci ke všem funkcím
- Napište testy pro nové funkce

## 📄 Licence

Distributed under the MIT License. See `LICENSE` for more information.

## 👨‍💻 Autor

**Roman Denev (Sumixon)**

- GitHub: [@Sumixon](https://github.com/Sumixon)
- Email: romna.denev@gmail.com

## 🙏 Poděkování

- [Python Software Foundation](https://www.python.org/) za skvělý programovací jazyk
- [Pandas](https://pandas.pydata.org/) za výkonné data processing
- [OpenPyXL](https://openpyxl.readthedocs.io/) za Excel export funkcionalitu
- [Tkinter](https://docs.python.org/3/library/tkinter.html) jako základní GUI toolkit
- [customtkinter](https://github.com/TomSchimansky/CustomTkinter) za moderní vzhled aplikace

## 📈 Changelog

### v2.1.0 (2026-01-08)

- ✅ Migrace GUI na customtkinter (moderní vzhled, dark/light režim)
- ✅ Nové rozložení pomocí karet (CTkFrame) a CTkTabview
- ✅ Vylepšené čitelné zobrazení tabulky (Treeview) v tmavém režimu

### v2.0.0 (2025-01-01)

- ✅ Kompletně přepracované moderní UI
- ✅ Vylepšený parser TXT souborů s lepším error handlingem
- ✅ Rozšířená podpora všech parametrů drsnosti
- ✅ Optimalizované zpracování velkých souborů
- ✅ Přidány záložky pro lepší organizaci

### v1.0.0 (2024-12-01)

- ✅ První verze aplikace
- ✅ Základní import/export funkcionalita
- ✅ Tkinter GUI s základním designem

## 🔗 Užitečné odkazy

- [Mitutoyo SJ-412 Manual](https://mitutoyo.com/)
- [Python Documentation](https://docs.python.org/3/)
- [Pandas Documentation](https://pandas.pydata.org/docs/)
- [Tkinter Tutorial](https://docs.python.org/3/library/tkinter.html)

---

**Vytvořeno s ❤️ pro přesné měření drsnosti povrchu**
