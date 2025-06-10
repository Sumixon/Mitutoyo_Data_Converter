import tkinter
from tkinter import *
from tkinter import ttk, PhotoImage, filedialog, messagebox, scrolledtext
import os
import re
import pandas as pd
from datetime import datetime
import json  # Pro přehlednější výpis dat
import shutil  # Pro kopírování souborů

window = Tk()
window.geometry("800x600")  # Zvětšení okna pro lepší zobrazení
window.resizable(False, False)
window.config(background="grey")
window.title("Převod txt souboru do xls formátu - SJ412 Mitutoyo")

#Definice barev a písma
main_font = ("Helvetica", 12)
bg_color = "grey"

# Import složky pro TXT soubory - použijeme absolutní cestu
script_dir = os.path.dirname(os.path.abspath(__file__))
import_dir = os.path.join(script_dir, "import")
os.makedirs(import_dir, exist_ok=True)




# Funkce pro zpracování dat
def parse_txt_file(file_path, debug=False):
    """Zpracuje TXT soubor z Mitutoyo SJ-412 a vrátí slovník s hodnotami."""
    data = {}
    section = None
    raw_content = []
    
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                line = line.strip()
                raw_content.append(line)
                
                if not line:
                    continue
                    
                # Detekce sekce
                if line.startswith("//"):
                    section_name = line.replace("//", "").strip()
                    if section_name:  # Pokud není prázdný řetězec
                        section = section_name
                        data[section] = {}
                        if debug:
                            print(f"Nalezena sekce: {section}")
                    continue
                
                # Zpracování dat
                if section:
                    parts = line.split(';')
                    if len(parts) >= 2:
                        key = parts[0].strip()
                        value = parts[1].strip() if parts[1].strip() not in ["Err110", "Err116"] else "N/A"
                        unit = parts[2].strip() if len(parts) > 2 and parts[2] else ""
                        
                        if key:  # Pokud není prázdný klíč
                            data[section][key] = {"value": value, "unit": unit}
                            
                            # Pro některé důležité hodnoty je přidáme i do kořene slovníku pro snadnější přístup
                            if key in ["Ra", "Rz", "Rq", "Date"]:
                                data[key] = {"value": value, "unit": unit}
        
        # Přidáme název souboru do dat
        data["FileName"] = os.path.basename(file_path)
        data["_raw_content"] = raw_content  # Uložíme surový obsah pro ladění
        
        if debug:
            print(f"Zpracován soubor: {os.path.basename(file_path)}")
            print(f"Nalezené sekce: {list(data.keys())}")
        
        return data
    except Exception as e:
        messagebox.showerror("Chyba při zpracování souboru", f"Soubor {os.path.basename(file_path)} nelze zpracovat: {str(e)}")
        return None

def load_txt_files():
    """Načte TXT soubory z vybraného adresáře."""
    files = filedialog.askopenfilenames(
        title="Vyberte TXT soubory z měřicího přístroje",
        filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
    )
    
    
    
    if not files:
        return
    
    # Vyčistíme tabulku
    for item in file_table.get_children():
        file_table.delete(item)
    
    # Vyčistíme import složku před importem nových souborů
    try:
        existing_files = [os.path.join(import_dir, f) for f in os.listdir(import_dir) 
                         if f.lower().endswith('.txt')]
        for file_path in existing_files:
            os.remove(file_path)
    except Exception as e:
        print(f"Chyba při mazání starých souborů: {e}")

    # Zkopírujeme soubory do import adresáře a přidáme je do tabulky
    imported_files = []
    
    # Zajistíme, že adresář existuje před kopírováním
    os.makedirs(import_dir, exist_ok=True)
    
    
    for file_path in files:
        file_name = os.path.basename(file_path)
        destination = os.path.join(import_dir, file_name)
        
        try:
            # Zkusíme použít prostý copy místo copy2
            shutil.copy(file_path, destination)
            
            
            # Kontrola oprávnění
            if not os.access(import_dir, os.W_OK):
                print(f"VAROVÁNÍ: Nemáte oprávnění pro zápis do adresáře {import_dir}")
            
            data = parse_txt_file(destination)
            if data:
                imported_files.append(data)
                
                # Základní údaje pro zobrazení v tabulce
                ra_value = find_value_in_data(data, "Ra")
                rz_value = find_value_in_data(data, "Rz")
                date_value = find_value_in_data(data, "Date")
                
                file_table.insert("", "end", values=(
                    file_name,
                    date_value,
                    ra_value,
                    rz_value
                ))
        except Exception as e:
            messagebox.showerror("Chyba při importu", f"Soubor {file_name} nelze importovat: {str(e)}")
            print(f"Chyba při importu {file_name}: {e}")
    
    # Kontrola souborů v adresáři po importu
    try:
        files_in_dir = os.listdir(import_dir)
        txt_files = [f for f in files_in_dir if f.lower().endswith('.txt')]
    except Exception as e:
        print(f"Chyba při čtení adresáře: {e}")
    
    
    
    # Zobrazíme ladící informace o prvním souboru, pokud existuje
    if imported_files:
        export_btn.config(state=NORMAL)


def export_to_excel():
    """Exportuje data do Excel formátu."""
    try:
        # Zkontrolujeme, že adresář existuje
        if not os.path.exists(import_dir):
            messagebox.showerror("Chyba", f"Adresář {import_dir} neexistuje!")
            return
            
        # Výpis všech souborů v adresáři (bez filtrace)
        all_files = os.listdir(import_dir)
        
        
        # Načtení všech souborů ve složce import
        files = []
        for f in all_files:
            if f.lower().endswith('.txt'):
                full_path = os.path.join(import_dir, f)
                if os.path.isfile(full_path):
                    files.append(full_path)
        
        print(f"Soubory pro export (s plnou cestou): {files}")
        
        if not files:
            messagebox.showinfo("Info", f"Žádné soubory k exportu v adresáři {import_dir}")
            return
        
        # Zpracování všech souborů
        all_data = []
        for file_path in files:
            print(f"Zpracovávám soubor: {file_path}")
            print(f"Soubor existuje: {os.path.exists(file_path)}")
            data = parse_txt_file(file_path)
            if data:
                all_data.append(data)
        
        
        
        if not all_data:
            messagebox.showinfo("Info", "Žádná data k exportu")
            return
            
        
        
        # Vytvoření DataFrame pro Excel
        excel_data = []
        for data in all_data:
            # Funkce pro hledání hodnoty v různých částech dat
            row = {"Soubor": data.get("FileName", "")}
            
            row["Datum"] = find_value_in_data(data, "Date")
            row["Ra [μm]"] = find_value_in_data(data, "Ra")
            row["Rq [μm]"] = find_value_in_data(data, "Rq")
            row["Rz [μm]"] = find_value_in_data(data, "Rz")
            row["Rp [μm]"] = find_value_in_data(data, "Rp")
            row["Rv [μm]"] = find_value_in_data(data, "Rv")
            row["Rsk [μm]"] = find_value_in_data(data, "Rsk")
            row["Rku [μm]"] = find_value_in_data(data, "Rku")
            row["Rc [μm]"] = find_value_in_data(data, "Rc")
            row["RPc [/cm]"] = find_value_in_data(data, "RPc")
            row["RSm [μm]"] = find_value_in_data(data, "RSm")
            row["RDq [μm]"] = find_value_in_data(data, "RDq")
            row["Rmr [%]"] = find_value_in_data(data, "Rmr")
            row["Rdc [μm]"] = find_value_in_data(data, "Rdc")            
            row["Rt [μm]"] = find_value_in_data(data, "Rt")
            row["Rz1max [μm]"] = find_value_in_data(data, "Rz1max")
            row["Rk [μm]"] = find_value_in_data(data, "Rk")
            row["Rpk [μm]"] = find_value_in_data(data, "Rpk")
            row["Rvk [μm]"] = find_value_in_data(data, "Rvk")
            row["Mr1 [%]"] = find_value_in_data(data, "Mr1")
            row["Mr2 [%]"] = find_value_in_data(data, "Mr2")
            row["A1 []"] = find_value_in_data(data, "A1")
            row["A2 []"] = find_value_in_data(data, "A2")
            
            excel_data.append(row)
        
        df = pd.DataFrame(excel_data)
        
        # Výběr kam uložit XLSX soubor
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"mitutoyo_data_{now}.xlsx"
        )
        
        if file_path:
            try:
                df.to_excel(file_path, index=False)
                messagebox.showinfo("Export úspěšný", f"Data byla uložena do souboru:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Chyba při exportu", f"Nelze uložit Excel soubor: {str(e)}")
                print(f"Chyba při exportu: {e}")
    except Exception as e:
        messagebox.showerror("Chyba", f"Nastala neočekávaná chyba: {str(e)}")
        print(f"Chyba při exportu: {e}")


def clear_files():
    """Vymaže všechny TXT soubory z import složky."""
    try:
        # Načteme všechny TXT soubory (bez ohledu na velikost písmen)
        files = [os.path.join(import_dir, f) for f in os.listdir(import_dir) 
                if f.lower().endswith('.txt')]
        
        if not files:
            messagebox.showinfo("Info", "Žádné soubory ke smazání")
            return
        
        if messagebox.askyesno("Potvrdit smazání", "Opravdu chcete smazat všechny TXT soubory?"):
            for file_path in files:
                try:
                    os.remove(file_path)
                except Exception as e:
                    messagebox.showerror("Chyba při mazání", 
                        f"Soubor {os.path.basename(file_path)} nelze smazat: {str(e)}")
            
            # Vyčistíme tabulku
            for item in file_table.get_children():
                file_table.delete(item)
            
            # Aktualizujeme počet souborů a stav tlačítek
            export_btn.config(state=DISABLED)
            
            # Kontrola po smazání
            remaining = [f for f in os.listdir(import_dir) if f.lower().endswith('.txt')]
            
            
            messagebox.showinfo("Hotovo", "Soubory byly úspěšně smazány")
    except Exception as e:
        print(f"Chyba při mazání souborů: {e}")
        messagebox.showerror("Chyba", f"Nastala neočekávaná chyba: {str(e)}")
       



def find_value_in_data(data, key):
    """Pomocná funkce pro hledání hodnoty v různých částech dat"""
    if key in data:
        return data[key].get("value", "")
    for section in ["CalcResult", "Header", "Condition-A"]:
        if section in data and key in data[section]:
            return data[section][key].get("value", "")
    return ""

# Záložky programu
nb = ttk.Notebook(window)
nb.place(x=0, y=0, width=800, height=600)
tab1 = tkinter.Frame(window, borderwidth=1, highlightcolor="black", background=bg_color)
nb.add(tab1, text="Import")
tab2 = tkinter.Frame(window, borderwidth=1, highlightcolor="black", background=bg_color)
nb.add(tab2, text="Nastavení")
tab3 = tkinter.Frame(window, borderwidth=1, highlightcolor="black", background=bg_color)
nb.add(tab3, text="O programu")

# ---------- ZÁLOŽKA IMPORT ----------
# Frame pro tlačítka
button_frame = Frame(tab1, bg=bg_color)
button_frame.pack(pady=10, fill=X)

# Tlačítka pro import a export
import_btn = Button(button_frame, text="Importovat TXT soubory", command=load_txt_files, font=main_font)
import_btn.pack(side=LEFT, padx=10)

export_btn = Button(button_frame, text="Exportovat do Excel", command=export_to_excel, font=main_font, state=DISABLED)
export_btn.pack(side=LEFT, padx=10)

clear_btn = Button(button_frame, text="Vymazat soubory", command=clear_files, font=main_font)
clear_btn.pack(side=LEFT, padx=10)


# Tabulka pro zobrazení importovaných souborů
table_frame = Frame(tab1, bg=bg_color)
table_frame.pack(pady=10, padx=10, fill=BOTH, expand=True)

# Scrollbar pro tabulku
table_scroll = Scrollbar(table_frame)
table_scroll.pack(side=RIGHT, fill=Y)

# Tabulka
columns = ("file", "date", "ra", "rz")
file_table = ttk.Treeview(table_frame, columns=columns, show="headings", yscrollcommand=table_scroll.set)

# Konfigurace sloupců
file_table.heading("file", text="Soubor")
file_table.heading("date", text="Datum")
file_table.heading("ra", text="Ra [μm]")
file_table.heading("rz", text="Rz [μm]")

file_table.column("file", width=150)
file_table.column("date", width=150)
file_table.column("ra", width=100)
file_table.column("rz", width=100)

file_table.pack(fill=BOTH, expand=True)
table_scroll.config(command=file_table.yview)

# Přidáme událost pro kliknutí na řádek
#file_table.bind("<Double-1>", show_file_details)

# Frame pro instrukce
info_frame = Frame(tab1, bg=bg_color)
info_frame.pack(pady=10, fill=X)

info_text = Label(info_frame, 
                 text="Instrukce: Klikněte na 'Importovat TXT soubory' pro výběr souborů z měřicího přístroje.\n" +
                      "Po importu můžete exportovat data do Excelu pomocí tlačítka 'Exportovat do Excel'.\n" +
                      "Pro zobrazení detailů souboru klikněte dvakrát na řádek v tabulce.",
                 font=main_font, bg=bg_color, justify=LEFT)
info_text.pack(padx=10)

# ---------- ZÁLOŽKA O PROGRAMU ----------
nb.select(tab1)  # Změnil jsem výchozí záložku na Import
nb.enable_traversal()

# Cesta k adresáři se souborem
current_dir = os.path.dirname(os.path.abspath(__file__))
img_path = os.path.join(current_dir, "img", "logo_male130x50.png")

os.makedirs(os.path.join(current_dir, "img"), exist_ok=True)

try:
    # Načtení loga Miele
    logo_img = PhotoImage(file=img_path)
    logo_label = Label(tab3, image=logo_img, bg=bg_color)
    logo_label.pack(pady=10)
    
except Exception as e:
    # Záložní řešení, pokud by soubor s logem nebyl nalezen
    print(f"Nepodařilo se načíst logo: {e}")
    
    # Zobrazení textu místo loga
    logo_label = Label(tab3, text="Miele", font=("Arial", 18, "bold"), fg="white", bg="darkred", 
                       width=15, height=2)
    logo_label.pack(pady=10)

# Přidání textu aplikace
app_label = Label(tab3, text="Aplikace pro převod dat\nMitutoyo-SJ 412", font=("Arial", 14, "bold"), bg=bg_color)
app_label.pack(pady=10)

# Přidání dalších informací
info_text = "Verze: 1.0\nVytvořeno: 2025\nAutor: Roman Denev\n\n" \
            "Aplikace slouží k převodu dat z měřicího přístroje Mitutoyo SJ 412 do formátu XLS.\n" \
            "Pro použití aplikace je nutné mít nainstalovaný Python a potřebné knihovny.\n\n" \
            "Pokud máte jakékoli dotazy nebo potřebujete pomoc, neváhejte kontaktovat autora.\n\n" \
            "Děkujeme za použití naší aplikace!"
info_label = Label(tab3, text=info_text, bg=bg_color, font=main_font)
info_label.pack(pady=20)



window.mainloop()