import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import re
import pandas as pd
from datetime import datetime
import json
import shutil

class ModernApp:
    def __init__(self):
        self.window = tk.Tk()
        self.setup_window()
        self.setup_style()
        self.setup_variables()
        self.create_widgets()
        
    def setup_window(self):
        """Nastavení hlavního okna"""
        self.window.geometry("1000x700")
        self.window.minsize(800, 600)
        self.window.title("Převod txt souboru do xls formátu - SJ412 Mitutoyo")
        self.window.configure(bg='#f0f0f0')
        
        # Centrování okna
        self.center_window()
        
    def center_window(self):
        """Vycentruje okno na obrazovce"""
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (1000 // 2)
        y = (self.window.winfo_screenheight() // 2) - (700 // 2)
        self.window.geometry(f"1000x700+{x}+{y}")
        
    def setup_style(self):
        """Nastavení moderního stylu"""
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Moderní barvy
        colors = {
            'primary': '#2563eb',      # Modrá
            'primary_hover': '#1d4ed8',
            'success': '#059669',      # Zelená
            'danger': '#dc2626',       # Červená
            'secondary': '#6b7280',    # Šedá
            'background': '#f8fafc',   # Světle šedá
            'surface': '#ffffff',      # Bílá
            'text': '#1f2937',         # Tmavá
            'text_light': '#6b7280'    # Světle šedá
        }
        
        # Styl pro tlačítka
        self.style.configure('Primary.TButton',
                           background=colors['primary'],
                           foreground='white',
                           borderwidth=0,
                           focuscolor='none',
                           font=('Segoe UI', 10, 'bold'),
                           padding=(15, 8))
        
        self.style.map('Primary.TButton',
                      background=[('active', colors['primary_hover']),
                                ('pressed', colors['primary_hover'])])
        
        self.style.configure('Success.TButton',
                           background=colors['success'],
                           foreground='white',
                           borderwidth=0,
                           focuscolor='none',
                           font=('Segoe UI', 10, 'bold'),
                           padding=(15, 8))
        
        self.style.configure('Danger.TButton',
                           background=colors['danger'],
                           foreground='white',
                           borderwidth=0,
                           focuscolor='none',
                           font=('Segoe UI', 10, 'bold'),
                           padding=(15, 8))
        
        # Styl pro Notebook
        self.style.configure('TNotebook',
                           background=colors['background'],
                           borderwidth=0)
        
        self.style.configure('TNotebook.Tab',
                           background=colors['surface'],
                           foreground=colors['text'],
                           padding=(20, 12),
                           font=('Segoe UI', 10, 'bold'))
        
        self.style.map('TNotebook.Tab',
                      background=[('selected', colors['primary']),
                                ('active', colors['background'])],
                      foreground=[('selected', 'white')])
        
        # Styl pro Treeview
        self.style.configure('Modern.Treeview',
                           background=colors['surface'],
                           foreground=colors['text'],
                           rowheight=30,
                           fieldbackground=colors['surface'],
                           font=('Segoe UI', 9)
                           )
        
        self.style.configure('Modern.Treeview.Heading',
                           background=colors['primary'],
                           foreground='white',
                           font=('Segoe UI', 10, 'bold'))
        
    def setup_variables(self):
        """Nastavení proměnných"""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.import_dir = os.path.join(script_dir, "import")
        os.makedirs(self.import_dir, exist_ok=True)
        
    def create_widgets(self):
        """Vytvoření widgets"""
        # Hlavní container
        main_container = ttk.Frame(self.window)
        main_container.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Nadpis
        title_frame = ttk.Frame(main_container)
        title_frame.pack(fill='x', pady=(0, 20))
        
        title_label = ttk.Label(title_frame, 
                               text="Převod dat z Mitutoyo SJ-412",
                               font=('Segoe UI', 20, 'bold'))
        title_label.pack()
        
        subtitle_label = ttk.Label(title_frame,
                                  text="Aplikace pro zpracování měřicích dat",
                                  font=('Segoe UI', 10),
                                  foreground='#6b7280')
        subtitle_label.pack()
        
        # Notebook pro záložky
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill='both', expand=True)
        
        # Záložky
        self.create_import_tab()
        self.create_settings_tab()
        self.create_about_tab()
        
    def create_import_tab(self):
        """Vytvoření záložky Import"""
        import_tab = ttk.Frame(self.notebook)
        self.notebook.add(import_tab, text="📁 Import")
        
        # Container pro obsah
        content_frame = ttk.Frame(import_tab)
        content_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Panel s tlačítky
        button_panel = ttk.Frame(content_frame)
        button_panel.pack(fill='x', pady=(0, 20))
        
        # Card styl pro tlačítka
        card_frame = ttk.Frame(button_panel, relief='solid', borderwidth=1)
        card_frame.pack(fill='x', pady=10)
        card_frame.configure(style='Card.TFrame')
        
        button_container = ttk.Frame(card_frame)
        button_container.pack(padx=20, pady=15)
        
        self.import_btn = ttk.Button(button_container,
                                   text="📂 Importovat soubory",
                                   command=self.load_txt_files,
                                   style='Primary.TButton')
        self.import_btn.pack(side='left', padx=(0, 10))
        
        self.export_btn = ttk.Button(button_container,
                                   text="📊 Exportovat do Excel",
                                   command=self.export_to_excel,
                                   style='Success.TButton',
                                   state='disabled')
        self.export_btn.pack(side='left', padx=(0, 10))
        
        self.clear_btn = ttk.Button(button_container,
                                  text="🗑️ Vymazat soubory",
                                  command=self.clear_files,
                                  style='Danger.TButton')
        self.clear_btn.pack(side='left')
        
        # Tabulka
        table_frame = ttk.Frame(content_frame)
        table_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # Nadpis tabulky
        table_title = ttk.Label(table_frame,
                               text="Importované soubory",
                               font=('Segoe UI', 12, 'bold'))
        table_title.pack(anchor='w', pady=(0, 10))
        
        # Treeview s moderním stylem
        tree_container = ttk.Frame(table_frame)
        tree_container.pack(fill='both', expand=True)
        
        # Scrollbary
        v_scrollbar = ttk.Scrollbar(tree_container, orient='vertical')
        h_scrollbar = ttk.Scrollbar(tree_container, orient='horizontal')
        
        columns = ("file", "date", "ra", "rz")
        self.file_table = ttk.Treeview(tree_container,
                                     columns=columns,
                                     show="headings",
                                     style='Modern.Treeview',
                                     yscrollcommand=v_scrollbar.set,
                                     xscrollcommand=h_scrollbar.set)
        
        # Konfigurace sloupců
        self.file_table.heading("file", text="📄 Soubor")
        self.file_table.heading("date", text="📅 Datum")
        self.file_table.heading("ra", text="📏 Ra [μm]")
        self.file_table.heading("rz", text="📐 Rz [μm]")
        
        self.file_table.column("file", width=200, minwidth=150, anchor='w')
        self.file_table.column("date", width=150, minwidth=100, anchor='center')
        self.file_table.column("ra", width=100, minwidth=80, anchor='center')
        self.file_table.column("rz", width=100, minwidth=80, anchor='center')
        self.file_table.tag_configure('oddrow', background='#f0f0f0')
        self.file_table.tag_configure('evenrow', background='#ffffff')

        # Umístění treeview a scrollbarů
        self.file_table.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')

        
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)
        
        v_scrollbar.config(command=self.file_table.yview)
        h_scrollbar.config(command=self.file_table.xview)
        
        # Instrukce
        info_frame = ttk.Frame(content_frame)
        info_frame.pack(fill='x', pady=(10, 0))
        
        info_text = ttk.Label(info_frame,
                            text="💡 Tip: Vyberte TXT soubory z měřicího přístroje a exportujte je do Excelu",
                            font=('Segoe UI', 9),
                            foreground='#6b7280')
        info_text.pack()
        
    def create_settings_tab(self):
        """Vytvoření záložky Nastavení"""
        settings_tab = ttk.Frame(self.notebook)
        self.notebook.add(settings_tab, text="⚙️ Nastavení")
        
        content_frame = ttk.Frame(settings_tab)
        content_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        title = ttk.Label(content_frame,
                         text="Nastavení aplikace",
                         font=('Segoe UI', 16, 'bold'))
        title.pack(pady=(0, 20))
        
        # Zde můžete přidat nastavení podle potřeby
        placeholder = ttk.Label(content_frame,
                               text="Nastavení budou přidána v další verzi",
                               font=('Segoe UI', 10),
                               foreground='#6b7280')        
        placeholder.pack()
        
    def create_about_tab(self):
        """Vytvoření záložky O programu"""
        about_tab = ttk.Frame(self.notebook)
        self.notebook.add(about_tab, text="ℹ️ O programu")
        
        content_frame = ttk.Frame(about_tab)
        content_frame.pack(fill='both', expand=True, padx=40, pady=40)
        
        # Logo nebo ikona
        logo_frame = ttk.Frame(content_frame)
        logo_frame.pack(pady=(0, 20))
        
        logo_label = ttk.Label(logo_frame,
                              text="🔧",
                              font=('Segoe UI', 48))
        logo_label.pack()
        
        # Informace o aplikaci
        app_title = ttk.Label(content_frame,
                             text="Mitutoyo Data Converter",
                             font=('Segoe UI', 18, 'bold'))
        app_title.pack()
        
        version_label = ttk.Label(content_frame,
                                 text="Verze 2.0 - Moderní edice",
                                 font=('Segoe UI', 12),
                                 foreground='#6b7280')
        version_label.pack(pady=(5, 20))
        
        info_text = """Moderní aplikace pro převod dat z měřicího přístroje Mitutoyo SJ-412 do formátu Excel.

✨ Funkce:
• Import TXT souborů z měřicího přístroje
• Automatické zpracování a analýza dat
• Export do Excel formátu
• Podpora různých měřicích parametrů

👨‍💻 Autor: Roman Denev
📅 Vytvořeno: 2025
🐍 Technologie: Python, Tkinter, Pandas"""
        
        info_label = ttk.Label(content_frame,
                              text=info_text,
                              font=('Segoe UI', 10),
                              justify='left')
        info_label.pack()
        
    # Původní funkce s malými úpravami pro kompatibilitu
    def load_txt_files(self):
        """Načte TXT soubory z vybraného adresáře."""
        files = filedialog.askopenfilenames(
            title="Vyberte TXT soubory z měřicího přístroje",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if not files:
            return
        
        # Vyčistíme tabulku
        for item in self.file_table.get_children():
            self.file_table.delete(item)
        
        # Vyčistíme import složku před importem nových souborů
        try:
            existing_files = [os.path.join(self.import_dir, f) for f in os.listdir(self.import_dir) 
                             if f.lower().endswith('.txt')]
            for file_path in existing_files:
                os.remove(file_path)
        except Exception as e:
            print(f"Chyba při mazání starých souborů: {e}")

        # Zkopírujeme soubory do import adresáře a přidáme je do tabulky
        imported_files = []
        
        # Zajistíme, že adresář existuje před kopírováním
        os.makedirs(self.import_dir, exist_ok=True)
        
        for file_path in files:
            file_name = os.path.basename(file_path)
            destination = os.path.join(self.import_dir, file_name)
            
            try:
                shutil.copy(file_path, destination)
                
                data = self.parse_txt_file(destination)
                if data:
                    imported_files.append(data)
                    
                    # Základní údaje pro zobrazení v tabulce
                    ra_value = self.find_value_in_data(data, "Ra")
                    rz_value = self.find_value_in_data(data, "Rz")
                    date_value = self.find_value_in_data(data, "Date")
                    
                    self.file_table.insert("", "end", values=(
                        file_name,
                        date_value,
                        ra_value,
                        rz_value
                    ))
            except Exception as e:
                messagebox.showerror("Chyba při importu", f"Soubor {file_name} nelze importovat: {str(e)}")
                print(f"Chyba při importu {file_name}: {e}")
        
        if imported_files:
            self.export_btn.config(state='normal')

    def export_to_excel(self):
        """Exportuje data do Excel formátu."""
        try:
            if not os.path.exists(self.import_dir):
                messagebox.showerror("Chyba", f"Adresář {self.import_dir} neexistuje!")
                return
                
            all_files = os.listdir(self.import_dir)
            files = []
            for f in all_files:
                if f.lower().endswith('.txt'):
                    full_path = os.path.join(self.import_dir, f)
                    if os.path.isfile(full_path):
                        files.append(full_path)
            
            if not files:
                messagebox.showinfo("Info", f"Žádné soubory k exportu v adresáři {self.import_dir}")
                return
            
            # Zpracování všech souborů
            all_data = []
            for file_path in files:
                data = self.parse_txt_file(file_path)
                if data:
                    all_data.append(data)
            
            if not all_data:
                messagebox.showinfo("Info", "Žádná data k exportu")
                return
                
            # Vytvoření DataFrame pro Excel
            excel_data = []
            for data in all_data:
                row = {"Soubor": data.get("FileName", "")}
                
                row["Datum"] = self.find_value_in_data(data, "Date")
                row["Ra [μm]"] = self.find_value_in_data(data, "Ra")
                row["Rq [μm]"] = self.find_value_in_data(data, "Rq")
                row["Rz [μm]"] = self.find_value_in_data(data, "Rz")
                row["Rp [μm]"] = self.find_value_in_data(data, "Rp")
                row["Rv [μm]"] = self.find_value_in_data(data, "Rv")
                row["Rsk [μm]"] = self.find_value_in_data(data, "Rsk")
                row["Rku [μm]"] = self.find_value_in_data(data, "Rku")
                row["Rc [μm]"] = self.find_value_in_data(data, "Rc")
                row["RPc [/cm]"] = self.find_value_in_data(data, "RPc")
                row["RSm [μm]"] = self.find_value_in_data(data, "RSm")
                row["RDq [μm]"] = self.find_value_in_data(data, "RDq")
                row["Rmr [%]"] = self.find_value_in_data(data, "Rmr")
                row["Rdc [μm]"] = self.find_value_in_data(data, "Rdc")            
                row["Rt [μm]"] = self.find_value_in_data(data, "Rt")
                row["Rz1max [μm]"] = self.find_value_in_data(data, "Rz1max")
                row["Rk [μm]"] = self.find_value_in_data(data, "Rk")
                row["Rpk [μm]"] = self.find_value_in_data(data, "Rpk")
                row["Rvk [μm]"] = self.find_value_in_data(data, "Rvk")
                row["Mr1 [%]"] = self.find_value_in_data(data, "Mr1")
                row["Mr2 [%]"] = self.find_value_in_data(data, "Mr2")
                row["A1 []"] = self.find_value_in_data(data, "A1")
                row["A2 []"] = self.find_value_in_data(data, "A2")
                
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

    def clear_files(self):
        """Vymaže všechny TXT soubory z import složky."""
        try:
            files = [os.path.join(self.import_dir, f) for f in os.listdir(self.import_dir) 
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
                for item in self.file_table.get_children():
                    self.file_table.delete(item)
                
                self.export_btn.config(state='disabled')
                messagebox.showinfo("Hotovo", "Soubory byly úspěšně smazány")
        except Exception as e:
            print(f"Chyba při mazání souborů: {e}")
            messagebox.showerror("Chyba", f"Nastala neočekávaná chyba: {str(e)}")

    def parse_txt_file(self, file_path, debug=False):
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
                        if section_name:
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
                            
                            if key:
                                data[section][key] = {"value": value, "unit": unit}
                                
                                if key in ["Ra", "Rz", "Rq", "Date"]:
                                    data[key] = {"value": value, "unit": unit}
            
            data["FileName"] = os.path.basename(file_path)
            data["_raw_content"] = raw_content
            
            if debug:
                print(f"Zpracován soubor: {os.path.basename(file_path)}")
                print(f"Nalezené sekce: {list(data.keys())}")
            
            return data
        except Exception as e:
            messagebox.showerror("Chyba při zpracování souboru", f"Soubor {os.path.basename(file_path)} nelze zpracovat: {str(e)}")
            return None

    def find_value_in_data(self, data, key):
        """Pomocná funkce pro hledání hodnoty v různých částech dat"""
        if key in data:
            return data[key].get("value", "")
        for section in ["CalcResult", "Header", "Condition-A"]:
            if section in data and key in data[section]:
                return data[section][key].get("value", "")
        return ""

    def run(self):
        """Spuštění aplikace"""
        self.window.mainloop()

# Spuštění aplikace
if __name__ == "__main__":
    app = ModernApp()
    app.run()