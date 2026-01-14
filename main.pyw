import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import re
import pandas as pd
from datetime import datetime
import json
import shutil
import customtkinter as ctk

try:
    from PIL import Image
except Exception:
    Image = None

class ModernApp:
    def __init__(self):
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        self.window = ctk.CTk()
        self.setup_variables()
        self.setup_window()
        self.setup_style()
        self.create_widgets()
        self.refresh_file_table()
        
        
    def setup_window(self):
        """Nastavení hlavního okna"""
        self.window.geometry("1000x800")
        self.window.minsize(800, 800)
        self.window.title(self.t("app.window_title"))

        
        # Centrování okna
        self.center_window()
        
    def center_window(self):
        """Vycentruje okno na obrazovce"""
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (1000 // 2)
        y = (self.window.winfo_screenheight() // 2) - (750 // 2)
        self.window.geometry(f"1000x700+{x}+{y}")
        
    def setup_style(self):
        """Nastavení moderního stylu"""
        self.style = ttk.Style()
        try:
            self.style.theme_use('clam')
        except Exception:
            pass
        
        # Barvy podle vzhledu customtkinter (Light / Dark)
        appearance = ctk.get_appearance_mode()
        if appearance == "Dark":
            colors = {
                'primary': '#3b82f6',      # Modrá
                'primary_hover': '#2563eb',
                'success': '#22c55e',      # Zelená
                'danger': '#ef4444',       # Červená
                'secondary': '#9ca3af',    # Šedá
                'background': '#020617',   # Velmi tmavé pozadí
                'surface': '#0f172a',      # Tmavá karta
                'text': '#e5e7eb',         # Světlý text
                'text_light': '#9ca3af',   # Světle šedý text
                'row_odd': '#0b1220',
                'row_even': '#0f172a'
            }
        else:
            colors = {
                'primary': '#2563eb',      # Modrá
                'primary_hover': '#1d4ed8',
                'success': '#059669',      # Zelená
                'danger': '#dc2626',       # Červená
                'secondary': '#6b7280',    # Šedá
                'background': '#f3f4f6',   # Světle šedá
                'surface': '#e5e7eb',      # Mírně šedá (ne čistě bílá)
                'text': '#111827',         # Tmavý text
                'text_light': '#6b7280',   # Světle šedá
                'row_odd': '#f3f4f6',
                'row_even': '#ffffff'
            }

        self.colors = colors

        # Základní styl pro ttk Label (platí jen pro ttk widgety, ne pro customtkinter CTkLabel)
        self.style.configure(
            'TLabel',
            font=('Segoe UI', 16),
            background=self.colors['background'],
            foreground=self.colors['text']
        )
        
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
        
        # Styl pro Treeview – aby nelískal bílou v Dark režimu
        self.style.configure(
            'Modern.Treeview',
            background=colors['surface'],
            foreground=colors['text'],
            rowheight=30,
            fieldbackground=colors['surface'],
            font=('Segoe UI', 9),
            borderwidth=0
        )

        self.style.configure(
            'Modern.Treeview.Heading',
            background=colors['primary'],
            foreground='white',
            font=('Segoe UI', 14, 'bold')
        )

        self.style.map(
            'Modern.Treeview',
            background=[('selected', colors['primary'])],
            foreground=[('selected', 'white')]
        )
        
    def setup_variables(self):
        """Nastavení proměnných"""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.script_dir = script_dir

        self.settings_path = os.path.join(script_dir, "settings.json")
        self.translations_path = os.path.join(script_dir, "locales", "translations.json")

        self.translations = self.load_translations()
        self.language = self.load_settings().get("language", "cs")

        self.import_dir = os.path.join(script_dir, "import")
        os.makedirs(self.import_dir, exist_ok=True)

        # Vlajky se vkládají staticky do projektu (žádné automatické stahování).
        # Podporované cesty (preferované): img/flags/*.png, fallback: img/*.png
        self.flag_dir = os.path.join(script_dir, "img", "flags")
        os.makedirs(self.flag_dir, exist_ok=True)

        self.flag_files = {
            "cs": [
                os.path.join(self.flag_dir, "cz.png"),
                os.path.join(script_dir, "img", "cz.png")
            ],
            "en": [
                os.path.join(self.flag_dir, "en.png"),
                os.path.join(script_dir, "img", "en.png")
            ],
            "de": [
                os.path.join(self.flag_dir, "de.png"),
                os.path.join(script_dir, "img", "de.png")
            ]
        }

        self.flag_images = self.load_flag_images()
        
    def create_widgets(self):
        """Vytvoření widgets"""
        # Hlavní container
        main_container = ctk.CTkFrame(self.window, fg_color="transparent")
        main_container.pack(fill='both', expand=True, padx=20, pady=20)

        # Header (nadpis + přepínač jazyka)
        header_frame = ctk.CTkFrame(main_container, fg_color="transparent")
        header_frame.pack(fill='x', pady=(0, 20))

        header_left = ctk.CTkFrame(header_frame, fg_color="transparent")
        header_left.pack(side='left', fill='x', expand=True)

        header_right = ctk.CTkFrame(header_frame, fg_color="transparent")
        header_right.pack(side='right')

        self.title_label = ctk.CTkLabel(
            header_left,
            text=self.t("app.title"),
            font=('Segoe UI', 32, 'bold')
        )
        self.title_label.pack(anchor='w')

        self.subtitle_label = ctk.CTkLabel(
            header_left,
            text=self.t("app.subtitle"),
            font=('Segoe UI', 20),
            text_color=getattr(self, "colors", {}).get('text_light', '#6b7280')
        )
        self.subtitle_label.pack(anchor='w')

        self.create_language_switcher(header_right)
        
        # Moderní záložky
        self.tabview = ctk.CTkTabview(main_container)
        self.tabview.pack(fill='both', expand=True)

        # Záložky
        self.import_tab = self.tabview.add(self.t("tabs.import"))
        self.settings_tab = self.tabview.add(self.t("tabs.settings"))
        self.about_tab = self.tabview.add(self.t("tabs.about"))

        self.create_import_tab(self.import_tab)
        self.create_settings_tab(self.settings_tab)
        self.create_about_tab(self.about_tab)
        
    def create_import_tab(self, parent):
        """Vytvoření záložky Import"""
        # Container pro obsah
        content_frame = ctk.CTkFrame(parent, fg_color="transparent")
        content_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Panel s tlačítky
        button_panel = ctk.CTkFrame(content_frame, fg_color="transparent")
        button_panel.pack(fill='x', pady=(0, 20))
        
        # Card styl pro tlačítka
        card_frame = ctk.CTkFrame(button_panel, corner_radius=12)
        card_frame.pack(fill='x', pady=10)
        
        button_container = ctk.CTkFrame(card_frame, fg_color="transparent")
        button_container.pack(padx=20, pady=15, fill='x')
        
        self.import_btn = ctk.CTkButton(
            button_container,
            text=self.t("buttons.import_files"),
            command=self.load_txt_files
        )
        self.import_btn.pack(side='left', padx=(0, 10))
        
        self.export_btn = ctk.CTkButton(
            button_container,
            text=self.t("buttons.export_excel"),
            command=self.export_to_excel,
            state='disabled'
        )
        self.export_btn.pack(side='left', padx=(0, 10))
        
        self.clear_btn = ctk.CTkButton(
            button_container,
            text=self.t("buttons.clear_files"),
            command=self.clear_files,
            fg_color="#dc2626",
            hover_color="#b91c1c"
        )
        self.clear_btn.pack(side='left')
        
        # Tabulka
        table_card = ctk.CTkFrame(content_frame, corner_radius=12)
        table_card.pack(fill='both', expand=True, pady=(0, 10))

        table_frame = ctk.CTkFrame(table_card, fg_color="transparent")
        table_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # Nadpis tabulky
        table_title = ctk.CTkLabel(
            table_frame,
            text=self.t("table.title"),
            font=('Segoe UI', 16, 'bold')
        )
        table_title.pack(anchor='w', padx=16, pady=(16, 10))
        
        # Treeview s moderním stylem
        tree_container = ttk.Frame(table_frame)
        tree_container.pack(fill='both', expand=True, padx=16, pady=(0, 16))
        
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
        self.file_table.heading("file", text=self.t("table.col.file"))
        self.file_table.heading("date", text=self.t("table.col.date"))
        self.file_table.heading("ra", text=self.t("table.col.ra"))
        self.file_table.heading("rz", text=self.t("table.col.rz"))
        
        self.file_table.column("file", width=200, minwidth=150, anchor='w')
        self.file_table.column("date", width=150, minwidth=100, anchor='center')
        self.file_table.column("ra", width=100, minwidth=80, anchor='center')
        self.file_table.column("rz", width=100, minwidth=80, anchor='center')
        odd = getattr(self, "colors", {}).get('row_odd', '#f0f0f0')
        even = getattr(self, "colors", {}).get('row_even', '#ffffff')
        self.file_table.tag_configure('oddrow', background=odd)
        self.file_table.tag_configure('evenrow', background=even)

        # Umístění treeview a scrollbarů
        self.file_table.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')

        
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)
        
        v_scrollbar.config(command=self.file_table.yview)
        h_scrollbar.config(command=self.file_table.xview)
        
        # Instrukce
        info_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        info_frame.pack(fill='x', pady=(10, 0))
        
        info_text = ctk.CTkLabel(
            info_frame,
            text=self.t("tip"),
            font=('Segoe UI', 16),
            text_color=getattr(self, "colors", {}).get('text_light', '#6b7280')
        )
        info_text.pack()
        
    def create_settings_tab(self, parent):
        """Vytvoření záložky Nastavení"""
        content_frame = ctk.CTkFrame(parent, fg_color="transparent")
        content_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        card = ctk.CTkFrame(content_frame, corner_radius=12)
        card.pack(fill='x', pady=(0, 20))

        title = ctk.CTkLabel(
            card,
            text=self.t("settings.title"),
            font=('Segoe UI', 16, 'bold')
        )
        title.pack(padx=16, pady=(16, 8), anchor='w')
        
        # Zde můžete přidat nastavení podle potřeby
        placeholder = ctk.CTkLabel(
            card,
            text=self.t("settings.placeholder"),
            font=('Segoe UI', 10),
            text_color=getattr(self, "colors", {}).get('text_light', '#6b7280')
        )
        placeholder.pack(padx=16, pady=(0, 16), anchor='w')
        
    def create_about_tab(self, parent):
        """Vytvoření záložky O programu"""
        content_frame = ctk.CTkFrame(parent, fg_color="transparent")
        content_frame.pack(fill='both', expand=True, padx=20, pady=20)

        card = ctk.CTkFrame(content_frame, corner_radius=12)
        card.pack(fill='both', expand=True)
        
        # Logo nebo ikona
        logo_frame = ctk.CTkFrame(card, fg_color="transparent")
        logo_frame.pack(pady=(16, 10))
        
        logo_label = ctk.CTkLabel(
            logo_frame,
            text="🔧",
            font=('Segoe UI', 48)
        )
        logo_label.pack()
        
        # Informace o aplikaci
        app_title = ctk.CTkLabel(
            card,
            text=self.t("about.title"),
            font=('Segoe UI', 18, 'bold')
        )
        app_title.pack()
        
        version_label = ctk.CTkLabel(
            card,
            text=self.t("about.version"),
            font=('Segoe UI', 12),
            text_color=getattr(self, "colors", {}).get('text_light', '#6b7280')
        )
        version_label.pack(pady=(5, 14))
        
        info_text = self.t("about.info")
        
        info_label = ctk.CTkLabel(
            card,
            text=info_text,
            font=('Segoe UI', 18),
            justify='left'
        )
        info_label.pack(padx=16, pady=(0, 16), anchor='w')
        
    # Původní funkce s malými úpravami pro kompatibilitu
    def load_txt_files(self):
        """Načte TXT soubory z vybraného adresáře."""
        files = filedialog.askopenfilenames(
            title=self.t("dialogs.select_files.title"),
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
            except Exception as e:
                messagebox.showerror(
                    self.t("import.error.title"),
                    self.t("import.error.message", file=file_name, err=str(e))
                )
                print(f"Chyba při importu {file_name}: {e}")

        self.refresh_file_table()

    def export_to_excel(self):
        """Exportuje data do Excel formátu."""
        try:
            if not os.path.exists(self.import_dir):
                messagebox.showerror(self.t("common.error"), self.t("export.dir_missing", dir=self.import_dir))
                return
                
            all_files = os.listdir(self.import_dir)
            files = []
            for f in all_files:
                if f.lower().endswith('.txt'):
                    full_path = os.path.join(self.import_dir, f)
                    if os.path.isfile(full_path):
                        files.append(full_path)
            
            if not files:
                messagebox.showinfo(self.t("common.info"), self.t("export.no_files", dir=self.import_dir))
                return
            
            # Zpracování všech souborů
            all_data = []
            for file_path in files:
                data = self.parse_txt_file(file_path)
                if data:
                    all_data.append(data)
            
            if not all_data:
                messagebox.showinfo(self.t("common.info"), self.t("export.no_data"))
                return
                
            # Vytvoření DataFrame pro Excel
            excel_data = []
            for data in all_data:
                row = {self.t("excel.col.file"): data.get("FileName", "")}
                
                row[self.t("excel.col.date")] = self.find_value_in_data(data, "Date")
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
                    messagebox.showinfo(
                        self.t("export.success.title"),
                        self.t("export.success.message", file=file_path)
                    )
                except Exception as e:
                    messagebox.showerror(self.t("common.error"), self.t("export.save_error", err=str(e)))
                    print(f"Chyba při exportu: {e}")
        except Exception as e:
            messagebox.showerror(self.t("common.error"), self.t("unexpected.error", err=str(e)))
            print(f"Chyba při exportu: {e}")

    def clear_files(self):
        """Vymaže všechny TXT soubory z import složky."""
        try:
            files = [os.path.join(self.import_dir, f) for f in os.listdir(self.import_dir) 
                    if f.lower().endswith('.txt')]
            
            if not files:
                messagebox.showinfo(self.t("common.info"), self.t("delete.no_files"))
                return
            
            if messagebox.askyesno(self.t("confirm.delete.title"), self.t("confirm.delete.message")):
                for file_path in files:
                    try:
                        os.remove(file_path)
                    except Exception as e:
                        messagebox.showerror(
                            self.t("delete.file_error.title"),
                            self.t("delete.file_error.message", file=os.path.basename(file_path), err=str(e))
                        )
                
                # Vyčistíme tabulku
                for item in self.file_table.get_children():
                    self.file_table.delete(item)
                
                self.export_btn.configure(state='disabled')
                messagebox.showinfo(self.t("common.done"), self.t("delete.success"))
        except Exception as e:
            print(f"Chyba při mazání souborů: {e}")
            messagebox.showerror(self.t("common.error"), self.t("unexpected.error", err=str(e)))

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
            messagebox.showerror(
                self.t("parse.error.title"),
                self.t("parse.error.message", file=os.path.basename(file_path), err=str(e))
            )
            return None

    def t(self, key: str, **kwargs) -> str:
        """Překladový helper."""
        lang_map = self.translations.get(self.language, {})
        value = lang_map.get(key) or self.translations.get("cs", {}).get(key) or key
        try:
            return value.format(**kwargs)
        except Exception:
            return value

    def load_translations(self) -> dict:
        """Načte překlady z locales/translations.json."""
        try:
            if os.path.exists(self.translations_path):
                with open(self.translations_path, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception as e:
            print(f"Chyba při načítání překladů: {e}")
        return {}

    def load_settings(self) -> dict:
        try:
            if os.path.exists(self.settings_path):
                with open(self.settings_path, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception as e:
            print(f"Chyba při načítání nastavení: {e}")
        return {}

    def save_settings(self) -> None:
        try:
            with open(self.settings_path, "w", encoding="utf-8") as f:
                json.dump({"language": self.language}, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Chyba při ukládání nastavení: {e}")

    def set_language(self, lang_code: str) -> None:
        if lang_code not in ("cs", "en", "de"):
            return
        if getattr(self, "language", "cs") == lang_code:
            return
        self.language = lang_code
        self.save_settings()
        self.window.title(self.t("app.window_title"))
        self.rebuild_ui()

    def rebuild_ui(self) -> None:
        """Znovu postaví celé UI (spolehlivé přepnutí jazyků)."""
        for child in self.window.winfo_children():
            child.destroy()
        self.create_widgets()
        self.refresh_file_table()

    def get_txt_files_in_import_dir(self):
        if not os.path.exists(self.import_dir):
            return []
        all_files = os.listdir(self.import_dir)
        files = []
        for name in all_files:
            if name.lower().endswith('.txt'):
                full_path = os.path.join(self.import_dir, name)
                if os.path.isfile(full_path):
                    files.append(full_path)
        return sorted(files)

    def refresh_file_table(self) -> None:
        """Obnoví tabulku podle aktuálního obsahu import/ adresáře."""
        if not hasattr(self, "file_table"):
            return

        for item in self.file_table.get_children():
            self.file_table.delete(item)

        files = self.get_txt_files_in_import_dir()
        for idx, file_path in enumerate(files):
            try:
                data = self.parse_txt_file(file_path)
                if not data:
                    continue

                file_name = os.path.basename(file_path)
                ra_value = self.find_value_in_data(data, "Ra")
                rz_value = self.find_value_in_data(data, "Rz")
                date_value = self.find_value_in_data(data, "Date")

                tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
                self.file_table.insert("", "end", values=(file_name, date_value, ra_value, rz_value), tags=(tag,))
            except Exception as e:
                print(f"Chyba při obnově tabulky ({file_path}): {e}")

        if files:
            self.export_btn.configure(state='normal')
        else:
            self.export_btn.configure(state='disabled')

    def load_flag_images(self) -> dict:
        """Načte vlajky do CTkImage. Při chybě vrátí prázdný dict a UI použije text."""
        images = {}
        if Image is None:
            return images

        for lang, candidates in getattr(self, "flag_files", {}).items():
            try:
                path = None
                for p in candidates:
                    if os.path.exists(p):
                        path = p
                        break
                if not path:
                    continue

                pil_img = Image.open(path)
                images[lang] = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(28, 18))
            except Exception as e:
                print(f"Nelze načíst vlajku pro {lang}: {e}")
        return images

    def create_language_switcher(self, parent):
        """Vytvoří přepínač jazyků pomocí vlajek."""
        switcher = ctk.CTkFrame(parent, fg_color="transparent")
        switcher.pack(anchor='e')

        self.language_buttons = {}

        def add_btn(lang_code: str, fallback_text: str):
            img = getattr(self, "flag_images", {}).get(lang_code)
            selected = self.language == lang_code
            btn = ctk.CTkButton(
                switcher,
                text="" if img else fallback_text,
                image=img,
                width=40,
                height=26,
                corner_radius=8,
                fg_color="transparent",
                hover_color=getattr(self, "colors", {}).get('surface', '#e5e7eb'),
                border_width=2 if selected else 0,
                border_color=getattr(self, "colors", {}).get('primary', '#2563eb'),
                command=lambda: self.set_language(lang_code)
            )
            btn.pack(side='left', padx=4)
            self.language_buttons[lang_code] = btn

        add_btn("cs", "CZ")
        add_btn("en", "EN")
        add_btn("de", "DE")

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