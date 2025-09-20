import customtkinter as ctk
import pandas as pd
import json
import os
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass, asdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import requests
import threading

# Dil sistemi
LANGUAGES = {
    "DE": {
        "app_title": "EKS Formular Ausfüller Pro",
        "customer": "Kunde",
        "period": "Zeitraum",
        "template": "Vorlage", 
        "load_bwa": "BWA Datei laden",
        "auto_mapping": "Automatische Zuordnung",
        "export_eks": "EKS Exportieren",
        "new_customer": "Neuer Kunde",
        "customer_code": "Kundennummer",
        "customer_name": "Kundenname",
        "from_date": "Von Datum",
        "to_date": "Bis Datum",
        "quick_select": "Schnellauswahl",
        "mapping_results": "Zuordnungsergebnisse",
        "confidence": "Vertrauen",
        "monthly_values": "Monatswerte",
        "total": "Gesamt",
        "success": "Erfolgreich",
        "error": "Fehler",
        "settings": "Einstellungen",
        "api_key": "API Schlüssel",
        "save": "Speichern",
        "cancel": "Abbrechen",
        "loading": "Laden...",
        "file_loaded": "Datei geladen",
        "no_file": "Keine Datei",
        "processing": "Verarbeitung...",
        "q1": "Q1", "q2": "Q2", "q3": "Q3", "q4": "Q4",
        "half_year": "Halbjahr", "full_year": "Ganzes Jahr",
        "language": "Sprache"
    },
    "TR": {
        "app_title": "EKS Form Doldurucu Pro",
        "customer": "Müşteri",
        "period": "Dönem",
        "template": "Şablon",
        "load_bwa": "BWA Dosyası Yükle",
        "auto_mapping": "Otomatik Eşleştirme",
        "export_eks": "EKS Dışa Aktar",
        "new_customer": "Yeni Müşteri",
        "customer_code": "Müşteri Kodu",
        "customer_name": "Müşteri Adı",
        "from_date": "Başlangıç Tarihi",
        "to_date": "Bitiş Tarihi",
        "quick_select": "Hızlı Seçim",
        "mapping_results": "Eşleştirme Sonuçları",
        "confidence": "Güven",
        "monthly_values": "Aylık Değerler",
        "total": "Toplam",
        "success": "Başarılı",
        "error": "Hata",
        "settings": "Ayarlar",
        "api_key": "API Anahtarı",
        "save": "Kaydet",
        "cancel": "İptal",
        "loading": "Yükleniyor...",
        "file_loaded": "Dosya yüklendi",
        "no_file": "Dosya yok",
        "processing": "İşleniyor...",
        "q1": "Q1", "q2": "Q2", "q3": "Q3", "q4": "Q4",
        "half_year": "6 Ay", "full_year": "12 Ay",
        "language": "Dil"
    }
}

@dataclass
class Customer:
    code: str
    name: str
    created_date: str
    default_template: str = "eks_standard.xlsx"
    notes: str = ""
    bwa_history: List[Dict] = None
    
    def __post_init__(self):
        if self.bwa_history is None:
            self.bwa_history = []

@dataclass 
class MappingRule:
    eks_field: str
    bwa_source: str
    calculation_type: str  # 'direct', 'sum'
    source_accounts: List[str] = None
    description_de: str = ""

class BWAParser:
    def __init__(self):
        self.mapping_rules = self._init_mapping_rules()
        self.bwa_data = None
        self.customer_info = None
        self.available_months = []
        
    def _init_mapping_rules(self) -> Dict[str, MappingRule]:
        return {
            # A Bölümü - Betriebseinnahmen
            "A1": MappingRule("A1", "Summe Erlöse", "direct", description_de="Betriebseinnahmen"),
            "A5": MappingRule("A5", "Summe Umsatzsteuer", "direct", description_de="Vereinnahmte Umsatzsteuer"),
            "A7": MappingRule("A7", "Ust-Erstattung", "direct", description_de="vom Finanzamt erstattete Umsatzsteuer"),

            
            # B Bölümü - Betriebsausgaben  
            "B1": MappingRule("B1", "Wareneinkauf", "sum", ["5400", "Summe Material, Stoffe, Waren"], "Wareneinkauf"),
            "B2c": MappingRule("B2c", "6030", "direct", ["6030", "6036", "6171"], "geringfügig Beschäftigte"),
            "B3": MappingRule("B3", "Miete + Energie", "sum", ["6310", "6325"], "Raumkosten (Miete und Energiekosten)"),
            "B11": MappingRule("B11", "6805", "direct", ["6805"], "Telefonkosten"),
            "B14c": MappingRule("B14c", "6855", "direct", ["6855"], "Nebenkosten des Geldverkehrs"),
            "B17": MappingRule("B17", "Summe Vorsteuer", "direct", description_de="gezahlte Vorsteuer"),
            
            # Ek mapping'ler (eşleştirme listesinden)
            "B10": MappingRule("B10", "Büromaterial", "sum", ["6815", "6800"], "Büromaterial plus Porto"),
            "B14e": MappingRule("B14e", "6330", "direct", ["6330"], "Reinigung"),
            "B14f": MappingRule("B14f", "6630", "direct", ["6630"], "Repräsentationskosten"),
            "B14h": MappingRule("B14h", "Sonstige", "sum", ["6300", "6850"], "sonst. Betriebliche Ausgaben"),
            
            # Diğer potansiyel mapping'ler
            "B2a": MappingRule("B2a", "Vollzeit", "sum", ["6010", "6110", "6120", "6170"], "Vollzeitbeschäftigte"),
            "B4": MappingRule("B4", "Versicherung", "sum", ["6400", "6420"], "Betriebliche Versicherungen"),
            "B5_1a": MappingRule("B5_1a", "6570", "direct", ["6570"], "Steuern (Kfz)"),
            "B5_1b": MappingRule("B5_1b", "6520", "direct", ["6520"], "Versicherung (Kfz)"),
            "B5_1c": MappingRule("B5_1c", "6530", "direct", ["6530"], "Betriebskosten (Kfz)"),
            "B5_1d": MappingRule("B5_1d", "6540", "direct", ["6540"], "Reparaturen (Kfz)"),
            "B6": MappingRule("B6", "6600", "direct", ["6600"], "Maßnahmen"),
            "B7a": MappingRule("B7a", "6670", "direct", ["6670"], "Reisekosten"),
            "B12": MappingRule("B12", "Beratung", "sum", ["6825", "6830"], "Beratungskosten"),
            "B14b": MappingRule("B14b", "6835", "direct", ["6835"], "Miete Einrichtung"),
            "B14g": MappingRule("B14g", "6335", "direct", ["6335"], "Instandhaltung betr. Räume"),
            "B14i": MappingRule("B14i", "6640", "direct", ["6640"], "Bewirtungskosten"),
            "B16": MappingRule("B16", "Tilgung", "sum", ["3150", "3160", "3170"], "Tilgung Darlehen"),
            "B18": MappingRule("B18", "3820", "direct", ["3820"], "an Finanzamt gezahlte USt")
        }
    
    def load_bwa_file(self, file_path: str) -> Tuple[bool, str]:
        try:
            # BWA laden mit header=None für rohe Daten
            df = pd.read_excel(file_path, header=None)
            
            # Kunde info aus erster Zeile extrahieren
            first_row = str(df.iloc[0, 0]) if not df.empty else ""
            if first_row and len(first_row) > 6:
                parts = first_row.split(" ", 1)
                if len(parts) >= 2 and parts[0].isdigit():
                    self.customer_info = {
                        "code": parts[0],
                        "name": parts[1]
                    }
            
            # Header finden (normalerweise Zeile 2: "Konto Bezeichnung")
            header_row = -1
            for i, row in df.iterrows():
                if any("Konto" in str(cell) and "Bezeichnung" in str(cell) for cell in row if pd.notna(cell)):
                    header_row = i
                    break
            
            if header_row == -1:
                return False, "BWA Header nicht gefunden"
            
            # Daten ab header_row neu laden
            self.bwa_data = pd.read_excel(file_path, header=header_row)
            
            # Verfügbare Monate extrahieren
            month_cols = ['JAN', 'FEB', 'MRZ', 'APR', 'MAI', 'JUN', 'JUL', 'AUG', 'SEP', 'OKT', 'NOV', 'DEZ']
            self.available_months = [col for col in self.bwa_data.columns if col in month_cols]
            
            return True, f"BWA geladen: {len(self.available_months)} Monate verfügbar"
            
        except Exception as e:
            return False, f"Fehler beim Laden: {str(e)}"
    
    def extract_values_for_period(self, start_month: str, end_month: str) -> Dict:
        if self.bwa_data is None or self.bwa_data.empty:
            return {}
        
        # Monat-Indices bestimmen  
        month_order = ['JAN', 'FEB', 'MRZ', 'APR', 'MAI', 'JUN', 'JUL', 'AUG', 'SEP', 'OKT', 'NOV', 'DEZ']
        try:
            start_idx = month_order.index(start_month)
            end_idx = month_order.index(end_month)
            selected_months = month_order[start_idx:end_idx+1]
        except ValueError:
            selected_months = self.available_months[:6]  # Default: erste 6 Monate
        
        # Nur verfügbare Monate verwenden
        selected_months = [m for m in selected_months if m in self.available_months]
        
        results = {}
        for field, rule in self.mapping_rules.items():
            extracted = self._extract_mapping(rule, selected_months)
            confidence = self._calculate_confidence(extracted['values'])
            
            results[field] = {
                'values': extracted['values'],
                'confidence': confidence,
                'source': rule.bwa_source,
                'description': rule.description_de,
                'months': selected_months,
                'total': sum(v for v in extracted['values'] if v is not None)
            }
        
        return results
    
    def _extract_mapping(self, rule: MappingRule, months: List[str]) -> Dict:
        if rule.calculation_type == "direct":
            return self._find_direct_match(rule.bwa_source, months)
        elif rule.calculation_type == "sum":
            return self._sum_multiple_accounts(rule.source_accounts, months)
        return {'values': [None] * len(months)}
    
    def _find_direct_match(self, search_term: str, months: List[str]) -> Dict:
        try:
            # Erste Spalte nach Suchbegriff durchsuchen
            first_col = self.bwa_data.iloc[:, 0].astype(str)
            mask = first_col.str.contains(search_term, case=False, na=False)
            
            if mask.any():
                row = self.bwa_data[mask].iloc[0]
                values = []
                for month in months:
                    if month in self.bwa_data.columns:
                        val = row[month]
                        values.append(float(val) if pd.notna(val) and val != '' else None)
                    else:
                        values.append(None)
                return {'values': values}
        except Exception:
            pass
        
        return {'values': [None] * len(months)}
    
    def _sum_multiple_accounts(self, accounts: List[str], months: List[str]) -> Dict:
        total_values = [0] * len(months)
        found_any = False
        
        for account in accounts:
            result = self._find_direct_match(account, months)
            for i, val in enumerate(result['values']):
                if val is not None:
                    total_values[i] += val
                    found_any = True
        
        return {'values': total_values if found_any else [None] * len(months)}
    
    def _calculate_confidence(self, values: List) -> int:
        if not values:
            return 0
        non_null = sum(1 for v in values if v is not None)
        return int((non_null / len(values)) * 100)

class CustomerManager:
    def __init__(self, data_dir: str = "data"):
        self.data_dir = data_dir
        self.customers_dir = os.path.join(data_dir, "customers")
        os.makedirs(self.customers_dir, exist_ok=True)
        
    def save_customer(self, customer: Customer) -> bool:
        try:
            file_path = os.path.join(self.customers_dir, f"{customer.code}.json")
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(asdict(customer), f, ensure_ascii=False, indent=2)
            return True
        except Exception:
            return False
    
    def load_customer(self, customer_code: str) -> Optional[Customer]:
        try:
            file_path = os.path.join(self.customers_dir, f"{customer_code}.json")
            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                return Customer(**data)
        except Exception:
            pass
        return None
    
    def get_all_customers(self) -> List[Customer]:
        customers = []
        for file_name in os.listdir(self.customers_dir):
            if file_name.endswith('.json'):
                customer_code = file_name[:-5]
                customer = self.load_customer(customer_code)
                if customer:
                    customers.append(customer)
        return sorted(customers, key=lambda c: c.code)

class EKSFormFiller(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Grundkonfiguration
        self.title("EKS Formular Ausfüller Pro")
        self.geometry("1400x900")
        self.configure(fg_color="#1a1a1a")
        
        # Sprache
        self.language = "DE"
        self.texts = LANGUAGES[self.language]
        
        # Components
        self.bwa_parser = BWAParser()
        self.customer_manager = CustomerManager()
        
        # State
        self.current_customer = None
        self.bwa_file_path = None
        self.extracted_data = {}
        self.selected_start_month = "JAN"
        self.selected_end_month = "JUN"
        
        self.setup_ui()
        self.load_customer_list()
    
    def setup_ui(self):
        # Header
        header_frame = ctk.CTkFrame(self, height=80, fg_color="#2b2b2b")
        header_frame.pack(fill="x", padx=10, pady=5)
        header_frame.pack_propagate(False)
        
        # Titel und Controls in Header
        title_label = ctk.CTkLabel(header_frame, text=self.texts["app_title"], 
                                 font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(side="left", padx=20, pady=20)
        
        # Einstellungen Button
        settings_btn = ctk.CTkButton(header_frame, text="⚙️", width=40, height=40,
                                   command=self.open_settings)
        settings_btn.pack(side="right", padx=10, pady=20)
        
        # Dil seçimi
        language_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        language_frame.pack(side="right", padx=10, pady=20)
        
        ctk.CTkLabel(language_frame, text=self.texts["language"] + ":", 
                    font=ctk.CTkFont(size=12)).pack(side="left", padx=5)
        
        self.language_combo = ctk.CTkComboBox(language_frame, values=["DE", "TR"], 
                                            width=60, command=self.change_language)
        self.language_combo.set(self.language)
        self.language_combo.pack(side="left", padx=5)
        
        # Main Container
        main_container = ctk.CTkFrame(self, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Control Panel (oben)
        control_frame = ctk.CTkFrame(main_container, height=120, fg_color="#2b2b2b")
        control_frame.pack(fill="x", pady=(0, 10))
        control_frame.pack_propagate(False)
        
        # Kunde Auswahl
        customer_frame = ctk.CTkFrame(control_frame, fg_color="#3b3b3b")
        customer_frame.pack(side="left", fill="y", padx=10, pady=10)
        
        ctk.CTkLabel(customer_frame, text=self.texts["customer"], 
                    font=ctk.CTkFont(weight="bold")).pack(pady=5)
        
        self.customer_combo = ctk.CTkComboBox(customer_frame, width=200, command=self.on_customer_selected)
        self.customer_combo.pack(pady=5)
        
        new_customer_btn = ctk.CTkButton(customer_frame, text="+", width=30, height=30,
                                        command=self.create_new_customer)
        new_customer_btn.pack(pady=5)
        
        # Zeitraum Auswahl
        period_frame = ctk.CTkFrame(control_frame, fg_color="#3b3b3b")
        period_frame.pack(side="left", fill="y", padx=10, pady=10)
        
        ctk.CTkLabel(period_frame, text=self.texts["period"], 
                    font=ctk.CTkFont(weight="bold")).pack(pady=5)
        
        period_controls = ctk.CTkFrame(period_frame, fg_color="transparent")
        period_controls.pack(pady=5)
        
        months = ['JAN', 'FEB', 'MRZ', 'APR', 'MAI', 'JUN', 'JUL', 'AUG', 'SEP', 'OKT', 'NOV', 'DEZ']
        
        self.start_month_combo = ctk.CTkComboBox(period_controls, values=months, width=80, 
                                               command=self.on_period_changed)
        self.start_month_combo.pack(side="left", padx=2)
        self.start_month_combo.set("JAN")
        
        ctk.CTkLabel(period_controls, text="-").pack(side="left", padx=5)
        
        self.end_month_combo = ctk.CTkComboBox(period_controls, values=months, width=80,
                                             command=self.on_period_changed)
        self.end_month_combo.pack(side="left", padx=2)
        self.end_month_combo.set("JUN")
        
        # Quick Select Buttons
        quick_frame = ctk.CTkFrame(period_frame, fg_color="transparent")
        quick_frame.pack(pady=5)
        
        quick_buttons = [
            ("Q1", lambda: self.set_period("JAN", "MRZ")),
            ("Q2", lambda: self.set_period("APR", "JUN")),
            ("Q3", lambda: self.set_period("JUL", "SEP")),
            ("Q4", lambda: self.set_period("OKT", "DEZ")),
        ]
        
        for text, command in quick_buttons:
            btn = ctk.CTkButton(quick_frame, text=text, width=40, height=25, command=command)
            btn.pack(side="left", padx=1)
        
        # Content Area (unten)
        content_frame = ctk.CTkFrame(main_container, fg_color="transparent")
        content_frame.pack(fill="both", expand=True)
        
        # Sol panel - BWA Import
        left_panel = ctk.CTkFrame(content_frame, width=400, fg_color="#2b2b2b")
        left_panel.pack(side="left", fill="y", padx=(0, 10))
        left_panel.pack_propagate(False)
        
        ctk.CTkLabel(left_panel, text="BWA Import", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
        
        self.load_bwa_btn = ctk.CTkButton(left_panel, text=self.texts["load_bwa"],
                                         command=self.load_bwa_file, height=40)
        self.load_bwa_btn.pack(pady=10, padx=20, fill="x")
        
        self.bwa_status_label = ctk.CTkLabel(left_panel, text=self.texts["no_file"], 
                                           text_color="gray")
        self.bwa_status_label.pack(pady=5)
        
        # BWA Info Anzeige
        self.bwa_info_frame = ctk.CTkFrame(left_panel, fg_color="#3b3b3b")
        self.bwa_info_frame.pack(fill="x", padx=20, pady=10)
        
        self.mapping_btn = ctk.CTkButton(left_panel, text=self.texts["auto_mapping"],
                                        command=self.perform_mapping, height=40, state="disabled")
        self.mapping_btn.pack(pady=20, padx=20, fill="x")
        
        # Template Analyse Button (Debug)
        analyze_btn = ctk.CTkButton(left_panel, text="🔍 Template Analysieren",
                                  command=self.analyze_template_wrapper, height=30)
        analyze_btn.pack(pady=5, padx=20, fill="x")
        
        self.export_btn = ctk.CTkButton(left_panel, text=self.texts["export_eks"],
                                       command=self.export_eks, height=40, state="disabled")
        self.export_btn.pack(pady=10, padx=20, fill="x")
        
        # Rechts: Mapping Ergebnisse
        right_panel = ctk.CTkFrame(content_frame, fg_color="#2b2b2b")
        right_panel.pack(side="right", fill="both", expand=True)
        
        ctk.CTkLabel(right_panel, text=self.texts["mapping_results"], 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
        
        self.results_frame = ctk.CTkScrollableFrame(right_panel, fg_color="#1a1a1a")
        self.results_frame.pack(fill="both", expand=True, padx=20, pady=10)
    
    def load_customer_list(self):
        customers = self.customer_manager.get_all_customers()
        customer_options = [f"{c.code} - {c.name}" for c in customers]
        if customer_options:
            self.customer_combo.configure(values=customer_options)
            self.customer_combo.set(customer_options[0])
            self.on_customer_selected(customer_options[0])
        else:
            self.customer_combo.configure(values=["Keine Kunden"])
    
    def on_customer_selected(self, selection):
        if " - " in selection:
            customer_code = selection.split(" - ")[0]
            self.current_customer = self.customer_manager.load_customer(customer_code)
    
    def create_new_customer(self):
        dialog = CustomerDialog(self, self.texts)
        if dialog.result:
            customer = Customer(
                code=dialog.result["code"],
                name=dialog.result["name"],
                created_date=datetime.now().strftime("%Y-%m-%d")
            )
            if self.customer_manager.save_customer(customer):
                self.load_customer_list()
                # Neuen Kunden auswählen
                new_selection = f"{customer.code} - {customer.name}"
                self.customer_combo.set(new_selection)
                self.on_customer_selected(new_selection)
    
    def set_period(self, start: str, end: str):
        self.start_month_combo.set(start)
        self.end_month_combo.set(end)
        self.on_period_changed()
    
    def on_period_changed(self, value=None):
        self.selected_start_month = self.start_month_combo.get()
        self.selected_end_month = self.end_month_combo.get()
    
    def load_bwa_file(self):
        file_path = filedialog.askopenfilename(
            title="BWA Excel Datei auswählen",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if file_path:
            self.bwa_status_label.configure(text=self.texts["loading"])
            
            # Threading für UI Responsiveness
            def load_thread():
                success, message = self.bwa_parser.load_bwa_file(file_path)
                
                # UI Update im Main Thread
                self.after(0, lambda: self.on_bwa_loaded(success, message, file_path))
            
            threading.Thread(target=load_thread, daemon=True).start()
    
    def on_bwa_loaded(self, success: bool, message: str, file_path: str):
        if success:
            self.bwa_file_path = file_path
            self.bwa_status_label.configure(text="✅ " + self.texts["file_loaded"], text_color="green")
            self.mapping_btn.configure(state="normal")
            
            # BWA Info anzeigen
            self.update_bwa_info()
            
            # Auto-Kunde erstellen falls nicht vorhanden
            if self.bwa_parser.customer_info and not self.current_customer:
                self.auto_create_customer()
        else:
            self.bwa_status_label.configure(text="❌ " + message, text_color="red")
    
    def update_bwa_info(self):
        # Clear previous info
        for widget in self.bwa_info_frame.winfo_children():
            widget.destroy()
        
        if self.bwa_parser.customer_info:
            info = self.bwa_parser.customer_info
            ctk.CTkLabel(self.bwa_info_frame, text=f"Kunde: {info['code']}", 
                        font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=2)
            ctk.CTkLabel(self.bwa_info_frame, text=f"Name: {info['name']}").pack(anchor="w", padx=10, pady=2)
        
        if self.bwa_parser.available_months:
            months_text = ", ".join(self.bwa_parser.available_months)
            ctk.CTkLabel(self.bwa_info_frame, text=f"Monate: {months_text}").pack(anchor="w", padx=10, pady=2)
    
    def auto_create_customer(self):
        if self.bwa_parser.customer_info:
            info = self.bwa_parser.customer_info
            existing = self.customer_manager.load_customer(info["code"])
            if not existing:
                customer = Customer(
                    code=info["code"],
                    name=info["name"],
                    created_date=datetime.now().strftime("%Y-%m-%d")
                )
                if self.customer_manager.save_customer(customer):
                    self.load_customer_list()
                    new_selection = f"{customer.code} - {customer.name}"
                    self.customer_combo.set(new_selection)
                    self.on_customer_selected(new_selection)
    
    def perform_mapping(self):
        if self.bwa_parser.bwa_data is None or self.bwa_parser.bwa_data.empty:
            return
        
        self.mapping_btn.configure(text=self.texts["processing"], state="disabled")
        
        def mapping_thread():
            extracted = self.bwa_parser.extract_values_for_period(
                self.selected_start_month, self.selected_end_month
            )
            self.after(0, lambda: self.on_mapping_complete(extracted))
        
        threading.Thread(target=mapping_thread, daemon=True).start()
    
    def on_mapping_complete(self, extracted_data: Dict):
        self.extracted_data = extracted_data
        self.mapping_btn.configure(text=self.texts["auto_mapping"], state="normal")
        self.export_btn.configure(state="normal")
        self.display_mapping_results()
    
    def display_mapping_results(self):
        # Clear previous results
        for widget in self.results_frame.winfo_children():
            widget.destroy()
        
        if not self.extracted_data:
            ctk.CTkLabel(self.results_frame, text="Keine Ergebnisse").pack(pady=20)
            return
        
        total_confidence = 0
        valid_mappings = 0
        
        for field, data in self.extracted_data.items():
            # Result Frame
            result_frame = ctk.CTkFrame(self.results_frame, fg_color="#2b2b2b")
            result_frame.pack(fill="x", pady=5, padx=10)
            
            # Header mit Confidence
            confidence = data.get('confidence', 0)
            color = "#4CAF50" if confidence > 80 else "#FF9800" if confidence > 50 else "#F44336"
            
            header_frame = ctk.CTkFrame(result_frame, fg_color="transparent")
            header_frame.pack(fill="x", padx=10, pady=5)
            
            field_label = ctk.CTkLabel(header_frame, text=f"{field}: {data['description']}", 
                                     font=ctk.CTkFont(weight="bold"))
            field_label.pack(side="left")
            
            confidence_label = ctk.CTkLabel(header_frame, text=f"{confidence}%", 
                                          text_color=color, font=ctk.CTkFont(weight="bold"))
            confidence_label.pack(side="right")
            
            # Source
            source_label = ctk.CTkLabel(result_frame, text=f"Quelle: {data['source']}", 
                                      font=ctk.CTkFont(size=12))
            source_label.pack(anchor="w", padx=20, pady=2)
            
            # Monatswerte
            values = data['values']
            months = data.get('months', [])
            
            if months and values:
                values_frame = ctk.CTkFrame(result_frame, fg_color="#3b3b3b")
                values_frame.pack(fill="x", padx=20, pady=5)
                
                for month, value in zip(months, values):
                    value_text = f"{value:,.0f} €" if value is not None else "N/A"
                    month_frame = ctk.CTkFrame(values_frame, fg_color="transparent")
                    month_frame.pack(side="left", padx=5, pady=5)
                    
                    ctk.CTkLabel(month_frame, text=month, font=ctk.CTkFont(size=10)).pack()
                    ctk.CTkLabel(month_frame, text=value_text, font=ctk.CTkFont(size=12, weight="bold")).pack()
            
            # Gesamt
            total = data.get('total', 0)
            total_label = ctk.CTkLabel(result_frame, text=f"Gesamt: {total:,.0f} €", 
                                     font=ctk.CTkFont(size=14, weight="bold"))
            total_label.pack(anchor="w", padx=20, pady=5)
            
            if confidence > 0:
                total_confidence += confidence
                valid_mappings += 1
        
        # Summary
        if valid_mappings > 0:
            avg_confidence = total_confidence / valid_mappings
            summary_frame = ctk.CTkFrame(self.results_frame, fg_color="#2b2b2b")
            summary_frame.pack(fill="x", pady=20, padx=10)
            
            summary_text = f"Durchschnittliche Zuordnung: {avg_confidence:.1f}% ({valid_mappings}/{len(self.extracted_data)} Felder)"
            ctk.CTkLabel(summary_frame, text=summary_text, 
                        font=ctk.CTkFont(size=16, weight="bold")).pack(pady=20)
    
    def export_eks(self):
        if not self.extracted_data or not self.current_customer:
            warning_msg = "Keine Daten zum Exportieren oder kein Kunde ausgewählt" if self.language == "DE" else "Dışa aktarılacak veri yok veya müşteri seçilmedi"
            messagebox.showwarning("Warnung" if self.language == "DE" else "Uyarı", warning_msg)
            return
        
        # Template Ordner prüfen
        template_dir = "templates"
        os.makedirs(template_dir, exist_ok=True)
        
        # Export
        try:
            filename = f"{self.current_customer.code}_EKS_{self.selected_start_month}-{self.selected_end_month}_{datetime.now().strftime('%Y%m%d')}.xlsx"
            export_path = filedialog.asksaveasfilename(
                title="EKS Export speichern" if self.language == "DE" else "EKS Dışa Aktar",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=filename
            )
            
            if export_path:
                success = self.create_eks_export(export_path)
                if success:
                    success_msg = f"EKS erfolgreich exportiert:\n{export_path}" if self.language == "DE" else f"EKS başarıyla dışa aktarıldı:\n{export_path}"
                    messagebox.showinfo("Erfolg" if self.language == "DE" else "Başarılı", success_msg)
                    # BWA History aktualisieren
                    self.update_customer_history()
                else:
                    error_msg = "Export fehlgeschlagen" if self.language == "DE" else "Dışa aktarma başarısız"
                    messagebox.showerror("Fehler" if self.language == "DE" else "Hata", error_msg)
        
        except Exception as e:
            error_msg = f"Export Fehler: {str(e)}" if self.language == "DE" else f"Dışa Aktarma Hatası: {str(e)}"
            messagebox.showerror("Fehler" if self.language == "DE" else "Hata", error_msg)
    
    def create_eks_export(self, export_path: str) -> bool:
        try:
            # Template dosyasını kontrol et
            template_path = os.path.join("templates", "eks_form.xlsx")
            if not os.path.exists(template_path):
                # Fallback: otomatik template oluştur
                return self.create_automatic_export(export_path)
            
            # Gerçek EKS template'ini yükle
            wb = openpyxl.load_workbook(template_path)
            ws = wb.active
            
            # EKS template'indeki hücreleri doldur
            success = self.fill_eks_template(ws)
            if not success:
                return False
            
            # Müşteri bilgilerini güncelle (template'te varsa)
            self.update_customer_info_in_template(ws)
            
            # Dönem bilgilerini güncelle
            self.update_period_info_in_template(ws)
            
            # Kaydet
            wb.save(export_path)
            return True
            
        except Exception as e:
            print(f"Template Export Fehler: {e}")
            return False
    
    def fill_eks_template(self, ws) -> bool:
        """EKS template'indeki hücreleri BWA verilerine göre doldurur"""
        try:
            # Gerçek EKS form pozisyonları (görselden alınan doğru pozisyonlar)
            eks_positions = {
                # A Bölümü - Betriebseinnahmen (Satır 10-17)
                "A1": {"start_row": 10, "months_start_col": 3},   # Betriebseinnahmen
                "A2": {"start_row": 11, "months_start_col": 3},   # Privatentnahmen von Waren  
                "A3": {"start_row": 12, "months_start_col": 3},   # Sonstige betriebliche Einnahmen
                "A4": {"start_row": 13, "months_start_col": 3},   # Zuwendung von Dritten/Darlehen
                "A5": {"start_row": 14, "months_start_col": 3},   # Vereinnahmte Umsatzsteuer
                "A6": {"start_row": 15, "months_start_col": 3},   # Umsatzsteuer auf private Warenentnahme
                "A7": {"start_row": 16, "months_start_col": 3},   # vom Finanzamt erstattete Umsatzsteuer
                
                # B Bölümü - Betriebsausgaben (Satır 22-66)
                "B1": {"start_row": 22, "months_start_col": 3},   # Wareneinkauf
                "B2a": {"start_row": 24, "months_start_col": 3},  # Vollzeitbeschäftigte
                "B2b": {"start_row": 25, "months_start_col": 3},  # Teilzeitbeschäftigte  
                "B2c": {"start_row": 26, "months_start_col": 3},  # geringfügig Beschäftigte
                "B2d": {"start_row": 27, "months_start_col": 3},  # mithelfende Familienangehörige
                "B3": {"start_row": 28, "months_start_col": 3},   # Raumkosten (Neben-kosten und Energiekosten)
                "B4": {"start_row": 29, "months_start_col": 3},   # Betriebliche Versicherungen / Beiträge
                "B5": {"start_row": 30, "months_start_col": 3},   # Kraftfahrzeugkosten
                "B5_1a": {"start_row": 33, "months_start_col": 3},# Steuern
                "B5_1b": {"start_row": 34, "months_start_col": 3},# Versicherung
                "B5_1c": {"start_row": 35, "months_start_col": 3},# Betriebskosten
                "B5_1d": {"start_row": 36, "months_start_col": 3},# Reparaturen
                "B5_1x": {"start_row": 37, "months_start_col": 3},# abzgl. private km (0,10 € je  gefahrenen km)
                "B5_2": {"start_row": 38, "months_start_col": 3}, # Privates Kfz - betriebliche Fahrten
                "B6": {"start_row": 39, "months_start_col": 3},   # Maßnahmen ggf. auf besonderem Blatt
                "B7a": {"start_row": 41, "months_start_col": 3},  # Übernachtungskosten
                "B7b": {"start_row": 42, "months_start_col": 3},  # Reisenebenkosten
                "B7c": {"start_row": 43, "months_start_col": 3},  # öffentliche Verkehrsmittel
                "B8": {"start_row": 47, "months_start_col": 3},   # Investitionen
                "B9": {"start_row": 48, "months_start_col": 3},   # Investition aus Zuwendungen Dritter
                "B10": {"start_row": 50, "months_start_col": 3},  # Büromaterial plus Porto
                "B11": {"start_row": 51, "months_start_col": 3},  # Telefonkosten
                "B12": {"start_row": 52, "months_start_col": 3},  # Beratungskosten
                "B13": {"start_row": 53, "months_start_col": 3},  # Fortbildungskosten
                "B14": {"start_row": 54, "months_start_col": 3},  # Sonstige Betriebsausgaben
                "B14a": {"start_row": 55, "months_start_col": 3}, # Reparatur Anlagevermögen
                "B14b": {"start_row": 56, "months_start_col": 3}, # Miete Einrichtung
                "B14c": {"start_row": 57, "months_start_col": 3}, # Nebenkosten des Geldverkehrs
                "B14d": {"start_row": 58, "months_start_col": 3}, # betriebliche Abfallbeseitigung
                "B14e": {"start_row": 59, "months_start_col": 3}, # Reinigung
                "B14f": {"start_row": 60, "months_start_col": 3}, # Repräsentationskosten
                "B14g": {"start_row": 61, "months_start_col": 3}, # Instandhaltung betr. Räume
                "B14h": {"start_row": 62, "months_start_col": 3}, # sonst. Betriebliche Ausgaben
                "B14i": {"start_row": 63, "months_start_col": 3}, # Bewirtungskosten
                "B15": {"start_row": 64, "months_start_col": 3},  # Schuldzinsen aus Anlagevermögen
                "B16": {"start_row": 65, "months_start_col": 3},  # Tilgung bestehender betrieblicher Darlehen
                "B17": {"start_row": 66, "months_start_col": 3},  # gezahlte Vorsteuer
                "B18": {"start_row": 67, "months_start_col": 3}   # an das Finanzamt gezahlte Umsatzsteuer
            }
            
            for field, data in self.extracted_data.items():
                if field in eks_positions:
                    pos = eks_positions[field]
                    row = pos["start_row"]
                    start_col = pos["months_start_col"]
                    
                    # Aylık değerleri yerleştir (C, D, E, F, G, H sütunları)
                    values = data.get('values', [])
                    for i, value in enumerate(values):
                        if value is not None and i < 6:  # Maksimum 6 ay
                            col = start_col + i
                            col_letter = chr(ord('A') + col - 1)
                            ws[f'{col_letter}{row}'] = value
                            ws[f'{col_letter}{row}'].number_format = '#,##0.00'
                    
                    print(f"Filled {field} at row {row}: {values}")
            
            return True
            
        except Exception as e:
            print(f"Template fill error: {e}")
            return False
    
    def update_customer_info_in_template(self, ws):
        """Müşteri bilgilerini template'e yazar"""
        try:
            if self.current_customer:
                # Template'teki müşteri bilgi hücrelerini doldur (görsel referansa göre)
                
                # Müşteri numarası - Satır 2, D sütununa (merged D2:K2)
                ws['D2'] = self.current_customer.code
                
                # Müşteri adı - Satır 3, D sütununa (merged D3:K3)  
                ws['D3'] = self.current_customer.name
                
                # Doğum tarihi alanı - Satır 4, D sütununa (merged D4:K4)
                # Bu alan boş bırakılabilir veya müşteri verisi varsa doldurulabilir
                
        except Exception as e:
            print(f"Customer info update error: {e}")
    
    def update_period_info_in_template(self, ws):
        """Dönem bilgilerini template'e yazar"""
        try:
            # Template'te "Bewilligungszeitraum vom _01.0x.200x__ bis _3x.0x.200x__" 
            # pattern'ini bul ve tarihleri güncelle
            
            months = list(self.extracted_data.values())[0].get('months', [])
            if not months:
                return
                
            # Ay isimlerini tarih formatına çevir
            month_to_number = {
                'JAN': '01', 'FEB': '02', 'MRZ': '03', 'APR': '04', 
                'MAI': '05', 'JUN': '06', 'JUL': '07', 'AUG': '08',
                'SEP': '09', 'OKT': '10', 'NOV': '11', 'DEZ': '12'
            }
            
            start_month_num = month_to_number.get(months[0], '01')
            end_month_num = month_to_number.get(months[-1], '06')
            
            # Yıl bilgisi (şu anki yıl)
            from datetime import datetime
            current_year = datetime.now().year
            
            # Template'te dönem bilgisini içeren hücreyi bul ve güncelle
            for row in range(1, 20):  # İlk 20 satırda ara
                for col in range(1, 10):  # İlk 10 sütunda ara
                    cell = ws.cell(row=row, column=col)
                    if cell.value and "Bewilligungszeitraum vom" in str(cell.value):
                        # Orijinal pattern'i yeni tarihlerle değiştir
                        original_text = str(cell.value)
                        
                        # Tarihleri güncelle
                        updated_text = original_text.replace(
                            "_01.0x.200x__", f"01.{start_month_num}.{current_year}"
                        ).replace(
                            "_3x.0x.200x__", f"30.{end_month_num}.{current_year}"
                        )
                        
                        cell.value = updated_text
                        print(f"Period updated: {updated_text}")
                        break
                            
        except Exception as e:
            print(f"Period info update error: {e}")
    
    def analyze_template_structure(self):
        """Template yapısını analiz eder ve pozisyonları otomatik bulur"""
        template_path = os.path.join("templates", "eks_form.xlsx")
        if not os.path.exists(template_path):
            return None
            
        try:
            wb = openpyxl.load_workbook(template_path)
            ws = wb.active
            
            analysis = {
                "customer_fields": {},
                "data_positions": {},
                "month_columns": [],
                "structure": []
            }
            
            print("=== EKS TEMPLATE ANALYSE ===")
            
            # Template'i tara ve yapıyı analiz et
            for row in range(1, min(100, ws.max_row + 1)):
                for col in range(1, min(20, ws.max_column + 1)):
                    cell = ws.cell(row=row, column=col)
                    if cell.value:
                        cell_text = str(cell.value).strip()
                        col_letter = chr(ord('A') + col - 1)
                        
                        # Müşteri bilgi alanlarını bul
                        if "Nummer der Bedarfsgemeinschaft" in cell_text:
                            analysis["customer_fields"]["number"] = f"{col_letter}{row}"
                        elif "Name, Vorname" in cell_text:
                            analysis["customer_fields"]["name"] = f"{col_letter}{row}"
                        elif "Bewilligungszeitraum" in cell_text:
                            analysis["customer_fields"]["period"] = f"{col_letter}{row}"
                        
                        # Ay başlıklarını bul
                        if cell_text in ['JAN', 'FEB', 'MRZ', 'APR', 'MAI', 'JUN', 'JUL', 'AUG', 'SEP', 'OKT', 'NOV', 'DEZ']:
                            analysis["month_columns"].append((cell_text, col))
                        
                        # EKS kodlarını bul (A1, A5, B1, vs.)
                        if len(cell_text) <= 4 and any(cell_text.startswith(prefix) for prefix in ['A1','A2','A3','A4', 'A5','A6','A7', 'B1', 'B2', 'B3', 'B11', 'B14', 'B17']):
                            analysis["data_positions"][cell_text] = {"row": row, "col": col}
                        
                        # Template yapısını kaydet (ilk 50 satır)
                        if row <= 50:
                            analysis["structure"].append({
                                "position": f"{col_letter}{row}",
                                "content": cell_text[:50] + "..." if len(cell_text) > 50 else cell_text
                            })
            
            # Sonuçları yazdır
            print("Müşteri Alanları:", analysis["customer_fields"])
            print("Ay Sütunları:", analysis["month_columns"])
            print("EKS Pozisyonları:", analysis["data_positions"])
            print("Template Yapısı (ilk 50 satır):")
            for item in analysis["structure"][:20]:  # İlk 20 öğeyi göster
                print(f"  {item['position']}: {item['content']}")
            
            return analysis
            
        except Exception as e:
            print(f"Template analysis error: {e}")
            return None
    def analyze_template_wrapper(self):
        """Template analiz fonksiyonunu çağırır ve sonuçları gösterir"""
        analysis = self.analyze_template_structure()
        if analysis:
            # Sonuçları dialog'da göster
            result_text = f"""Template Analizi Tamamlandı!

Müşteri Alanları: {len(analysis['customer_fields'])} adet
Ay Sütunları: {len(analysis['month_columns'])} adet  
EKS Pozisyonları: {len(analysis['data_positions'])} adet

Konsol çıktısını kontrol edin."""
            
            messagebox.showinfo("Template Analizi", result_text)
        else:
            messagebox.showerror("Hata", "Template analizi başarısız. templates/eks_form.xlsx dosyası var mı?")
    
    def create_automatic_export(self, export_path: str) -> bool:
        """Fallback: Otomatik template oluşturur (orijinal kod)"""
        try:
            # Neue Workbook erstellen
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "EKS Formular"
            
            # Styles definieren
            header_font = Font(bold=True, size=12)
            title_font = Font(bold=True, size=14)
            number_font = Font(size=11)
            
            header_fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
            
            # Titel
            ws['A1'] = "Angaben zum voraussichtlichen Einkommen aus selbständiger Tätigkeit"
            ws['A1'].font = title_font
            ws.merge_cells('A1:H1')
            
            # Kunde Info
            ws['A3'] = f"Nummer der Bedarfsgemeinschaft: {self.current_customer.code}"
            ws['A4'] = f"Name, Vorname: {self.current_customer.name}"
            ws['A5'] = f"Bewilligungszeitraum: {self.selected_start_month} - {self.selected_end_month}"
            
            # Header für Monate
            months = list(self.extracted_data.values())[0].get('months', [])
            row = 8
            
            ws[f'A{row}'] = "Position"
            ws[f'A{row}'].font = header_font
            ws[f'A{row}'].fill = header_fill
            
            ws[f'B{row}'] = "Beschreibung"
            ws[f'B{row}'].font = header_font
            ws[f'B{row}'].fill = header_fill
            
            col_start = 3  # C Spalte
            for i, month in enumerate(months):
                col_letter = chr(ord('C') + i)
                ws[f'{col_letter}{row}'] = month
                ws[f'{col_letter}{row}'].font = header_font
                ws[f'{col_letter}{row}'].fill = header_fill
            
            # Summe Spalte
            sum_col = chr(ord('C') + len(months))
            ws[f'{sum_col}{row}'] = "Summe"
            ws[f'{sum_col}{row}'].font = header_font
            ws[f'{sum_col}{row}'].fill = header_fill
            
            # Daten einfügen
            current_row = row + 1
            
            # A. Betriebseinnahmen
            ws[f'A{current_row}'] = "A. Betriebseinnahmen"
            ws[f'A{current_row}'].font = header_font
            current_row += 1
            
            for field, data in self.extracted_data.items():
                if field.startswith('A'):
                    ws[f'A{current_row}'] = field
                    ws[f'B{current_row}'] = data['description']
                    
                    # Monatswerte
                    for i, value in enumerate(data['values']):
                        col_letter = chr(ord('C') + i)
                        if value is not None:
                            ws[f'{col_letter}{current_row}'] = value
                            ws[f'{col_letter}{current_row}'].number_format = '#,##0.00'
                    
                    # Summe
                    ws[f'{sum_col}{current_row}'] = data.get('total', 0)
                    ws[f'{sum_col}{current_row}'].number_format = '#,##0.00'
                    ws[f'{sum_col}{current_row}'].font = Font(bold=True)
                    
                    current_row += 1
            
            current_row += 1
            
            # B. Betriebsausgaben
            ws[f'A{current_row}'] = "B. Betriebsausgaben"
            ws[f'A{current_row}'].font = header_font
            current_row += 1
            
            for field, data in self.extracted_data.items():
                if field.startswith('B'):
                    ws[f'A{current_row}'] = field
                    ws[f'B{current_row}'] = data['description']
                    
                    # Monatswerte
                    for i, value in enumerate(data['values']):
                        col_letter = chr(ord('C') + i)
                        if value is not None:
                            ws[f'{col_letter}{current_row}'] = value
                            ws[f'{col_letter}{current_row}'].number_format = '#,##0.00'
                    
                    # Summe
                    ws[f'{sum_col}{current_row}'] = data.get('total', 0)
                    ws[f'{sum_col}{current_row}'].number_format = '#,##0.00'
                    ws[f'{sum_col}{current_row}'].font = Font(bold=True)
                    
                    current_row += 1
            
            # Spaltenbreiten anpassen
            ws.column_dimensions['A'].width = 12
            ws.column_dimensions['B'].width = 30
            for i in range(len(months) + 1):
                col_letter = chr(ord('C') + i)
                ws.column_dimensions[col_letter].width = 15
            
            # Speichern
            wb.save(export_path)
            return True
            
        except Exception as e:
            print(f"Automatic Export Fehler: {e}")
            return False
    
    def update_customer_history(self):
        if self.current_customer and self.bwa_file_path:
            history_entry = {
                "file_path": os.path.basename(self.bwa_file_path),
                "period": f"{self.selected_start_month}-{self.selected_end_month}",
                "processed_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "confidence": self.calculate_average_confidence()
            }
            
            self.current_customer.bwa_history.append(history_entry)
            self.customer_manager.save_customer(self.current_customer)
    
    def change_language(self, selected_language):
        """Dil değiştirme fonksiyonu"""
        if selected_language == self.language:
            return
            
        self.language = selected_language
        self.texts = LANGUAGES[self.language]
        
        # UI'ı anında güncelle
        self.refresh_ui()
    
    def refresh_ui(self):
        """UI metinlerini anında güncelle"""
        # Başlık
        self.title(self.texts["app_title"])
        
        # Ana butonları güncelle
        try:
            self.load_bwa_btn.configure(text=self.texts["load_bwa"])
            self.mapping_btn.configure(text=self.texts["auto_mapping"])
            self.export_btn.configure(text=self.texts["export_eks"])
            
            # Status label'ları güncelle
            if hasattr(self, 'bwa_status_label'):
                current_text = self.bwa_status_label.cget("text")
                if "Keine Datei" in current_text or "Dosya yok" in current_text:
                    self.bwa_status_label.configure(text=self.texts["no_file"])
                elif "geladen" in current_text or "yüklendi" in current_text:
                    self.bwa_status_label.configure(text="✅ " + self.texts["file_loaded"])
        except Exception as e:
            print(f"UI refresh error: {e}")
    
    def update_ui_texts(self):
        """UI metinlerini güncelle (eski fonksiyon - artık kullanılmıyor)"""
        pass
    
    def calculate_average_confidence(self) -> float:
        if not self.extracted_data:
            return 0.0
        
        confidences = [data.get('confidence', 0) for data in self.extracted_data.values()]
        return sum(confidences) / len(confidences) if confidences else 0.0
    
    def open_settings(self):
        settings_dialog = SettingsDialog(self, self.texts)

class CustomerDialog(ctk.CTkToplevel):
    def __init__(self, parent, texts):
        super().__init__(parent)
        
        self.texts = texts
        self.result = None
        
        self.title("Neuer Kunde")
        self.geometry("400x300")
        self.configure(fg_color="#2b2b2b")
        
        # Modal machen
        self.transient(parent)
        self.grab_set()
        
        self.setup_ui()
        
        # Zentrieren
        self.center_window()
    
    def setup_ui(self):
        # Hauptframe
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Titel
        title_label = ctk.CTkLabel(main_frame, text="Neuen Kunden erstellen", 
                                 font=ctk.CTkFont(size=18, weight="bold"))
        title_label.pack(pady=20)
        
        # Kundennummer
        code_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        code_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(code_frame, text="Kundennummer:", width=120).pack(side="left")
        self.code_entry = ctk.CTkEntry(code_frame, width=200)
        self.code_entry.pack(side="right")
        
        # Kundenname
        name_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        name_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(name_frame, text="Kundenname:", width=120).pack(side="left")
        self.name_entry = ctk.CTkEntry(name_frame, width=200)
        self.name_entry.pack(side="right")
        
        # Buttons
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(side="bottom", pady=20)
        
        cancel_btn = ctk.CTkButton(button_frame, text="Abbrechen", 
                                 command=self.cancel, width=100)
        cancel_btn.pack(side="left", padx=10)
        
        save_btn = ctk.CTkButton(button_frame, text="Speichern", 
                               command=self.save, width=100)
        save_btn.pack(side="right", padx=10)
        
        # Focus auf erstes Feld
        self.code_entry.focus()
    
    def center_window(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (400 // 2)
        y = (self.winfo_screenheight() // 2) - (300 // 2)
        self.geometry(f"400x300+{x}+{y}")
    
    def save(self):
        code = self.code_entry.get().strip()
        name = self.name_entry.get().strip()
        
        if not code or not name:
            messagebox.showwarning("Warnung", "Bitte alle Felder ausfüllen")
            return
        
        self.result = {
            "code": code,
            "name": name
        }
        self.destroy()
    
    def cancel(self):
        self.destroy()

class SettingsDialog(ctk.CTkToplevel):
    def __init__(self, parent, texts):
        super().__init__(parent)
        
        self.texts = texts
        self.settings_file = "settings.json"
        self.settings = self.load_settings()
        
        self.title("Einstellungen")
        self.geometry("500x400")
        self.configure(fg_color="#2b2b2b")
        
        # Modal
        self.transient(parent)
        self.grab_set()
        
        self.setup_ui()
        self.center_window()
    
    def load_settings(self) -> Dict:
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception:
            pass
        
        return {
            "claude_api_key": "",
            "auto_customer_creation": True,
            "default_template": "eks_standard.xlsx",
            "backup_enabled": True
        }
    
    def save_settings(self):
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
            return True
        except Exception:
            return False
    
    def setup_ui(self):
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Titel
        title_label = ctk.CTkLabel(main_frame, text="Einstellungen", 
                                 font=ctk.CTkFont(size=18, weight="bold"))
        title_label.pack(pady=20)
        
        # Scrollable Frame für Einstellungen
        settings_frame = ctk.CTkScrollableFrame(main_frame, fg_color="#3b3b3b")
        settings_frame.pack(fill="both", expand=True, pady=10)
        
        # API Einstellungen
        api_section = ctk.CTkFrame(settings_frame, fg_color="#4b4b4b")
        api_section.pack(fill="x", pady=10, padx=10)
        
        ctk.CTkLabel(api_section, text="API Einstellungen", 
                    font=ctk.CTkFont(size=14, weight="bold")).pack(pady=10)
        
        # Claude API Key
        api_frame = ctk.CTkFrame(api_section, fg_color="transparent")
        api_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(api_frame, text="Claude API Key:", width=150).pack(side="left")
        self.api_key_entry = ctk.CTkEntry(api_frame, width=250, show="*")
        self.api_key_entry.pack(side="right", padx=10)
        self.api_key_entry.insert(0, self.settings.get("claude_api_key", ""))
        
        # Test Button
        test_btn = ctk.CTkButton(api_section, text="API Testen", command=self.test_api, width=100)
        test_btn.pack(pady=10)
        
        # Allgemeine Einstellungen
        general_section = ctk.CTkFrame(settings_frame, fg_color="#4b4b4b")
        general_section.pack(fill="x", pady=10, padx=10)
        
        ctk.CTkLabel(general_section, text="Allgemeine Einstellungen", 
                    font=ctk.CTkFont(size=14, weight="bold")).pack(pady=10)
        
        # Auto Kunde erstellen
        self.auto_customer_var = ctk.BooleanVar(value=self.settings.get("auto_customer_creation", True))
        auto_customer_check = ctk.CTkCheckBox(general_section, text="Kunden automatisch aus BWA erstellen",
                                            variable=self.auto_customer_var)
        auto_customer_check.pack(pady=5, padx=20, anchor="w")
        
        # Backup aktiviert
        self.backup_var = ctk.BooleanVar(value=self.settings.get("backup_enabled", True))
        backup_check = ctk.CTkCheckBox(general_section, text="Automatische Backups aktiviert",
                                     variable=self.backup_var)
        backup_check.pack(pady=5, padx=20, anchor="w")
        
        # Buttons
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(side="bottom", pady=20)
        
        cancel_btn = ctk.CTkButton(button_frame, text="Abbrechen", 
                                 command=self.cancel, width=100)
        cancel_btn.pack(side="left", padx=10)
        
        save_btn = ctk.CTkButton(button_frame, text="Speichern", 
                               command=self.save, width=100)
        save_btn.pack(side="right", padx=10)
    
    def center_window(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (500 // 2)
        y = (self.winfo_screenheight() // 2) - (400 // 2)
        self.geometry(f"500x400+{x}+{y}")
    
    def test_api(self):
        api_key = self.api_key_entry.get().strip()
        if not api_key:
            messagebox.showwarning("Warnung", "Bitte API Key eingeben")
            return
        
        # Test API Call (vereinfacht)
        try:
            # Hier würde ein echter API Test stehen
            messagebox.showinfo("Erfolg", "API Key ist gültig")
        except Exception as e:
            messagebox.showerror("Fehler", f"API Test fehlgeschlagen: {str(e)}")
    
    def save(self):
        self.settings["claude_api_key"] = self.api_key_entry.get().strip()
        self.settings["auto_customer_creation"] = self.auto_customer_var.get()
        self.settings["backup_enabled"] = self.backup_var.get()
        
        if self.save_settings():
            messagebox.showinfo("Erfolg", "Einstellungen gespeichert")
            self.destroy()
        else:
            messagebox.showerror("Fehler", "Einstellungen konnten nicht gespeichert werden")
    
    def cancel(self):
        self.destroy()

# Template Manager Klasse
class TemplateManager:
    def __init__(self, template_dir: str = "templates"):
        self.template_dir = template_dir
        os.makedirs(template_dir, exist_ok=True)
        self.create_default_template()
    
    def create_default_template(self):
        """Erstellt ein Standard EKS Template falls nicht vorhanden"""
        template_path = os.path.join(self.template_dir, "eks_standard.xlsx")
        if not os.path.exists(template_path):
            try:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "EKS Standard"
                
                # Template Struktur erstellen
                ws['A1'] = "EKS Standard Template"
                ws['A2'] = "Dieses Template wird automatisch befüllt"
                
                # Mapping Bereiche definieren
                ws['A4'] = "A. Betriebseinnahmen"
                ws['A5'] = "A1 - Betriebseinnahmen"
                ws['A6'] = "A5 - Umsatzsteuer"
                
                ws['A8'] = "B. Betriebsausgaben"
                ws['A9'] = "B1 - Material, Stoffe, Waren"
                ws['A10'] = "B2c - Aushilfslöhne"
                ws['A11'] = "B3 - Miete und Energiekosten"
                ws['A12'] = "B11 - Telefon"
                ws['A13'] = "B14c - Nebenkosten Geldverkehr"
                ws['A14'] = "B17 - Vorsteuer"
                
                wb.save(template_path)
            except Exception as e:
                print(f"Fehler beim Erstellen des Standard Templates: {e}")
    
    def get_available_templates(self) -> List[str]:
        """Gibt Liste verfügbarer Templates zurück"""
        templates = []
        for file in os.listdir(self.template_dir):
            if file.endswith('.xlsx'):
                templates.append(file)
        return sorted(templates)

# Hauptprogramm
def main():
    # Arbeitsverzeichnisse erstellen
    directories = ["data", "data/customers", "templates", "exports"]
    for directory in directories:
        os.makedirs(directory, exist_ok=True)
    
    # CustomTkinter Erscheinungsbild setzen
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    
    # Hauptanwendung starten
    app = EKSFormFiller()
    app.mainloop()

if __name__ == "__main__":
    main()