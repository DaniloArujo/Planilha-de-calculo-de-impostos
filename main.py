import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from typing import Dict, List, Optional, Union
import locale
from dataclasses import dataclass
from enum import Enum

# Configuração de locale para pt_BR
def configure_locale():
    try:
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    except locale.Error:
        try:
            locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')
        except locale.Error:
            print("Não foi possível configurar o locale para pt_BR. Usando padrão do sistema.")

configure_locale()

# ==================== MODELO ====================

@dataclass
class TaxConfig:
    """Configurações de impostos padrão"""
    ICMS: float = 20.0
    PIS: float = 0.65
    COFINS: float = 3
    IRPJ: float = 1.2
    CSLL: float = 1.08

class TaxCalculator:
    """Classe responsável por calcular os impostos e valores relacionados"""
    
    @staticmethod
    def calculate_taxes(row: Dict[str, float], tax_config: TaxConfig) -> Dict[str, float]:
        """Calcula todos os impostos e valores derivados"""
        calculations = {}
        
        # Valores básicos
        unit_cost = row['Valor Unitário de Custo (R$)']
        quantity = row['Quantidade']
        profit_margin = row['Margem de Lucro Bruto (%)'] / 100
        
        calculations['Valor Total de Custo (R$)'] = unit_cost * quantity
        calculations['Valor Unitário de Venda (R$)'] = unit_cost * (1 + profit_margin)
        calculations['Valor Total de Venda (R$)'] = calculations['Valor Unitário de Venda (R$)'] * quantity
        
        # Cálculos de impostos
        taxes = {
            'ICMS': row.get('ICMS (%)', tax_config.ICMS),
            'PIS': row.get('PIS (%)', tax_config.PIS),
            'COFINS': row.get('COFINS (%)', tax_config.COFINS),
            'IRPJ': row.get('IRPJ (%)', tax_config.IRPJ),
            'CSLL': row.get('CSLL (%)', tax_config.CSLL)
        }
        
        for tax_name, rate in taxes.items():
            calculations[f'Valor unit. {tax_name}'] = calculations['Valor Unitário de Venda (R$)'] * (rate/100)
            calculations[f'Valor Total {tax_name} (R$)'] = calculations[f'Valor unit. {tax_name}'] * quantity
            calculations[f'{tax_name} (%)'] = rate
        
        # Totais consolidados
        total_taxes = sum(calculations[f'Valor Total {tax_name} (R$)'] for tax_name in taxes.keys())
        calculations['Valor Total de impostos'] = total_taxes
        calculations['Valor Total Unitário'] = calculations['Valor Unitário de Venda (R$)'] + sum(
            calculations[f'Valor unit. {tax_name}'] for tax_name in taxes.keys()
        )
        calculations['Valor Total'] = calculations['Valor Total de Venda (R$)'] + total_taxes
        calculations['Total Alíquota Impostos (%)'] = sum(taxes.values())
        
        return calculations

class DataModel:
    """Classe responsável por gerenciar os dados da aplicação"""
    def __init__(self):
        self.columns = [
            "Item", "Descrição", "Valor Unitário de Custo (R$)", "Quantidade", 
            "Valor Total de Custo (R$)", "Margem de Lucro Bruto (%)", 
            "Valor Unitário de Venda (R$)", "Valor Total de Venda (R$)", 
            "Estado de Destino", "ICMS (%)", "Valor unit. ICMS", 
            "Valor Total ICMS (R$)", "PIS (%)", "Valor unit. PIS", 
            "Valor Total PIS (R$)", "COFINS (%)", "Valor unit. COFINS", 
            "Valor Total COFINS (R$)", "IRPJ (%)", "Valor unit. IRPJ", 
            "Valor Total IRPJ (R$)", "CSLL (%)", "Valor unit. CSLL", 
            "Valor Total CSLL (R$)", "Valor Total de impostos", 
            "Valor Total Unitário", "Valor Total", "Total Alíquota Impostos (%)"
        ]
        self.data = pd.DataFrame(columns=self.columns)
        self.current_file = None
        self.tax_config = TaxConfig()
        self.next_item_number = 1
        
        self.state_icms_table = {
            "AC": 17, "AL": 18, "AP": 18, "AM": 20, "BA": 20.5, 
            "CE": 20, "DF": 20, "ES": 17, "GO": 17, "MA": 22, 
            "MT": 17, "MS": 17, "MG": 18, "PA": 19, "PB": 18, 
            "PR": 19.5, "PE": 20.5, "PI": 21, "RJ": 20, "RN": 18, 
            "RS": 17, "RO": 17.5, "RR": 17, "SC": 17, "SP": 17, 
            "SE": 18, "TO": 18
        }
    
    def add_item(self, item_data: Dict[str, Union[str, float]]) -> None:
        """Adiciona um novo item ao DataFrame com validação completa"""
        try:
            processed_data = {}
            for key, value in item_data.items():
                if key in ['Descrição', 'Estado de Destino']:
                    processed_data[key] = str(value)
                else:
                    if isinstance(value, str):
                        processed_data[key] = float(value.replace('.', '').replace(',', '.'))
                    else:
                        processed_data[key] = float(value)
        
            if processed_data['Estado de Destino'] not in self.state_icms_table:
                raise ValueError("Estado inválido")
            if processed_data['Valor Unitário de Custo (R$)'] <= 0:
                raise ValueError("Valor unitário de custo deve ser positivo")
            if processed_data['Quantidade'] <= 0:
                raise ValueError("Quantidade deve ser positiva")
            if not processed_data['Descrição']:
                raise ValueError("Descrição do item é obrigatória")
            
            tax_calculations = TaxCalculator.calculate_taxes(processed_data, self.tax_config)
            new_row = {**processed_data, **tax_calculations}
            new_row['Item'] = self.next_item_number
            
            new_df = pd.DataFrame([new_row])
            self.data = pd.concat([self.data, new_df], ignore_index=True)
            self.next_item_number += 1
            
        except Exception as e:
            raise ValueError(f"Erro ao adicionar item: {str(e)}")
    
    def update_item(self, index: int, column: str, new_value: Union[str, float]) -> None:
        """Atualiza um valor específico e recalcula os dependentes"""
        try:
            if column not in ['Descrição', 'Estado de Destino', 'Item']:
                if isinstance(new_value, str):
                    new_value = float(new_value.replace('.', '').replace(',', '.'))
            
            self.data.at[index, column] = new_value
            
            if column in ['Valor Unitário de Custo (R$)', 'Quantidade', 'Margem de Lucro Bruto (%)', 
                         'ICMS (%)', 'PIS (%)', 'COFINS (%)', 'IRPJ (%)', 'CSLL (%)', 'Estado de Destino']:
                item_data = self.data.iloc[index].to_dict()
                tax_calculations = TaxCalculator.calculate_taxes(item_data, self.tax_config)
                
                for col, value in tax_calculations.items():
                    self.data.at[index, col] = value
                    
        except Exception as e:
            raise ValueError(f"Erro ao atualizar item: {str(e)}")
    
    def delete_items(self, indices: List[int]) -> None:
        """Remove itens do DataFrame pelos índices"""
        self.data = self.data.drop(indices).reset_index(drop=True)
        self.data['Item'] = range(1, len(self.data) + 1)
        self.next_item_number = len(self.data) + 1
    
    def calculate_totals(self) -> Dict[str, float]:
        """Calcula os totais consolidados automaticamente"""
        if self.data.empty:
            return {}
        
        columns_to_sum = [
            "Valor Unitário de Custo (R$)", 
            "Quantidade",
            "Valor Total de Custo (R$)", 
            "Valor Total de Venda (R$)",
            "Valor Total ICMS (R$)",
            "Valor Total PIS (R$)",
            "Valor Total COFINS (R$)",
            "Valor Total IRPJ (R$)",
            "Valor Total CSLL (R$)",
            "Valor Total de impostos",
            "Valor Total"
        ]
        
        # Garantir que todas as colunas existem no DataFrame
        valid_columns = [col for col in columns_to_sum if col in self.data.columns]
        
        # Calcular totais com tratamento para dados ausentes
        totals = self.data[valid_columns].apply(lambda x: pd.to_numeric(x, errors='coerce')).sum().to_dict()
        
        # Adicionar campos não numéricos
        totals['Descrição'] = "TOTAL"
        totals['Item'] = ""
        totals['Estado de Destino'] = ""
        totals['Margem de Lucro Bruto (%)'] = ""
        
        # Garantir todas as colunas necessárias
        for col in self.columns:
            if col not in totals:
                totals[col] = ""
        
        return totals

    def clear_data(self) -> None:
        self.data = pd.DataFrame(columns=self.columns)
        self.next_item_number = 1
    
    def load_from_file(self, filepath: str) -> None:
        """Carrega dados removendo totais existentes do arquivo"""
        try:
            if filepath.endswith('.xlsx'):
                df = pd.read_excel(filepath)
            elif filepath.endswith('.csv'):
                df = pd.read_csv(filepath)
            else:
                raise ValueError("Formato de arquivo não suportado")
            
            # Remove linha de totais se existir
            if not df.empty and df.iloc[-1]['Descrição'] == "TOTAL":
                df = df.iloc[:-1]  # Remove última linha
            
            # Garante todas as colunas necessárias
            for col in self.columns:
                if col not in df.columns:
                    df[col] = 0.0 if col not in ['Descrição', 'Estado de Destino', 'Item'] else ""
            
            self.data = df
            if not self.data.empty:
                self.next_item_number = self.data['Item'].max() + 1
            else:
                self.next_item_number = 1
            
            self.current_file = filepath
        except Exception as e:
            raise ValueError(f"Erro ao carregar arquivo: {str(e)}")   
    
    def save_to_file(self, filepath: str) -> None:
        """Salva os dados incluindo os totais como última linha"""
        try:
            # Cria uma cópia dos dados para adicionar os totais
            data_to_save = self.data.copy()
            totals = self.calculate_totals()
            
            if totals:
                totals_df = pd.DataFrame([totals])
                data_to_save = pd.concat([data_to_save, totals_df], ignore_index=True)
            
            if filepath.endswith('.xlsx'):
                data_to_save.to_excel(filepath, index=False)
            elif filepath.endswith('.csv'):
                data_to_save.to_csv(filepath, index=False)
            else:
                raise ValueError("Formato de arquivo não suportado")
            
            self.current_file = filepath
        except Exception as e:
            raise ValueError(f"Erro ao salvar arquivo: {str(e)}")

   

# ==================== VISUALIZAÇÃO ====================

class ColorScheme(Enum):
    PRIMARY = "#2c3e50"
    SECONDARY = "#34495e"
    ACCENT = "#3498db"
    BACKGROUND = "#ecf0f1"
    TEXT = "#2c3e50"
    SUCCESS = "#27ae60"
    WARNING = "#f39c12"
    ERROR = "#e74c3c"
    LIGHT_GRAY = "#bdc3c7"
    WHITE = "#ffffff"
    HIGHLIGHT = "#2980b9"
    ROW_EVEN = "#CBD8DE"
    ROW_ODD = "#ffffff"

class Fonts(Enum):
    TITLE = ("Segoe UI", 14, "bold")
    HEADER = ("Segoe UI", 12, "bold")
    BODY = ("Segoe UI", 10)
    SMALL = ("Segoe UI", 9)

class ICMSEditorWindow:
    def __init__(self, parent, state_icms_table: Dict[str, float], update_callback):
        self.parent = parent
        self.state_icms_table = state_icms_table.copy()
        self.update_callback = update_callback
        
        self.window = tk.Toplevel(parent)
        self.window.title("Editar Tabela de ICMS por Estado")
        self.window.geometry("500x600")
        self.window.configure(bg=ColorScheme.BACKGROUND.value)
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.configure_styles()
        self.create_widgets()
        self.populate_table()
    
    def configure_styles(self) -> None:
        self.style.configure("Treeview", 
                           font=Fonts.BODY.value,
                           rowheight=28,
                           background=ColorScheme.WHITE.value,
                           fieldbackground=ColorScheme.WHITE.value,
                           foreground=ColorScheme.TEXT.value,
                           bordercolor=ColorScheme.LIGHT_GRAY.value)
        self.style.configure("Treeview.Heading", 
                           font=Fonts.HEADER.value,
                           background=ColorScheme.PRIMARY.value,
                           foreground=ColorScheme.WHITE.value,
                           relief="flat")
        self.style.map("Treeview",
                      background=[('selected', ColorScheme.HIGHLIGHT.value)],
                      foreground=[('selected', ColorScheme.WHITE.value)])
        self.style.configure("Accent.TButton",
                           background=ColorScheme.ACCENT.value,
                           foreground=ColorScheme.WHITE.value)
    
    def create_widgets(self) -> None:
        self.main_frame = ttk.Frame(self.window, padding=15)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        self.tree = ttk.Treeview(self.main_frame, columns=("Estado", "ICMS"), show="headings")
        self.tree.heading("Estado", text="Estado", anchor=tk.CENTER)
        self.tree.heading("ICMS", text="ICMS (%)", anchor=tk.CENTER)
        self.tree.column("Estado", width=150, anchor=tk.CENTER)
        self.tree.column("ICMS", width=150, anchor=tk.CENTER)
        
        scrollbar = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        save_btn = ttk.Button(
            self.main_frame, 
            text="Salvar Alterações", 
            command=self.save_changes,
            style="Accent.TButton"
        )
        
        self.tree.grid(row=0, column=0, sticky="nsew", pady=(0, 15))
        scrollbar.grid(row=0, column=1, sticky="ns", pady=(0, 15))
        save_btn.grid(row=1, column=0, columnspan=2, sticky="ew")
        
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        self.tree.bind('<Double-1>', self.edit_cell)
    
    def populate_table(self) -> None:
        for estado, icms in sorted(self.state_icms_table.items()):
            self.tree.insert("", tk.END, values=(estado, f"{icms:.2f}"), iid=estado)
    
    def edit_cell(self, event) -> None:
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        
        if not item or column != '#2':
            return
        
        current_value = self.tree.set(item, "ICMS")
        
        x, y, width, height = self.tree.bbox(item, column)
        entry = ttk.Entry(self.tree, font=Fonts.BODY.value)
        entry.place(x=x, y=y, width=width, height=height, anchor=tk.NW)
        entry.insert(0, current_value)
        entry.select_range(0, tk.END)
        entry.focus()
        
        def save_edit():
            try:
                new_value = float(entry.get().replace(',', '.'))
                if 0 <= new_value <= 100:
                    self.tree.set(item, "ICMS", f"{new_value:.2f}")
                    self.state_icms_table[item] = new_value
                    entry.destroy()
                else:
                    messagebox.showerror("Erro", "O valor do ICMS deve estar entre 0 e 100%")
            except ValueError:
                messagebox.showerror("Erro", "Por favor, insira um valor numérico válido.")
        
        entry.bind("<FocusOut>", lambda e: save_edit())
        entry.bind("<Return>", lambda e: save_edit())
    
    def save_changes(self) -> None:
        self.update_callback(self.state_icms_table)
        self.window.destroy()

class MainView:
    def __init__(self, root, model: DataModel, controller):
        self.root = root
        self.model = model
        self.controller = controller
        self.configure_styles()
        self.setup_ui()
    
    def configure_styles(self) -> None:
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        self.style.configure(".", 
                           font=Fonts.BODY.value,
                           background=ColorScheme.BACKGROUND.value,
                           foreground=ColorScheme.TEXT.value)
        
        self.style.configure("TButton", 
                           font=Fonts.BODY.value,
                           padding=6,
                           background=ColorScheme.PRIMARY.value,
                           foreground=ColorScheme.WHITE.value,
                           borderwidth=1)
        self.style.configure("Accent.TButton",
                           background=ColorScheme.ACCENT.value)
        self.style.map("TButton",
                      background=[('active', ColorScheme.SECONDARY.value)],
                      relief=[('pressed', 'sunken'), ('!pressed', 'raised')])
        self.style.map("Accent.TButton",
                     background=[('active', ColorScheme.HIGHLIGHT.value)])
        
        self.style.configure("TLabel", font=Fonts.BODY.value)
        self.style.configure("Title.TLabel", 
                           font=Fonts.TITLE.value,
                           foreground=ColorScheme.WHITE.value,
                           background=ColorScheme.PRIMARY.value)
        self.style.configure("Subtitle.TLabel",
                           font=Fonts.HEADER.value,
                           foreground=ColorScheme.PRIMARY.value)
        
        self.style.configure("TFrame", background=ColorScheme.BACKGROUND.value)
        self.style.configure("Header.TFrame", background=ColorScheme.PRIMARY.value)
        self.style.configure("TLabelframe", 
                           background=ColorScheme.BACKGROUND.value,
                           bordercolor=ColorScheme.LIGHT_GRAY.value)
        self.style.configure("TLabelframe.Label", 
                           font=Fonts.HEADER.value,
                           foreground=ColorScheme.PRIMARY.value)
        
        self.style.configure("TCombobox", 
                           font=Fonts.BODY.value, 
                           padding=5,
                           fieldbackground=ColorScheme.WHITE.value)
        
        self.style.configure("Treeview", 
                           font=Fonts.BODY.value,
                           rowheight=28,
                           background=ColorScheme.WHITE.value,
                           fieldbackground=ColorScheme.WHITE.value,
                           foreground=ColorScheme.TEXT.value,
                           bordercolor=ColorScheme.LIGHT_GRAY.value)
        self.style.configure("Treeview.Heading", 
                           font=Fonts.HEADER.value,
                           background=ColorScheme.PRIMARY.value,
                           foreground=ColorScheme.WHITE.value,
                           relief="flat")
        self.style.map("Treeview",
                      background=[('selected', ColorScheme.HIGHLIGHT.value)],
                      foreground=[('selected', ColorScheme.WHITE.value)])
        
        self.style.configure("Status.TLabel", 
                           font=Fonts.SMALL.value,
                           background=ColorScheme.SECONDARY.value,
                           foreground=ColorScheme.WHITE.value,
                           padding=5,
                           relief=tk.SUNKEN)
    
    def setup_ui(self) -> None:
        self.root.title("Sistema de Cálculo de Custos e Impostos")
        self.root.state('zoomed')
        self.root.configure(bg=ColorScheme.BACKGROUND.value)
        
        try:
            self.root.iconbitmap("calculator_icon.ico")
        except:
            pass
        
        self.brazilian_states = list(self.model.state_icms_table.keys())
        
        self.create_header()
        self.create_input_frame()
        self.create_table_frame()
        self.create_status_bar()
        self.create_menu()
        self.tree.bind('<Double-1>', self.controller.edit_cell)
        
        self.set_default_values()
    
    def create_header(self) -> None:
        header_frame = ttk.Frame(self.root, style="Header.TFrame", padding=(15, 10, 15, 10))
        header_frame.pack(fill=tk.X)
        
        title_label = ttk.Label(
            header_frame, 
            text="Sistema de Cálculo de Custos e Impostos", 
            style="Title.TLabel"
        )
        title_label.pack(side=tk.LEFT)
        
        ttk.Frame(header_frame).pack(side=tk.LEFT, expand=True, fill=tk.X)
        
    def create_input_frame(self) -> None:
        self.input_frame = ttk.LabelFrame(
            self.root, 
            text=" Adicionar Item ", 
            padding=(15, 10),
            style="TLabelframe"
        )
        self.input_frame.pack(fill=tk.X, padx=15, pady=(10, 5), ipady=5)
        
        self.input_frame.grid_columnconfigure(1, weight=1)
        self.input_frame.grid_columnconfigure(3, weight=1)
        self.input_frame.grid_columnconfigure(5, minsize=10)
        
        left_fields = [
            ("Descrição:", "description", 30),
            ("Valor Unitário de Custo (R$):", "unit_cost", 15),
            ("Quantidade:", "quantity", 10),
            ("Margem de Lucro Bruto (%):", "profit_margin", 10),
            ("Estado de Destino:", "state", 5)
        ]
        
        right_fields = [
            ("ICMS (%):", "icms", 10),
            ("PIS (%):", "pis", 10),
            ("COFINS (%):", "cofins", 10),
            ("IRPJ (%):", "irpj", 10),
            ("CSLL (%):", "csll", 10)
        ]
        
        self.input_widgets = {}
        
        for row, (label, name, width) in enumerate(left_fields):
            lbl = ttk.Label(self.input_frame, text=label)
            lbl.grid(row=row, column=0, sticky=tk.W, padx=5, pady=3)
            
            if name == "state":
                widget = ttk.Combobox(self.input_frame, 
                                    values=self.brazilian_states, 
                                    width=width,
                                    state="readonly")
                widget.bind("<<ComboboxSelected>>", self.controller.update_icms_by_state)
            else:
                widget = ttk.Entry(self.input_frame, width=width)
            
            widget.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=3)
            self.input_widgets[name] = widget
        
        for row, (label, name, width) in enumerate(right_fields):
            lbl = ttk.Label(self.input_frame, text=label)
            lbl.grid(row=row, column=3, sticky=tk.W, padx=(15,5), pady=3)
            
            widget = ttk.Entry(self.input_frame, width=width)
            widget.grid(row=row, column=4, sticky=tk.EW, padx=5, pady=3)
            self.input_widgets[name] = widget
        
        button_frame = ttk.Frame(self.input_frame)
        button_frame.grid(row=5, column=0, columnspan=5, pady=(10, 0), sticky=tk.EW)
        
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=0)
        button_frame.grid_columnconfigure(2, weight=0)
        button_frame.grid_columnconfigure(3, weight=1)
        
        add_btn = ttk.Button(
            button_frame, 
            text="Adicionar Item", 
            command=self.controller.add_item,
            style="Accent.TButton"
        )
        add_btn.grid(row=0, column=1, padx=5, sticky=tk.EW)
        
        delete_btn = ttk.Button(
            button_frame, 
            text="Excluir Item", 
            command=self.controller.delete_selected,
            style="TButton"
        )
        delete_btn.grid(row=0, column=2, padx=5, sticky=tk.EW)

    def create_table_frame(self) -> None:
        self.table_frame = ttk.LabelFrame(
            self.root, 
            text=" Itens da Planilha ", 
            padding=(10, 5),
            style="TLabelframe"
        )
        self.table_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        
        container = ttk.Frame(self.table_frame)
        container.pack(fill=tk.BOTH, expand=True)
        
        self.tree = ttk.Treeview(
            container, 
            columns=self.model.columns,
            show="headings",
            selectmode="extended",
            style="Treeview"
        )
        
        column_widths = {
            "Item": 60,
            "Descrição": 250,
            "Valor Unitário de Custo (R$)": 230,
            "Quantidade": 100,
            "Valor Total de Custo (R$)": 210,
            "Margem de Lucro Bruto (%)": 310,
            "Valor Unitário de Venda (R$)": 300,
            "Valor Total de Venda (R$)": 300,
            "Estado de Destino": 200,
            "ICMS (%)": 300,
            "Valor unit. ICMS": 300,
            "Valor Total ICMS (R$)": 300,
            "PIS (%)": 300,
            "Valor unit. PIS": 300,
            "Valor Total PIS (R$)": 300,
            "COFINS (%)": 300,
            "Valor unit. COFINS": 300,
            "Valor Total COFINS (R$)": 300,
            "IRPJ (%)": 300,
            "Valor unit. IRPJ": 300,
            "Valor Total IRPJ (R$)": 300,
            "CSLL (%)": 300,
            "Valor unit. CSLL": 300,
            "Valor Total CSLL (R$)": 300,
            "Valor Total de impostos": 300,
            "Valor Total Unitário": 300,
            "Valor Total": 300,
            "Total Alíquota Impostos (%)": 300
        }
        
        for col in self.model.columns:
            self.tree.heading(col, text=col, anchor=tk.CENTER)
            width = column_widths.get(col, 100)
            self.tree.column(col, width=width, anchor=tk.CENTER, stretch=False)
        
        vsb = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        
        self.tree.tag_configure('evenrow', background=ColorScheme.ROW_EVEN.value)
        self.tree.tag_configure('oddrow', background=ColorScheme.ROW_ODD.value)
        self.tree.tag_configure('total', background=ColorScheme.LIGHT_GRAY.value, font=Fonts.HEADER.value)
        
        def handle_column_resize(event):
            column_id = self.tree.identify_column(event.x)
            col_index = int(column_id[1:]) - 1
            col_name = self.model.columns[col_index]
            
            bbox = self.tree.bbox(self.tree.identify_row(event.y), column_id)
            if bbox:
                new_width = max(20, event.x - bbox[0])
                self.tree.column(col_name, width=new_width)
        
        self.tree.bind("<B1-Motion>", handle_column_resize)
        self.tree.bind("<ButtonRelease-1>", lambda e: self.tree.config(cursor=""))
        self.tree.bind("<Button-1>", self.start_column_resize)
        self.tree.bind('<Double-1>', self.controller.edit_cell)

    def start_column_resize(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "separator":
            self.tree.config(cursor="sb_h_double_arrow")
        else:
            self.tree.config(cursor="")
    
    def create_status_bar(self) -> None:
        self.status_bar = ttk.Label(
            self.root, 
            text="Pronto", 
            style="Status.TLabel",
            anchor=tk.W
        )
        self.status_bar.pack(fill=tk.X, padx=0, pady=0)
    
    def create_menu(self) -> None:
        menubar = tk.Menu(self.root, bg=ColorScheme.BACKGROUND.value, fg=ColorScheme.TEXT.value)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Novo", command=self.controller.new_file, accelerator="Ctrl+N")
        file_menu.add_command(label="Abrir", command=self.controller.open_file, accelerator="Ctrl+O")
        file_menu.add_command(label="Salvar", command=self.controller.save_file, accelerator="Ctrl+S")
        file_menu.add_command(label="Salvar Como", command=self.controller.save_file_as)
        file_menu.add_separator()
        file_menu.add_command(label="Sair", command=self.root.quit, accelerator="Alt+F4")
        menubar.add_cascade(label="Arquivo", menu=file_menu)
        
        action_menu = tk.Menu(menubar, tearoff=0)
        # Removido o item "Calcular Totais"
        action_menu.add_command(label="Limpar Planilha", command=self.controller.clear_spreadsheet)
        action_menu.add_command(label="Excluir Item Selecionado", command=self.controller.delete_selected, accelerator="Del")
        menubar.add_cascade(label="Ações", menu=action_menu)
        
        config_menu = tk.Menu(menubar, tearoff=0)
        config_menu.add_command(label="Editar Tabela de ICMS por Estado", command=self.controller.edit_icms_table)
        menubar.add_cascade(label="Configurações", menu=config_menu)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Sobre", command=self.controller.show_about)
        help_menu.add_command(label="Ajuda", command=self.controller.show_help)
        menubar.add_cascade(label="Informações", menu=help_menu)
        
        self.root.config(menu=menubar)
        
        self.root.bind("<Control-n>", lambda e: self.controller.new_file())
        self.root.bind("<Control-o>", lambda e: self.controller.open_file())
        self.root.bind("<Control-s>", lambda e: self.controller.save_file())
        self.root.bind("<Delete>", lambda e: self.controller.delete_selected())
    
    def set_default_values(self) -> None:
        self.input_widgets['unit_cost'].insert(0, "1,00")
        self.input_widgets['quantity'].insert(0, "1,00")
        self.input_widgets['profit_margin'].insert(0, "0,00")
        self.input_widgets['state'].set("SP")
        self.input_widgets['icms'].insert(0, locale.format_string('%.2f', self.model.tax_config.ICMS, grouping=True))
        self.input_widgets['pis'].insert(0, locale.format_string('%.2f', self.model.tax_config.PIS, grouping=True))
        self.input_widgets['cofins'].insert(0, locale.format_string('%.2f', self.model.tax_config.COFINS, grouping=True))
        self.input_widgets['irpj'].insert(0, locale.format_string('%.2f', self.model.tax_config.IRPJ, grouping=True))
        self.input_widgets['csll'].insert(0, locale.format_string('%.2f', self.model.tax_config.CSLL, grouping=True))
    
    def update_table(self) -> None:
        # Limpa a tabela existente
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Adiciona os itens normais
        for index, row in self.model.data.iterrows():
            formatted_values = []
            for col in self.model.columns:
                value = row[col]
                if pd.isna(value):
                    formatted_values.append("")
                elif isinstance(value, (float, int)):
                    if col == 'Item':
                        formatted_values.append(str(int(value)))
                    else:
                        formatted_values.append(locale.format_string('%.2f', value, grouping=True))
                else:
                    formatted_values.append(str(value))
            
            tag = 'evenrow' if index % 2 == 0 else 'oddrow'
            self.tree.insert("", tk.END, values=formatted_values, iid=str(index), tags=(tag,))
        
        # Adiciona linha de totais automaticamente
        totals = self.model.calculate_totals()
        if totals:
            formatted_totals = []
            for col in self.model.columns:
                value = totals.get(col, "")
                
                # Formatação especial para totais
                if col in ["Valor Total de Custo (R$)", "Valor Total de Venda (R$)", 
                        "Valor Total ICMS (R$)", "Valor Total PIS (R$)", 
                        "Valor Total COFINS (R$)", "Valor Total IRPJ (R$)", 
                        "Valor Total CSLL (R$)", "Valor Total de impostos", 
                        "Valor Total"]:
                    formatted_value = locale.currency(value, grouping=True, symbol=False) if value else ""
                elif isinstance(value, (float, int)):
                    formatted_value = locale.format_string('%.2f', value, grouping=True)
                else:
                    formatted_value = str(value)
                
                formatted_totals.append(formatted_value)
            
        self.tree.insert("", tk.END, values=formatted_totals, tags=('total',))
    
    def update_row_in_table(self, row_index: int, item: str) -> None:
        formatted_values = []
        for col in self.model.columns:
            value = self.model.data.at[row_index, col]
            if pd.isna(value):
                formatted_values.append("")
            elif isinstance(value, (float, int)):
                if col == 'Item':
                    formatted_values.append(str(int(value)))
                else:
                    formatted_values.append(locale.format_string('%.2f', value, grouping=True))
            else:
                formatted_values.append(str(value))
        
        self.tree.item(item, values=formatted_values)

class Controller:
    def __init__(self, root):
        self.model = DataModel()
        self.view = MainView(root, self.model, self)
    
    def add_item(self) -> None:
        try:
            item_data = {
                'Descrição': self.view.input_widgets['description'].get(),
                'Valor Unitário de Custo (R$)': self.view.input_widgets['unit_cost'].get(),
                'Quantidade': self.view.input_widgets['quantity'].get(),
                'Margem de Lucro Bruto (%)': self.view.input_widgets['profit_margin'].get(),
                'Estado de Destino': self.view.input_widgets['state'].get(),
                'ICMS (%)': self.view.input_widgets['icms'].get(),
                'PIS (%)': self.view.input_widgets['pis'].get(),
                'COFINS (%)': self.view.input_widgets['cofins'].get(),
                'IRPJ (%)': self.view.input_widgets['irpj'].get(),
                'CSLL (%)': self.view.input_widgets['csll'].get()
            }
            
            self.model.add_item(item_data)
            self.view.update_table()
            
            self.view.input_widgets['description'].delete(0, tk.END)
            self.view.input_widgets['unit_cost'].delete(0, tk.END)
            self.view.input_widgets['unit_cost'].insert(0, "1,00")
            self.view.input_widgets['quantity'].delete(0, tk.END)
            self.view.input_widgets['quantity'].insert(0, "1,00")
            
            self.view.status_bar.config(text="Item adicionado com sucesso!")
            
        except ValueError as e:
            messagebox.showerror("Erro", f"Não foi possível adicionar o item:\n{str(e)}")
    
    def update_icms_by_state(self, event=None) -> None:
        state = self.view.input_widgets['state'].get()
        if state in self.model.state_icms_table:
            self.view.input_widgets['icms'].delete(0, tk.END)
            self.view.input_widgets['icms'].insert(0, locale.format_string('%.2f', self.model.state_icms_table[state], grouping=True))
    
    def edit_cell(self, event) -> None:
        item = self.view.tree.identify_row(event.y)
        column = self.view.tree.identify_column(event.x)
        
        # Block editing for totals row
        if 'total' in self.view.tree.item(item, 'tags'):
            return
        
        if not item or column == '#0':
            return
        
        col_name = self.model.columns[int(column[1:])-1]
        
        # Prevent editing the 'Item' column
        if col_name == 'Item':
            return
        
        try:
            row_index = int(item)
        except ValueError:
            return  # Invalid item (e.g., totals row)
        
        try:
            current_value = self.model.data.at[row_index, col_name]
        except KeyError:
            return  # Invalid column name
        
        # Create inplace editor
        self._create_inplace_editor(item, column, col_name, row_index, current_value)
        
        
    def _create_inplace_editor(self, item, column, col_name, row_index, current_value):
        # Obtém as coordenadas da célula
        bbox = self.view.tree.bbox(item, column)
        if not bbox:
            return  # Célula não está visível
        
        x, y, width, height = bbox

        # Obtém o container pai da Treeview (onde as scrollbars estão)
        container = self.view.tree.master

        # Cria o editor no container pai
        if col_name == 'Estado de Destino':
            entry = ttk.Combobox(
                container,
                values=self.model.state_icms_table.keys(),
                state="readonly"
            )
        else:
            entry = ttk.Entry(container)

        entry.place(x=x, y=y, width=width, height=height)

        # Configura valor inicial e foco
        if isinstance(current_value, (float, int)) and col_name != 'Item':
            display_value = locale.format_string('%.2f', current_value, grouping=True)
        else:
            display_value = str(current_value)
        
        entry.insert(0, display_value)
        entry.select_range(0, tk.END)
        entry.focus()

        def save_edit():
            try:
                new_value = entry.get()
                
                if col_name in ['Descrição', 'Estado de Destino', 'Item']:
                    pass  # Keep as string
                else:
                    new_value = new_value.replace('.', '').replace(',', '.')
                    new_value = float(new_value)
                
                self.model.update_item(row_index, col_name, new_value)
                
                if col_name == 'Estado de Destino':
                    new_state = new_value
                    icms_rate = self.model.state_icms_table.get(new_state, self.model.tax_config.ICMS)
                    self.model.update_item(row_index, 'ICMS (%)', icms_rate)
                
                self.view.update_table()
                self.view.status_bar.config(text="Item atualizado com sucesso!")
                
            except ValueError as e:
                messagebox.showerror("Erro", f"Valor inválido: {str(e)}")
            finally:
                entry.destroy()

        entry.bind("<FocusOut>", lambda e: save_edit())
        entry.bind("<Return>", lambda e: save_edit())

    def edit_icms_table(self) -> None:
        ICMSEditorWindow(
            self.view.root,
            self.model.state_icms_table,
            self.update_icms_table
        )
    
    def update_icms_table(self, new_table: Dict[str, float]) -> None:
        self.model.state_icms_table = new_table
        self.view.status_bar.config(text="Tabela de ICMS atualizada com sucesso!")
    
    def delete_selected(self) -> None:
        selected_items = self.view.tree.selection()
        if not selected_items:
            messagebox.showwarning("Aviso", "Nenhum item selecionado para excluir")
            return
        
        if messagebox.askyesno("Confirmar", f"Deseja excluir {len(selected_items)} item(ns)?"):
            indices = [int(item) for item in selected_items]
            self.model.delete_items(indices)
            self.view.update_table()  # This triggers automatic total update
            self.view.status_bar.config(text=f"{len(selected_items)} item(ns) excluído(s) com sucesso!")
    
    
    def new_file(self) -> None:
        if not self.model.data.empty:
            if messagebox.askyesno("Novo Arquivo", "Deseja salvar as alterações antes de criar um novo arquivo?"):
                self.save_file()
        
        self.model.clear_data()
        self.view.update_table()
        self.model.current_file = None
        self.view.status_bar.config(text="Novo arquivo criado.")
    
    def open_file(self) -> None:
        filepath = filedialog.askopenfilename(
            title="Abrir Arquivo",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")],
            defaultextension=".xlsx"
        )
        
        if filepath:
            try:
                self.model.load_from_file(filepath)
                self.view.update_table()
                self.view.status_bar.config(text=f"Arquivo carregado: {os.path.basename(filepath)}")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível abrir o arquivo.\nErro: {str(e)}")
    
    def save_file(self) -> None:
        if self.model.current_file:
            self.save_to_file(self.model.current_file)
        else:
            self.save_file_as()
    
    def save_file_as(self) -> None:
        filepath = filedialog.asksaveasfilename(
            title="Salvar Como",
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Arquivos CSV", "*.csv")]
        )
        
        if filepath:
            self.save_to_file(filepath)
    
    def save_to_file(self, filepath: str) -> None:
        try:
            self.model.save_to_file(filepath)
            self.view.status_bar.config(text=f"Arquivo salvo: {os.path.basename(filepath)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível salvar o arquivo.\nErro: {str(e)}")
    
    def clear_spreadsheet(self) -> None:
        if self.model.data.empty:
            return
            
        if messagebox.askyesno("Limpar Planilha", "Tem certeza que deseja limpar toda a planilha?\nTodos os dados serão perdidos."):
            self.model.clear_data()
            self.view.update_table()
            self.view.status_bar.config(text="Planilha limpa.")
            
    
    def show_help(self) -> None:
        help_text = """Sistema de Cálculo de Custos e Impostos

Como usar:
1. Preencha os campos no painel "Adicionar Item"
2. Clique em "Adicionar Item" para incluir na planilha
3. Edite itens clicando duas vezes sobre eles
4. Calcule totais usando o botão ou F9
5. Salve sua planilha quando terminar

Atalhos:
- Ctrl+N: Novo arquivo
- Ctrl+O: Abrir arquivo
- Ctrl+S: Salvar arquivo
- F9: Calcular totais
- Del: Excluir itens selecionados"""
        
        messagebox.showinfo("Ajuda", help_text)
    
    def show_about(self) -> None:
        about_text = """Sistema de Cálculo de Custos e Impostos

Versão: 1.0 BETA
Desenvolvido por: Danilo Araujo Mota
Contato: daniloaraujomota@gmail.com
Data: 2025

MIT License

Copyright (c) 2025 Danilo Araujo Mota

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

"""
        
        messagebox.showinfo("Sobre", about_text)

if __name__ == "__main__":
    root = tk.Tk()
    app = Controller(root)
    root.mainloop()