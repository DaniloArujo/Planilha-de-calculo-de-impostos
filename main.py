import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from typing import Dict, List, Optional, Union
import locale

# Configurar locale para pt_BR (padrão brasileiro)
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')
    except:
        print("Não foi possível configurar o locale para pt_BR. Usando padrão do sistema.")

class TaxCalculator:
    """Classe responsável por calcular os impostos e valores relacionados"""
    @staticmethod
    def calculate_taxes(row: Dict[str, float], tax_rates: Dict[str, float]) -> Dict[str, float]:
        """Calcula todos os impostos e valores derivados"""
        calculations = {}
        
        # Valores básicos
        calculations['Valor Total de Custo (R$)'] = row['Valor Unitário de Custo (R$)'] * row['Quantidade']
        calculations['Valor Unitário de Venda (R$)'] = row['Valor Unitário de Custo (R$)'] * (1 + row['Margem de Lucro Bruto (%)']/100)
        calculations['Valor Total de Venda (R$)'] = calculations['Valor Unitário de Venda (R$)'] * row['Quantidade']
        
        # Cálculos de impostos
        taxes = {
            'ICMS': tax_rates['ICMS (%)'],
            'PIS': tax_rates['PIS (%)'],
            'COFINS': tax_rates['COFINS (%)'],
            'IRPJ': tax_rates['IRPJ (%)'],
            'CSLL': tax_rates['CSLL (%)']
        }
        
        for tax_name, rate in taxes.items():
            calculations[f'Valor unit. {tax_name}'] = calculations['Valor Unitário de Venda (R$)'] * (rate/100)
            calculations[f'Valor Total {tax_name} (R$)'] = calculations[f'Valor unit. {tax_name}'] * row['Quantidade']
        
        # Totais
        total_taxes = sum(calculations[f'Valor Total {tax_name} (R$)'] for tax_name in taxes.keys())
        calculations['Valor Total de impostos'] = total_taxes
        calculations['Valor Total Unitário'] = (calculations['Valor Unitário de Venda (R$)'] + 
                                             sum(calculations[f'Valor unit. {tax_name}'] for tax_name in taxes.keys()))
        calculations['Valor Total'] = calculations['Valor Total de Venda (R$)'] + total_taxes
        calculations['Total Alíquota Impostos (%)'] = sum(taxes.values())
        
        return calculations

class DataModel:
    """Classe responsável por gerenciar os dados da aplicação"""
    def __init__(self):
        self.columns = [
            "Item",  # Nova coluna para número do item
            "Descrição", "Valor Unitário de Custo (R$)", "Quantidade", "Valor Total de Custo (R$)", 
            "Margem de Lucro Bruto (%)", "Valor Unitário de Venda (R$)", "Valor Total de Venda (R$)", 
            "Estado de Destino", "ICMS (%)", "Valor unit. ICMS", "Valor do item ICMS (R$)", 
            "PIS (%)", "Valor unit. PIS", "Valor Total PIS (R$)", "COFINS (%)", 
            "Valor unit. COFINS", "Valor Total COFINS (R$)", "IRPJ (%)", 
            "Valor Unit. IRRPJ", "Valor Total IRPJ (R$)", "CSLL (%)", 
            "Valor Unit. CSLL", "Valor Total CSLL (R$)", "Valor Total de impostos", 
            "Valor Total Unitário", "Valor Total", "Total Alíquota Impostos (%)"
        ]
        self.data = pd.DataFrame(columns=self.columns)
        self.current_file = None
        self.tax_rates = {
            'ICMS (%)': 18.0,
            'PIS (%)': 1.65,
            'COFINS (%)': 7.6,
            'IRPJ (%)': 1.2,
            'CSLL (%)': 1.08
        }
        self.next_item_number = 1  # Contador para números de itens
        
        # Tabela de ICMS por estado
        self.state_icms_table = {
            "AC": 17, "AL": 18, "AP": 18, "AM": 20, "BA": 20.5, 
            "CE": 20, "DF": 20, "ES": 17, "GO": 17, "MA": 22, 
            "MT": 17, "MS": 17, "MG": 18, "PA": 19, "PB": 18, 
            "PR": 19.5, "PE": 20.5, "PI": 21, "RJ": 20, "RN": 18, 
            "RS": 17, "RO": 17.5, "RR": 17, "SC": 17, "SP": 17, 
            "SE": 18, "TO": 18
        }
    
    def add_item(self, item_data: Dict[str, Union[str, float]]) -> None:
        """Adiciona um novo item ao DataFrame"""
        tax_calculations = TaxCalculator.calculate_taxes(item_data, self.tax_rates)
        new_row = {**item_data, **tax_calculations}
        new_row['Item'] = self.next_item_number  # Adiciona número do item
        self.next_item_number += 1  # Incrementa para o próximo item
        
        self.data.loc[len(self.data)] = new_row
    
    def update_item(self, index: int, column: str, new_value: Union[str, float]) -> None:
        """Atualiza um valor específico e recalcula os dependentes"""
        self.data.at[index, column] = new_value
        
        if column in ['Valor Unitário de Custo (R$)', 'Quantidade', 'Margem de Lucro Bruto (%)', 
                     'ICMS (%)', 'PIS (%)', 'COFINS (%)', 'IRPJ (%)', 'CSLL (%)']:
            item_data = self.data.iloc[index].to_dict()
            tax_calculations = TaxCalculator.calculate_taxes(item_data, self.tax_rates)
            
            for col, value in tax_calculations.items():
                self.data.at[index, col] = value
    
    def delete_items(self, indices: List[int]) -> None:
        """Remove itens do DataFrame pelos índices"""
        self.data = self.data.drop(indices).reset_index(drop=True)
        # Atualiza os números dos itens após exclusão
        self.data['Item'] = range(1, len(self.data) + 1)
        self.next_item_number = len(self.data) + 1
    
    def calculate_totals(self) -> Dict[str, float]:
        """Calcula os totais das colunas numéricas"""
        if self.data.empty:
            return {}
        
        numeric_cols = [col for col in self.data.columns if col not in ['Item', 'Descrição', 'Estado de Destino']]
        totals = self.data[numeric_cols].sum().to_dict()
        totals['Descrição'] = "TOTAIS"
        totals['Item'] = ""
        return totals
    
    def clear_data(self) -> None:
        """Limpa todos os dados"""
        self.data = pd.DataFrame(columns=self.columns)
        self.next_item_number = 1  # Reseta o contador de itens
    
    def load_from_file(self, filepath: str) -> None:
        """Carrega dados de um arquivo"""
        if filepath.endswith('.xlsx'):
            self.data = pd.read_excel(filepath)
        elif filepath.endswith('.csv'):
            self.data = pd.read_csv(filepath)
        else:
            raise ValueError("Formato de arquivo não suportado")
        
        # Atualiza o contador de itens
        if not self.data.empty:
            self.next_item_number = self.data['Item'].max() + 1
        
        self.current_file = filepath
    
    def save_to_file(self, filepath: str) -> None:
        """Salva dados em um arquivo"""
        if filepath.endswith('.xlsx'):
            self.data.to_excel(filepath, index=False)
        elif filepath.endswith('.csv'):
            self.data.to_csv(filepath, index=False)
        else:
            raise ValueError("Formato de arquivo não suportado")
        
        self.current_file = filepath

class ICMSEditorWindow:
    """Janela para edição da tabela de ICMS por estado"""
    def __init__(self, parent, state_icms_table: Dict[str, float], update_callback):
        self.parent = parent
        self.state_icms_table = state_icms_table
        self.update_callback = update_callback
        
        self.window = tk.Toplevel(parent)
        self.window.title("Editar Tabela de ICMS por Estado")
        self.window.geometry("500x600")
        
        self.create_widgets()
        self.populate_table()
    
    def create_widgets(self) -> None:
        """Cria os widgets da janela"""
        self.frame = ttk.Frame(self.window)
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Treeview
        self.tree = ttk.Treeview(self.frame, columns=("Estado", "ICMS"), show="headings")
        self.tree.heading("Estado", text="Estado")
        self.tree.heading("ICMS", text="ICMS (%)")
        self.tree.column("Estado", width=100)
        self.tree.column("ICMS", width=100)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(self.frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Layout
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Botão de salvar
        ttk.Button(
            self.frame, 
            text="Salvar Alterações", 
            command=self.save_changes
        ).grid(row=1, column=0, columnspan=2, pady=10)
        
        self.frame.grid_rowconfigure(0, weight=1)
        self.frame.grid_columnconfigure(0, weight=1)
        
        # Bindings
        self.tree.bind('<Double-1>', self.edit_cell)
    
    def populate_table(self) -> None:
        """Preenche a tabela com os valores atuais"""
        for estado, icms in sorted(self.state_icms_table.items()):
            self.tree.insert("", tk.END, values=(estado, icms), iid=estado)
    
    def edit_cell(self, event) -> None:
        """Permite editar o valor do ICMS para um estado"""
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        
        if not item or column != '#2':  # Coluna ICMS
            return
        
        current_value = self.tree.set(item, "ICMS")
        
        # Criar entrada para edição
        x, y, width, height = self.tree.bbox(item, column)
        entry = ttk.Entry(self.tree)
        entry.place(x=x, y=y, width=width, height=height, anchor=tk.NW)
        entry.insert(0, current_value)
        entry.focus()
        
        def save_edit():
            try:
                new_value = float(entry.get().replace(',', '.'))  # Converte vírgula para ponto
                self.tree.set(item, "ICMS", f"{new_value:.2f}".replace('.', ','))  # Formata para padrão BR
                self.state_icms_table[item] = new_value
                entry.destroy()
            except ValueError:
                messagebox.showerror("Erro", "Por favor, insira um valor numérico válido.")
        
        entry.bind("<FocusOut>", lambda e: save_edit())
        entry.bind("<Return>", lambda e: save_edit())
    
    def save_changes(self) -> None:
        """Salva as alterações e fecha a janela"""
        self.update_callback(self.state_icms_table)
        self.window.destroy()

class ItemEditorWindow:
    """Janela para edição de um item específico"""
    def __init__(self, parent, item_data: Dict[str, Union[str, float]], columns: List[str], update_callback):
        self.parent = parent
        self.item_data = item_data
        self.columns = columns
        self.update_callback = update_callback
        self.index = None
        
        self.window = tk.Toplevel(parent)
        self.window.title("Editar Item")
        self.window.transient(parent)
        self.window.grab_set()
        
        self.create_widgets()
    
    def create_widgets(self) -> None:
        """Cria os widgets da janela de edição"""
        self.frame = ttk.Frame(self.window)
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.entries = {}
        for i, col in enumerate(self.columns):
            ttk.Label(self.frame, text=f"{col}:").grid(row=i, column=0, sticky=tk.W, padx=5, pady=2)
            
            entry = ttk.Entry(self.frame)
            entry.grid(row=i, column=1, sticky=tk.W, padx=5, pady=2)
            
            # Formata o valor para exibição (padrão brasileiro)
            if isinstance(self.item_data[col], (float, int)):
                entry.insert(0, locale.format_string('%.2f', self.item_data[col], grouping=True))
            else:
                entry.insert(0, str(self.item_data[col]))
            
            self.entries[col] = entry
        
        # Botão de salvar
        ttk.Button(
            self.frame, 
            text="Salvar", 
            command=self.save_changes
        ).grid(row=len(self.columns), column=0, columnspan=2, pady=10)
    
    def save_changes(self) -> None:
        """Salva as alterações no item"""
        new_values = {}
        for col, entry in self.entries.items():
            try:
                if col in ['Descrição', 'Estado de Destino', 'Item']:
                    new_values[col] = entry.get()
                else:
                    # Converte de padrão brasileiro (vírgula decimal) para float
                    value_str = entry.get().replace('.', '').replace(',', '.')
                    new_values[col] = float(value_str)
            except ValueError:
                messagebox.showerror("Erro", f"Valor inválido para {col}")
                return
        
        self.update_callback(new_values)
        self.window.destroy()

class MainView:
    """Classe principal da interface gráfica"""
    def __init__(self, root, model: DataModel):
        self.root = root
        self.model = model
        self.setup_ui()
    
    def setup_ui(self) -> None:
        """Configura a interface do usuário"""
        self.root.title("Sistema de Cálculo de Custos e Impostos Completo")
        self.root.geometry("1500x800")
        
        # Estados brasileiros
        self.brazilian_states = list(self.model.state_icms_table.keys())
        
        # Criar widgets
        self.create_input_frame()
        self.create_table_frame()
        self.create_status_bar()
        self.create_menu()
        
        # Configurar valores padrão
        self.set_default_values()
    
    def create_input_frame(self) -> None:
        """Cria o frame de entrada de dados"""
        self.input_frame = ttk.LabelFrame(self.root, text="Adicionar Item", padding=10)
        self.input_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Campos de entrada
        fields = [
            ("Descrição:", "description", 30),
            ("Valor Unitário de Custo (R$):", "unit_cost", 10),
            ("Quantidade:", "quantity", 10),
            ("Margem de Lucro Bruto (%):", "profit_margin", 10),
            ("Estado de Destino:", "state", 5),
            ("ICMS (%):", "icms", 10),
            ("PIS (%):", "pis", 10),
            ("COFINS (%):", "cofins", 10),
            ("IRPJ (%):", "irpj", 10),
            ("CSLL (%):", "csll", 10)
        ]
        
        self.input_widgets = {}
        for i, (label, name, width) in enumerate(fields):
            ttk.Label(self.input_frame, text=label).grid(row=i//2, column=(i%2)*2, sticky=tk.W, padx=5, pady=2)
            
            if name == "state":
                widget = ttk.Combobox(self.input_frame, values=self.brazilian_states, width=width)
                widget.bind("<<ComboboxSelected>>", self.update_icms_by_state)
            else:
                widget = ttk.Entry(self.input_frame, width=width)
            
            widget.grid(row=i//2, column=(i%2)*2+1, sticky=tk.W, padx=5, pady=2)
            self.input_widgets[name] = widget
        
        # Botão para adicionar item
        ttk.Button(
            self.input_frame, 
            text="Adicionar Item", 
            command=self.add_item
        ).grid(row=5, column=0, columnspan=4, pady=10)
    
    def create_table_frame(self) -> None:
        """Cria o frame da tabela de dados"""
        self.table_frame = ttk.LabelFrame(self.root, text="Itens da Planilha", padding=10)
        self.table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Container para treeview e scrollbars
        container = ttk.Frame(self.table_frame)
        container.pack(fill=tk.BOTH, expand=True)
        
        # Treeview
        self.tree = ttk.Treeview(
            container, 
            columns=self.model.columns,
            show="headings",
            selectmode="extended"
        )
        
        # Configurar colunas
        column_widths = {
            "Item": 50,  # Coluna para número do item
            "Descrição": 300,
            "Valor Unitário de Custo (R$)": 150,
            "Quantidade": 80,
            "Valor Total de Custo (R$)": 150,
            "Valor Unitário de Venda (R$)": 150,
            "Estado de Destino": 80,
            "ICMS (%)": 80,
            "Valor unit. ICMS": 120,
            "Valor do item ICMS (R$)": 150,
            "PIS (%)": 80,
            "Valor unit. PIS": 120,
            "Valor Total PIS (R$)": 150,
            "COFINS (%)": 80,
            "Valor unit. COFINS": 120,
            "Valor Total COFINS (R$)": 150,
            "IRPJ (%)": 80,
            "Valor Unit. IRRPJ": 120,
            "Valor Total IRPJ (R$)": 150,
            "CSLL (%)": 80,
            "Valor Unit. CSLL": 120,
            "Valor Total CSLL (R$)": 150,
            "Valor Total de impostos": 150,
            "Valor Total Unitário": 150,
            "Valor Total": 150,
            "Total Alíquota Impostos (%)": 150
        }
        
        for col in self.model.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths.get(col, 100), anchor=tk.CENTER)
        
        # Scrollbars
        vsb = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Layout
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        
        # Eventos
        self.tree.bind('<Double-1>', self.edit_cell)
    
    def create_status_bar(self) -> None:
        """Cria a barra de status"""
        self.status_bar = ttk.Label(self.root, text="Pronto", relief=tk.SUNKEN)
        self.status_bar.pack(fill=tk.X, padx=10, pady=5)
    
    def create_menu(self) -> None:
        """Cria o menu principal"""
        menubar = tk.Menu(self.root)
        
        # Menu Arquivo
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Novo", command=self.new_file)
        file_menu.add_command(label="Abrir", command=self.open_file)
        file_menu.add_command(label="Salvar", command=self.save_file)
        file_menu.add_command(label="Salvar Como", command=self.save_file_as)
        file_menu.add_separator()
        file_menu.add_command(label="Sair", command=self.root.quit)
        menubar.add_cascade(label="Arquivo", menu=file_menu)
        
        # Menu Ações
        action_menu = tk.Menu(menubar, tearoff=0)
        action_menu.add_command(label="Calcular Totais", command=self.calculate_totals)
        action_menu.add_command(label="Limpar Planilha", command=self.clear_spreadsheet)
        action_menu.add_command(label="Excluir Item Selecionado", command=self.delete_selected)
        menubar.add_cascade(label="Ações", menu=action_menu)
        
        # Menu Configurações
        config_menu = tk.Menu(menubar, tearoff=0)
        config_menu.add_command(label="Editar Tabela de ICMS por Estado", command=self.edit_icms_table)
        menubar.add_cascade(label="Configurações", menu=config_menu)
        
        self.root.config(menu=menubar)
    
    def set_default_values(self) -> None:
        """Define valores padrão para os campos de entrada"""
        self.input_widgets['unit_cost'].insert(0, "1,00")
        self.input_widgets['quantity'].insert(0, "1,00")
        self.input_widgets['profit_margin'].insert(0, "30,00")
        self.input_widgets['state'].set("SP")
        self.input_widgets['icms'].insert(0, locale.format_string('%.2f', self.model.tax_rates['ICMS (%)'], grouping=True))
        self.input_widgets['pis'].insert(0, locale.format_string('%.2f', self.model.tax_rates['PIS (%)'], grouping=True))
        self.input_widgets['cofins'].insert(0, locale.format_string('%.2f', self.model.tax_rates['COFINS (%)'], grouping=True))
        self.input_widgets['irpj'].insert(0, locale.format_string('%.2f', self.model.tax_rates['IRPJ (%)'], grouping=True))
        self.input_widgets['csll'].insert(0, locale.format_string('%.2f', self.model.tax_rates['CSLL (%)'], grouping=True))
    
    def update_icms_by_state(self, event=None) -> None:
        """Atualiza o ICMS com base no estado selecionado"""
        state = self.input_widgets['state'].get()
        if state in self.model.state_icms_table:
            self.input_widgets['icms'].delete(0, tk.END)
            self.input_widgets['icms'].insert(0, locale.format_string('%.2f', self.model.state_icms_table[state], grouping=True))
    
    def add_item(self) -> None:
        """Adiciona um novo item à planilha"""
        try:
            # Converte os valores de padrão brasileiro para float
            def parse_br_number(value: str) -> float:
                return float(value.replace('.', '').replace(',', '.'))
            
            item_data = {
                'Descrição': self.input_widgets['description'].get(),
                'Valor Unitário de Custo (R$)': parse_br_number(self.input_widgets['unit_cost'].get()),
                'Quantidade': parse_br_number(self.input_widgets['quantity'].get()),
                'Margem de Lucro Bruto (%)': parse_br_number(self.input_widgets['profit_margin'].get()),
                'Estado de Destino': self.input_widgets['state'].get(),
                'ICMS (%)': parse_br_number(self.input_widgets['icms'].get()),
                'PIS (%)': parse_br_number(self.input_widgets['pis'].get()),
                'COFINS (%)': parse_br_number(self.input_widgets['cofins'].get()),
                'IRPJ (%)': parse_br_number(self.input_widgets['irpj'].get()),
                'CSLL (%)': parse_br_number(self.input_widgets['csll'].get())
            }
            
            self.model.add_item(item_data)
            self.update_table()
            
            # Limpar campos e redefinir valores padrão
            self.input_widgets['description'].delete(0, tk.END)
            self.input_widgets['unit_cost'].delete(0, tk.END)
            self.input_widgets['unit_cost'].insert(0, "1,00")
            self.input_widgets['quantity'].delete(0, tk.END)
            self.input_widgets['quantity'].insert(0, "1,00")
            
            self.status_bar.config(text="Item adicionado com sucesso!")
            
        except ValueError as e:
            messagebox.showerror("Erro", f"Por favor, insira valores válidos.\nErro: {str(e)}")
    
    def update_table(self) -> None:
        """Atualiza a exibição da tabela com os dados atuais"""
        # Limpar tabela
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Adicionar dados formatados no padrão brasileiro
        for index, row in self.model.data.iterrows():
            formatted_values = []
            for col in self.model.columns:
                value = row[col]
                if isinstance(value, (float, int)) and not pd.isna(value):
                    if col == 'Item':  # Número do item sem casas decimais
                        formatted_values.append(str(int(value)))
                    else:
                        formatted_values.append(locale.format_string('%.2f', value, grouping=True))
                else:
                    formatted_values.append(str(value))
            
            self.tree.insert("", tk.END, values=formatted_values, iid=str(index))
    
    def edit_cell(self, event) -> None:
        """Abre a janela de edição para uma célula específica"""
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        
        if not item or column == '#0':
            return
        
        # Obter valor atual
        col_name = self.model.columns[int(column[1:])-1]
        row_index = int(item)
        current_value = self.model.data.at[row_index, col_name]
        
        # Criar janela de edição
        ItemEditorWindow(
            self.root,
            {col_name: current_value},
            [col_name],
            lambda new_values: self.save_edit(row_index, col_name, new_values[col_name], item)
        )
    
    def save_edit(self, row_index: int, col_name: str, new_value: Union[str, float], item: str) -> None:
        """Salva as alterações feitas na edição de uma célula"""
        try:
            self.model.update_item(row_index, col_name, new_value)
            self.update_row_in_table(row_index, item)
            self.status_bar.config(text="Item atualizado com sucesso!")
        except ValueError:
            messagebox.showerror("Erro", "Por favor, insira um valor válido.")
    
    def update_row_in_table(self, row_index: int, item: str) -> None:
        """Atualiza uma linha específica na tabela"""
        formatted_values = []
        for col in self.model.columns:
            value = self.model.data.at[row_index, col]
            if isinstance(value, (float, int)) and not pd.isna(value):
                if col == 'Item':  # Número do item sem casas decimais
                    formatted_values.append(str(int(value)))
                else:
                    formatted_values.append(locale.format_string('%.2f', value, grouping=True))
            else:
                formatted_values.append(str(value))
        
        self.tree.item(item, values=formatted_values)
    
    def edit_icms_table(self) -> None:
        """Abre a janela para edição da tabela de ICMS por estado"""
        ICMSEditorWindow(
            self.root,
            self.model.state_icms_table,
            self.update_icms_table
        )
    
    def update_icms_table(self, new_table: Dict[str, float]) -> None:
        """Atualiza a tabela de ICMS com os novos valores"""
        self.model.state_icms_table = new_table
        self.status_bar.config(text="Tabela de ICMS atualizada com sucesso!")
    
    def delete_selected(self) -> None:
        """Exclui os itens selecionados"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Aviso", "Nenhum item selecionado para excluir")
            return
        
        if messagebox.askyesno("Confirmar", f"Deseja excluir {len(selected_items)} item(ns)?"):
            indices = [int(item) for item in selected_items]
            self.model.delete_items(indices)
            self.update_table()
            self.status_bar.config(text=f"{len(selected_items)} item(ns) excluído(s) com sucesso!")
    
    def calculate_totals(self) -> None:
        """Calcula e exibe os totais"""
        if not self.model.data.empty:
            totals = self.model.calculate_totals()
            
            # Formata os totais no padrão brasileiro
            formatted_totals = []
            for col in self.model.columns:
                value = totals.get(col, "")
                if isinstance(value, (float, int)) and col != 'Item':
                    formatted_totals.append(locale.format_string('%.2f', value, grouping=True))
                else:
                    formatted_totals.append(str(value))
            
            # Adicionar linha de totais
            self.tree.insert("", tk.END, values=formatted_totals, tags=('total',))
            self.tree.tag_configure('total', background='#f0f0f0', font=('Helvetica', 10, 'bold'))
            
            self.status_bar.config(text="Totais calculados com sucesso!")
        else:
            messagebox.showwarning("Aviso", "Não há dados para calcular totais.")
    
    def new_file(self) -> None:
        """Cria um novo arquivo"""
        if not self.model.data.empty:
            if messagebox.askyesno("Novo Arquivo", "Deseja salvar as alterações antes de criar um novo arquivo?"):
                self.save_file()
        
        self.model.clear_data()
        self.update_table()
        self.model.current_file = None
        self.status_bar.config(text="Novo arquivo criado.")
    
    def open_file(self) -> None:
        """Abre um arquivo existente"""
        filepath = filedialog.askopenfilename(
            title="Abrir Arquivo",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")]
        )
        
        if filepath:
            try:
                self.model.load_from_file(filepath)
                self.update_table()
                self.status_bar.config(text=f"Arquivo carregado: {os.path.basename(filepath)}")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível abrir o arquivo.\nErro: {str(e)}")
    
    def save_file(self) -> None:
        """Salva o arquivo atual"""
        if self.model.current_file:
            self.save_to_file(self.model.current_file)
        else:
            self.save_file_as()
    
    def save_file_as(self) -> None:
        """Salva o arquivo com um novo nome"""
        filepath = filedialog.asksaveasfilename(
            title="Salvar Como",
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Arquivos CSV", "*.csv")]
        )
        
        if filepath:
            self.save_to_file(filepath)
    
    def save_to_file(self, filepath: str) -> None:
        """Salva os dados no arquivo especificado"""
        try:
            self.model.save_to_file(filepath)
            self.status_bar.config(text=f"Arquivo salvo: {os.path.basename(filepath)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível salvar o arquivo.\nErro: {str(e)}")
    
    def clear_spreadsheet(self) -> None:
        """Limpa toda a planilha"""
        if messagebox.askyesno("Limpar Planilha", "Tem certeza que deseja limpar toda a planilha?"):
            self.model.clear_data()
            self.update_table()
            self.status_bar.config(text="Planilha limpa.")

class Application:
    """Classe principal da aplicação"""
    def __init__(self, root):
        self.model = DataModel()
        self.view = MainView(root, self.model)

if __name__ == "__main__":
    root = tk.Tk()
    app = Application(root)
    root.mainloop()