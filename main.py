import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

class PlanilhaCustosApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Cálculo de Custos e Impostos Completo")
        self.root.geometry("1500x800")
        
        # Colunas da planilha
        self.colunas = [
            "Descrição", "Valor Unitário de Custo (R$)", "Quantidade", "Valor Total de Custo (R$)", 
            "Margem de Lucro Bruto (%)", "Valor Unitário de Venda (R$)", "Valor Total de Venda (R$)", 
            "Estado de Destino", "ICMS (%)", "Valor unit. ICMS", "Valor do item ICMS (R$)", 
            "PIS (%)", "Valor unit. PIS", "Valor Total PIS (R$)", "COFINS (%)", 
            "Valor unit. COFINS", "Valor Total COFINS (R$)", "IRPJ (%)", 
            "Valor Unit. IRRPJ", "Valor Total IRPJ (R$)", "CSLL (%)", 
            "Valor Unit. CSLL", "Valor Total CSLL (R$)", "Valor Total de impostos", 
            "Valor Total Unitário", "Valor Total", "Total Alíquota Impostos (%)"
        ]
        
        self.dados = pd.DataFrame(columns=self.colunas)
        self.current_file = None
        
        # Estados brasileiros para combobox
        self.estados_brasileiros = [
            "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", 
            "MA", "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", 
            "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO"
        ]
        
        # Tabela de ICMS por estado (valores padrão)
        self.tabela_icms_estados = {
            "AC": 17, "AL": 17, "AP": 17, "AM": 17, "BA": 17, 
            "CE": 17, "DF": 17, "ES": 17, "GO": 17, "MA": 17, 
            "MT": 17, "MS": 17, "MG": 18, "PA": 17, "PB": 17, 
            "PR": 18, "PE": 17, "PI": 17, "RJ": 20, "RN": 17, 
            "RS": 18, "RO": 17, "RR": 17, "SC": 17, "SP": 18, 
            "SE": 17, "TO": 17
        }
        
        # Criar interface
        self.criar_widgets()
        
    def criar_widgets(self):
        # Frame de entrada de dados
        frame_entrada = ttk.LabelFrame(self.root, text="Adicionar Item", padding=10)
        frame_entrada.pack(fill=tk.X, padx=10, pady=5)
        
        # Campos de entrada
        campos = [
            ("Descrição:", "entry_descricao", 30),
            ("Valor Unitário de Custo (R$):", "entry_valor_custo", 10),
            ("Quantidade:", "entry_quantidade", 10),
            ("Margem de Lucro Bruto (%):", "entry_margem", 10),
            ("Estado de Destino:", "combo_estado", 5),
            ("ICMS (%):", "entry_icms", 10),
            ("PIS (%):", "entry_pis", 10),
            ("COFINS (%):", "entry_cofins", 10),
            ("IRPJ (%):", "entry_irpj", 10),
            ("CSLL (%):", "entry_csll", 10)
        ]
        
        for i, (label, var_name, width) in enumerate(campos):
            ttk.Label(frame_entrada, text=label).grid(row=i//2, column=(i%2)*2, sticky=tk.W, padx=5, pady=2)
            
            if var_name.startswith("entry"):
                entry = ttk.Entry(frame_entrada, width=width)
                entry.grid(row=i//2, column=(i%2)*2+1, sticky=tk.W, padx=5, pady=2)
                setattr(self, var_name, entry)
            elif var_name.startswith("combo"):
                combo = ttk.Combobox(frame_entrada, values=self.estados_brasileiros, width=width)
                combo.grid(row=i//2, column=(i%2)*2+1, sticky=tk.W, padx=5, pady=2)
                combo.set("SP")  # Valor padrão
                setattr(self, var_name, combo)
                combo.bind("<<ComboboxSelected>>", self.atualizar_icms_por_estado)
        
        # Valores padrão
        self.entry_valor_custo.insert(0, "1")
        self.entry_quantidade.insert(0, "1")
        self.entry_margem.insert(0, "30")
        self.entry_icms.insert(0, "18")
        self.entry_pis.insert(0, "1.65")
        self.entry_cofins.insert(0, "7.6")
        self.entry_irpj.insert(0, "1.2")
        self.entry_csll.insert(0, "1.08")
        
        # Botão para adicionar item
        btn_adicionar = ttk.Button(frame_entrada, text="Adicionar Item", command=self.adicionar_item)
        btn_adicionar.grid(row=5, column=0, columnspan=4, pady=10)
        
        # Frame da tabela
        frame_tabela = ttk.LabelFrame(self.root, text="Itens da Planilha", padding=10)
        frame_tabela.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Treeview com barra de rolagem
        container = ttk.Frame(frame_tabela)
        container.pack(fill=tk.BOTH, expand=True)
        
        # Treeview
        self.tree = ttk.Treeview(
            container, 
            columns=self.colunas,
            show="headings",
            selectmode="extended"
        )
        
        # Configurar colunas 
        colunas_principais = {
            "Descrição": 300,
            "Valor Unitário de Custo (R$)": 300,
            "Quantidade": 300,
            "Valor Total de Custo (R$)": 300,
            "Valor Unitário de Venda (R$)": 120,
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
        
        for col in self.colunas:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=colunas_principais.get(col, 100), anchor=tk.CENTER)
        
        # Barra de rolagem
        vsb = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Layout
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        
        # Barra de status
        self.status_bar = ttk.Label(self.root, text="Pronto", relief=tk.SUNKEN)
        self.status_bar.pack(fill=tk.X, padx=10, pady=5)
        
        # Menu
        self.criar_menu()
        
        # Eventos
        self.tree.bind('<Double-1>', self.editar_celula)

    def atualizar_icms_por_estado(self, event=None):
        """Atualiza o campo ICMS com o valor correspondente ao estado selecionado"""
        estado = self.combo_estado.get()
        if estado in self.tabela_icms_estados:
            self.entry_icms.delete(0, tk.END)
            self.entry_icms.insert(0, str(self.tabela_icms_estados[estado]))

    def adicionar_item(self):
        try:
            # Obter valores dos campos
            descricao = self.entry_descricao.get()
            valor_custo = float(self.entry_valor_custo.get())
            quantidade = float(self.entry_quantidade.get())
            margem = float(self.entry_margem.get())
            estado = self.combo_estado.get()
            icms = float(self.entry_icms.get())
            pis = float(self.entry_pis.get())
            cofins = float(self.entry_cofins.get())
            irpj = float(self.entry_irpj.get())
            csll = float(self.entry_csll.get())
            
            # Cálculos básicos
            valor_total_custo = valor_custo * quantidade
            valor_unitario_venda = valor_custo * (1 + margem/100)
            valor_total_venda = valor_unitario_venda * quantidade
            
            # Cálculos de impostos
            valor_unit_icms = valor_unitario_venda * (icms/100)
            valor_item_icms = valor_unit_icms * quantidade
            
            valor_unit_pis = valor_unitario_venda * (pis/100)
            valor_total_pis = valor_unit_pis * quantidade
            
            valor_unit_cofins = valor_unitario_venda * (cofins/100)
            valor_total_cofins = valor_unit_cofins * quantidade
            
            valor_unit_irpj = valor_unitario_venda * (irpj/100)
            valor_total_irpj = valor_unit_irpj * quantidade
            
            valor_unit_csll = valor_unitario_venda * (csll/100)
            valor_total_csll = valor_unit_csll * quantidade
            
            # Totais
            valor_total_impostos = (valor_item_icms + valor_total_pis + valor_total_cofins + 
                                  valor_total_irpj + valor_total_csll)
            
            valor_total_unitario = valor_unitario_venda + valor_unit_icms + valor_unit_pis + valor_unit_cofins + valor_unit_irpj + valor_unit_csll
            valor_total = valor_total_venda + valor_total_impostos
            
            total_aliquota = icms + pis + cofins + irpj + csll
            
            # Criar novo item
            novo_item = {
                "Descrição": descricao,
                "Valor Unitário de Custo (R$)": valor_custo,
                "Quantidade": quantidade,
                "Valor Total de Custo (R$)": valor_total_custo,
                "Margem de Lucro Bruto (%)": margem,
                "Valor Unitário de Venda (R$)": valor_unitario_venda,
                "Valor Total de Venda (R$)": valor_total_venda,
                "Estado de Destino": estado,
                "ICMS (%)": icms,
                "Valor unit. ICMS": valor_unit_icms,
                "Valor do item ICMS (R$)": valor_item_icms,
                "PIS (%)": pis,
                "Valor unit. PIS": valor_unit_pis,
                "Valor Total PIS (R$)": valor_total_pis,
                "COFINS (%)": cofins,
                "Valor unit. COFINS": valor_unit_cofins,
                "Valor Total COFINS (R$)": valor_total_cofins,
                "IRPJ (%)": irpj,
                "Valor Unit. IRRPJ": valor_unit_irpj,
                "Valor Total IRPJ (R$)": valor_total_irpj,
                "CSLL (%)": csll,
                "Valor Unit. CSLL": valor_unit_csll,
                "Valor Total CSLL (R$)": valor_total_csll,
                "Valor Total de impostos": valor_total_impostos,
                "Valor Total Unitário": valor_total_unitario,
                "Valor Total": valor_total,
                "Total Alíquota Impostos (%)": total_aliquota
            }
            
            # Adicionar ao DataFrame
            self.dados.loc[len(self.dados)] = novo_item
            self.atualizar_tabela()
            
            # Limpar campos
            self.entry_descricao.delete(0, tk.END)
            self.entry_valor_custo.delete(0, tk.END)
            self.entry_valor_custo.insert(0, "1")  
            self.entry_quantidade.delete(0, tk.END)
            self.entry_quantidade.insert(0, "1") 
        

            self.status_bar.config(text="Item adicionado com sucesso!")
            
        except ValueError as e:
            messagebox.showerror("Erro", f"Por favor, insira valores válidos.\nErro: {str(e)}")

    def atualizar_tabela(self):
        # Limpar tabela
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Adicionar dados formatados
        for index, row in self.dados.iterrows():
            valores_formatados = [
                f"{row[col]:.2f}" if isinstance(row[col], (float, int)) and not pd.isna(row[col]) else str(row[col])
                for col in self.colunas
            ]
            self.tree.insert("", tk.END, values=valores_formatados, iid=str(index))

    def editar_celula(self, event):
        # Identificar item e coluna clicados
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        
        if not item or column == '#0':
            return
        
        # Obter valor atual
        col_name = self.colunas[int(column[1:])-1]
        row_index = int(item)
        current_value = self.dados.at[row_index, col_name]
        
        # Criar janela de edição
        self.janela_edicao = tk.Toplevel(self.root)
        self.janela_edicao.title(f"Editar {col_name}")
        self.janela_edicao.transient(self.root)
        self.janela_edicao.grab_set()
        
        tk.Label(self.janela_edicao, text=col_name).pack(padx=10, pady=5)
        
        self.entry_edicao = ttk.Entry(self.janela_edicao)
        self.entry_edicao.insert(0, str(current_value))
        self.entry_edicao.pack(padx=10, pady=5)
        self.entry_edicao.focus()
        
        btn_salvar = ttk.Button(
            self.janela_edicao, 
            text="Salvar", 
            command=lambda: self.salvar_edicao(row_index, col_name, item)
        )
        btn_salvar.pack(pady=10)

    def salvar_edicao(self, row_index, col_name, item):
        try:
            novo_valor = self.entry_edicao.get()
            
            # Converter para o tipo apropriado
            if col_name in ["Descrição", "Estado de Destino"]:
                self.dados.at[row_index, col_name] = novo_valor
            else:
                self.dados.at[row_index, col_name] = float(novo_valor)
            
            # Fechar janela de edição
            self.janela_edicao.destroy()
            
            # Recalcular TODOS os valores dependentes
            self.recalcular_todos_os_valores(row_index)
            
            # Atualizar exibição
            self.atualizar_linha_na_tabela(row_index, item)
            
            self.status_bar.config(text="Item atualizado com sucesso!")
            
        except ValueError:
            messagebox.showerror("Erro", "Por favor, insira um valor válido.")

    def recalcular_todos_os_valores(self, row_index):
        """Recalcula todos os valores dependentes após uma edição"""
        row = self.dados.iloc[row_index]
        
        try:
            # Obter valores básicos (alguns podem ter sido editados)
            valor_custo = row["Valor Unitário de Custo (R$)"]
            quantidade = row["Quantidade"]
            margem = row["Margem de Lucro Bruto (%)"]
            icms = row["ICMS (%)"]
            pis = row["PIS (%)"]
            cofins = row["COFINS (%)"]
            irpj = row["IRPJ (%)"]
            csll = row["CSLL (%)"]
            
            # Recalcular valores dependentes
            valor_total_custo = valor_custo * quantidade
            valor_unitario_venda = valor_custo * (1 + margem/100)
            valor_total_venda = valor_unitario_venda * quantidade
            
            # Cálculos de impostos
            valor_unit_icms = valor_unitario_venda * (icms/100)
            valor_item_icms = valor_unit_icms * quantidade
            
            valor_unit_pis = valor_unitario_venda * (pis/100)
            valor_total_pis = valor_unit_pis * quantidade
            
            valor_unit_cofins = valor_unitario_venda * (cofins/100)
            valor_total_cofins = valor_unit_cofins * quantidade
            
            valor_unit_irpj = valor_unitario_venda * (irpj/100)
            valor_total_irpj = valor_unit_irpj * quantidade
            
            valor_unit_csll = valor_unitario_venda * (csll/100)
            valor_total_csll = valor_unit_csll * quantidade
            
            # Atualizar DataFrame
            self.dados.at[row_index, "Valor Total de Custo (R$)"] = valor_total_custo
            self.dados.at[row_index, "Valor Unitário de Venda (R$)"] = valor_unitario_venda
            self.dados.at[row_index, "Valor Total de Venda (R$)"] = valor_total_venda
            self.dados.at[row_index, "Valor unit. ICMS"] = valor_unit_icms
            self.dados.at[row_index, "Valor do item ICMS (R$)"] = valor_item_icms
            self.dados.at[row_index, "Valor unit. PIS"] = valor_unit_pis
            self.dados.at[row_index, "Valor Total PIS (R$)"] = valor_total_pis
            self.dados.at[row_index, "Valor unit. COFINS"] = valor_unit_cofins
            self.dados.at[row_index, "Valor Total COFINS (R$)"] = valor_total_cofins
            self.dados.at[row_index, "Valor Unit. IRRPJ"] = valor_unit_irpj
            self.dados.at[row_index, "Valor Total IRPJ (R$)"] = valor_total_irpj
            self.dados.at[row_index, "Valor Unit. CSLL"] = valor_unit_csll
            self.dados.at[row_index, "Valor Total CSLL (R$)"] = valor_total_csll
            
            # Totais
            valor_total_impostos = (valor_item_icms + valor_total_pis + valor_total_cofins + 
                                  valor_total_irpj + valor_total_csll)
            
            valor_total_unitario = valor_unitario_venda + valor_unit_icms + valor_unit_pis + valor_unit_cofins + valor_unit_irpj + valor_unit_csll
            valor_total = valor_total_venda + valor_total_impostos
            
            total_aliquota = icms + pis + cofins + irpj + csll
            
            self.dados.at[row_index, "Valor Total de impostos"] = valor_total_impostos
            self.dados.at[row_index, "Valor Total Unitário"] = valor_total_unitario
            self.dados.at[row_index, "Valor Total"] = valor_total
            self.dados.at[row_index, "Total Alíquota Impostos (%)"] = total_aliquota
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao recalcular valores: {str(e)}")

    def atualizar_linha_na_tabela(self, row_index, item):
        """Atualiza apenas a linha modificada na tabela"""
        valores_formatados = [
            f"{self.dados.at[row_index, col]:.2f}" if isinstance(self.dados.at[row_index, col], (float, int)) and not pd.isna(self.dados.at[row_index, col]) 
            else str(self.dados.at[row_index, col])
            for col in self.colunas
        ]
        self.tree.item(item, values=valores_formatados)

    def criar_menu(self):
        menubar = tk.Menu(self.root)
        
        # Menu Arquivo
        menu_arquivo = tk.Menu(menubar, tearoff=0)
        menu_arquivo.add_command(label="Novo", command=self.novo_arquivo)
        menu_arquivo.add_command(label="Abrir", command=self.abrir_arquivo)
        menu_arquivo.add_command(label="Salvar", command=self.salvar_arquivo)
        menu_arquivo.add_command(label="Salvar Como", command=self.salvar_arquivo_como)
        menu_arquivo.add_separator()
        menu_arquivo.add_command(label="Sair", command=self.root.quit)
        menubar.add_cascade(label="Arquivo", menu=menu_arquivo)
        
        # Menu Ações
        menu_acoes = tk.Menu(menubar, tearoff=0)
        menu_acoes.add_command(label="Calcular Totais", command=self.calcular_totais)
        menu_acoes.add_command(label="Limpar Planilha", command=self.limpar_planilha)
        menu_acoes.add_command(label="Excluir Item Selecionado", command=self.excluir_selecionado)
        menubar.add_cascade(label="Ações", menu=menu_acoes)
        
        # Menu Configurações
        menu_config = tk.Menu(menubar, tearoff=0)
        menu_config.add_command(label="Editar Tabela de ICMS por Estado", command=self.editar_tabela_icms)
        menubar.add_cascade(label="Configurações", menu=menu_config)
        
        self.root.config(menu=menubar)
    
    def editar_tabela_icms(self):
        """Abre uma janela para editar os valores de ICMS por estado"""
        janela_icms = tk.Toplevel(self.root)
        janela_icms.title("Editar Tabela de ICMS por Estado")
        janela_icms.geometry("500x600")
        
        # Frame principal
        frame_principal = ttk.Frame(janela_icms)
        frame_principal.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Treeview para exibir e editar os valores
        tree = ttk.Treeview(frame_principal, columns=("Estado", "ICMS"), show="headings")
        tree.heading("Estado", text="Estado")
        tree.heading("ICMS", text="ICMS (%)")
        tree.column("Estado", width=100)
        tree.column("ICMS", width=100)
        
        # Adicionar barra de rolagem
        scrollbar = ttk.Scrollbar(frame_principal, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Layout
        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        frame_principal.grid_rowconfigure(0, weight=1)
        frame_principal.grid_columnconfigure(0, weight=1)
        
        # Preencher a treeview com os valores atuais
        for estado, icms in sorted(self.tabela_icms_estados.items()):
            tree.insert("", tk.END, values=(estado, icms), iid=estado)
        
        # Função para editar célula
        def editar_celula(event):
            item = tree.identify_row(event.y)
            column = tree.identify_column(event.x)
            
            if not item or column == '#0':
                return
            
            # Obter valor atual
            if column == '#1':  # Coluna Estado (não editável)
                return
            else:  # Coluna ICMS
                current_value = tree.set(item, "ICMS")
                
                # Criar entrada para edição
                x, y, width, height = tree.bbox(item, column)
                entry = ttk.Entry(tree)
                entry.place(x=x, y=y, width=width, height=height, anchor=tk.NW)
                entry.insert(0, current_value)
                entry.focus()
                
                def salvar_edicao():
                    try:
                        novo_valor = float(entry.get())
                        tree.set(item, "ICMS", f"{novo_valor:.2f}")
                        self.tabela_icms_estados[item] = novo_valor
                        entry.destroy()
                    except ValueError:
                        messagebox.showerror("Erro", "Por favor, insira um valor numérico válido.")
                
                entry.bind("<FocusOut>", lambda e: salvar_edicao())
                entry.bind("<Return>", lambda e: salvar_edicao())
        
        tree.bind('<Double-1>', editar_celula)
        
        # Botão para salvar
        btn_salvar = ttk.Button(
            frame_principal, 
            text="Salvar Alterações", 
            command=janela_icms.destroy
        )
        btn_salvar.grid(row=1, column=0, columnspan=2, pady=10)
    
    def excluir_selecionado(self):
        selecionados = self.tree.selection()
        if not selecionados:
            messagebox.showwarning("Aviso", "Nenhum item selecionado para excluir")
            return
        
        if messagebox.askyesno("Confirmar", f"Deseja excluir {len(selecionados)} item(ns)?"):
            indices = [int(item) for item in selecionados]
            self.dados = self.dados.drop(indices)
            self.atualizar_tabela()
            self.status_bar.config(text=f"{len(selecionados)} item(ns) excluído(s) com sucesso!")

    def calcular_totais(self):
        if not self.dados.empty:
            # Criar linha de totais
            totais = {
                "Descrição": "TOTAIS",
                "Quantidade": self.dados["Quantidade"].sum(),
                "Valor Total de Custo (R$)": self.dados["Valor Total de Custo (R$)"].sum(),
                "Valor Total de Venda (R$)": self.dados["Valor Total de Venda (R$)"].sum(),
                "Valor do item ICMS (R$)": self.dados["Valor do item ICMS (R$)"].sum(),
                "Valor Total PIS (R$)": self.dados["Valor Total PIS (R$)"].sum(),
                "Valor Total COFINS (R$)": self.dados["Valor Total COFINS (R$)"].sum(),
                "Valor Total IRPJ (R$)": self.dados["Valor Total IRPJ (R$)"].sum(),
                "Valor Total CSLL (R$)": self.dados["Valor Total CSLL (R$)"].sum(),
                "Valor Total de impostos": self.dados["Valor Total de impostos"].sum(),
                "Valor Total": self.dados["Valor Total"].sum()
            }
            
            # Adicionar linha de totais
            self.dados.loc["TOTAL"] = totais
            self.atualizar_tabela()
            
            # Remover a linha de totais para futuras adições
            self.dados = self.dados.drop("TOTAL", errors="ignore")
            
            self.status_bar.config(text="Totais calculados com sucesso!")
        else:
            messagebox.showwarning("Aviso", "Não há dados para calcular totais.")
    
    def novo_arquivo(self):
        if not self.dados.empty:
            if messagebox.askyesno("Novo Arquivo", "Deseja salvar as alterações antes de criar um novo arquivo?"):
                self.salvar_arquivo()
        
        self.dados = pd.DataFrame(columns=self.colunas)
        self.atualizar_tabela()
        self.current_file = None
        self.status_bar.config(text="Novo arquivo criado.")
    
    def abrir_arquivo(self):
        filepath = filedialog.askopenfilename(
            title="Abrir Arquivo",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")]
        )
        
        if filepath:
            try:
                if filepath.endswith('.xlsx'):
                    self.dados = pd.read_excel(filepath)
                elif filepath.endswith('.csv'):
                    self.dados = pd.read_csv(filepath)
                else:
                    messagebox.showerror("Erro", "Formato de arquivo não suportado.")
                    return
                
                # Garantir que todas as colunas necessárias existam
                for col in self.colunas:
                    if col not in self.dados.columns:
                        self.dados[col] = 0  # ou pd.NA conforme necessário
                
                self.atualizar_tabela()
                self.current_file = filepath
                self.status_bar.config(text=f"Arquivo carregado: {os.path.basename(filepath)}")
                
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível abrir o arquivo.\nErro: {str(e)}")
    
    def salvar_arquivo(self):
        if self.current_file:
            self.salvar_como(self.current_file)
        else:
            self.salvar_arquivo_como()
    
    def salvar_arquivo_como(self):
        filepath = filedialog.asksaveasfilename(
            title="Salvar Como",
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Arquivos CSV", "*.csv")]
        )
        
        if filepath:
            self.salvar_como(filepath)
    
    def salvar_como(self, filepath):
        try:
            if filepath.endswith('.xlsx'):
                self.dados.to_excel(filepath, index=False)
            elif filepath.endswith('.csv'):
                self.dados.to_csv(filepath, index=False)
            
            self.current_file = filepath
            self.status_bar.config(text=f"Arquivo salvo: {os.path.basename(filepath)}")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível salvar o arquivo.\nErro: {str(e)}")
    
    def limpar_planilha(self):
        if messagebox.askyesno("Limpar Planilha", "Tem certeza que deseja limpar toda a planilha?"):
            self.dados = pd.DataFrame(columns=self.colunas)
            self.atualizar_tabela()
            self.status_bar.config(text="Planilha limpa.")

if __name__ == "__main__":
    root = tk.Tk()
    app = PlanilhaCustosApp(root)
    root.mainloop()