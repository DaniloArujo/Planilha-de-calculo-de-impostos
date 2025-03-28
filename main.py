import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

class PlanilhaCustosApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Cálculo de Custos e Impostos Completo")
        self.root.geometry("1500x800")
        
        # Dados da planilha com todas as colunas solicitadas
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
        
        # Valores padrão
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
        
        # Configurar colunas (apenas as principais terão largura fixa)
        colunas_principais = {
            "Descrição": 200,
            "Valor Unitário de Custo (R$)": 120,
            "Quantidade": 80,
            "Valor Total de Custo (R$)": 120,
            "Valor Unitário de Venda (R$)": 120,
            "Estado de Destino": 80
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
            self.tree.insert("", tk.END, values=valores_formatados)

    def editar_celula(self, event):
        # Identificar item e coluna clicados
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        
        if not item or column == '#0':
            return
        
        # Obter valor atual
        col_name = self.colunas[int(column[1:])-1]
        row_index = int(self.tree.index(item))
        current_value = self.dados.at[row_index, col_name]
        
        # Criar janela de edição
        self.janela_edicao = tk.Toplevel(self.root)
        self.janela_edicao.title(f"Editar {col_name}")
        
        tk.Label(self.janela_edicao, text=col_name).pack(padx=10, pady=5)
        
        self.entry_edicao = ttk.Entry(self.janela_edicao)
        self.entry_edicao.insert(0, str(current_value))
        self.entry_edicao.pack(padx=10, pady=5)
        self.entry_edicao.focus()
        
        btn_salvar = ttk.Button(
            self.janela_edicao, 
            text="Salvar", 
            command=lambda: self.salvar_edicao(row_index, col_name)
        )
        btn_salvar.pack(pady=10)

    def salvar_edicao(self, row_index, col_name):
        try:
            novo_valor = self.entry_edicao.get()
            
            # Converter para o tipo apropriado
            if col_name in ["Descrição", "Estado de Destino"]:
                self.dados.at[row_index, col_name] = novo_valor
            else:
                self.dados.at[row_index, col_name] = float(novo_valor)
            
            # Recalcular todos os valores dependentes
            self.recalcular_linha(row_index)
            
            self.janela_edicao.destroy()
            self.atualizar_tabela()
            
        except ValueError:
            messagebox.showerror("Erro", "Por favor, insira um valor válido.")

    def recalcular_linha(self, row_index):
        # Obter a linha atualizada
        row = self.dados.iloc[row_index]
        
        try:
            # Recalcular todos os valores
            valor_custo = row["Valor Unitário de Custo (R$)"]
            quantidade = row["Quantidade"]
            margem = row["Margem de Lucro Bruto (%)"]
            icms = row["ICMS (%)"]
            pis = row["PIS (%)"]
            cofins = row["COFINS (%)"]
            irpj = row["IRPJ (%)"]
            csll = row["CSLL (%)"]
            
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
            
            # Atualizar valores no DataFrame
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
        
        self.root.config(menu=menubar)
    
    def excluir_selecionado(self):
        selecionados = self.tree.selection()
        if not selecionados:
            messagebox.showwarning("Aviso", "Nenhum item selecionado para excluir")
            return
        
        if messagebox.askyesno("Confirmar", f"Deseja excluir {len(selecionados)} item(ns)?"):
            indices = [int(self.tree.index(item)) for item in selecionados]
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