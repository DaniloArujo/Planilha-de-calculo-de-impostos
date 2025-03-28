import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

class PlanilhaCustosApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Cálculo de Custos e Impostos")
        self.root.geometry("1000x600")
        
        # Dados da planilha
        self.dados = pd.DataFrame(columns=[
            "Item", "Descrição", "Valor Unitário", "Quantidade", 
            "Subtotal", "Imposto (%)", "Valor Imposto", "Custos Adicionais", "Total"
        ])
        
        self.current_file = None
        self.editing = False
        self.edit_entry = None
        
        # Criar interface
        self.criar_widgets()
        
    def criar_widgets(self):
        # Frame de entrada de dados
        frame_entrada = ttk.LabelFrame(self.root, text="Adicionar Item", padding=10)
        frame_entrada.pack(fill=tk.X, padx=10, pady=5)
        
        # Campos de entrada
        ttk.Label(frame_entrada, text="Item:").grid(row=0, column=0, sticky=tk.W)
        self.entry_item = ttk.Entry(frame_entrada, width=30)
        self.entry_item.grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(frame_entrada, text="Descrição:").grid(row=1, column=0, sticky=tk.W)
        self.entry_descricao = ttk.Entry(frame_entrada, width=30)
        self.entry_descricao.grid(row=1, column=1, padx=5, pady=2)
        
        ttk.Label(frame_entrada, text="Valor Unitário:").grid(row=2, column=0, sticky=tk.W)
        self.entry_valor = ttk.Entry(frame_entrada, width=15)
        self.entry_valor.grid(row=2, column=1, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(frame_entrada, text="Quantidade:").grid(row=3, column=0, sticky=tk.W)
        self.entry_quantidade = ttk.Entry(frame_entrada, width=15)
        self.entry_quantidade.grid(row=3, column=1, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(frame_entrada, text="Imposto (%):").grid(row=4, column=0, sticky=tk.W)
        self.entry_imposto = ttk.Entry(frame_entrada, width=15)
        self.entry_imposto.grid(row=4, column=1, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(frame_entrada, text="Custos Adicionais:").grid(row=5, column=0, sticky=tk.W)
        self.entry_custos = ttk.Entry(frame_entrada, width=15)
        self.entry_custos.grid(row=5, column=1, padx=5, pady=2, sticky=tk.W)
        
        # Botão para adicionar item
        btn_adicionar = ttk.Button(frame_entrada, text="Adicionar Item", command=self.adicionar_item)
        btn_adicionar.grid(row=6, column=0, columnspan=2, pady=10)
        
        # Frame da tabela
        frame_tabela = ttk.LabelFrame(self.root, text="Itens da Planilha", padding=10)
        frame_tabela.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Tabela para exibir os itens
        self.tree = ttk.Treeview(frame_tabela, columns=list(self.dados.columns), show="headings")
        
        for col in self.dados.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor=tk.CENTER)
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Configurar edição
        self.tree.bind('<Double-1>', self.editar_celula)
        self.tree.bind('<Return>', self.editar_celula)
        
        # Barra de rolagem
        scrollbar = ttk.Scrollbar(self.tree, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Barra de status
        self.status_bar = ttk.Label(self.root, text="Pronto", relief=tk.SUNKEN)
        self.status_bar.pack(fill=tk.X, padx=10, pady=5)
        
        # Menu
        self.criar_menu()
    
    def editar_celula(self, event):
        # Identificar item e coluna clicados
        rowid = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        
        if not rowid or column == '#0':  # Clicou no cabeçalho ou área vazia
            return
        
        # Obter posição e valor atual
        x, y, width, height = self.tree.bbox(rowid, column)
        value = self.tree.set(rowid, column)
        
        # Criar entrada de edição
        self.edit_entry = ttk.Entry(self.tree, width=width)
        self.edit_entry.place(x=x, y=y, width=width, height=height)
        self.edit_entry.insert(0, value)
        self.edit_entry.select_range(0, tk.END)
        self.edit_entry.focus()
        
        # Configurar bindings para finalizar edição
        self.edit_entry.bind('<FocusOut>', lambda e: self.finalizar_edicao(rowid, column))
        self.edit_entry.bind('<Return>', lambda e: self.finalizar_edicao(rowid, column))
        self.edit_entry.bind('<Escape>', lambda e: self.cancelar_edicao())
        
        self.editing = True

    def finalizar_edicao(self, rowid, column):
        if not self.editing or not self.edit_entry:
            return
        
        # Obter novo valor
        new_value = self.edit_entry.get()
        
        # Atualizar treeview
        column_index = int(column[1:]) - 1
        column_name = self.dados.columns[column_index]
        self.tree.set(rowid, column, new_value)
        
        # Atualizar DataFrame
        row_index = self.tree.index(rowid)
        self.dados.at[row_index, column_name] = self.converter_valor(new_value, column_name)
        
        # Recalcular valores dependentes se necessário
        if column_name in ["Valor Unitário", "Quantidade", "Imposto (%)", "Custos Adicionais"]:
            self.recalcular_linha(row_index)
        
        self.cancelar_edicao()

    def cancelar_edicao(self):
        if self.edit_entry:
            self.edit_entry.destroy()
            self.edit_entry = None
        self.editing = False

    def converter_valor(self, value, column_name):
        # Converter para o tipo apropriado
        if column_name in ["Item", "Descrição"]:
            return str(value)
        try:
            return float(value) if '.' in value else int(value)
        except ValueError:
            return 0 if column_name not in ["Item", "Descrição"] else value

    def recalcular_linha(self, row_index):
        # Recalcular valores da linha
        row = self.dados.iloc[row_index]
        
        try:
            valor_unitario = float(row["Valor Unitário"])
            quantidade = float(row["Quantidade"])
            imposto_perc = float(row["Imposto (%)"])
            custos_adicionais = float(row["Custos Adicionais"])
            
            subtotal = valor_unitario * quantidade
            valor_imposto = subtotal * (imposto_perc / 100)
            total = subtotal + valor_imposto + custos_adicionais
            
            self.dados.at[row_index, "Subtotal"] = subtotal
            self.dados.at[row_index, "Valor Imposto"] = valor_imposto
            self.dados.at[row_index, "Total"] = total
            
            # Atualizar exibição
            valores = [self.dados.at[row_index, col] for col in self.dados.columns]
            self.tree.item(self.tree.get_children()[row_index], values=valores)
            
        except (ValueError, TypeError):
            pass  # Manter valores atuais se houver erro na conversão

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
        menubar.add_cascade(label="Ações", menu=menu_acoes)
        
        # Menu Ajuda
        menu_ajuda = tk.Menu(menubar, tearoff=0)
        menu_ajuda.add_command(label="Sobre", command=self.mostrar_sobre)
        menubar.add_cascade(label="Ajuda", menu=menu_ajuda)
        
        self.root.config(menu=menubar)
    
    def adicionar_item(self):
        try:
            item = self.entry_item.get()
            descricao = self.entry_descricao.get()
            valor_unitario = float(self.entry_valor.get())
            quantidade = float(self.entry_quantidade.get())
            imposto_perc = float(self.entry_imposto.get())
            custos_adicionais = float(self.entry_custos.get() if self.entry_custos.get() else 0)
            
            subtotal = valor_unitario * quantidade
            valor_imposto = subtotal * (imposto_perc / 100)
            total = subtotal + valor_imposto + custos_adicionais
            
            novo_item = {
                "Item": item,
                "Descrição": descricao,
                "Valor Unitário": valor_unitario,
                "Quantidade": quantidade,
                "Subtotal": subtotal,
                "Imposto (%)": imposto_perc,
                "Valor Imposto": valor_imposto,
                "Custos Adicionais": custos_adicionais,
                "Total": total
            }
            
            self.dados.loc[len(self.dados)] = novo_item
            self.atualizar_tabela()
            
            # Limpar campos de entrada
            self.entry_item.delete(0, tk.END)
            self.entry_descricao.delete(0, tk.END)
            self.entry_valor.delete(0, tk.END)
            self.entry_quantidade.delete(0, tk.END)
            self.entry_imposto.delete(0, tk.END)
            self.entry_custos.delete(0, tk.END)
            
            self.status_bar.config(text="Item adicionado com sucesso!")
            
        except ValueError as e:
            messagebox.showerror("Erro", f"Por favor, insira valores válidos.\nErro: {str(e)}")
    
    def atualizar_tabela(self):
        # Limpar tabela
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Preencher com novos dados
        for _, row in self.dados.iterrows():
            valores = [row[col] for col in self.dados.columns]
            self.tree.insert("", tk.END, values=valores)
    
    def calcular_totais(self):
        if not self.dados.empty:
            totais = {
                "Item": "TOTAL",
                "Descrição": "",
                "Valor Unitário": "",
                "Quantidade": self.dados["Quantidade"].sum(),
                "Subtotal": self.dados["Subtotal"].sum(),
                "Imposto (%)": "",
                "Valor Imposto": self.dados["Valor Imposto"].sum(),
                "Custos Adicionais": self.dados["Custos Adicionais"].sum(),
                "Total": self.dados["Total"].sum()
            }
            
            # Adicionar linha de totais
            self.dados.loc[len(self.dados)] = totais
            self.atualizar_tabela()
            
            # Remover a linha de totais para futuras adições
            self.dados = self.dados.iloc[:-1]
            
            self.status_bar.config(text="Totais calculados com sucesso!")
        else:
            messagebox.showwarning("Aviso", "Não há dados para calcular totais.")
    
    def novo_arquivo(self):
        if not self.dados.empty:
            if messagebox.askyesno("Novo Arquivo", "Deseja salvar as alterações antes de criar um novo arquivo?"):
                self.salvar_arquivo()
        
        self.dados = pd.DataFrame(columns=self.dados.columns)
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
            self.dados = pd.DataFrame(columns=self.dados.columns)
            self.atualizar_tabela()
            self.status_bar.config(text="Planilha limpa.")
    
    def mostrar_sobre(self):
        messagebox.showinfo("Sobre", "Sistema de Cálculo de Custos e Impostos\nVersão 1.0\n\nDesenvolvido para gerenciamento de projetos.")

if __name__ == "__main__":
    root = tk.Tk()
    app = PlanilhaCustosApp(root)
    root.mainloop()