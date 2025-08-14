import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import ttk as ttk_tree  # Treeview vem daqui

class FaturamentoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Consulta de Faturamento")
        self.root.geometry("1100x600")
        self.root.resizable(False, False)

        self.df = None

        # Título
        ttk.Label(root, text="Consulta de Faturamento Anual de Firmas", font=("Segoe UI", 18, "bold")).pack(pady=15)

        # Botão para escolher arquivo
        self.btn_arquivo = ttk.Button(root, text="Escolher arquivo Excel", command=self.carregar_arquivo, bootstyle=PRIMARY)
        self.btn_arquivo.pack(pady=10)

        # Campo de busca
        frame_busca = ttk.Frame(root)
        frame_busca.pack(pady=10)

        ttk.Label(frame_busca, text="Nome da empresa:").pack(side=tk.LEFT, padx=5)
        self.entry_empresa = ttk.Entry(frame_busca, width=40)
        self.entry_empresa.pack(side=tk.LEFT, padx=5)

        self.btn_buscar = ttk.Button(frame_busca, text="Buscar", command=self.buscar_empresa, bootstyle=SUCCESS)
        self.btn_buscar.pack(side=tk.LEFT, padx=5)

        # Área da tabela
        frame_tabela = ttk.Frame(root)
        frame_tabela.pack(fill="both", expand=True, pady=10)

        self.tree = ttk_tree.Treeview(frame_tabela, columns=[], show="headings", height=15)
        self.tree.pack(side=tk.LEFT, fill="both", expand=True)

        # Barra de rolagem horizontal
        scrollbar_x = ttk.Scrollbar(frame_tabela, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=scrollbar_x.set)
        scrollbar_x.pack(side=tk.BOTTOM, fill="x")

        # Barra de rolagem vertical
        scrollbar_y = ttk.Scrollbar(frame_tabela, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar_y.set)
        scrollbar_y.pack(side=tk.RIGHT, fill="y")

    def carregar_arquivo(self):
        caminho = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
        if caminho:
            try:
                # Lê todas as abas do Excel
                planilhas = pd.read_excel(caminho, sheet_name=None)

                # Concatena todas as abas em um único DataFrame
                self.df = pd.concat(planilhas.values(), ignore_index=True)

                messagebox.showinfo("Sucesso", "Arquivo carregado com sucesso, meu chapa!")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível abrir o arquivo:\n{e}")

    def buscar_empresa(self):
        if self.df is None:
            messagebox.showwarning("Atenção", "Por favor, carregue um arquivo primeiro.\nNão vem colocar o carro na frente dos bois!")
            return
       
        empresa = self.entry_empresa.get().strip().upper()
        if not empresa:
            messagebox.showwarning("Atenção", "Digite o nome da empresa para buscar.")
            return

        # Filtra linhas da empresa
        filtro = self.df[self.df['CLIENTE'].astype(str).str.upper().str.contains(empresa, na=False)]

        if filtro.empty:
            messagebox.showinfo("Resultado", "Nenhuma empresa encontrada.\nPelo menos escreve o nome direito né?! ¬¬'")
            return
       
        # Criar uma visão só para a tabela, sem as duas últimas colunas
        if filtro.shape[1] > 2:
            filtro_view = filtro.iloc[:, :-2]   # remove as 2 últimas colunas da EXIBIÇÃO
        else:
            filtro_view = filtro.copy()

        # Limpar linhas existentes
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Redefinir colunas
        self.tree["columns"] = ()
        colunas_tabela = list(filtro_view.columns)
        self.tree["columns"] = colunas_tabela

        for col in colunas_tabela:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor="center")

        # Inserir dados formatados
        for _, row in filtro_view.iterrows():
            linha = []
            for col in colunas_tabela:
                valor = row[col]
                eh_numero = isinstance(valor, (int, float))
                if eh_numero and col.upper() != "MESES FATURADOS":
                    linha.append(f"R$ {valor:,.2f}")
                else:
                    linha.append(str(valor))
            self.tree.insert("", tk.END, values=linha)

        # Ajustar larguras
        for col in colunas_tabela:
            self.tree.column(col, width=max(100, len(col) * 10))


if __name__ == "__main__":
    app = ttk.Window(themename="minty")
    FaturamentoApp(app)
    app.mainloop()