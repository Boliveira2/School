import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import datetime
import re

from src.relatorio import gerar_relatorioMensal
from src.transferencias import gerar_transferencias
from src.comunicacao import gerar_emails

class GestorReceitasGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestor de Receitas - Associação de Pais")

        self.caminhos = {
            "alunos": "",
            "precos": "",
            "entradas": "",
            "transferencias": "",
            "relatorio_mensal": "",
            "pasta_emails": "",
            "mes": "",
            # ano_letivo removido
        }

        self.criar_widgets()
        self.mostrar_debug = True  # Controla se mostramos debug

    def criar_widgets(self):
        frame = ttk.Frame(self.root, padding=20)
        frame.grid(row=0, column=0, sticky="nsew")

        self.campos = {}
        campos_labels = [
            ("alunos", "Ficheiro alunos.csv"),
            ("precos", "Ficheiro precos.csv"),
            ("entradas", "Ficheiro entradas.xlsx"),
            ("transferencias", "Ficheiro transferencias.xlsx"),
            ("relatorio_mensal", "Ficheiro relatorioMensal.xlsx"),
            ("pasta_emails", "Pasta para emails"),
            ("mes", "Mês"),
            # ano_letivo removido
        ]

        for i, (chave, label_texto) in enumerate(campos_labels):
            ttk.Label(frame, text=label_texto).grid(row=i, column=0, sticky="w", pady=5)

            if chave == "mes":
                meses = ["setembro", "outubro", "novembro", "dezembro",
                         "janeiro", "fevereiro", "março", "abril",
                         "maio", "junho", "julho", "agosto"]
                combo = ttk.Combobox(frame, values=meses, state="readonly")
                combo.grid(row=i, column=1, pady=5, padx=(0, 10))
                self.campos[chave] = combo
            else:
                entrada = ttk.Entry(frame, width=50)
                entrada.grid(row=i, column=1, pady=5, padx=(0, 10))
                self.campos[chave] = entrada

            if chave in ("pasta_emails", "alunos", "precos", "entradas", "transferencias", "relatorio_mensal"):
                ttk.Button(frame, text="Procurar", command=lambda c=chave: self.procurar_ficheiro_ou_pasta(c)).grid(row=i, column=2, pady=5)

        ttk.Button(frame, text="Gerar relatório início ano letivo", command=self.confirmar_relatorio_ano_letivo).grid(row=len(campos_labels), column=0, columnspan=3, pady=10, sticky="ew")
        ttk.Button(frame, text="Gerar Relatório Mensal", command=self.acao_relatorio).grid(row=len(campos_labels)+1, column=0, columnspan=3, pady=10, sticky="ew")
        ttk.Button(frame, text="Gerar Transferências", command=self.acao_transferencias).grid(row=len(campos_labels)+2, column=0, columnspan=3, pady=10, sticky="ew")
        ttk.Button(frame, text="Gerar Emails", command=self.acao_emails).grid(row=len(campos_labels)+3, column=0, columnspan=3, pady=10, sticky="ew")
        ttk.Button(frame, text="Sair", command=self.root.quit).grid(row=len(campos_labels)+4, column=0, columnspan=3, pady=20, sticky="ew")

        # Caixa de texto para debug
        self.debug_text = tk.Text(frame, height=10, width=80)
        self.debug_text.grid(row=len(campos_labels)+5, column=0, columnspan=3, pady=10)

    def procurar_ficheiro_ou_pasta(self, chave):
        initial_dir = os.path.dirname(os.path.abspath(__file__))  # pasta do script gui.py
        if chave == "pasta_emails":
            caminho = filedialog.askdirectory(initialdir=initial_dir)
        else:
            caminho = filedialog.askopenfilename(initialdir=initial_dir)
        if caminho:
            self.campos[chave].delete(0, tk.END)
            self.campos[chave].insert(0, caminho)

    def obter_caminhos(self,new_file):
        for chave in self.caminhos:
            widget = self.campos.get(chave)
            if isinstance(widget, ttk.Combobox):
                self.caminhos[chave] = widget.get().strip()
            elif widget:
                self.caminhos[chave] = widget.get().strip()

        # Extrair ano letivo do nome do ficheiro relatorio_mensal
        
        caminho_relatorio = self.caminhos.get("relatorio_mensal", "")
        if new_file is False:
            ano_letivo = self.parse_ano_letivo_do_ficheiro(caminho_relatorio)
            if ano_letivo:
                self.caminhos["ano_letivo"] = ano_letivo
                self.log_debug(f"Ano letivo extraído do ficheiro: {ano_letivo}")
            else:
                self.caminhos["ano_letivo"] = None
                self.log_debug("Não foi possível extrair o ano letivo do nome do ficheiro relatorioMensal.xlsx.")
        else :
            self.caminhos["ano_letivo"] = None
        return self.caminhos

    def parse_ano_letivo_do_ficheiro(self, caminho):
        # Procura padrão: RelatorioMensal_2024_2025.xlsx ou algo semelhante
        if not caminho:
            return None
        nome = os.path.basename(caminho)
        match = re.search(r"(\d{4})_(\d{4})", nome)
        if match:
            ano_inicio = match.group(1)
            ano_fim = match.group(2)
            return f"{ano_inicio}_{ano_fim}"
        return None

    def log_debug(self, mensagem):
        if self.mostrar_debug:
            self.debug_text.insert(tk.END, mensagem + "\n")
            self.debug_text.see(tk.END)

    def confirmar_relatorio_ano_letivo(self):
        resposta1 = messagebox.askyesno("Confirmação", "Tem a certeza que quer gerar o relatório do início do ano letivo?")
        if resposta1:
            resposta2 = messagebox.askyesno("Confirmação final", "Confirma que pretende criar o relatório do início do ano letivo? Esta ação pode substituir dados existentes.")
            if resposta2:
                self.gerar_relatorio_ano_letivo()

    def gerar_relatorio_ano_letivo(self):
        caminhos = self.obter_caminhos(True)
        try:
            from src.relatorio import gerar_relatorio_base_ano_letivo

            ano_atual = datetime.datetime.now().year
            ano_seguinte = ano_atual + 1
            nome_sugerido = f"RelatorioMensal_{ano_atual}_{ano_seguinte}.xlsx"

            initial_dir = os.path.dirname(os.path.abspath(__file__))
            caminho = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Ficheiros Excel", "*.xlsx")],
                initialfile=nome_sugerido,
                initialdir=initial_dir,
                title="Guardar Relatório Inicial Ano Letivo"
            )
            if not caminho:
                self.log_debug("Operação cancelada pelo utilizador.")
                return

            gerar_relatorio_base_ano_letivo(caminhos, caminho)

            messagebox.showinfo("Sucesso", "Relatório base do ano letivo gerado com sucesso.")
            self.log_debug(f"Relatório base do ano letivo criado: {caminho}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar relatório base: {e}")
            self.log_debug(f"Erro ao gerar relatório base: {e}")

    def acao_relatorio(self):
        caminhos = self.obter_caminhos(False)
        try:
            gerar_relatorioMensal(caminhos)
            messagebox.showinfo("Sucesso", "Relatório mensal gerado com sucesso.")
            self.log_debug("Relatório mensal gerado com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar relatório: {e}")
            self.log_debug(f"Erro ao gerar relatório: {e}")

    def acao_transferencias(self):
        caminhos = self.obter_caminhos(False)
        try:
            gerar_transferencias(caminhos)
            messagebox.showinfo("Sucesso", "Transferências geradas com sucesso.")
            self.log_debug("Transferências geradas com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar transferências: {e}")
            self.log_debug(f"Erro ao gerar transferências: {e}")

    def acao_emails(self):
        caminhos = self.obter_caminhos(False)
        try:
            gerar_emails(caminhos)
            messagebox.showinfo("Sucesso", "Emails gerados com sucesso.")
            self.log_debug("Emails gerados com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar emails: {e}")
            self.log_debug(f"Erro ao gerar emails: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = GestorReceitasGUI(root)
    root.mainloop()
