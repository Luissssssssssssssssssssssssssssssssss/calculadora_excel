import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from openpyxl import Workbook, load_workbook
from datetime import datetime
import math
import os


# ---------- Função para salvar no Excel ----------
def salvar_resultado_excel(expressao, resultado, arquivo):
    """Cria (ou atualiza) um arquivo Excel com os resultados."""
    if not arquivo:
        return

    if os.path.exists(arquivo):
        wb = load_workbook(arquivo)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Resultados"
        ws.append(["Expressão", "Resultado", "Data/Hora"])

    ws.append([expressao, resultado, datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    wb.save(arquivo)
    print(f"✅ Resultado salvo em: {os.path.abspath(arquivo)}")


# ---------- Lógica da Calculadora ----------
class Calculadora:
    def avaliar(self, expressao):
        """Avalia expressões matemáticas de forma segura."""
        try:
            permitido = {k: v for k, v in math.__dict__.items() if not k.startswith("__")}
            resultado = eval(expressao, {"__builtins__": None}, permitido)
            return resultado
        except Exception:
            raise ValueError("Expressão inválida")

    def juros_simples(self, capital, taxa, tempo):
        """Retorna o montante com juros simples."""
        return capital * (1 + (taxa / 100) * tempo)

    def juros_compostos(self, capital, taxa, tempo):
        """Retorna o montante com juros compostos."""
        return capital * ((1 + (taxa / 100)) ** tempo)


# ---------- Interface Gráfica ----------
class CalculadoraGUI:
    def __init__(self, root):
        self.calc = Calculadora()
        self.root = root
        self.root.title("🧮 Calculadora com Registro em Excel")
        self.root.geometry("400x640")
        self.root.resizable(False, False)

        self.arquivo_excel = None  # Caminho do Excel
        self.tema_atual = "cyborg"

        # ---------- Campo de exibição ----------
        self.display = ttk.Entry(root, justify="right", font=("Arial", 22))
        self.display.pack(padx=10, pady=15, fill="x")

        # ---------- Grade de botões ----------
        frame = ttk.Frame(root)
        frame.pack()

        botoes = [
            ("7", 1, 0), ("8", 1, 1), ("9", 1, 2), ("/", 1, 3),
            ("4", 2, 0), ("5", 2, 1), ("6", 2, 2), ("*", 2, 3),
            ("1", 3, 0), ("2", 3, 1), ("3", 3, 2), ("-", 3, 3),
            ("0", 4, 0), (".", 4, 1), ("(", 4, 2), (")", 4, 3),
            ("C", 5, 0), ("**", 5, 1), ("+", 5, 2), ("=", 5, 3),
        ]

        for (texto, linha, coluna) in botoes:
            ttk.Button(
                frame,
                text=texto,
                width=7,
                bootstyle=INFO,
                command=lambda t=texto: self.on_click(t)
            ).grid(row=linha, column=coluna, padx=5, pady=5)

        # ---------- Histórico ----------
        ttk.Label(root, text="Histórico de Cálculos", font=("Arial", 12, "bold")).pack(pady=8)
        self.lista_resultados = tk.Listbox(root, height=8, font=("Consolas", 11))
        self.lista_resultados.pack(padx=10, pady=5, fill="both")

        # ---------- Botões Extras ----------
        frame_funcoes = ttk.Frame(root)
        frame_funcoes.pack(pady=10)

        ttk.Button(
            frame_funcoes,
            text="💰 Juros Simples",
            bootstyle=WARNING,
            width=18,
            command=self.calcular_juros_simples
        ).grid(row=0, column=0, padx=5, pady=5)

        ttk.Button(
            frame_funcoes,
            text="📈 Juros Compostos",
            bootstyle=PRIMARY,
            width=18,
            command=self.calcular_juros_compostos
        ).grid(row=0, column=1, padx=5, pady=5)

        ttk.Button(
            frame_funcoes,
            text="🎨 Trocar Tema",
            bootstyle=(SECONDARY, OUTLINE),
            width=18,
            command=self.trocar_tema
        ).grid(row=1, column=0, columnspan=2, pady=5)

        ttk.Button(
            root,
            text="📁 Escolher arquivo Excel",
            bootstyle=(SUCCESS, OUTLINE),
            command=self.escolher_arquivo
        ).pack(pady=10)

    # ---------- Lógica dos botões ----------
    def on_click(self, char):
        if char == "C":
            self.display.delete(0, tk.END)
        elif char == "=":
            expressao = self.display.get()
            try:
                resultado = self.calc.avaliar(expressao)
                self.display.delete(0, tk.END)
                self.display.insert(tk.END, str(resultado))
                self.lista_resultados.insert(tk.END, f"{expressao} = {resultado}")

                if self.arquivo_excel:
                    salvar_resultado_excel(expressao, resultado, self.arquivo_excel)
                else:
                    messagebox.showwarning("Aviso", "Escolha um arquivo Excel antes de salvar!")

            except ValueError:
                messagebox.showerror("Erro", "Expressão inválida!")
        else:
            self.display.insert(tk.END, char)

    # ---------- Escolher arquivo ----------
    def escolher_arquivo(self):
        caminho = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Planilhas Excel", "*.xlsx")],
            title="Salvar resultados como..."
        )
        if caminho:
            self.arquivo_excel = caminho
            messagebox.showinfo("Arquivo selecionado", f"Resultados serão salvos em:\n{caminho}")

    # ---------- Juros Simples ----------
    def calcular_juros_simples(self):
        try:
            c = float(simpledialog.askstring("Juros Simples", "Digite o capital inicial (R$):"))
            i = float(simpledialog.askstring("Juros Simples", "Digite a taxa (% ao período):"))
            t = float(simpledialog.askstring("Juros Simples", "Digite o tempo (em períodos):"))
            resultado = self.calc.juros_simples(c, i, t)
            expressao = f"Juros Simples: C={c}, i={i}%, t={t}"
            self.display.delete(0, tk.END)
            self.display.insert(tk.END, str(round(resultado, 2)))
            self.lista_resultados.insert(tk.END, f"{expressao} → M={round(resultado,2)}")

            if self.arquivo_excel:
                salvar_resultado_excel(expressao, round(resultado, 2), self.arquivo_excel)
            else:
                messagebox.showwarning("Aviso", "Escolha um arquivo Excel antes de salvar!")

        except:
            messagebox.showerror("Erro", "Valores inválidos para cálculo de juros simples!")

    # ---------- Juros Compostos ----------
    def calcular_juros_compostos(self):
        try:
            c = float(simpledialog.askstring("Juros Compostos", "Digite o capital inicial (R$):"))
            i = float(simpledialog.askstring("Juros Compostos", "Digite a taxa (% ao período):"))
            t = float(simpledialog.askstring("Juros Compostos", "Digite o tempo (em períodos):"))
            resultado = self.calc.juros_compostos(c, i, t)
            expressao = f"Juros Compostos: C={c}, i={i}%, t={t}"
            self.display.delete(0, tk.END)
            self.display.insert(tk.END, str(round(resultado, 2)))
            self.lista_resultados.insert(tk.END, f"{expressao} → M={round(resultado,2)}")

            if self.arquivo_excel:
                salvar_resultado_excel(expressao, round(resultado, 2), self.arquivo_excel)
            else:
                messagebox.showwarning("Aviso", "Escolha um arquivo Excel antes de salvar!")

        except:
            messagebox.showerror("Erro", "Valores inválidos para cálculo de juros compostos!")

    # ---------- Troca de tema ----------
    def trocar_tema(self):
        temas = ["cyborg", "flatly", "darkly", "solar", "morph", "superhero", "pulse"]
        indice_atual = temas.index(self.tema_atual)
        novo_tema = temas[(indice_atual + 1) % len(temas)]
        self.root.style.theme_use(novo_tema)
        self.tema_atual = novo_tema
        messagebox.showinfo("Tema alterado", f"Tema atual: {novo_tema}")


# ---------- Execução ----------
if __name__ == "__main__":
    app = ttk.Window(themename="cyborg")
    CalculadoraGUI(app)
    app.mainloop()

