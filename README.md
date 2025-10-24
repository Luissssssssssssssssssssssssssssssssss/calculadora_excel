# calculadora_excel
#criar pasta com um arquivo no excel para salvar o histórico da calculadora
import tkinter as tk
from tkinter import filedialog, messagebox
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


# ---------- Interface Gráfica ----------
class CalculadoraGUI:
    def __init__(self, root):
        self.calc = Calculadora()
        self.root = root
        self.root.title("🧮 Calculadora com Registro em Excel")
        self.root.geometry("400x580")
        self.root.resizable(False, False)

        self.arquivo_excel = None  # Caminho escolhido pelo usuário

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

        # ---------- Botão de salvar ----------
        ttk.Button(
            root,
            text="📁 Escolher local do arquivo Excel",
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

    def escolher_arquivo(self):
        """Abre janela para o usuário escolher ou criar o arquivo Excel."""
        caminho = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Planilhas Excel", "*.xlsx")],
            title="Salvar resultados como..."
        )
        if caminho:
            self.arquivo_excel = caminho
            messagebox.showinfo("Arquivo selecionado", f"Resultados serão salvos em:\n{caminho}")


# ---------- Execução ----------
if __name__ == "__main__":
    app = ttk.Window(themename="cyborg")  # Tema moderno (pode trocar por 'flatly', 'darkly', 'solar', etc.)
    CalculadoraGUI(app)
    app.mainloop()
