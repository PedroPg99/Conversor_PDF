import tkinter as tk
from tkinter import filedialog, messagebox
from docx2pdf import convert
import os

def selecionar_arquivo():
    caminho = filedialog.askopenfilename(
        title="Selecione um arquivo .docx",
        filetypes=[("Documentos Word", "*.docx")]
    )
    entrada.set(caminho)

def converter_para_pdf():
    arquivo_docx = entrada.get()
    if not arquivo_docx.endswith(".docx"):
        messagebox.showerror("Erro", "Por favor, selecione um arquivo .docx válido.")
        return
    try:
        convert(arquivo_docx)
        messagebox.showinfo("Sucesso", "Arquivo convertido com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha na conversão:\n{str(e)}")



# Criar janela principal
janela = tk.Tk()
janela.title("Conversor DOCX para PDF")
janela.geometry("600x150")
janela.resizable(False, False)

# Interface
entrada = tk.StringVar()

tk.Label(janela, text="Selecione um arquivo .docx").pack(pady=10)
tk.Entry(janela, textvariable=entrada, width=50).pack(padx=10)
tk.Button(janela, text="Procurar", command=selecionar_arquivo).pack(pady=5)
tk.Button(janela, text="Converter para PDF", command=converter_para_pdf).pack(pady=10)

# Iniciar interface
janela.mainloop()


# Para usar o programa standalone usa a  pyinstaller e ico para colocar um icone --onefile --icon=meu_icone.ico conversordocxparapdf.py