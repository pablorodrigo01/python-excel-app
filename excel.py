import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl import load_workbook

# Função para atualizar o arquivo
def atualizar_arquivo():
    try:
        # Lê o arquivo de entrada e extrai as duas colunas desejadas
        df = pd.read_excel(file_input.get(), usecols='H, G', skiprows=1)

        # Aplica o aumento percentual sobre a coluna "Custo"
        porcentagem = float(porcentagem_entry.get().replace(",", ".")) # Converte a porcentagem para float
        df["Custo"] = df["Custo"] * (1 + porcentagem/100)

        # Abre o arquivo de destino em modo de leitura
        book = load_workbook(file_output.get())

        # Seleciona a planilha "Anúncios" para atualização
        sheet = book["Anúncios"]

        # Atualiza as células C3:CX com os valores da coluna "Nome do produto" do arquivo de entrada
        for i, estoque in enumerate(df["Estoque"], start=4):
            sheet.cell(row=i, column=5, value=estoque)

        # Atualiza as células F3:FX com os valores da coluna "Custo" do arquivo de entrada
        for i, custo in enumerate(df["Custo"], start=4):
            sheet.cell(row=i, column=6, value=custo)

        # Salva as alterações no arquivo de destino
        book.save(file_output.get())

        # Mostra uma mensagem de sucesso
        messagebox.showinfo("Sucesso", "Arquivo atualizado com sucesso!")

    except Exception as e:
        # Mostra uma mensagem de erro
        messagebox.showerror("Erro", str(e))

    finally:
        # Fecha a janela principal
        root.destroy()

# Cria uma janela
root = tk.Tk()
root.title("Atualização de Arquivo")

# Cria os widgets para seleção dos arquivos
tk.Label(root, text="Arquivo de Entrada:").grid(row=0, column=0, sticky="w")
file_input = tk.Entry(root)
file_input.grid(row=0, column=1)
tk.Button(root, text="Selecionar", command=lambda: file_input.insert(
    0, filedialog.askopenfilename())).grid(row=0, column=2)

tk.Label(root, text="Arquivo de Saída:").grid(row=1, column=0, sticky="w")
file_output = tk.Entry(root)
file_output.grid(row=1, column=1)
tk.Button(root, text="Selecionar", command=lambda: file_output.insert(
    0, filedialog.asksaveasfilename(defaultextension=".xlsx"))).grid(row=1, column=2)

# Cria um widget para escolher a porcentagem de aumento
tk.Label(root, text="Porcentagem de Aumento:").grid(row=2, column=0, sticky="w", pady=10)
porcentagem_entry = tk.Entry(root, width=5)
porcentagem_entry.grid(row=2, column=1)
tk.Label(root, text="\uFF05", anchor="w").grid(row=2, column=2)

# Cria um widget vazio para adicionar espaço entre o campo de porcentagem e o botão de atualizar
tk.Label(root, text="").grid(row=3, column=1)

# Cria um botão para atualizar o arquivo
tk.Button(root, text="Atualizar", command=atualizar_arquivo).grid(row=3, column=1, pady=10)

# Inicia o loop principal da interface gráfica
root.mainloop()