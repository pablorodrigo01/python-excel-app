from flask import Flask, render_template, request, redirect
import pandas as pd
from openpyxl import load_workbook
from flask_frozen import Freezer

app = Flask(__name__)
freezer = Freezer(app)

# Página inicial
@app.route('/')
def index():
    return render_template('index.html')

# Função para atualizar o arquivo
@app.route('/atualizar-arquivo', methods=['POST'])
def atualizar_arquivo():
    try:
        # Lê o arquivo de entrada e extrai as duas colunas desejadas
        df = pd.read_excel(request.form['file_input'], usecols='H, G', skiprows=1)

        # Aplica o aumento percentual sobre a coluna "Custo"
        porcentagem = float(request.form['porcentagem_entry'].replace(",", ".")) # Converte a porcentagem para float
        df["Custo"] = df["Custo"] * (1 + porcentagem/100)

        # Abre o arquivo de destino em modo de leitura
        book = load_workbook(request.form['file_output'])

        # Seleciona a planilha "Anúncios" para atualização
        sheet = book["Anúncios"]

        # Atualiza as células C3:CX com os valores da coluna "Nome do produto" do arquivo de entrada
        for i, estoque in enumerate(df["Estoque"], start=4):
            sheet.cell(row=i, column=5, value=estoque)

        # Atualiza as células F3:FX com os valores da coluna "Custo" do arquivo de entrada
        for i, custo in enumerate(df["Custo"], start=4):
            sheet.cell(row=i, column=6, value=custo)

        # Salva as alterações no arquivo de destino
        book.save(request.form['file_output'])

        return redirect('/')
    
    except Exception as e:
        # Mostra uma mensagem de erro
        return redirect('/')
    
if __name__ == '__main__':
    freezer.freeze()