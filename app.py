from flask import Flask, render_template, request
import pandas as pd
from openpyxl import load_workbook

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/result', methods=['POST'])
def result():
    try:
        # Lê o arquivo de entrada e extrai as duas colunas desejadas
        df = pd.read_excel(request.files['file_input'], usecols='H, G', skiprows=1)

        # Aplica o aumento percentual sobre a coluna "Custo"
        porcentagem = float(request.form['porcentagem_entry'].replace(",", ".")) # Converte a porcentagem para float
        df["Custo"] = df["Custo"] * (1 + porcentagem/100)

        # Abre o arquivo de destino em modo de leitura
        book = load_workbook(request.files['file_output'])

        # Seleciona a planilha "Anúncios" para atualização
        sheet = book["Anúncios"]

        # Atualiza as células E3:EX com os valores da coluna "Estoque" do arquivo de entrada
        for i, estoque in enumerate(df["Estoque"], start=4):
            sheet.cell(row=i, column=5, value=estoque)

        # Atualiza as células F3:FX com os valores da coluna "Custo" do arquivo de entrada
        for i, custo in enumerate(df["Custo"], start=4):
            sheet.cell(row=i, column=6, value=custo)

        # Salva as alterações no arquivo de destino
        book.save(request.files['file_output'])

        # Retorna uma mensagem de sucesso
        return render_template('result.html', success=True)
    

    except Exception as e:
        # Retorna uma mensagem de erro
        return render_template('result.html', success=False, error=str(e))

if __name__ == '__main__':
    app.run(debug=True)