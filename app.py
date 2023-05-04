from flask import Flask, render_template, request, redirect, flash
import pandas as pd

app = Flask(__name__)
app.secret_key = b'_5#y2L"F4Q8z\n\xec]/'

# Página inicial
@app.route('/')
def index():
    return render_template('index.html')

# Função para atualizar o arquivo
@app.route('/atualizar-arquivo', methods=['POST'])
def atualizar_arquivo():
    try:
        # Lê o arquivo de entrada e extrai as duas colunas desejadas
        df = pd.read_excel(request.files['file_input'], usecols='H, G', skiprows=1)

        # Aplica o aumento percentual sobre a coluna "Custo"
        porcentagem = float(request.form['porcentagem_entry'].replace(",", ".")) # Converte a porcentagem para float
        df["Custo"] = df["Custo"] * (1 + porcentagem/100)

        # Carrega o arquivo de destino com o pandas
        df_up = pd.read_excel(request.files['file_output'], sheet_name='Anúncios')

        # Atualiza as células E3:EX com os valores da coluna "Estoque" do arquivo de entrada
        for i, estoque in enumerate(df["Estoque"], start=2):
            df_up.loc[i, 'Quantidade\n(Obligatorio)'] = estoque

        # Atualiza as células F3:FX com os valores da coluna "Custo" do arquivo de entrada
        for i, custo in enumerate(df["Custo"], start=2):
            df_up.loc[i, 'Preço\n(Obligatorio)'] = custo

        # Escreve o DataFrame atualizado de volta para o arquivo
        df_up.to_excel(request.files['file_output'], sheet_name='Anúncios', index=False)

        # Mostra uma mensagem de sucesso
        flash("Arquivo atualizado com sucesso!", "success")

        return redirect('/')
    
    except Exception as e:
        # Mostra uma mensagem de erro
        flash("Ocorreu um erro ao atualizar o arquivo: " + str(e), "error")
        return redirect('/')

if __name__ == '__main__':
    app.run(debug=True)