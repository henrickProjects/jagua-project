from flask import Flask, request, send_file, jsonify
from flask_cors import CORS # Permite que o site na InfinityFree conecte aqui
import os
import shutil
# Importe suas outras bibliotecas (requests, openpyxl, re...)

app = Flask(__name__)
# Libera o acesso para qualquer site chamar a sua API
CORS(app) 

# Coloque suas funções antigas aqui (consultar_cnpj, formatar, preencher excel...)
# ...

@app.route('/gerar_unico', methods=['POST'])
def rota_unico():
    dados = request.json
    cnpj_limpo = dados.get('cnpj')
    
    # --- AQUI VAI SUA LÓGICA ---
    # 1. Consulta a API da cnpja
    # 2. Carrega o openpyxl e edita o modelo
    # 3. Salva com o nome do cliente em uma pasta temporária
    caminho_arquivo_gerado = f"./temp/{cnpj_limpo}.xlsx" 
    
    # Envia o arquivo de volta para o site baixar
    return send_file(caminho_arquivo_gerado, as_attachment=True, download_name='cadastro.xlsx')

@app.route('/gerar_multiplos', methods=['POST'])
def rota_multiplos():
    dados = request.json
    lista_cnpjs = dados.get('cnpjs', [])
    
    # --- AQUI VAI SUA LÓGICA ---
    # 1. Faz um loop for pela lista_cnpjs
    # 2. Gera todos os Excel dentro de uma pasta temporária
    # 3. Zipa a pasta inteira usando o shutil.make_archive
    caminho_zip = f"./temp/cadastros.zip"
    
    # Envia o ZIP de volta para o site
    return send_file(caminho_zip, as_attachment=True, download_name='cadastros.zip')

if __name__ == '__main__':
    # Cria pastas temporárias
    os.makedirs("./temp", exist_ok=True)
    app.run(host='0.0.0.0', port=5000)
