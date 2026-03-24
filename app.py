from flask import Flask, render_template, request, redirect, url_for, send_from_directory, session
import pandas as pd
import math  
import os
import json
from datetime import datetime
from werkzeug.utils import secure_filename  

app = Flask(__name__)
app.secret_key = 'sua_chave_secreta_aqui'  # Adicione uma chave secreta para sessões

@app.before_request
def require_login():
    allowed_routes = ['home', 'login', 'logout', 'static']
    if request.endpoint not in allowed_routes and 'user' not in session:
        return redirect(url_for('home'))

# Configurações de upload
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
HISTORICO_FILE = 'historico.json'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER

# Criar pastas se não existirem
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

# O nosso banco de dados de mentirinha
USUARIOS_VALIDOS = {
    "admin": "1234",
    "joao": "senha123"
}

# Funções para histórico
def carregar_historico():
    if os.path.exists(HISTORICO_FILE):
        with open(HISTORICO_FILE, 'r') as f:
            return json.load(f)
    return []

def salvar_historico(historico):
    with open(HISTORICO_FILE, 'w') as f:
        json.dump(historico, f, indent=4)


def remover_arquivo_se_existir(caminho_arquivo):
    if caminho_arquivo and os.path.exists(caminho_arquivo):
        os.remove(caminho_arquivo)

# Configurações do arquivo Excel
NOME_ARQUIVO = 'RELATORIO VALORES ACUMULADOS HPOD GESTÃOVT MERCADO TORRE - MARÇO 2026 .xlsx'
COLUNAS = [
    'Matricula', 'Nome', 'Seção', 'Função', 'CPF', 
    'Nùmero do Cartão', 'Uso diário', 'D', 'F', 
    'Pedido Inicial', 'Total Acumulado', 'Valor economizado', 
    'Pedido Final', 'Status', 'Filial'
]


def obter_arquivo_excel_atual():
    """Retorna o caminho do último arquivo processado pelo admin.

    Ordem de prioridade:
    1) Último arquivo com data_processamento no histórico
    2) Arquivo .xlsx mais recente da pasta processed
    3) NOME_ARQUIVO (fallback legado)
    """
    historico = carregar_historico()
    candidatos_historico = []

    for item in historico:
        arquivo_processado = item.get('arquivo_processado')
        if not arquivo_processado:
            continue

        caminho_processado = os.path.join(PROCESSED_FOLDER, arquivo_processado)
        if not os.path.exists(caminho_processado):
            continue

        data_processamento = item.get('data_processamento')
        try:
            if data_processamento:
                instante = datetime.strptime(data_processamento, '%Y-%m-%d %H:%M:%S')
            else:
                instante = datetime.fromtimestamp(os.path.getmtime(caminho_processado))
        except ValueError:
            instante = datetime.fromtimestamp(os.path.getmtime(caminho_processado))

        candidatos_historico.append((instante, caminho_processado))

    if candidatos_historico:
        candidatos_historico.sort(key=lambda x: x[0], reverse=True)
        return candidatos_historico[0][1]

    # Fallback: pega o .xlsx mais recente da pasta de processados
    arquivos_processados = []
    for nome in os.listdir(PROCESSED_FOLDER):
        if nome.lower().endswith('.xlsx'):
            caminho = os.path.join(PROCESSED_FOLDER, nome)
            if os.path.isfile(caminho):
                arquivos_processados.append(caminho)

    if arquivos_processados:
        return max(arquivos_processados, key=os.path.getmtime)

    # Fallback legado para manter compatibilidade
    if os.path.exists(NOME_ARQUIVO):
        return NOME_ARQUIVO

    raise FileNotFoundError('Nenhum arquivo processado pelo admin foi encontrado para leitura.')

@app.route('/')
def home():
    return render_template('login.html')

@app.route('/login', methods=['POST'])
def login():
    usuario_digitado = request.form.get('usuario')
    senha_digitada = request.form.get('senha')

    if usuario_digitado in USUARIOS_VALIDOS and USUARIOS_VALIDOS[usuario_digitado] == senha_digitada:
        session['user'] = usuario_digitado
        # Agora redireciona para a lista!
        return redirect(url_for('lista'))
    else:
        return render_template('login.html', erro="Usuário ou senha incorretos.")

# --- MUDAMOS O NOME DA ROTA DE /dashboard PARA /lista ---
@app.route('/lista')
def lista():
    filtro_nome = request.args.get('nome', '')
    filtro_cpf = request.args.get('cpf', '')
    filtro_filial = request.args.get('filial', '')
    filtro_cartao = request.args.get('cartao', '')
    
    try:
        pagina_atual = int(request.args.get('page', 1))
    except ValueError:
        pagina_atual = 1

    itens_por_pagina = 50

    try:
        arquivo_excel = obter_arquivo_excel_atual()
        df = pd.read_excel(arquivo_excel, sheet_name='CONSOLIDADO', header=2, usecols=COLUNAS, dtype=str)
        df = df.fillna('')

        if filtro_nome: df = df[df['Nome'].str.contains(filtro_nome, case=False)]
        if filtro_cpf: df = df[df['CPF'].str.contains(filtro_cpf, case=False)]
        if filtro_filial: df = df[df['Filial'].astype(str).str.strip() == filtro_filial.strip()]
        if filtro_cartao: df = df[df['Nùmero do Cartão'].astype(str).str.contains(filtro_cartao, case=False)]
        
        total_paginas = math.ceil(len(df) / itens_por_pagina)
        if pagina_atual < 1: pagina_atual = 1
        if pagina_atual > total_paginas and total_paginas > 0: pagina_atual = total_paginas

        df_pagina = df.iloc[(pagina_atual - 1) * itens_por_pagina : pagina_atual * itens_por_pagina]
        
        # Agora renderiza lista.html em vez de dashboard.html
        return render_template('lista.html', dados=df_pagina.to_dict(orient='records'), colunas=COLUNAS, 
                               nome_pesquisado=filtro_nome, cpf_pesquisado=filtro_cpf, 
                               filial_pesquisada=filtro_filial, cartao_pesquisado=filtro_cartao, 
                               pagina_atual=pagina_atual, total_paginas=total_paginas)
    except Exception as e:
        return f"<h1>Erro ao carregar o Excel:</h1><p>{e}</p>"

# --- NOVAS ROTAS (VAZIAS POR ENQUANTO) ---
def _to_float(value):
    if value is None or (isinstance(value, str) and value.strip() == ''):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    s = str(value).strip().replace('R$', '').replace(' ', '')
    if s == '':
        return 0.0

    has_dot = '.' in s
    has_comma = ',' in s

    if has_dot and has_comma:
        # BR format: 1.234.567,89
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            # US format: 1,234,567.89
            s = s.replace(',', '')
    elif has_comma:
        # assume decimal comma format
        s = s.replace('.', '').replace(',', '.')
    elif has_dot:
        # assume dot decimal format, não remover pontos
        # se houver agrupamento com ponto como 1.000.000, sem vírgula, isso pode ser milhares, mas não vamos transformar
        # conta usando valores normais de ponto
        pass

    try:
        return float(s)
    except ValueError:
        return 0.0

@app.route('/dashboard')
def dashboard_graficos():
    try:
        arquivo_excel = obter_arquivo_excel_atual()
        df = pd.read_excel(arquivo_excel, sheet_name='CONSOLIDADO', header=2, usecols=COLUNAS, dtype=str)
        df = df.fillna('')

        # Remover a linha de total/resumo que adiciona soma duplicada no Excel
        mask_total = (
            df['Nome'].astype(str).str.strip().str.upper().eq('TOTAL') |
            df['Matricula'].astype(str).str.strip().str.upper().eq('TOTAL') |
            df['Nome'].astype(str).str.strip().str.upper().str.contains('TOTAL', na=False) |
            df['Matricula'].astype(str).str.strip().str.upper().str.contains('TOTAL', na=False)
        )
        df = df[~mask_total]

        # Somar usando parsing robusto para formatos BR e EN
        pedido_inicial = df['Pedido Inicial'].map(_to_float).sum()
        pedido_final = df['Pedido Final'].map(_to_float).sum()
        economia_2 = pedido_inicial - pedido_final
        hpod_12 = economia_2 * 0.12
        glosa_nov2025 = 0.0
        hpod_nfse = hpod_12 - glosa_nov2025

        status_counts = df['Status'].str.strip().str.upper().value_counts()
        recarga_normal = int(status_counts.get('RECARGA NORMAL', 0))
        complemento = int(status_counts.get('COMPLEMENTO', 0))
        nao_efetuar = int(status_counts.get('NÃO EFETUAR RECARGA', 0) or status_counts.get('NAO EFETUAR RECARGA', 0))

        return render_template('dashboard.html',
                               pedido_inicial=pedido_inicial,
                               pedido_final=pedido_final,
                               economia_2=economia_2,
                               hpod_12=hpod_12,
                               glosa_nov2025=glosa_nov2025,
                               hpod_nfse=hpod_nfse,
                               recarga_normal=recarga_normal,
                               complemento=complemento,
                               nao_efetuar=nao_efetuar)
    except Exception as e:
        return f"<h1>Erro ao gerar dashboard:</h1><p>{e}</p>"

@app.route('/arquivos', methods=['GET', 'POST'])
def arquivos():
    user = session.get('user')
    if not user:
        return redirect(url_for('home'))
    
    if request.method == 'POST':
        if user == 'joao':
            if 'arquivo' not in request.files:
                return "Nenhum arquivo enviado", 400
            file = request.files['arquivo']
            if file.filename == '':
                return "Arquivo não selecionado", 400
            if file and file.filename.endswith('.xlsx'):
                filename = secure_filename(file.filename)
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                uploaded_filename = f"{timestamp}_{filename}"
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], uploaded_filename)
                file.save(file_path)
                
                # Atualizar histórico
                historico = carregar_historico()
                novo_item = {
                    'id': len(historico) + 1,
                    'usuario': user,
                    'data_envio': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'arquivo_enviado': uploaded_filename,
                    'arquivo_processado': None
                }
                historico.append(novo_item)
                salvar_historico(historico)
                
                return redirect(url_for('arquivos'))
            else:
                return "Arquivo inválido. Apenas .xlsx permitido", 400
        elif user == 'admin':
            # Admin uploading processed file
            item_id = request.form.get('item_id')
            if not item_id:
                return "Selecione um item para anexar o arquivo processado", 400
            if 'arquivo' not in request.files:
                return "Nenhum arquivo enviado", 400
            file = request.files['arquivo']
            if file.filename == '':
                return "Arquivo não selecionado", 400
            if file and file.filename.endswith('.xlsx'):
                filename = secure_filename(file.filename)
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                processed_filename = f"processed_{timestamp}_{filename}"
                processed_path = os.path.join(app.config['PROCESSED_FOLDER'], processed_filename)
                file.save(processed_path)
                
                # Atualizar histórico
                historico = carregar_historico()
                for item in historico:
                    if str(item['id']) == item_id and item['arquivo_processado'] is None:
                        item['arquivo_processado'] = processed_filename
                        item['data_processamento'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        break
                salvar_historico(historico)
                
                return redirect(url_for('arquivos'))
            else:
                return "Arquivo inválido. Apenas .xlsx permitido", 400
    
    historico = carregar_historico()
    if user == 'joao':
        # Joao vê apenas seus envios
        historico_filtrado = [h for h in historico if h.get('usuario') == 'joao']
    else:
        # Admin vê todos
        historico_filtrado = historico
    
    return render_template('arquivos.html', historico=historico_filtrado, user=user)

@app.route('/download/<int:item_id>')
def download(item_id):
    historico = carregar_historico()
    item = next((i for i in historico if i['id'] == item_id), None)
    if item and item['arquivo_processado']:
        return send_from_directory(app.config['PROCESSED_FOLDER'], item['arquivo_processado'], as_attachment=True)
    return "Arquivo não encontrado", 404

@app.route('/download_uploaded/<int:item_id>')
def download_uploaded(item_id):
    historico = carregar_historico()
    item = next((i for i in historico if i['id'] == item_id), None)
    if item and item['arquivo_enviado']:
        return send_from_directory(app.config['UPLOAD_FOLDER'], item['arquivo_enviado'], as_attachment=True)
    return "Arquivo não encontrado", 404


@app.route('/admin/limpar_item/<int:item_id>', methods=['POST'])
def limpar_item_admin(item_id):
    user = session.get('user')
    if user != 'admin':
        return "Acesso negado", 403

    historico = carregar_historico()
    item = next((i for i in historico if i['id'] == item_id), None)

    if not item:
        return "Item não encontrado", 404

    caminho_enviado = os.path.join(app.config['UPLOAD_FOLDER'], item.get('arquivo_enviado', ''))
    caminho_processado = os.path.join(app.config['PROCESSED_FOLDER'], item.get('arquivo_processado', ''))

    remover_arquivo_se_existir(caminho_enviado)
    remover_arquivo_se_existir(caminho_processado)

    historico_atualizado = [i for i in historico if i.get('id') != item_id]

    # Reorganiza IDs para manter sequência após exclusão
    for idx, registro in enumerate(historico_atualizado, start=1):
        registro['id'] = idx

    salvar_historico(historico_atualizado)
    return redirect(url_for('arquivos'))


@app.route('/admin/limpar_processados', methods=['POST'])
def limpar_todos_processados_admin():
    user = session.get('user')
    if user != 'admin':
        return "Acesso negado", 403

    historico = carregar_historico()
    historico_atualizado = []

    for item in historico:
        if item.get('arquivo_processado'):
            caminho_enviado = os.path.join(app.config['UPLOAD_FOLDER'], item.get('arquivo_enviado', ''))
            caminho_processado = os.path.join(app.config['PROCESSED_FOLDER'], item.get('arquivo_processado', ''))

            remover_arquivo_se_existir(caminho_enviado)
            remover_arquivo_se_existir(caminho_processado)
            continue

        historico_atualizado.append(item)

    # Reorganiza IDs para manter sequência após exclusão em lote
    for idx, registro in enumerate(historico_atualizado, start=1):
        registro['id'] = idx

    salvar_historico(historico_atualizado)
    return redirect(url_for('arquivos'))

def processar_excel(input_path, output_path):
    # Exemplo de processamento: adicionar coluna "Processado em"
    df = pd.read_excel(input_path)
    df['Processado em'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    df.to_excel(output_path, index=False)

@app.route('/logout')
def logout():
    session.pop('user', None)
    return redirect(url_for('home'))

if __name__ == '__main__':
    app.run(debug=True)