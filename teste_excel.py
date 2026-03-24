import pandas as pd
import sys

# 1. O nome exato do seu arquivo
nome_arquivo = 'RELATORIO VALORES ACUMULADOS HPOD GESTÃOVT MERCADO TORRE - MARÇO 2026 .xlsx'

# 2. A lista de colunas, agora com a escrita EXATA do seu Excel
# Atenção aos detalhes de escrita e quebras de linha (\n)!
colunas_corretas = [
    'Matricula',       # <--- Correção: O Excel tem uma quebra de linha aqui
    'Nome', 
    'Seção', 
    'Função', 
    'CPF', 
    'Nùmero do Cartão', # <--- Correção: De 'Nùmero' para 'Número' com acento
    'Uso diário', 'D', 'F', 
    'Pedido Inicial', 'Total Acumulado', 'Valor economizado', 
    'Pedido Final', 'Status', 'Filial'
]

print("Tentando carregar a planilha...")

try:
    # 3. Lendo a planilha (aba CONSOLIDADO)
    # A chave do problema está aqui: header=2 diz para ler a terceira linha (índice 2).
    # E usecols=colunas_corretas garante que só puxaremos as colunas que você quer.
    df = pd.read_excel(nome_arquivo, sheet_name='CONSOLIDADO', header=2, usecols=colunas_corretas, dtype=str)
    
    # 4. Verificação de Sucesso
    # Se o DataFrame não estiver vazio, tudo funcionou.
    if not df.empty:
        print("\n=== Planilha carregada com sucesso! ===")
        print(f"Total de registros encontrados: {len(df)}")
        
        # Mostra os 5 primeiros funcionários para confirmar visualmente
        print("\nPrimeiros 5 registros carregados:")
        print(df.head())
    else:
        print("Erro: A planilha foi carregada, mas parece estar vazia.")

except FileNotFoundError:
    print(f"Erro: O arquivo '{nome_arquivo}' não foi encontrado. Verifique se ele está na mesma pasta deste script.")
except Exception as e:
    print(f"Ocorreu um erro ao ler o arquivo: {e}")