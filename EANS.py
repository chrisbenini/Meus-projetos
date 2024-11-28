# IMPORTANDO AS BIBLIOTECAS QUE SERAM USADAS NO CÓDIGO

from datetime import datetime, timedelta
import pandas as pd
import numpy as np
import sqlite3
import pyodbc
import sys
import re
import os

# FAZENDO A CONEXÃO COM O BANCO DE DADOS 
dados_conexao = ''
conexao = pyodbc.connect(dados_conexao)

# EXTRAINDO INFORMAÇÕES DO BANCO COM SQL
planilha_EAN = '''
SELECT 
    prd.cd_prod AS "CODIGO",
    COALESCE(NULLIF(prd.cd_barra, ''), '') AS "EAN",
	prd.descricao AS "PRODUTO",
    fab.descricao AS "FABRICANTE"
FROM 
    produto prd
INNER JOIN 
    preco prc ON prd.cd_prod = prc.cd_prod
LEFT JOIN 
    fabric fab ON prd.cd_fabric = fab.cd_fabric
WHERE 
    prd.cd_barra IS NOT NULL
    AND prd.cd_barra != ''
    AND prc.cd_tabela = 'PADRAO';
'''

# CARREGAR A CONSULTA DO SQL
codigo_df = pd.read_sql_query(planilha_EAN, conexao)

# FECHANDO A CONEXÃO COM O BANCO
conexao.close()

# LIMPEZA DOS DADOS SQL
codigo_df['EAN'] = dismed_df['EAN'].astype(str).str.strip()
codigo_df['EAN'] = dismed_df['EAN'].apply(lambda x: re.sub(r'\D', '', x))
codigo_df['EAN'] = pd.to_numeric(dismed_df['EAN'], errors='coerce')

# FILTRANDO PARA EXCLUIR FABRICANTES "BRINDES"
filtered_df = codigo_df[codigo_df['FABRICANTE'] != 'BRINDES']

# ADICIONANDO A COLUNA "DIGITOS EAN"
filtered_df['DIGITOS EAN'] = filtered_df['EAN'].apply(lambda x: len(str(int(x))) if pd.notnull(x) else 0)

# FILTRANDO EANS COM ERRO (MENOS DE 13 OU MAIS DE 13 DÍGITOS)
alerta_df = filtered_df[(filtered_df['DIGITOS EAN'] != 13)]

# IDENTIFICANDO A DATA ATUAL
data_atual = datetime.now().strftime('%d-%m-%Y')

# NOME DOS ARQUIVOS DE SAÍDA
alerta1 = r"P:\\TESTE\\TESTE\\EAN_ALERTA {data_atual}.xlsx"
alerta2 = r"P:\\TESTE\\TESTE\\TESTE\\TESTE\\TESTE\\EAN_ALERTA {data_atual}.xlsx"
alerta3 = r"P:\\TESTE\\TESTE\\TESTE\\EAN_ALERTA {data_atual}.xlsx"

# CRIAR DIRETORIOS SE ELES NAO EXISTIREM
os.makedirs(os.path.dirname(alerta1), exist_ok=True)
os.makedirs(os.path.dirname(alerta2), exist_ok=True)
os.makedirs(os.path.dirname(alerta3), exist_ok=True)

# SALVANDO O DATAFRAME ATUALIZADO EM AMBOS OS ARQUIVOS
alerta_df.to_excel(alerta1.format(data_atual=data_atual), index=False)
alerta_df.to_excel(alerta2.format(data_atual=data_atual), index=False)
alerta_df.to_excel(alerta2.format(data_atual=data_atual), index=False)

# RESULTADO FINAL
print(alerta_df)

# MENSAGENS DE CONFIRMAÇÃO
print(f"DADOS ATUALIZADOS FORAM SALVOS NO ARQUIVO: {alerta1.format(data_atual=data_atual)}")
print(f"DADOS ATUALIZADOS FORAM SALVOS NO ARQUIVO: {alerta2.format(data_atual=data_atual)}")




