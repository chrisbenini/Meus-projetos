
# IMPORTANDO AS BIBLIOTECAS QUE SERAM USADAS NO CÓDIGO

from datetime import datetime, timedelta
import pandas as pd
import numpy as np
import sqlite3
import pyodbc
import sys
import re
import os

# LOCAL DAS PLANILHAS BASES
planilhas_pmc = r"P:\\INTELIGENCIA\\PRECIFICACAO\\Python\\ALERTA PMC_PMPF\\BASES\\PMC.xlsx"
planilhas_pmpf = r"P:\\INTELIGENCIA\\PRECIFICACAO\\Python\\ALERTA PMC_PMPF\\BASES\\PMPF.xlsx"

# FAZENDO A CONEXÃO COM O BANCO DE DADOS 
dados_conexao = 'DRIVER={SQL Server};SERVER=SRVDELL;DATABASE=MOINHO;UID=servicefarma;PWD=sf@2023#d'
conexao = pyodbc.connect(dados_conexao)

# EXTRAINDO INFORMAÇÕES DO BANCO COM SQL
planilha_dismed = '''
SET NOCOUNT ON

SELECT 
    pro.cd_prod AS 'CODIGO',
    pro.cd_barra AS 'EAN',
    pro.descricao AS 'DESCRICAO',
    fab.descricao AS 'FABRICANTE',
    pro.cd_prod_ncm AS 'NCM',
    ISNULL(pmc.estado,'') AS 'UF',
    CONVERT(NUMERIC(15,2),ISNULL(pmc.vl_preco,0)) AS 'PMC',
    CONVERT(NUMERIC(15,2),ISNULL(pmc.ValorPMPF,0)) AS 'PMPF',
    CONVERT(INT,est.qtde) AS 'ESTOQUE',
    CONVERT(CHAR(10),GETDATE(),103) + ' ' + CONVERT(CHAR(5),GETDATE(),108) AS 'DATA HORA CONSULTA'
FROM
    produto pro
    JOIN fabric fab ON fab.cd_fabric = pro.cd_fabric
    JOIN estoque est WITH (NOLOCK) ON pro.cd_prod = est.cd_prod
    LEFT JOIN prc_max_prod pmc ON pro.cd_prod = pmc.cd_prod
WHERE
    est.cd_local = 'CENTRAL'
    AND est.cd_emp = 1
    AND est.qtde > 0
    AND pro.cd_linha NOT IN ('109','110','111')
    AND ISNULL(pmc.vl_preco,0) = 0
    AND pmc.estado = 'SP'
    AND (pro.cd_prod_ncm LIKE '3003%' OR pro.cd_prod_ncm LIKE '3004%') 
ORDER BY
    fab.descricao,
    pro.descricao;
'''

# CARREGAR A CONSULTA DO SQL
dismed_df = pd.read_sql_query(planilha_dismed, conexao)

# FECHANDO A CONEXÃO COM O BANCO
conexao.close()

# IDENTIFICANDO A DATA ATUAL
data_atual = datetime.now().strftime('%d-%m-%Y')

# LIMPEZA DOS DADOS SQL
dismed_df['EAN'] = dismed_df['EAN'].astype(str).str.strip()
dismed_df['EAN'] = dismed_df['EAN'].apply(lambda x: re.sub(r'\D', '', x))
dismed_df['EAN'] = pd.to_numeric(dismed_df['EAN'], errors='coerce')

# VERIFICA O NOME CORRETO DA COLUNA EAN NA PLANILHA PMC
pmc_df = pd.read_excel(planilhas_pmc)

# AJUSTANDO O NOME DA COLUNA EAN SE NECESSÁRIO NA PLANILHA PMC
ean_col_pmc = 'EAN 1' if 'EAN 1' in pmc_df.columns else 'EAN'
pmc_df[ean_col_pmc] = pmc_df[ean_col_pmc].astype(str).str.strip()  # REMOVENDO ESPAÇOS
pmc_df[ean_col_pmc] = pmc_df[ean_col_pmc].apply(lambda x: re.sub(r'\D', '', x))  # REMOVENDO CARACTERES NÃO NUMÉRICOS
pmc_df['EAN'] = pd.to_numeric(pmc_df[ean_col_pmc], errors='coerce')  # CONVERTENDO PARA NUMÉRICO E CRIANDO COLUNA 'EAN' PADRONIZADA

# AJUSTANDO O NOME DA COLUNA EAN SE NECESSÁRIO NA PLANILHA PMPF
pmpf_df = pd.read_excel(planilhas_pmpf)

ean_col_pmpf = 'EAN 1' if 'EAN 1' in pmpf_df.columns else 'EAN'
pmpf_df[ean_col_pmpf] = pmpf_df[ean_col_pmpf].astype(str).str.strip()  # REMOVENDO ESPAÇOS
pmpf_df[ean_col_pmpf] = pmpf_df[ean_col_pmpf].apply(lambda x: re.sub(r'\D', '', x))  # REMOVENDO CARACTERES NÃO NUMÉRICOS
pmpf_df['EAN'] = pd.to_numeric(pmpf_df[ean_col_pmpf], errors='coerce')  # CONVERTENDO PARA NUMÉRICO E CRIANDO COLUNA 'EAN' PADRONIZADA

# ATUALIZANDO PMC COM BASE NO MERGE
if 'PMC 18%' in pmc_df.columns:
    dismed_df = dismed_df.merge(pmc_df[['EAN', 'PMC 18%']], on='EAN', how='left')  # MERGE COM PMC
    dismed_df['PMC'] = dismed_df.apply(lambda row: row['PMC 18%'] if row['PMC'] == 0.00 else row['PMC'], axis=1)  # ATUALIZANDO PMC
    dismed_df.drop(columns=['PMC 18%'], inplace=True)  # REMOVENDO A COLUNA AUXILIAR

# ATUALIZANDO PMPF COM BASE NO MERGE
if 'PMPF' in pmpf_df.columns:
    dismed_df = dismed_df.merge(pmpf_df[['EAN', 'PMPF']], on='EAN', how='left', suffixes=('', '_new'))  # MERGE COM PMPF
    dismed_df['PMPF'] = dismed_df.apply(lambda row: row['PMPF_new'] if row['PMPF'] == 0.00 else row['PMPF'], axis=1)  # ATUALIZANDO PMPF
    dismed_df.drop(columns=['PMPF_new'], inplace=True)  # REMOVENDO A COLUNA AUXILIAR

# NOME DOS ARQUIVOS DE SAÍDA
alerta1 = r"P:\\ESCRITA_FISCAL\\ALERTA_PMC_PMPF\\ALERTA {data_atual}.xlsx"
alerta2 = r"P:\\DEPTO_FARMACEUTICO\\ALERTA_PMC_PMPF\\ALERTA {data_atual}.xlsx"
alerta3 = r"P:\\INTELIGENCIA\\PRECIFICACAO\\Python\\ALERTA PMC_PMPF\\ALERTA\\ALERTA {data_atual}.xlsx"


# CRIAR DIRETORIOS SE ELES NAO EXISTIREM
os.makedirs(os.path.dirname(alerta1), exist_ok=True)
os.makedirs(os.path.dirname(alerta2), exist_ok=True)
os.makedirs(os.path.dirname(alerta3), exist_ok=True)

# SALVANDO O DATAFRAME ATUALIZADO EM AMBOS OS ARQUIVOS
dismed_df.to_excel(alerta1.format(data_atual=data_atual), index=False)
dismed_df.to_excel(alerta2.format(data_atual=data_atual), index=False)
dismed_df.to_excel(alerta3.format(data_atual=data_atual), index=False)  

# RESULTADO FINAL
print(dismed_df)

# MENSAGENS DE CONFIRMAÇÃO
print(f"DADOS ATUALIZADOS FORAM SALVOS NO ARQUIVO: {alerta1.format(data_atual=data_atual)}")
print(f"DADOS ATUALIZADOS FORAM SALVOS NO ARQUIVO: {alerta2.format(data_atual=data_atual)}")
print(f"DADOS ATUALIZADOS FORAM SALVOS NO ARQUIVO: {alerta3.format(data_atual=data_atual)}")



