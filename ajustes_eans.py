"""
ajustes_eans.py

Identificação de EANs inválidos (diferente de 13 dígitos)
a partir de dados do banco (SQL Server).
"""

from datetime import datetime
import pandas as pd
import pyodbc
import os
import re

# String de conexão (preencher conforme ambiente)
dados_conexao = ""

SQL_EAN = """
SELECT 
    prd.cd_prod AS CODIGO,
    COALESCE(NULLIF(prd.cd_barra, ''), '') AS EAN,
    prd.descricao AS PRODUTO,
    fab.descricao AS FABRICANTE
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
"""

def limpar_ean_serie(serie: pd.Series) -> pd.Series:
    """Remove espaços, caracteres não numéricos e converte para número."""
    serie = serie.astype(str).str.strip()
    serie = serie.apply(lambda x: re.sub(r"\D", "", x))
    return pd.to_numeric(serie, errors="coerce")

def main() -> None:
    # 1) Buscar dados do banco
    conexao = pyodbc.connect(dados_conexao)
    df = pd.read_sql_query(SQL_EAN, conexao)
    conexao.close()

    # 2) Limpeza do EAN
    df["EAN"] = limpar_ean_serie(df["EAN"])

    # 3) Remover fabricante "BRINDES"
    df = df[df["FABRICANTE"] != "BRINDES"]

    # 4) Adicionar coluna com quantidade de dígitos
    df["DIGITOS EAN"] = df["EAN"].apply(
        lambda x: len(str(int(x))) if pd.notnull(x) else 0
    )

    # 5) Filtrar EANs com erro (≠ 13 dígitos)
    alerta_df = df[df["DIGITOS EAN"] != 13]

    data_atual = datetime.now().strftime("%d-%m-%Y")

    output_paths = [
        r"P:\TESTE\TESTE\EAN_ALERTA {data_atual}.xlsx",
        r"P:\TESTE\TESTE\TESTE\TESTE\TESTE\EAN_ALERTA {data_atual}.xlsx",
        r"P:\TESTE\TESTE\TESTE\EAN_ALERTA {data_atual}.xlsx",
    ]

    for path_template in output_paths:
        path = path_template.format(data_atual=data_atual)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        alerta_df.to_excel(path, index=False)
        print(f"DADOS ATUALIZADOS FORAM SALVOS NO ARQUIVO: {path}")

    print("Total de EANs com problema:", len(alerta_df))

if __name__ == "__main__":
    main()
