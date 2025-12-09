"""
ALERTAS.py

Ajuste de produtos sem PMC/PMPF e geração de planilhas de alerta
a partir de dados do banco (SQL Server) + planilhas base PMC/PMPF.
"""

from datetime import datetime
import pandas as pd
import pyodbc
import os
import re

# =========================
# CONFIGURAÇÕES / PARÂMETROS
# =========================

# Caminho das planilhas base de PMC e PMPF
planilha_pmc_path = r""  # ex: r"P:\TESTE\BASES\PMC.xlsx"
planilha_pmpf_path = r""  # ex: r"P:\TESTE\BASES\PMPF.xlsx"

# String de conexão com o banco (preencher de acordo com o ambiente)
dados_conexao = ""  # ex: "Driver={SQL Server};Server=...;Database=...;Trusted_Connection=yes;"


# =========================
# CONSULTA SQL
# =========================

SQL_ALERTAS = """
SET NOCOUNT ON;

SELECT 
    pro.cd_prod AS CODIGO,
    pro.cd_barra AS EAN,
    pro.descricao AS DESCRICAO,
    fab.descricao AS FABRICANTE,
    pro.cd_prod_ncm AS NCM,
    ISNULL(pmc.estado,'') AS UF,
    CONVERT(NUMERIC(15,2),ISNULL(pmc.vl_preco,0)) AS PMC,
    CONVERT(NUMERIC(15,2),ISNULL(pmc.ValorPMPF,0)) AS PMPF,
    CONVERT(INT,est.qtde) AS ESTOQUE,
    CONVERT(CHAR(10),GETDATE(),103) + ' ' + CONVERT(CHAR(5),GETDATE(),108) AS [DATA HORA CONSULTA]
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
"""

# =========================
# FUNÇÕES AUXILIARES
# =========================

def limpar_ean_serie(serie: pd.Series) -> pd.Series:
    """Remove espaços, caracteres não numéricos e converte para número."""
    serie = serie.astype(str).str.strip()
    serie = serie.apply(lambda x: re.sub(r"\D", "", x))
    return pd.to_numeric(serie, errors="coerce")

# =========================
# PIPELINE PRINCIPAL
# =========================

def main() -> None:
    # 1) Buscar dados do banco
    conexao = pyodbc.connect(dados_conexao)
    df = pd.read_sql_query(SQL_ALERTAS, conexao)
    conexao.close()

    # 2) Limpeza do EAN vindo do banco
    df["EAN"] = limpar_ean_serie(df["EAN"])

    # 3) Carregar planilha de PMC
    pmc_df = pd.read_excel(planilha_pmc_path)

    # Descobre automaticamente se a coluna é "EAN 1" ou "EAN"
    ean_col_pmc = "EAN 1" if "EAN 1" in pmc_df.columns else "EAN"
    pmc_df[ean_col_pmc] = limpar_ean_serie(pmc_df[ean_col_pmc])
    pmc_df.rename(columns={ean_col_pmc: "EAN"}, inplace=True)

    # 4) Carregar planilha de PMPF
    pmpf_df = pd.read_excel(planilha_pmpf_path)
    ean_col_pmpf = "EAN 1" if "EAN 1" in pmpf_df.columns else "EAN"
    pmpf_df[ean_col_pmpf] = limpar_ean_serie(pmpf_df[ean_col_pmpf])
    pmpf_df.rename(columns={ean_col_pmpf: "EAN"}, inplace=True)

    # 5) Atualizar PMC (quando vier 0 do banco)
    if "PMC 18%" in pmc_df.columns:
        df = df.merge(pmc_df[["EAN", "PMC 18%"]], on="EAN", how="left")
        df["PMC"] = df.apply(
            lambda row: row["PMC 18%"] if float(row["PMC"]) == 0.0 else row["PMC"],
            axis=1,
        )
        df.drop(columns=["PMC 18%"], inplace=True)

    # 6) Atualizar PMPF (quando vier 0 do banco)
    if "PMPF" in pmpf_df.columns:
        df = df.merge(
            pmpf_df[["EAN", "PMPF"]],
            on="EAN",
            how="left",
            suffixes=("", "_new"),
        )
        df["PMPF"] = df.apply(
            lambda row: row["PMPF_new"] if float(row["PMPF"]) == 0.0 else row["PMPF"],
            axis=1,
        )
        df.drop(columns=["PMPF_new"], inplace=True)

    # 7) Salvar saídas
    data_atual = datetime.now().strftime("%d-%m-%Y")

    output_paths = [
        r"P:\TESTE\TESTE\ALERTA {data_atual}.xlsx",
        r"P:\TESTE\TESTE\ALERTA {data_atual}.xlsx",
        r"P:\TESTE\TESTE\TESTE\TESTE\TESTE\ALERTA {data_atual}.xlsx",
    ]

    for path_template in output_paths:
        path = path_template.format(data_atual=data_atual)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        df.to_excel(path, index=False)
        print(f"DADOS ATUALIZADOS FORAM SALVOS NO ARQUIVO: {path}")

    print("Processo concluído com sucesso.")


if __name__ == "__main__":
    main()
