# IMPORTANDO AS BIBLIOTECAS QUE SERÃO USADAS NO CÓDIGO
from datetime import datetime
import pandas as pd
import os

# INFORMAÇÕES TABELA TESTE 
st_pad = 0.22
com_pis_pad = 0.3150
icms_sem_pis_pad = 0.40
icms_com_pis_pad = 0.4750
custo_operacional_pad = 0.12

# INFORMAÇÕES TABELA TESTE
st_mareb = 0.10
deb_cred = 0.20

# INFORMAÇÕES TABELA TESTE
st_cotef = 0.17
com_pis_cotef = 0.2650
icms_sem_pis_cotef = 0.35
icms_com_pis_cotef = 0.4250
custo_operacional_cotef = 0.12

# INFORMAÇÕES TABELA TESTE 
st_pe = 0.20
com_pis_pe = 0.2950
icms_sem_pis_pe = 0.38
icms_com_pis_pe = 0.4550
custo_operacional_pe = 0.12

# FUNÇÃO PARA CALCULAR OS VALORES DE ICMS, PIS, COFINS, COMISSÃO E MARGEM
def calcular_valores(df, custo_operacional, custo_col="CUSTO CAP"):
    
    # AGREGANDO PIS E COFINS
    df["PIS_COFS"] = df["PIS"] + df["COFINS"]  # Cria uma nova coluna com a soma do PIS e COFINS

    # CALCULO ICMS $
    df["ICMS VLR"] = df["VDA REAL C/ ST"] * df["ICMS VDA"]
    
    # CALCULO PIS + COFINS AGREGADOS NO PIS $
    df["PIS VLR"] = (df["VDA REAL C/ ST"] - df["ICMS VLR"]) * df["PIS_COFS"]  # Usa a nova coluna

    # CALCULO COMISSÃO $
    df["COMIS VLR"] = df["VDA REAL C/ ST"] * df["COMIS"]
    
    # CALCULO CUSTO OPERACIONAL $
    df["CUSTO OPERACIONAL VLR"] = df["VDA REAL C/ ST"] * custo_operacional
    
    # CALCULO MARGEM BRUTA
    df["MARGEM BRUTA"] = (df["VDA REAL C/ ST"] - df[custo_col]) / df["VDA REAL C/ ST"]
    
    # CALCULO MARGEM LÍQUIDA
    df["MARGEM LIQUIDA"] = (df["VDA REAL C/ ST"] - (df[custo_col] + df["ICMS VLR"] + df["PIS VLR"] + df["COMIS VLR"] + df["CUSTO OPERACIONAL VLR"])) / df["VDA REAL C/ ST"]
    
    # SITUAÇÃO COM BASE NA MARGEM LIQUIDA
    df["SITUAÇÃO"] = df["MARGEM LIQUIDA"].apply(lambda x: "AJUSTAR" if x < 0 else "REGULAR" if 0 <= x <= 0.02 else "BOM" if 0.03 <= x <= 0.04 else "EXCELENTE")

    # CALCULO DO PREÇO LÍQUIDO
    df["PREÇO LIQUIDO"] = df["VDA REAL C/ ST"] * df["MARGEM LIQUIDA"]
    
    # ADICIONA A COLUNA "CUSTO OPERACIONAL %" COM O VALOR DE 12% (0.12)
    df["CUSTO OPERACIONAL %"] = 0.12

    # ARREDONDAR OS VALORES NUMÉRICOS PARA 2 CASAS DECIMAIS
    #df = df.round(2)
    
    return df

# FUNÇÃO PARA FILTRAR OS PRODUTOS COM BASE NAS REGRAS DEFINIDAS
def filtrar_produtos(df):
    # FABRICANTES A SEREM REMOVIDOS COMPLETAMENTE   
    fabricantes_para_remover = ["TESTE", "TESTE", "TESTE", "TESTE"]

    # REMOVER PRODUTOS DOS FABRICANTES LISTADOS
    df = df[~df["FABRICANTE"].str.upper().isin(fabricantes_para_remover)]
    
    # REMOVER APENAS OS PRODUTOS 'TESTE' QUE TÊM O PREÇO 6,99 (CLASSICOS)
    df = df[~((df["FABRICANTE"].str.upper() == "TESTE") & (df["VDA REAL C/ ST"] == 6.99))]
    
    return df

def gerar_planilha_alerta(df, output_path, tipo="padrao"):
    if tipo == "mareb":
        df_alerta = df[(df["MARGEM LIQUIDA"] < 0)]
        df_alerta = df_alerta[["CÓDIGO", "PRODUTO", "FABRICANTE", "CUSTO BR", "VDA REAL C/ ST", "ICMS VDA", "PIS_COFS", "MARGEM BRUTA", "SITUAÇÃO"]]
    else:
        df_alerta = df[(df["MARGEM LIQUIDA"] < 0) | (df["MARGEM LIQUIDA"] <= 0.01)]
        df_alerta = df_alerta[[
            "CÓDIGO", "PRODUTO", "FABRICANTE", "EST DISP", "CUSTO CAP", "VDA REAL C/ ST", "ICMS VDA", "ICMS VLR", 
            "PIS_COFS", "PIS VLR", "COMIS", "COMIS VLR", "CUSTO OPERACIONAL %", "CUSTO OPERACIONAL VLR", 
            "MARGEM BRUTA", "MARGEM LIQUIDA", "PREÇO LIQUIDO", "SITUAÇÃO", "MENOR", "MEDIANA", "MAIOR"
        ]]

    # FILTRA OS PRODUTOS COM BASE NOS FABRICANTES E PREÇOS DEFINIDOS
    df_alerta = filtrar_produtos(df_alerta)    
    
    # FILTRA OS PRODUTOS COM VENCIMENTO CURTO (QUE TÊM "***" NO PRODUTO)
    df_vencimento_curto = df_alerta[df_alerta["PRODUTO"].str.contains("\*\*\*", na=False)]
    
    # REMOVE OS PRODUTOS COM VENCIMENTO CURTO DA PLANILHA DE ALERTA
    df_alerta = df_alerta[~df_alerta["PRODUTO"].str.contains("\*\*\*", na=False)]

    # ARREDONDAR VALORES E SALVAR A PLANILHA DE ALERTA
    df_alerta = df_alerta
    df_vencimento_curto = df_vencimento_curto
    
    # SALVAR AS PLANILHAS EM DUAS ABAS: "NORMAL" e "VNC CURTO"
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df_alerta.to_excel(writer, sheet_name="Alerta", index=False)
        df_vencimento_curto.to_excel(writer, sheet_name="VNC CURTO", index=False)

# FUNÇÃO PARA GERAR PLANILHA PRINCIPAL COM AS COLUNAS ESPECIFICADAS
def gerar_planilha_principal(dfs, output_path):
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        for df, sheet_name in dfs:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

# FUNÇÃO PARA EXTRAIR A DATA DO NOME DO ARQUIVO
def extrair_data_do_nome(nome_arquivo):
    # EXTRAI A DATA 
    try:
        data_str = nome_arquivo.split('_')[1].replace('.xlsx', '')
        return datetime.strptime(data_str, '%d-%m-%Y')
    except Exception as e:
        return None  # CASO NÃO CONSIGA EXTRAIR A DATA

# FUNÇÃO PARA LOCALIZAR O ARQUIVO DE COTAÇÃO MAIS RECENTE
def localizar_arquivo_mais_recente(diretorio, padrao_nome="COTACAO_"):
    arquivos = [f for f in os.listdir(diretorio) if f.startswith(padrao_nome) and f.endswith('.xlsx')]
    
    # FILTRA ARQUIVOS COM DATA VÁLIDA E ORDENA PELO MAIS RECENTE
    arquivos_com_data = [(arquivo, extrair_data_do_nome(arquivo)) for arquivo in arquivos]
    arquivos_com_data = [x for x in arquivos_com_data if x[1] is not None]  # REMOVE OS ARQUIVOS SEM DATA VÁLIDA
    
    # ORDENA PELA DATA (A MAIS RECENTE PRIMEIRO)
    arquivos_com_data.sort(key=lambda x: x[1], reverse=True)
    
    if arquivos_com_data:
        return os.path.join(diretorio, arquivos_com_data[0][0])  # RETORNA O ARQUIVO MAIS RECENTE
    else:
        raise FileNotFoundError("Nenhum arquivo de cotação recente encontrado no diretório.")

# FUNÇÃO PARA ASSOCIAR AS COLUNAS 'MENOR', 'MEDIANA', 'MAIOR' PELA COLUNA 'CÓDIGO'
def associar_colunas_cotacao(planilha_alerta, planilha_principal, caminho_cotacao):
    # CARREGA A PLANILHA DE COTAÇÃO
    cotacao = pd.read_excel(caminho_cotacao)
    print(cotacao.columns)

    # RENOMEIA A COLUNA "CODIGO" PARA "CÓDIGO" NA PLANILHA DE COTAÇÃO
    cotacao.rename(columns={'CODIGO': 'CÓDIGO'}, inplace=True)

    # GARANTE QUE AS COLUNAS DE CÓDIGO ESTÃO NO MESMO FORMATO
    planilha_alerta["CÓDIGO"] = planilha_alerta["CÓDIGO"].astype(str)
    planilha_principal["CÓDIGO"] = planilha_principal["CÓDIGO"].astype(str)
    cotacao["CÓDIGO"] = cotacao["CÓDIGO"].astype(str)

    # FAZ O MERGE TRAZENDO AS COLUNAS "MENOR", "MEDIANA" E "MAIOR"
    planilha_alerta = pd.merge(planilha_alerta, cotacao[["CÓDIGO", "MENOR", "MEDIANA", "MAIOR"]], on="CÓDIGO", how="left")
    planilha_principal = pd.merge(planilha_principal, cotacao[["CÓDIGO", "MENOR", "MEDIANA", "MAIOR"]], on="CÓDIGO", how="left")
    
    return planilha_alerta, planilha_principal

# FUNÇÃO PARA PROCESSAR AS PLANILHAS E GERAR AS TABELAS
def processar_planilhas():
    # CARREGAR AS PLANILHAS
    planilha_padrao = pd.read_excel(r"P:\\INTELIGENCIA\\COMERCIAL\\Simulador Margem\\Bases\\AdmPreço.xlsx")
    planilha_padrao.columns = planilha_padrao.columns.str.strip()  # REMOVE OS ESPAÇOS EXTRAS

    planilha_mareb = pd.read_excel(r"P:\\TESTE\\TESTE\\TESTE\\TESTE\\TESTE.xlsx")
    planilha_mareb.columns = planilha_mareb.columns.str.strip()  # REMOVE OS ESPAÇOS EXTRAS

    planilha_cotef = pd.read_excel(r"P:\\TESTE\\TESTE\\TESTE\\TESTE\\TESTE.xlsx")
    planilha_cotef.columns = planilha_cotef.columns.str.strip()  # REMOVE OS ESPAÇOS EXTRAS

    planilha_ped = pd.read_excel(r"P:\\TESTE\\TESTE\\TESTE\\TESTE\\TESTE.xlsx")
    planilha_ped.columns = planilha_ped.columns.str.strip()  # REMOVE OS ESPAÇOS EXTRAS

    # FILTRAR APENAS OS PRODUTOS COM ESTOQUE DISPONÍVEL
    planilha_padrao = planilha_padrao[planilha_padrao["EST DISP"] > 0]
    planilha_mareb = planilha_mareb[planilha_mareb["EST DISP"] > 0]
    planilha_cotef = planilha_cotef[planilha_cotef["EST DISP"] > 0]
    planilha_ped = planilha_ped[planilha_ped["EST DISP"] > 0]

    # EXCLUIR LINHAS ONDE O FABRICANTE É "BRINDES"
    planilha_padrao = planilha_padrao[planilha_padrao["FABRICANTE"].str.upper() != "TESTE"]
    planilha_mareb = planilha_mareb[planilha_mareb["FABRICANTE"].str.upper() != "TESTE"]
    planilha_cotef = planilha_cotef[planilha_cotef["FABRICANTE"].str.upper() != "TESTE"]
    planilha_ped = planilha_ped[planilha_ped["FABRICANTE"].str.upper() != "TESTE"]

    # CALCULAR VALORES PARA CADA PLANILHA
    planilha_padrao = calcular_valores(planilha_padrao, custo_operacional_pad)
    planilha_mareb = calcular_valores(planilha_mareb, custo_operacional_pad, custo_col="CUSTO BR")
    planilha_cotef = calcular_valores(planilha_cotef, custo_operacional_cotef)
    planilha_ped = calcular_valores(planilha_ped, custo_operacional_pe)

    # SELECIONAR APENAS AS COLUNAS ESPECIFICADAS PARA A PLANILHA PRINCIPAL
    planilha_padrao = planilha_padrao[[
        "CÓDIGO", "PRODUTO", "FABRICANTE", "EST DISP", "CUSTO CAP", "VDA REAL C/ ST", "ICMS VDA", "ICMS VLR", 
       "PIS_COFS", "PIS VLR", "COMIS", "COMIS VLR", "CUSTO OPERACIONAL %", "CUSTO OPERACIONAL VLR", 
        "MARGEM BRUTA", "MARGEM LIQUIDA", "PREÇO LIQUIDO", "SITUAÇÃO"
    ]]

    planilha_mareb = planilha_mareb[[
        "CÓDIGO", "PRODUTO", "FABRICANTE", "EST DISP", "CUSTO BR", "VDA REAL C/ ST", "ICMS VDA", "ICMS VLR", 
        "PIS_COFS", "PIS VLR", "COMIS", "COMIS VLR", "CUSTO OPERACIONAL %", "CUSTO OPERACIONAL VLR", 
        "MARGEM BRUTA", "MARGEM LIQUIDA", "PREÇO LIQUIDO", "SITUAÇÃO"
    ]]

    planilha_cotef = planilha_cotef[[
        "CÓDIGO", "PRODUTO", "FABRICANTE", "EST DISP", "CUSTO CAP", "VDA REAL C/ ST", "ICMS VDA", "ICMS VLR", 
        "PIS_COFS", "PIS VLR", "COMIS", "COMIS VLR", "CUSTO OPERACIONAL %", "CUSTO OPERACIONAL VLR", 
        "MARGEM BRUTA", "MARGEM LIQUIDA", "PREÇO LIQUIDO", "SITUAÇÃO"
    ]]

    planilha_ped = planilha_ped[[
        "CÓDIGO", "PRODUTO", "FABRICANTE", "EST DISP", "CUSTO CAP", "VDA REAL C/ ST", "ICMS VDA", "ICMS VLR", 
        "PIS_COFS", "PIS VLR", "COMIS", "COMIS VLR", "CUSTO OPERACIONAL %", "CUSTO OPERACIONAL VLR", 
        "MARGEM BRUTA", "MARGEM LIQUIDA", "PREÇO LIQUIDO", "SITUAÇÃO"
    ]]

    # LOCALIZAR O ARQUIVO DE COTAÇÃO MAIS RECENTE
    diretorio_cotacao = r"P:\\TESTE\\TESTE\\PRODUTOS_COTAÇÃO\\"
    caminho_cotacao = localizar_arquivo_mais_recente(diretorio_cotacao).strip()  # REMOVE OS ESPAÇOS EXTRAS

    # ASSOCIA AS COLUNAS "MENOR", "MEDIANA", "MAIOR" DA COTAÇÃO
    planilha_padrao, planilha_mareb = associar_colunas_cotacao(planilha_padrao, planilha_mareb, caminho_cotacao)
    planilha_cotef, planilha_ped = associar_colunas_cotacao(planilha_cotef, planilha_ped, caminho_cotacao)

    # NOME DO ARQUIVO PRINCIPAL COM DATA ATUAL
    data_atual = datetime.now().strftime("%d-%m-%Y")
    output_path_principal = f"P:\\TESTE\\TESTE\\TESTE\\TESTE\\TESTE\\Margens_produtos_{data_atual}.xlsx"

    # GERAR PLANILHA PRINCIPAL COM AS ABAS PARA CADA TABELA
    gerar_planilha_principal([
        (planilha_padrao, 'Padrão'),
        (planilha_mareb, 'Mare Beauty'),
        (planilha_cotef, 'Cotefácil'),
        (planilha_ped, 'Pedido Eletrônico')
    ], output_path_principal)

    # GERAR PLANILHAS DE ALERTA PARA CADA TABELA
    gerar_planilha_alerta(planilha_padrao, f"P:\\TESTE\\TESTE\\TESTE\\Checar_margem_padrao {data_atual}.xlsx")
    gerar_planilha_alerta(planilha_mareb, f"P:\\TESTE\\TESTE\\TESTE\\Checar_margem_mareb {data_atual}.xlsx", tipo="mareb")
    gerar_planilha_alerta(planilha_cotef, f"P:\\TESTE\\TESTE\\TESTE\\Checar_margem_cotef {data_atual}.xlsx")
    gerar_planilha_alerta(planilha_ped, f"P:\\TESTE\\TESTE\\TESTE\\Checar_margem_ped {data_atual}.xlsx")

    # MENSAGENS DE CONFIRMAÇÃO
    print(f"DADOS ATUALIZADOS FORAM SALVOS NO ARQUIVO: Checar_margem_padrao {data_atual}.xlsx")
    print(f"DADOS ATUALIZADOS FORAM SALVOS NO ARQUIVO: Checar_margem_mareb {data_atual}.xlsx")
    print(f"DADOS ATUALIZADOS FORAM SALVOS NO ARQUIVO: Checar_margem_cotef {data_atual}.xlsx")
    print(f"DADOS ATUALIZADOS FORAM SALVOS NO ARQUIVO: Checar_margem_ped {data_atual}.xlsx")

# EXECUTAR O PROCESSAMENTO
processar_planilhas()


