import xml.etree.ElementTree as ET
from tkinter import filedialog, messagebox
from pathlib import Path
from typing import List, Dict, Optional

import customtkinter as ctk
import pandas as pd


NFE_NAMESPACE = {"ns": "http://www.portalfiscal.inf.br/nfe"}


def extrair_informacoes_xml(xml_path: Path) -> pd.DataFrame:
    """
    Lê um arquivo XML de NF-e e extrai as informações dos itens
    para um DataFrame do pandas.

    :param xml_path: Caminho do arquivo XML.
    :return: DataFrame com os dados dos produtos.
    """
    try:
        tree = ET.parse(xml_path)
    except ET.ParseError as exc:
        raise ValueError(f"Erro ao ler o XML: {exc}") from exc

    root = tree.getroot()

    dados: List[Dict[str, Optional[float]]] = []

    for item in root.findall(".//ns:det", NFE_NAMESPACE):
        prod = item.find("ns:prod", NFE_NAMESPACE)
        imposto = item.find("ns:imposto", NFE_NAMESPACE)

        if prod is None:
            continue

        # Campos básicos do produto
        ean = _get_text(prod, "ns:cEAN")
        produto = _get_text(prod, "ns:xProd")
        ncm = _get_text(prod, "ns:NCM")
        cest = _get_text(prod, "ns:CEST")

        quantidade = _get_float(prod, "ns:qTrib")
        valor_total = _get_float(prod, "ns:vProd")
        desconto = _get_float(prod, "ns:vDesc")
        vlr_unit = _get_float(prod, "ns:vUnTrib")

        # Impostos
        ipi = _get_float_from_any(imposto, ".//ns:vIPI")
        vlr_total_st = _get_float_from_any(imposto, ".//ns:vICMSST")

        # Cálculos
        valor_unitario_st = vlr_total_st / quantidade if quantidade > 0 else 0.0
        valor_total_sem_st = valor_total
        valor_liquido = valor_total_sem_st - desconto
        valor_unitario_ipi = ipi / quantidade if quantidade > 0 else 0.0
        valor_total_liquido = valor_liquido + vlr_total_st + ipi
        valor_unitario_total = (
            valor_total_liquido / quantidade if quantidade > 0 else 0.0
        )

        linha = {
            "EAN": ean,
            "PRODUTO": produto,
            "NCM": ncm,
            "CEST": cest,
            "QUANTIDADE": quantidade,
            "VALOR UNITARIO": vlr_unit,
            "VALOR BRUTO": valor_total_sem_st,
            "DESCONTO": desconto,
            "VALOR SEM ST": valor_liquido,
            "VALOR UNITARIO ST": valor_unitario_st,
            "VALOR TOTAL ST": vlr_total_st,
            "VALOR UNITARIO IPI": valor_unitario_ipi,
            "VALOR TOTAL IPI": ipi,
            "VALOR TOTAL LIQUIDO": valor_total_liquido,
            "VALOR UNITARIO LIQUIDO": valor_unitario_total,
        }

        dados.append(linha)

    df = pd.DataFrame(dados)
    df = df.round(2)
    return df


def _get_text(element: ET.Element, path: str) -> Optional[str]:
    """Obtém o texto de um nó XML, retornando None se não existir."""
    child = element.find(path, NFE_NAMESPACE)
    return child.text if child is not None else None


def _get_float(element: ET.Element, path: str) -> float:
    """Obtém um valor float de um nó XML, retornando 0.0 se não existir."""
    child = element.find(path, NFE_NAMESPACE)
    if child is None or child.text is None:
        return 0.0
    try:
        return float(child.text.replace(",", "."))
    except ValueError:
        return 0.0


def _get_float_from_any(element: Optional[ET.Element], path: str) -> float:
    """Busca um valor float em um caminho a partir de um elemento, se existir."""
    if element is None:
        return 0.0
    child = element.find(path, NFE_NAMESPACE)
    if child is None or child.text is None:
        return 0.0
    try:
        return float(child.text.replace(",", "."))
    except ValueError:
        return 0.0


def executar(app: ctk.CTk) -> None:
    """
    Fluxo principal:
    1. Seleciona XML de NF-e.
    2. Extrai dados e converte para Excel.
    3. Solicita local para salvar.
    """
    arquivo_xml = filedialog.askopenfilename(
        title="Selecione o arquivo XML da NF-e",
        filetypes=[("Arquivos XML", "*.xml")],
    )

    if not arquivo_xml:
        return

    try:
        df_resultado = extrair_informacoes_xml(Path(arquivo_xml))
    except ValueError as exc:
        messagebox.showerror("Erro ao processar XML", str(exc))
        return
    except Exception as exc:  # fallback genérico
        messagebox.showerror("Erro inesperado", str(exc))
        return

    if df_resultado.empty:
        messagebox.showwarning(
            "Nenhum item encontrado",
            "Não foram encontrados itens de produtos na NF-e selecionada.",
        )
        return

    arquivo_excel = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Arquivos Excel", "*.xlsx")],
        title="Salvar planilha Excel como",
    )

    if not arquivo_excel:
        return

    try:
        df_resultado.to_excel(arquivo_excel, index=False)
    except Exception as exc:
        messagebox.showerror("Erro ao salvar Excel", str(exc))
        return

    messagebox.showinfo("Sucesso", "Documento Excel gerado com sucesso!")
    app.quit()


def main() -> None:
    """Inicializa e executa a aplicação desktop."""
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")

    app = ctk.CTk()
    app.title("Importador de XML de NF-e para Excel")
    app.geometry("800x400")
    app.resizable(False, False)
    app.minsize(800, 400)
    app.maxsize(800, 400)
    app.configure(fg_color="black")

    botao_importar = ctk.CTkButton(
        app,
        text="Importar XML e Gerar Excel",
        command=lambda: executar(app),
        fg_color="gray",
        text_color="white",
        font=ctk.CTkFont(size=16, weight="bold"),
        width=240,
        height=48,
    )
    botao_importar.place(relx=0.5, rely=0.5, anchor="center")

    app.mainloop()


if __name__ == "__main__":
    main()
