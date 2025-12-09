<h1 align="center">🧰 Meus-projetos (Python)</h1>

<p align="center">
  Scripts em Python para o dia a dia no varejo: verificação de EAN, alertas, cálculo de margem e conversão de XML de NF-e.
</p>

<p align="center">
  <img alt="repo size" src="https://img.shields.io/github/repo-size/chrisbenini/Meus-projetos?style=flat-square">
  <img alt="last commit" src="https://img.shields.io/github/last-commit/chrisbenini/Meus-projetos?style=flat-square">
  <img alt="issues" src="https://img.shields.io/github/issues/chrisbenini/Meus-projetos?style=flat-square">
  <img alt="license" src="https://img.shields.io/github/license/chrisbenini/Meus-projetos?style=flat-square">
  <img alt="python" src="https://img.shields.io/badge/python-3.10%2B-blue?style=flat-square&logo=python">
</p>

---

## ✨ O que tem aqui

- **`ALERTAS.py`** – Ajuste de produtos sem PMC/PMPF, cruzando base SQL + planilhas de PMC/PMPF e gerando planilhas de alerta em Excel.  
- **`EANS.py`** – Identificação de EANs inválidos (diferente de 13 dígitos) e geração de relatório de inconsistências.  
- **`Margem.py`** – Cálculo de margem bruta e líquida por produto, com planilhas de alerta e comparação com cotações.  
- **`convert_Xml.py`** – Aplicação desktop (GUI) que converte XML de NF-e em planilha Excel pronta para análise fiscal.

> Cada script é independente e pode ser usado em rotinas rápidas do dia a dia.

---

## 📂 Estrutura do repositório

```text
Meus-projetos/
├─ ALERTAS.py          # Ajuste de produtos sem PMC/PMPF e geração de alertas
├─ EANS.py             # Validação de EANs (dígitos incorretos)
├─ Margem.py           # Cálculo de margens (bruta e líquida) e planilhas de alerta
├─ convert_Xml.py      # Aplicação desktop para converter XML de NF-e em Excel
├─ requirements.txt    # Dependências do projeto
└─ README.md           # Esta documentação
