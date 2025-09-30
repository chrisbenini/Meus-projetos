<h1 align="center">🧰 Meus-projetos (Python)</h1>

<p align="center">
  Scripts para o dia a dia no varejo: verificação de EAN, alertas e cálculo de margem.
</p>

<p align="center">
  <img alt="Repo size" src="https://img.shields.io/github/repo-size/chrisbenini/Meus-projetos">
  <img alt="Last commit" src="https://img.shields.io/github/last-commit/chrisbenini/Meus-projetos">
  <img alt="Issues" src="https://img.shields.io/github/issues/chrisbenini/Meus-projetos">
  <img alt="License" src="https://img.shields.io/badge/license-MIT-green">
  <img alt="Python" src="https://img.shields.io/badge/python-3.10%2B-blue">
</p>

## ✨ O que tem aqui
- **ALERTAS.py** – Ajuste de produtos sem **PMC/PMPF** e geração de alertas
- **EANS.py** – Identificação de **EANs inválidos/duplicados**  
- **Margem.py** – **Cálculo de margem** de produtos

> *Cada script é independente para ser usado em rotinas rápidas.*

---

## 🚀 Como usar (bem simples)
```bash
# 1) Clone
git clone https://github.com/chrisbenini/Meus-projetos.git
cd Meus-projetos

# 2) (Opcional) Crie um ambiente
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate

# 3) Instale dependências (se existir requirements.txt)
pip install -r requirements.txt

# 4) Rode o que precisar
python ALERTAS.py
python EANS.py
python Margem.py
