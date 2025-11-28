import streamlit as st
import pandas as pd
from pathlib import Path

# ======================================
# CONFIG
# ======================================
st.set_page_config(page_title="Relatórios", layout="wide")
st.title("📊 Relatórios - Base Site")

# ======================================
# CAMINHO DO ARQUIVO
# ======================================
ARQUIVO = Path(
    r"C:\Users\2960007532\Documents\SITE STREAM LIT\Script Base Site\Site\Base site.xlsx"
)

# ======================================
# COLUNAS QUE VAMOS USAR
# ======================================
COLUNAS = [
    "Tipo de pedido",
    "Filial Destino",
    "oLPN",
    "Item",
    "Descrição",
    "Local de Picking",
    "Qtde. Peças Item",
    "Status oLPN",
    "BOX",
    "Audit",
    "Conferentes",
]

# ======================================
# CARREGAR ARQUIVO
# ======================================
if not ARQUIVO.exists():
    st.error(f"❌ Arquivo não encontrado!\n{ARQUIVO}")
    st.stop()

df = pd.read_excel(ARQUIVO, usecols=COLUNAS)

st.success("✅ Base carregada!")

# ======================================
# TABELA 1 — SEM AUDIT
# ======================================
st.subheader("📄 Tabela sem Audit")
df_sem = df.drop(columns=["Audit"])
st.dataframe(df_sem, use_container_width=True)

# ======================================
# TABELA 2 — COM AUDIT
# ======================================
st.subheader("📋 Tabela com Audit")
st.dataframe(df, use_container_width=True)
