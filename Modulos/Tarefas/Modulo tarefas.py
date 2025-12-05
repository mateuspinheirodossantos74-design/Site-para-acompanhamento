import streamlit as st

st.set_page_config(page_title="Módulo Tarefas", layout="wide")

st.title("🗂️ Módulo Tarefas")

# PASSO 1 — MENU SUSPENSO
menu = st.selectbox(
    "Selecione o modelos de coleta:",
    ["GRUPO", "FULL / PULL", "REP"]
)

st.write(f"Você selecionou: **{menu}**")
