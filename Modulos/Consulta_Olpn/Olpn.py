import streamlit as st
import pandas as pd
import base64
import os

# Define o título da aba e o ícone do site
st.set_page_config(page_title="Consulta Olpn", page_icon="🔍")

# ======== CAMINHO DA BASE ========
CAMINHO_BASE = r"C:\Users\2960007532\Documents\SITE STREAM LIT\Script Base Site\Site\Base site.xlsx"

# ======== ESTILO VISUAL (MESMO DO INÍCIO) ========
def estilo():
    st.markdown("""
        <style>
            body { background-color: black; color: white; }
            [data-testid="stSidebar"] { background-color: #111; color: white; }
            [data-testid="stSidebar"] * { color: white !important; }

            .stButton>button {
                background-color: #222;
                color: white;
                border: 1px solid white;
                border-radius: 10px;
                width: 100%;
                margin-bottom: 5px;
            }
            .stButton>button:hover {
                background-color: #444;
            }

            input {
                background-color: rgba(0,0,0,0.7) !important;
                color: white !important;
            }

            .bloco {
                padding: 15px;
                background-color: #1a1a1a;
                border-radius: 10px;
                border: 1px solid #333;
                margin-top: 15px;
            }
        </style>
    """, unsafe_allow_html=True)

# ======== FUNÇÃO PRINCIPAL ========
def consulta_olpn():

    estilo()
    st.title("🔍 Consulta oLPN")

    # Entrada do usuário
    olpn = st.text_input("Digite ou bip o oLPN:")

    if st.button("Consultar"):

        if not olpn.strip():
            st.warning("Digite um oLPN válido.")
            return

        try:
            df = pd.read_excel(CAMINHO_BASE)

            # Garantir que oLPN está como string
            df["oLPN"] = df["oLPN"].astype(str)
            olpn = str(olpn).strip()

            # Pesquisar na coluna K (oLPN)
            resultado = df.loc[df["oLPN"] == olpn]

            if resultado.empty:
                st.error("❌ oLPN não encontrado na base.")
                return

            # Pegar a linha encontrada
            info = resultado.iloc[0]

            # Mostrar o resultado formatado
            st.markdown("### ✅ Resultado encontrado")

            st.markdown(
                f"""
                <div class="bloco">
                    <h4>📦 Informações do oLPN</h4>
                    <b>Item:</b> {info['Item']}<br>
                    <b>Descrição:</b> {info['Descrição']}<br>
                    <b>Quantidade de Peças:</b> {info['Qtde. Peças Item']}<br>
                    <b>Audit:</b> {info['Audit']}<br>
                    <b>BOX:</b> {info['BOX']}<br>
                    
                </div>
                """,
                unsafe_allow_html=True
            )

        except Exception as e:
            st.error(f"Erro ao carregar a base: {e}")


# Executa caso for rodado sozinho
if __name__ == "__main__":
    consulta_olpn()
