import streamlit as st
from openai import OpenAI
import os

# ==================== LER CHAVE API ====================
CHAVE_PATH = r"C:\Users\2960007532\Documents\SITE STREAM LIT\Chave.txt"
with open(CHAVE_PATH, "r") as f:
    api_key = f.read().strip()

client = OpenAI(api_key=api_key)

# ==================== ESTILO ====================
def estilo():
    st.markdown("""
        <style>
            body { background-color: black; color: white; }
            [data-testid="stSidebar"] { background-color: #111; color: white; }
            [data-testid="stSidebar"] * { color: white !important; }

            .msg-user {
                background-color: #1e1e1e;
                padding: 10px;
                border-radius: 10px;
                text-align: right;
                margin-bottom: 8px;
            }
            .msg-bot {
                background-color: #333;
                padding: 10px;
                border-radius: 10px;
                margin-bottom: 8px;
            }
        </style>
    """, unsafe_allow_html=True)

# ==================== FUNÇÃO DE RESPOSTA ====================
def processar_mensagem(mensagem):
    if not mensagem.strip():
        return "Digite alguma coisa 😅"

    # Chamando a API da OpenAI
    try:
        resposta = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": mensagem}],
            temperature=0.7,
            max_tokens=500
        )
        texto = resposta.choices[0].message.content.strip()
        return texto
    except Exception as e:
        return f"❌ Erro na IA: {e}"

# ==================== INTERFACE ====================
def main():
    estilo()
    st.title("🤖 ChatBot / IA")

    if "chat" not in st.session_state:
        st.session_state.chat = []

    # Mostrar histórico
    for msg, tipo in st.session_state.chat:
        if tipo == "user":
            st.markdown(f"<div class='msg-user'>{msg}</div>", unsafe_allow_html=True)
        else:
            st.markdown(f"<div class='msg-bot'>{msg}</div>", unsafe_allow_html=True)

    # Campo de entrada
    usuario = st.text_input("Digite sua mensagem:")

    if st.button("Enviar"):
        if usuario.strip():
            st.session_state.chat.append((usuario, "user"))
            resposta = processar_mensagem(usuario)
            st.session_state.chat.append((resposta, "bot"))
            st.rerun()

if __name__ == "__main__":
    main()
