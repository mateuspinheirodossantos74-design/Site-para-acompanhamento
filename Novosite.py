import streamlit as st
import pandas as pd
import base64
from datetime import datetime
from io import BytesIO
import os
import time
from streamlit_autorefresh import st_autorefresh


# ======== CONFIGURAÇÕES DA PÁGINA ========
st.set_page_config(page_title="Novo Site - Setor", layout="wide")

# ======== CAMINHOS ========
caminho_imagem = r"C:\Users\2960007532\Documents\SITE STREAM LIT\Imagens\2.png"
caminho_excel = r"C:\Users\2960007532\Documents\SITE STREAM LIT\Matriculas ADM\Matriculas.xlsm"
pasta_sugestoes = r"\\10.129.10.6\cd1200\share2\abasteci0200\MATEUS\Sugestoes Site Stream Lt"
arquivo_sugestoes = os.path.join(pasta_sugestoes, "Sugestoes.xlsx")
os.makedirs(pasta_sugestoes, exist_ok=True)

# ======== FUNÇÃO ESTILO ========
def estilo_geral(fundo_imagem=False):
    if fundo_imagem:
        with open(caminho_imagem, "rb") as img_file:
            img_base64 = base64.b64encode(img_file.read()).decode()
        st.markdown(f"""
            <style>
                body {{
                    background: linear-gradient(rgba(0, 0, 0, 0.6), rgba(0,0,0,0.6)),
                                url("data:image/png;base64,{img_base64}") no-repeat center center fixed;
                    background-size: cover;
                    color: white;
                }}
                .stButton>button {{
                    background-color: #222;
                    color: white;
                    border-radius: 10px;
                    width: 100%;
                }}
                .stButton>button:hover {{
                    background-color: #444;
                }}
                input {{
                    background-color: rgba(0,0,0,0.7) !important;
                    color: white !important;
                }}
            </style>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
            <style>
                body { background-color: black; color: white; }
                [data-testid="stSidebar"] { background-color: #111; color: white; }
                [data-testid="stSidebar"] * { color: white !important; }
                .stButton>button { background-color: #222; color: white; border: 1px solid white;
                                   border-radius: 10px; width: 100%; margin-bottom: 5px; }
                .stButton>button:hover { background-color: #444; }
                input { background-color: rgba(0,0,0,0.7) !important; color: white !important; }
                @media print {
                    body { background-color: white !important; color: black !important; }
                    table, th, td { color: black !important; border: 1px solid #000; }
                    [data-testid="stSidebar"], [data-testid="stToolbar"], [data-testid="stHeader"] { display: none !important; }
                }
            </style>
        """, unsafe_allow_html=True)

# ======== FUNÇÃO LOGIN ========
def verificar_login(matricula, senha):
    try:
        df = pd.read_excel(caminho_excel, sheet_name="Matriculas")
        df.rename(columns={"MATRICULAS ADM": "Matricula", "NOME": "Nome", "Senha": "Senha"}, inplace=True)
        df["Matricula"] = df["Matricula"].astype(str)
        usuario = df.loc[df["Matricula"] == matricula]
        if not usuario.empty and str(usuario["Senha"].values[0]) == senha:
            return usuario["Nome"].values[0]
        else:
            return None
    except Exception as e:
        st.error(f"Erro ao verificar login: {e}")
        return None

# ======== CONTROLE DE SESSÃO ========
if "logado" not in st.session_state:
    st.session_state.logado = False
if "usuario" not in st.session_state:
    st.session_state.usuario = ""
if "matricula" not in st.session_state:
    st.session_state.matricula = ""

# ======== TELA DE LOGIN ========
if not st.session_state.logado:
    estilo_geral(fundo_imagem=True)
    st.markdown("<h1 style='text-align: center;'>🔐 Login do Sistema</h1>", unsafe_allow_html=True)
    input_matricula = st.text_input("Matrícula", key="input_matricula")
    input_senha = st.text_input("Senha", type="password", key="input_senha")
    if st.button("Entrar"):
        nome = verificar_login(input_matricula, input_senha)
        if nome:
            st.session_state.logado = True
            st.session_state.usuario = nome
            st.session_state.matricula = input_matricula
            st.rerun()
        else:
            st.error("Matrícula ou senha incorretas.")
    st.stop()

# ======== INTERFACE PRINCIPAL ========
estilo_geral(fundo_imagem=False)

# ======== RELÓGIO AO TOPO ========
st_autorefresh(interval=1000, key="clock_refresh")
col1, col2 = st.columns([4,1])
with col1:
    st.markdown(f"### Bem-vindo, **{st.session_state.usuario}** 👋")
with col2:
    st.markdown(f"🕒 {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

# ======== MENU LATERAL ========
st.sidebar.title("📋 Menu Principal")
menu = st.sidebar.radio(
    "Navegar para:",
    ["Início", "Acompanhamento", "Consulta oLPN", "Produtividade",
     "Relatórios", "Indicadores", "Tarefas", "Conferência", "Sugestões",
     "Chat bot / IA"]
)


st.sidebar.markdown("---")
if st.sidebar.button("🔄 Atualizar site"):
    st.rerun()
if st.sidebar.button("🚪 Sair"):
    st.session_state.logado = False
    st.session_state.usuario = ""
    st.session_state.matricula = ""
    st.rerun()

# ======== CONTEÚDO DAS TELAS ========
if menu == "Início":
    st.title("🏠 Início")
    st.write(f"Bem-vindo, **{st.session_state.usuario}**!")
    st.write("Aqui é o painel inicial do sistema do setor.")
    st.write("⚠️ Algumas funcionalidades ainda estão em desenvolvimento.")

elif menu == "Acompanhamento":
    st.title("📊 Acompanhamento de Processos")
    st.write("Visualize o status atualizado dos processos em andamento.")
    st.write("⚠️ Tela em desenvolvimento...")

elif menu == "Consulta oLPN":
    st.title("🔍 Consulta oLPN")
    st.write("Ferramenta para consultar OLPNs no sistema.")
    st.write("⚠️ Funcionalidade em desenvolvimento...")

elif menu == "Produtividade":
    st.title("📈 Produtividade")
    st.write("Painel de produtividade da equipe.")
    st.write("⚠️ Funcionalidade em desenvolvimento...")
    
elif menu == "CHAT BOT / IA":
    st.title("🤖 Chat Bot / IA")
    st.write("Assistente virtual para ajudar na navegação e tirar dúvidas sobre o sistema.")
    st.write("⚠️ Funcionalidade em desenvolvimento...")

elif menu == "Relatórios":
    st.title("📄 Relatórios")
    uploaded_file = st.file_uploader("Selecione o arquivo CSV ou Excel", type=["csv", "xls", "xlsx"])
    if uploaded_file:
        try:
            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file, sep=";", encoding="latin1")
            else:
                df = pd.read_excel(uploaded_file)
            
            # Ajustes de formato
            df = df.applymap(lambda x: str(int(x)) if isinstance(x,float) and x.is_integer() else x)
            if "Descrição" in df.columns:
                df["Descrição"] = df["Descrição"].astype(str).str.strip()
            if "Tipo de Pedido" in df.columns:
                df["Tipo de Pedido"] = df["Tipo de Pedido"].astype(str).str.strip()

            st.dataframe(df, use_container_width=True)

            # Botão para baixar a versão formatada
            buffer = BytesIO()
            df.to_excel(buffer, index=False)
            buffer.seek(0)
            st.download_button(
                label="💾 Baixar planilha formatada",
                data=buffer,
                file_name="Relatorio_Formatado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erro ao carregar o arquivo: {e}")

elif menu == "Indicadores":
    st.title("📊 Indicadores de Desempenho")
    st.write("Visualize os principais KPIs e métricas do setor.")
    st.write("⚠️ Funcionalidade em desenvolvimento...")

elif menu == "Tarefas":
    st.title("📝 Controle de Tarefas")
    st.write("Gerencie e acompanhe as tarefas da equipe.")
    st.write("⚠️ Tela em desenvolvimento...")

elif menu == "Conferência":
    st.title("✅ Conferência de Dados")
    st.write("Ferramentas para revisar e validar informações.")
    st.write("⚠️ Funcionalidade em desenvolvimento...")

elif menu == "Sugestões":
    st.title("💡 Envie uma Sugestão")
    sugestao = st.text_area("Digite sua sugestão:")
    if st.button("Enviar Sugestão"):
        if sugestao.strip():
            if os.path.exists(arquivo_sugestoes):
                df_sug = pd.read_excel(arquivo_sugestoes)
            else:
                df_sug = pd.DataFrame(columns=["Data","Usuario","Matrícula","Sugestao"])
            nova_linha = {
                "Data": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                "Usuario": st.session_state.usuario,
                "Matrícula": st.session_state.matricula,
                "Sugestao": sugestao.strip()
            }
            df_sug = pd.concat([df_sug, pd.DataFrame([nova_linha])], ignore_index=True)
            df_sug.to_excel(arquivo_sugestoes, index=False)
            st.success("Sugestão enviada e salva com sucesso! ✅")
        else:
            st.warning("Digite algo antes de enviar.")

# ======== RODAPÉ ========
st.markdown("---")
st.caption(f"© {datetime.now().year} - Sistema do Setor | Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
