# Novosite - Módulo "Início"
# Autor: Mateus Pinheiro
# Descrição: Módulo principal "Início" do site Novosite.
# Exibe cartões de status, filtros hierárquicos e botões de exportação de status.

import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path

# ----------------------
# TEMA VISUAL + CSS RESPONSIVO
# ----------------------
def _set_theme_black():
    """Tema preto com textos brancos, elementos ajustados e responsivos."""
    st.markdown(
        """
        <style>
        /* Fundo e texto */
        .stApp { background-color: #000000; color: #FFFFFF; }

        /* Cartões */
        .card { 
            background-color: rgba(255,255,255,0.05); 
            border-radius: 15px; 
            padding: 18px; 
            text-align: center; 
            box-shadow: 0 4px 15px rgba(0,0,0,0.5); 
        }
        .status-label { font-size: 28px; font-weight: bold; }
        .value-label { font-size: 28px; font-weight: bold; }
        .small-muted { color: rgba(255,255,255,0.8); font-size: 18px; }

        /* Botão principal */
        div.stButton > button:first-child {
            background-color: #1f1f1f;
            color: white;
            font-size: 18px;
            height: 50px;
            border-radius: 10px;
            border: 1px solid white;
            width: 100%;
        }

        /* Botões de download */
        div.stDownloadButton > button {
            background-color: #1f1f1f !important;
            color: white !important;
            border: 1px solid white !important;
            border-radius: 10px !important;
            padding: 10px 20px !important;
            font-size: 16px !important;
            width: 100% !important;
            transition: 0.3s;
        }
        div.stDownloadButton > button:hover {
            background-color: #333333 !important;
            border-color: #cccccc !important;
        }

        /* Labels e selects */
        label, .stMarkdown, div[data-baseweb="select"] span {
            color: white !important;
        }

        /* Tabela responsiva */
        .dataframe {
            overflow-x: auto;
            display: block;
        }

        /* -------------------- */
        /* RESPONSIVIDADE CELULAR */
        /* -------------------- */
        @media (max-width: 600px) {
            .status-label { font-size: 20px; }
            .value-label { font-size: 20px; }
            .small-muted { font-size: 14px; }
            h1, h2, h3 { font-size: 22px !important; text-align: center; }
            div.stButton > button:first-child { font-size: 16px; height: 45px; }
            div.stDownloadButton > button { font-size: 14px !important; padding: 8px 10px !important; }
            .card { padding: 10px; margin-bottom: 10px; }
            .stSelectbox, .stTextInput { font-size: 14px !important; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

def load_base_excel(path_file: str):
    """Lê a planilha Base site.xlsx com card estilizado quando o arquivo está em uso."""
    try:
        return pd.read_excel(path_file, sheet_name='BASE')

    except PermissionError:
        # === CSS DO CARD ESCURO + ANIMAÇÃO ===
        st.markdown("""
        <style>
            .erro-card {
                padding: 20px;
                border-radius: 12px;
                background: rgba(255,255,255,0.07);
                border: 1px solid rgba(255,255,255,0.15);
                text-align: center;
                color: white;
                margin-top: 20px;
                animation: pulse 1.5s infinite;
            }

            @keyframes pulse {
                0% { box-shadow: 0 0 0 rgba(255,255,255,0.1); }
                50% { box-shadow: 0 0 20px rgba(255,255,255,0.25); }
                100% { box-shadow: 0 0 0 rgba(255,255,255,0.1); }
            }

            .reload-text {
                margin-top: 15px;
                color: #cccccc;
                font-size: 14px;
            }

            .reload-btn {
                background-color: #1f1f1f !important;
                border: 1px solid white !important;
                color: white !important;
                padding: 8px 18px !important;
                border-radius: 8px !important;
                margin-top: 10px;
                font-size: 16px !important;
            }
        </style>
        """, unsafe_allow_html=True)

        # === CARD ESTILIZADO ===
        st.markdown("""
        <div class="erro-card">
            <h3>🚫 Base em atualização</h3>
            <p>A planilha está sendo atualizada pelo sistema automático.</p>
            <p>Aguarde um momento e tente novamente.</p>
            <div class="reload-text">
                A página será recarregada automaticamente em <b>10 segundos</b>.
            </div>
        </div>
        """, unsafe_allow_html=True)

        # === BOTÃO EXTRA MANUAL ===
        st.button("🔄 Atualizar agora", key="manual_reload", help="Clique para tentar novamente", type="primary")

        # === AUTO-RELOAD AUTOMÁTICO (10 segundos) ===
        st.markdown("""
            <meta http-equiv="refresh" content="10">
        """, unsafe_allow_html=True)

        st.stop()

    except Exception as e:
        st.error(f"⚠️ Erro inesperado ao carregar a base:<br>{e}", unsafe_allow_html=True)
        st.stop()



# ----------------------
# DOWNLOAD HELPER
# ----------------------
def gerar_excel_download(df, nome_arquivo):
    """Gera arquivo Excel e retorna como botão de download."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados')
    dados = output.getvalue()
    st.download_button(
        label=f"⬇️ Baixar {nome_arquivo}",
        data=dados,
        file_name=f"{nome_arquivo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# ----------------------
# FUNÇÃO PRINCIPAL
# ----------------------
def render_inicio(path_base_excel: str | Path):
    _set_theme_black()

    # Título
    st.markdown("<h1 class='centralizado'>CD 1200</h1>", unsafe_allow_html=True)
    st.markdown("---")

    # ----------------------
    # LOGO E BOTÃO ATUALIZAR SITE
    # ----------------------
    col1, col2, col3 = st.columns([6, 2, 2])
    with col3:
        st.image(r"C:\Users\2960007532\Documents\SITE STREAM LIT\Imagens\Logo.png", use_container_width=True)
        atualizar = st.button("🔄 Atualizar site")

    # Carregar base
    if "df_base" not in st.session_state or atualizar:
        st.session_state.df_base = load_base_excel(path_base_excel)

    df = st.session_state.df_base

    if df.empty:
        st.warning("Base vazia ou não carregada.")
        return

    # ----------------------
    # CONFIGURAÇÃO DE FILTROS
    # ----------------------
    col_setor = df.iloc[:, 31]      # Setor
    col_demanda = df.iloc[:, 34]    # Demanda
    col_filial = df.iloc[:, 6]      # Filial Destino
    col_box = df.iloc[:, 23]        # Box
    col_data = df.iloc[:, 26]       # Data Limite Expedição

    try:
        col_data = pd.to_datetime(col_data, errors='coerce').dt.date
    except:
        pass

    st.markdown("### 🔍 Filtros")

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown("<span style='color:white'>Data Limite Expedição</span>", unsafe_allow_html=True)
        filtro_data = st.selectbox("", ['Todas'] + sorted([d for d in col_data.dropna().unique()]))

    df_filt = df.copy()
    if filtro_data != 'Todas':
        df_filt = df_filt[col_data == filtro_data]

    col_setor_filt = df_filt.iloc[:, 31]
    col_demanda_filt = df_filt.iloc[:, 34]
    col_filial_filt = df_filt.iloc[:, 6]
    col_box_filt = df_filt.iloc[:, 23]

    with c2:
        st.markdown("<span style='color:white'>Setor</span>", unsafe_allow_html=True)
        filtro_setor = st.selectbox("", ['Todos'] + col_setor_filt.dropna().unique().tolist())

    df_filt2 = df_filt.copy()
    if filtro_setor != 'Todos':
        df_filt2 = df_filt2[col_setor_filt == filtro_setor]

    with c3:
        st.markdown("<span style='color:white'>Demanda</span>", unsafe_allow_html=True)
        filtro_demanda = st.selectbox("", ['Todos'] + df_filt2.iloc[:, 34].dropna().unique().tolist())

    with c4:
        st.markdown("<span style='color:white'>Filial Destino</span>", unsafe_allow_html=True)
        filtro_filial = st.selectbox("", ['Todos'] + df_filt2.iloc[:, 6].dropna().unique().tolist())

    with c5:
        st.markdown("<span style='color:white'>Box</span>", unsafe_allow_html=True)
        try:
            box_ordenado = sorted(df_filt2.iloc[:, 23].dropna().unique(), key=lambda x: int(''.join(filter(str.isdigit, str(x))) or 0))
        except:
            box_ordenado = df_filt2.iloc[:, 23].dropna().unique().tolist()
        filtro_box = st.selectbox("", ['Todos'] + box_ordenado)

    if filtro_demanda != 'Todos':
        df_filt2 = df_filt2[df_filt2.iloc[:, 34] == filtro_demanda]
    if filtro_filial != 'Todos':
        df_filt2 = df_filt2[df_filt2.iloc[:, 6] == filtro_filial]
    if filtro_box != 'Todos':
        df_filt2 = df_filt2[df_filt2.iloc[:, 23] == filtro_box]

    # ----------------------
    # CARTÕES DE STATUS
    # ----------------------
    st.markdown("---")
    st.markdown("<h2 class='centralizado'>Status por OLPNS</h2>", unsafe_allow_html=True)

    status_col = df_filt2.iloc[:, 21]
    qty_col = df_filt2.iloc[:, 19]
    status_list = ['Created', 'Loaded', 'Packed', 'Shipped']

    total_pecas = qty_col.sum()
    total_olpns = len(df_filt2)

    cols = st.columns(5)

    # Cartões dos status
    for i, status in enumerate(status_list):
        mask = status_col == status
        total_qty = qty_col[mask].sum()
        total_rows = mask.sum()
        total_fmt = f"{total_qty:,.0f}".replace(",", ".")
        cols[i].markdown(
            f"<div class='card'>"
            f"<div class='status-label'>{status}</div>"
            f"<div class='value-label'>{total_fmt} peças</div>"
            f"<div class='small-muted'>{total_rows} OLPNS</div>"
            f"</div>",
            unsafe_allow_html=True
        )

    # Total no final
    total_fmt = f"{total_pecas:,.0f}".replace(",", ".")
    cols[-1].markdown(
        f"<div class='card'>"
        f"<div class='status-label'>TOTAL</div>"
        f"<div class='value-label'>{total_fmt} peças</div>"
        f"<div class='small-muted'>{total_olpns} OLPNS</div>"
        f"</div>",
        unsafe_allow_html=True
    )

    # ----------------------
    # EXPORTAR EXCEL POR STATUS
    # ----------------------
    st.markdown("### 📁 Exportar dados filtrados")
    exp_col1, exp_col2, exp_col3 = st.columns(3)

    for col, status in zip([exp_col1, exp_col2, exp_col3], ['Created', 'Packed', 'Loaded']):
        with col:
            df_status = df_filt2[df_filt2.iloc[:, 21] == status]
            if not df_status.empty:
                gerar_excel_download(df_status, status)
            else:
                st.markdown(f"<div class='small-muted'>Sem dados de {status}</div>", unsafe_allow_html=True)

    # ----------------------
    # RODAPÉ
    # ----------------------
    st.markdown("---")
    st.markdown(
        "<div class='centralizado'>Criado por Mateus Pinheiro - Grupo Casas Bahia</div>",
        unsafe_allow_html=True,
    )

# ----------------------
# EXECUÇÃO STANDALONE
# ----------------------
if __name__ == "__main__":
    st.set_page_config(page_title="Novosite - Início", layout="wide")
    render_inicio(r"C:\Users\2960007532\Documents\SITE STREAM LIT\Script Base Site\Site\Base site.xlsx")

