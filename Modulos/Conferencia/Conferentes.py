import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path

# ============================
# CONFIG / PATH
# ============================
st.set_page_config(page_title="Conferentes", page_icon="📦", layout="wide")

BASE_PATH = (
    r"C:\Users\2960007532\Documents\SITE STREAM LIT\Script Base Site\Site\Base site.xlsx"
)
SHEET_NAME = "BASE"


# ============================
# TEMA PRETO + CARTÃO DE ERRO
# ============================
def _set_theme_black():
    st.markdown(
        """
        <style>
            .stApp { background-color: #000000; color: #FFFFFF; }
            [data-testid="stDataFrame"] div { font-size: 20px !important; color: white; }
            thead tr th { color: white; font-size: 18px !important; text-align: center; }
            tbody th, tbody td { text-align: center !important; vertical-align: middle; }

            div.stButton > button {
                background-color: #1f1f1f; color: white; width: 100%;
                border-radius: 8px; border: 1px solid white;
            }

            div.stDownloadButton > button {
                background-color: #1f1f1f !important; color: white !important;
                border: 1px solid white !important; border-radius: 8px !important;
            }

            .erro-card {
                padding: 18px; border-radius: 12px;
                background: rgba(255,255,255,0.07);
                border: 1px solid rgba(255,255,255,0.15);
                text-align: center; color: white;
                animation: pulse 1.5s infinite;
                margin-top: 12px;
            }

            @keyframes pulse {
                0% { box-shadow: 0 0 0 rgba(255,255,255,0.08); }
                50% { box-shadow: 0 0 18px rgba(255,255,255,0.18); }
                100% { box-shadow: 0 0 0 rgba(255,255,255,0.08); }
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _show_permission_card():
    st.markdown(
        """
        <div class="erro-card">
            <h3>🚫 Base em atualização</h3>
            <p>A planilha está sendo atualizada pelo sistema automático.</p>
            <p>Aguarde e tente novamente.</p>
            <div style="margin-top: 8px; color: #ccc">
                A página recarrega em <b>10 segundos</b>.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.button("🔄 Atualizar agora")
    st.markdown('<meta http-equiv="refresh" content="10">', unsafe_allow_html=True)
    st.stop()


# ============================
# UTILS
# ============================
def to_excel_bytes(df: pd.DataFrame, sheet_name="Dados") -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buffer.getvalue()


def _find_col(df, candidates):
    """Procura coluna independente de acentuação, espaços e maiúsc/minúsc."""
    def norm(s):
        return "".join(c.lower() for c in str(s) if c.isalnum())

    normalized = {norm(col): col for col in df.columns}

    for cand in candidates:
        key = norm(cand)
        if key in normalized:
            return normalized[key]
    return None


# ============================
# CARREGAR BASE
# ============================
@st.cache_data(show_spinner=False)
def carregar_base():
    path = Path(BASE_PATH)
    if not path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado em: {BASE_PATH}")

    try:
        df = pd.read_excel(path, sheet_name=SHEET_NAME)
        df.columns = df.columns.str.strip()
        return df

    except PermissionError:
        raise

    except Exception as e:
        raise RuntimeError(f"Erro ao carregar base: {e}")


# ============================
# FORMATAÇÃO (cores %)
# ============================
def colorir_percentual_str(valor):
    if not isinstance(valor, str) or not valor.endswith("%"):
        return ""

    try:
        num = int(valor.replace("%", ""))
    except:
        return ""

    return (
        "color: green; font-weight: bold;"
        if num >= 50
        else "color: red; font-weight: bold;"
    )


# ============================
# TABELA STATUS
# ============================
def gerar_tabela_status(df, col_conf, col_status, col_qtd):

    d = df[[col_conf, col_status, col_qtd]].copy()
    d[col_conf] = d[col_conf].astype(str).str.strip()
    d[col_status] = d[col_status].astype(str).str.upper().str.strip()
    d[col_qtd] = pd.to_numeric(d[col_qtd], errors="coerce").fillna(0)

    pivot = pd.pivot_table(
        d,
        index=col_conf,
        columns=col_status,
        values=col_qtd,
        aggfunc="count",
        fill_value=0,
    ).reset_index()

    for col in ["LOADED", "PACKED", "CREATED"]:
        if col not in pivot.columns:
            pivot[col] = 0

    pivot = pivot[[col_conf, "LOADED", "PACKED", "CREATED"]]
    pivot["TOTAL"] = pivot["LOADED"] + pivot["PACKED"] + pivot["CREATED"]
    pivot["% LOADED"] = ((pivot["LOADED"] / pivot["TOTAL"]) * 100).fillna(0).astype(int).astype(str) + "%"

    pivot = pivot.rename(columns={col_conf: "Conferente"})
    return pivot.sort_values("Conferente").reset_index(drop=True)


# ============================
# TABELA AUDIT
# ============================
def gerar_tabela_audit(df, col_conf, col_audit, col_qtd):

    d = df[[col_conf, col_audit, col_qtd]].copy()
    d[col_conf] = d[col_conf].astype(str).str.strip()
    d[col_audit] = d[col_audit].astype(str).str.upper().str.strip()
    d[col_qtd] = pd.to_numeric(d[col_qtd], errors="coerce").fillna(0)

    ajustes = {
        "OK": "AUDIT COMPLETO",
        "COMPLETO": "AUDIT COMPLETO",
        "INCOMPLETO": "AUDIT INCOMPLETO",
        "0": "AUDIT INCOMPLETO",
    }
    d[col_audit] = d[col_audit].replace(ajustes)

    pivot = pd.pivot_table(
        d,
        index=col_conf,
        columns=col_audit,
        values=col_qtd,
        aggfunc="count",
        fill_value=0,
    ).reset_index()

    for col in ["AUDIT COMPLETO", "AUDIT INCOMPLETO"]:
        if col not in pivot.columns:
            pivot[col] = 0

    pivot["TOTAL"] = pivot["AUDIT COMPLETO"] + pivot["AUDIT INCOMPLETO"]
    pivot["% AUDIT COMPLETO"] = (
        (pivot["AUDIT COMPLETO"] / pivot["TOTAL"]) * 100
    ).fillna(0).astype(int).astype(str) + "%"

    pivot = pivot.rename(columns={col_conf: "Conferente"})
    return pivot.sort_values("Conferente").reset_index(drop=True)


# ============================
# TOTAL GERAL
# ============================
def linha_total_geral(df_tab):
    total = {}

    for col in df_tab.columns:
        if col == "Conferente":
            total[col] = "TOTAL GERAL"

        elif df_tab[col].dtype in ("int64", "float64"):
            total[col] = int(df_tab[col].sum())

        else:
            total[col] = ""

    # porcentagem para status
    if "% LOADED" in df_tab.columns:
        total["% LOADED"] = (
            str(int((df_tab["LOADED"].sum() / df_tab["TOTAL"].sum()) * 100)) + "%"
        )

    if "% AUDIT COMPLETO" in df_tab.columns:
        total["% AUDIT COMPLETO"] = (
            str(int((df_tab["AUDIT COMPLETO"].sum() / df_tab["TOTAL"].sum()) * 100))
            + "%"
        )

    return pd.DataFrame([total])


# ============================
# RENDER
# ============================
def render():
    _set_theme_black()
    st.title("📦 Acompanhamento por Conferente")
    st.markdown("---")

    if st.button("🔄 Atualizar base"):
        st.cache_data.clear()
        st.rerun()

    # tentar carregar base
    try:
        df = carregar_base()
    except PermissionError:
        _show_permission_card()
        return
    except Exception as e:
        st.error(f"❌ {e}")
        return

    if df is None or df.empty:
        st.error("❌ Erro: base vazia.")
        return

    df.columns = df.columns.str.strip()

    # detectando colunas
    col_conf = _find_col(df, ["Conferentes", "Conferente", "AG"])
    col_status = _find_col(df, ["Status oLPN", "Status", "Status OLPn", "OLPN STATUS"])
    col_qtd = _find_col(df, ["Qtde. Peças Item", "Qtde", "Qtd", "Qtde Pecas"])
    col_audit = _find_col(df, ["Audit", "AUDIT", "Audit Status"])
    col_demanda = _find_col(df, ["Demanda"])
    col_setor = _find_col(df, ["Setor"])

    # validar colunas obrigatórias
    required = {
        "Conferente": col_conf,
        "Status": col_status,
        "Quantidade": col_qtd,
        "Audit": col_audit,
    }
    missing = [k for k, v in required.items() if v is None]

    if missing:
        st.error(
            f"Colunas obrigatórias não encontradas: {missing}<br>"
            f"Detectadas: {list(df.columns)}",
            unsafe_allow_html=True,
        )
        return

    df = df[df[col_conf].notna()].copy()

    # =============================
    # FILTROS
    # =============================
    c1, c2, c3 = st.columns(3)

    with c1:
        conferentes = ["TODOS"] + sorted(df[col_conf].astype(str).unique())
        filtro_conf = st.selectbox("Conferente:", conferentes)

    with c2:
        demandas = (
            ["TODOS"] + sorted(df[col_demanda].astype(str).unique())
            if col_demanda
            else ["TODOS"]
        )
        filtro_dem = st.selectbox("Demanda:", demandas)

    with c3:
        setores = (
            ["TODOS"] + sorted(df[col_setor].astype(str).unique())
            if col_setor
            else ["TODOS"]
        )
        filtro_setor = st.selectbox("Setor:", setores)

    if filtro_conf != "TODOS":
        df = df[df[col_conf] == filtro_conf]

    if col_demanda and filtro_dem != "TODOS":
        df = df[df[col_demanda] == filtro_dem]

    if col_setor and filtro_setor != "TODOS":
        df = df[df[col_setor] == filtro_setor]

    st.markdown("---")

    # =============================
    # EXTRATOR
    # =============================
    st.markdown("### 📤 Extrair dados filtrados")

    e1, e2, _ = st.columns([2, 2, 4])

    with e1:
        status_filter = st.selectbox("Status:", ["Nenhum", "LOADED", "PACKED", "CREATED"])
    with e2:
        audit_filter = st.selectbox(
            "Audit:", ["Nenhum", "AUDIT COMPLETO", "AUDIT INCOMPLETO"]
        )

    df_extr = df.copy()

    if status_filter != "Nenhum":
        df_extr = df_extr[df_extr[col_status].astype(str).str.upper() == status_filter]

    if audit_filter != "Nenhum":
        df_extr = df_extr[df_extr[col_audit].astype(str).str.upper() == audit_filter]

    st.download_button(
        "📥 Baixar Base Filtrada",
        data=to_excel_bytes(df_extr),
        file_name="Base_Filtrada.xlsx",
    )

    st.markdown("---")

    # =============================
    # TABELA STATUS
    # =============================
    st.markdown("## 🟦 Status da Conferência")

    tabela_status = gerar_tabela_status(df, col_conf, col_status, col_qtd)
    total_status = linha_total_geral(tabela_status)

    st.dataframe(
        tabela_status.style.applymap(colorir_percentual_str, subset=["% LOADED"]),
        use_container_width=True,
        height=520,
    )
    st.dataframe(
        total_status.style.set_properties(**{"font-weight": "bold"}),
        use_container_width=True,
        height=80,
    )

    # =============================
    # TABELA AUDIT
    # =============================
    st.markdown("## 🟩 Status Audit")

    tabela_audit = gerar_tabela_audit(df, col_conf, col_audit, col_qtd)
    total_audit = linha_total_geral(tabela_audit)

    st.dataframe(
        tabela_audit.style.applymap(colorir_percentual_str, subset=["% AUDIT COMPLETO"]),
        use_container_width=True,
        height=520,
    )
    st.dataframe(
        total_audit.style.set_properties(**{"font-weight": "bold"}),
        use_container_width=True,
        height=80,
    )

    st.markdown(
        "<div style='text-align:center;color:rgba(255,255,255,0.6)'>Criado por Mateus Pinheiro</div>",
        unsafe_allow_html=True,
    )


# ============================
# MAIN
# ============================
if __name__ == "__main__":
    render()
