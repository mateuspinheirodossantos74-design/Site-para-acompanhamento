import streamlit as st
import pandas as pd
from pathlib import Path
from io import BytesIO
import tempfile
import webbrowser
from datetime import datetime
import base64

# ===========================
# CONFIG
# ===========================
st.set_page_config(page_title="Relatórios", layout="wide")
st.title("📊 Relatórios - Base Site")

# ===========================
# CAMINHO DO ARQUIVO
# ===========================
ARQUIVO = Path(
    r"C:\Users\2960007532\Documents\SITE STREAM LIT\Script Base Site\Site\Base site.xlsx"
)

# ===========================
# COLUNAS
# ===========================
COLUNAS = [
    "Tipo de pedido","Filial Destino","oLPN","Item","Descrição",
    "Local de Picking","Qtde. Peças Item","Status oLPN","BOX",
    "Audit","Demanda","Conferentes"
]

# ===========================
# BOTÃO ATUALIZAR
# ===========================
if st.button("🔄 Atualizar Base"):
    try:
        st.cache_data.clear()
    except:
        pass
    st.rerun()

# ===========================
# CHECAR ARQUIVO
# ===========================
if not ARQUIVO.exists():
    st.error(f"❌ Arquivo não encontrado:\n{ARQUIVO}")
    st.stop()

# ===========================
# CARREGAR BASE
# ===========================
try:
    existing_cols = pd.read_excel(ARQUIVO, nrows=0).columns.tolist()
    usecols = [c for c in COLUNAS if c in existing_cols]
    df = pd.read_excel(ARQUIVO, usecols=usecols)
except Exception as e:
    st.error(f"Erro ao ler base: {e}")
    st.stop()

for c in COLUNAS:
    if c not in df.columns:
        df[c] = ""

for c in df.columns:
    if df[c].dtype == object:
        df[c] = df[c].astype(str)

# ===========================
# FILTROS
# ===========================
tabs = st.tabs(["Filtros adicionais"])
with tabs[0]:
    with st.expander("🔎 Filtros adicionais", expanded=False):
        f1, f2, f3 = st.columns(3)
        f4, f5 = st.columns(2)

        demandas = sorted(df["Demanda"].replace("nan", "").dropna().unique().tolist())
        f_demanda = f1.selectbox("Demanda:", ["Todos"] + demandas)

        boxes = sorted(
            df["BOX"].replace("nan", "").dropna().unique().tolist(),
            key=lambda x: int("".join(filter(str.isdigit, str(x))) or 0)
        )
        f_box = f2.selectbox("BOX:", ["Todos"] + boxes)

        filiais = sorted(df["Filial Destino"].replace("nan", "").dropna().unique().tolist())
        f_filial = f3.selectbox("Filial Destino:", ["Todos"] + filiais)

        olpns = sorted(df["oLPN"].replace("nan", "").dropna().unique().tolist())
        f_olpn = f4.selectbox("oLPN:", ["Todos"] + olpns)

        confs = sorted(df["Conferentes"].replace("nan", "").dropna().unique().tolist())
        f_conf = f5.selectbox("Conferentes:", ["Todos"] + confs)

# ===========================
# APLICAR FILTROS
# ===========================
df_f = df.copy()

if f_demanda and f_demanda != "Todos":
    df_f = df_f[df_f["Demanda"] == f_demanda]

if f_box and f_box != "Todos":
    df_f = df_f[df_f["BOX"] == f_box]

if f_filial and f_filial != "Todos":
    df_f = df_f[df_f["Filial Destino"] == f_filial]

if f_olpn and f_olpn != "Todos":
    df_f = df_f[df_f["oLPN"] == f_olpn]

if f_conf and f_conf != "Todos":
    df_f = df_f[df_f["Conferentes"] == f_conf]

# ===========================
# FUNÇÃO PARA ORDENAR BOX
# ===========================
def box_to_num(x):
    try:
        s = str(x)
        nums = ''.join(filter(str.isdigit, s))
        return int(nums) if nums else 0
    except:
        return 0

# ===========================
# TABELA SEM AUDIT
# ===========================
st.subheader("📄 Relatorio pakced")

statuses_sem = sorted(df_f["Status oLPN"].replace("nan", "").dropna().unique().tolist())
f_status_sem = st.multiselect("Status oLPN (somente tabela SEM Audit):", statuses_sem, key="status_sem")

df_sem = df_f.copy()
if f_status_sem:
    df_sem = df_sem[df_sem["Status oLPN"].isin(f_status_sem)]

df_sem = df_sem.assign(__BOX_ORD=df_sem["BOX"].map(box_to_num)).sort_values("__BOX_ORD").drop(columns="__BOX_ORD")

cols_display = [
    "Tipo de pedido","Filial Destino","oLPN","Item","Descrição",
    "Local de Picking","Qtde. Peças Item","Status oLPN","BOX","Conferentes"
]
view_sem = df_sem[cols_display].copy()
st.dataframe(view_sem, use_container_width=True, height=420)

# ===========================
# TABELA COM AUDIT
# ===========================
st.subheader("📋 Relatorio Audit")

statuses_com = sorted(df_f["Status oLPN"].replace("nan", "").dropna().unique().tolist())
audits_com = sorted(df_f["Audit"].replace("nan", "").dropna().unique().tolist())

c1, c2 = st.columns(2)
f_status_com = c1.multiselect("Status oLPN (tabela COM Audit):", statuses_com, key="status_com")
f_audit_com = c2.multiselect("Audit (tabela COM Audit):", audits_com, key="audit_com")

df_com = df_f.copy()
if f_status_com:
    df_com = df_com[df_com["Status oLPN"].isin(f_status_com)]
if f_audit_com:
    df_com = df_com[df_com["Audit"].isin(f_audit_com)]

df_com = df_com.assign(__BOX_ORD=df_com["BOX"].map(box_to_num)).sort_values("__BOX_ORD").drop(columns="__BOX_ORD")

cols_display_com = cols_display + ["Audit"]
view_com = df_com[cols_display_com].copy()
st.dataframe(view_com, use_container_width=True, height=420)

# ===========================
# EXPORTAR EXCEL
# ===========================
def export_excel_duas_sheets(df1, df2):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df1.to_excel(writer, index=False, sheet_name="SEM_AUDIT")
        df2.to_excel(writer, index=False, sheet_name="COM_AUDIT")
    return buffer.getvalue()

st.download_button(
    "📄 Exportar Excel (2 Sheets)",
    data=export_excel_duas_sheets(view_sem, view_com),
    file_name=f"Relatorio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# =====================================================
# FUNÇÃO AJUSTADA — TÍTULO REMOVIDO
# =====================================================
def gerar_html_tabelas(df_sem_audit, df_com_audit, conferente_nome=""):

    sem = df_sem_audit.copy()
    com = df_com_audit.copy()

    if "oLPN" not in sem.columns:
        sem["oLPN"] = ""
    sem.insert(sem.columns.get_loc("oLPN") + 1, "RMP", " ")

    for col in ["Tipo de pedido","Filial Destino","oLPN","Item","Descrição",
                "Local de Picking","Qtde. Peças Item","Status oLPN","BOX","Conferentes"]:
        if col not in sem.columns:
            sem[col] = ""
        if col not in com.columns:
            com[col] = ""
    if "Audit" not in com.columns:
        com["Audit"] = ""

    sem = sem[
        ["Tipo de pedido","Filial Destino","oLPN","RMP","Item","Descrição",
         "Local de Picking","Qtde. Peças Item","Status oLPN","BOX","Conferentes"]
    ]
    com = com[
        ["Tipo de pedido","Filial Destino","oLPN","Item","Descrição",
         "Local de Picking","Qtde. Peças Item","Status oLPN","BOX","Conferentes","Audit"]
    ]

    # ======================
    # CSS
    # ======================
    css = """
    <style>
    @page { margin: 0; size: landscape; }
    html, body {
        margin: 6px;
        padding: 0;
        font-family: Arial, Helvetica, sans-serif;
        font-size: 14px;
        color: #000;
    }
    h2, h3 { 
        color: #000;
        margin: 6px 0;
        font-size: 20px;
    }
    table { 
        border-collapse: collapse; 
        width: 100%; 
        table-layout: fixed; 
        font-size: 13px;
    }
    thead th {
        background-color: #e6e6e6;
        color: #000;
        font-weight: bold;
        padding: 6px;
        border: 1px solid #000;
        font-size: 14px;
    }
    td {
        border: 1px solid #bbb;
        padding: 4px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        font-size: 12px;
    }
    th.col-desc, td.col-desc { width: 420px; }
    th.col-conf, td.col-conf { width: 160px; }
    th.col-audit, td.col-audit { width: 160px; }
    </style>
    """

    def table_html(df_table, headers):
        thead = "<thead><tr>"
        for h in headers:
            thead += f"<th class='{h[1]}'>{h[0]}</th>"
        thead += "</tr></thead>"

        tbody = "<tbody>"
        for _, row in df_table.iterrows():
            tbody += "<tr>"
            for col in df_table.columns:

                cls = ""
                if col == "Descrição":
                    cls = "col-desc"
                elif col == "Conferentes":
                    cls = "col-conf"
                elif col == "Audit":
                    cls = "col-audit"

                raw = "" if pd.isna(row[col]) else str(row[col])

                try:
                    val = raw.encode("latin1").decode("utf-8")
                except:
                    val = raw

                tbody += f"<td class='{cls}'>{val}</td>"
            tbody += "</tr>"
        tbody += "</tbody>"

        return f"<table>{thead}{tbody}</table>"

    headers_packed = [
        ("Tipo",""),("Filial",""),("oLPN",""),("RMP",""),("Item",""),
        ("Descrição","col-desc"),("Picking",""),("Qtde",""),
        ("Status",""),("BOX",""),("Conferente","col-conf")
    ]

    headers_audit = [
        ("Tipo",""),("Filial",""),("oLPN",""),("Item",""),
        ("Descrição","col-desc"),("Picking",""),("Qtde",""),
        ("Status",""),("BOX",""),("Conferente","col-conf"),("Audit","col-audit")
    ]

    meta = f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
    if conferente_nome:
        meta += f" | Conferente: {conferente_nome}"

    # ======================
    # AQUI — título removido!
    # ======================
    html = f"""
    <html>
    <head>
    <meta charset="utf-8" />
    {css}
    </head>
    <body>

        <div style="margin-bottom:8px">{meta}</div>

        <h3>Packed</h3>
        {table_html(sem, headers_packed)}

        <br><br>

        <h3>Audit</h3>
        {table_html(com, headers_audit)}

        <script>
            window.onload = function() {{
                window.print();
            }};
        </script>

    </body>
    </html>
    """
    return html

# ===========================
# BOTÃO IMPRESSÃO — CORRIGIDO (UTF-8 SEM BUG)
# ===========================
import streamlit.components.v1 as components

st.markdown("---")
st.markdown("### 🖨️ Impressão")

if st.button("🖨️ Imprimir Relatório (tabelas filtradas)"):

    conferente_nome = f_conf if (f_conf and f_conf != "Todos") else ""

    # Gera o HTML
    html = gerar_html_tabelas(view_sem, view_com, conferente_nome=conferente_nome)

    # Remove título antigo
    html = html.replace("<h2>Relatório de Conferência</h2>", "")

    # Adiciona título sem acento
    titulo = "<h2 style='text-align:center;'>Relatorio de conferencia</h2>"
    html = titulo + html

    # Codifica em base64
    html_base64 = base64.b64encode(html.encode("utf-8")).decode("utf-8")

    # Javascript correto com Blob UTF-8 (SEM BUG DOS ACENTOS)
    js = f"""
        <script>
            const byteCharacters = atob("{html_base64}");
            const byteNumbers = new Array(byteCharacters.length);
            for (let i = 0; i < byteCharacters.length; i++) {{
                byteNumbers[i] = byteCharacters.charCodeAt(i);
            }}
            const byteArray = new Uint8Array(byteNumbers);
            const blob = new Blob([byteArray], {{type: 'text/html;charset=utf-8'}});
            const blobUrl = URL.createObjectURL(blob);
            window.open(blobUrl, '_blank');
        </script>
    """

    components.html(js, height=0)

    st.success("Relatório aberto em nova aba. Pode imprimir diretamente pelo navegador!")

