import os
import time
import datetime
import traceback
import win32com.client as win32

# ===============================================
# 🔧 CONFIGURAÇÕES
# ===============================================
ARQUIVO_EXCEL = r"C:\Users\2960007532\Documents\SITE STREAM LIT\Script Base Site\Site\Base site.xlsx"
ABA_NOME = "Base"
COLUNA_REF_DADOS = 11   # Coluna K (oLPN)
COLUNA_INICIO_FORMULAS = 32  # Coluna AF
LOG_PATH = r"C:\Users\2960007532\Documents\Automacao\BaseSite_Log.txt"

# Garante que a pasta de log existe
os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)

# ===============================================
# 🧾 FUNÇÃO DE LOG
# ===============================================
def registrar_log(msg):
    hora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    linha = f"[{hora}] {msg}"
    print(linha)
    with open(LOG_PATH, "a", encoding="utf-8") as log:
        log.write(linha + "\n")

# ===============================================
# 🚀 EXECUÇÃO PRINCIPAL
# ===============================================
excel = None
try:
    registrar_log("🕒 Aguardando 10 segundos após 311.py...")
    time.sleep(10)

    registrar_log("🚀 Iniciando atualização da Base Site...")

    # Cria instância isolada (não fecha outros Excels abertos)
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(ARQUIVO_EXCEL)
    ws = wb.Worksheets(ABA_NOME)

    # Atualiza consultas do Power Query
    registrar_log("🔄 Atualizando consultas do Excel...")
    wb.RefreshAll()
    time.sleep(20)

    # Aguarda cálculos pendentes
    while excel.CalculationState != 0:
        time.sleep(1)

    # Identifica a última linha com dados na coluna de referência (K)
    ultima_linha = ws.Cells(ws.Rows.Count, COLUNA_REF_DADOS).End(-4162).Row
    registrar_log(f"📊 Última linha com dados: {ultima_linha}")

    # Determina a última coluna com fórmula (usando a linha 2 como modelo)
    col_final = ws.Cells(2, ws.Columns.Count).End(-4159).Column
    primeira_celula_modelo = ws.Cells(2, COLUNA_INICIO_FORMULAS)
    ultima_celula_modelo = ws.Cells(2, col_final)
    range_modelo = ws.Range(primeira_celula_modelo, ultima_celula_modelo)

    # Define o destino onde as fórmulas serão aplicadas
    ultima_celula_destino = ws.Cells(ultima_linha, col_final)
    range_destino = ws.Range(primeira_celula_modelo, ultima_celula_destino)

    # Expande as fórmulas com AutoFill
    registrar_log(f"🧩 Expandindo fórmulas de {range_modelo.Address} até linha {ultima_linha}...")
    range_modelo.AutoFill(Destination=range_destino)
    registrar_log("✅ Fórmulas expandidas com sucesso.")

    # Salva alterações
    wb.Save()
    registrar_log(f"💾 Arquivo salvo com sucesso em: {ARQUIVO_EXCEL}")

    # Atualiza o timestamp (pra garantir leitura nova no site)
    try:
        os.utime(ARQUIVO_EXCEL, None)
        registrar_log("⏰ Data de modificação atualizada corretamente.")
    except Exception as e:
        registrar_log(f"⚠️ Falha ao atualizar timestamp: {e}")

    # Fecha apenas o arquivo aberto pelo script
    wb.Close(SaveChanges=True)
    excel.Quit()
    excel = None

    registrar_log("✅ Base Site atualizada e fechada com sucesso.")
    registrar_log("-" * 80)

except Exception as e:
    registrar_log("❌ ERRO durante a atualização da Base Site:")
    registrar_log(str(e))
    registrar_log(traceback.format_exc())
    registrar_log("-" * 80)
    if excel:
        try:
            excel.Quit()
        except:
            pass

