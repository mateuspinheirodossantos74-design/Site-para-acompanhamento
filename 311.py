# ============================================================
# 🔹 Automação ViaVP – Relatório 3.11 + CSV Base Atualizado
# Versão: 311_v3 — logs detalhados, timeout 90s, detecção robusta de download,
# profile temporário do Chrome e fechamento seguro apenas do Chrome criado.
# ============================================================

import os
import sys
import time
import tempfile
import shutil
import traceback
import logging
from datetime import datetime, timedelta

import psutil
import atexit

from selenium import webdriver  # type: ignore
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import (
    TimeoutException,
    WebDriverException,
    NoSuchElementException,
    StaleElementReferenceException
)

# ============================================================
# CONFIG / CONSTANTS
# ============================================================
USERPROFILE = os.environ.get("USERPROFILE") or os.path.expanduser("~")
DOWNLOADS_DIR = os.path.join(USERPROFILE, "Downloads")

# Lock/contador files
LOCK_FILE = os.path.join(USERPROFILE, "Documents", "Automacao", "311.lock")
CONTADOR_FILE = os.path.join(USERPROFILE, "Documents", "Automacao", "contador_execucoes.txt")

# CSV prefix (mantive o seu)
CSV_PREFIX = "3.11 - Status Wave + oLPN"

# Timeout / polling
MAX_WAIT_SECONDS = 90          # 1 minuto e meio
POLL_INTERVAL = 1.0            # checa a cada 1s
STABLE_CHECKS = 2              # quantas vezes o tamanho precisa ficar igual

# pasta destino para mover o CSV (mantive o seu)
PASTA_DESTINO = r"\\10.129.10.6\cd1200\share2\abasteci0200\MATEUS\BASE"

# ============================================================
# LOGGING
# ============================================================
logger = logging.getLogger("script311")
logger.setLevel(logging.DEBUG)
_console = logging.StreamHandler(sys.stdout)
_console.setLevel(logging.DEBUG)
_fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
_console.setFormatter(_fmt)
logger.addHandler(_console)

# adicional: também logar em arquivo local
log_dir = os.path.join(USERPROFILE, "Documents", "Automacao", "logs")
os.makedirs(log_dir, exist_ok=True)
fh = logging.FileHandler(os.path.join(log_dir, "script311.log"), encoding="utf-8")
fh.setLevel(logging.DEBUG)
fh.setFormatter(_fmt)
logger.addHandler(fh)

# ============================================================
# LOCK FILE – evita dupla execução
# ============================================================
def remover_lock():
    try:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)
            logger.info("🔓 Lock file removido automaticamente.")
    except Exception:
        logger.exception("Erro removendo lock file.")

atexit.register(remover_lock)
os.makedirs(os.path.dirname(LOCK_FILE), exist_ok=True)

if os.path.exists(LOCK_FILE):
    logger.warning("⚠️ Já existe uma execução em andamento. Abortando.")
    sys.exit()

open(LOCK_FILE, "w").close()
logger.info("🔒 Lock criado.")

# ============================================================
# CONTADOR DE EXECUÇÕES (sucesso / erro)
# ============================================================
def atualizar_contador(tipo):
    """Atualiza o contador diário de execuções (sucesso ou erro)."""
    hoje = datetime.now().strftime("%Y-%m-%d")
    dados = {"data": hoje, "sucesso": 0, "erro": 0}

    if os.path.exists(CONTADOR_FILE):
        try:
            with open(CONTADOR_FILE, "r", encoding="utf-8") as f:
                linhas = f.read().strip().split(",")
                if len(linhas) == 3 and linhas[0] == hoje:
                    dados["data"] = linhas[0]
                    dados["sucesso"] = int(linhas[1])
                    dados["erro"] = int(linhas[2])
        except Exception:
            logger.debug("Não foi possível ler contador atual, reiniciando para hoje.")

    if tipo == "sucesso":
        dados["sucesso"] += 1
    elif tipo == "erro":
        dados["erro"] += 1

    os.makedirs(os.path.dirname(CONTADOR_FILE), exist_ok=True)
    with open(CONTADOR_FILE, "w", encoding="utf-8") as f:
        f.write(f"{dados['data']},{dados['sucesso']},{dados['erro']}")

    return dados

# ============================================================
# CREDENCIAIS
# ============================================================
CRED_FOLDER = os.path.join(USERPROFILE, "Documents", "Usuario")
MAT_FILE = os.path.join(CRED_FOLDER, "Matricula.txt")
SENHA_FILE = os.path.join(CRED_FOLDER, "Senha.txt")

try:
    with open(MAT_FILE, "r", encoding="utf-8") as f:
        matricula = f.read().strip()
    with open(SENHA_FILE, "r", encoding="utf-8") as f:
        senha = f.read().strip()
    if not matricula or not senha:
        raise ValueError("Matrícula ou senha estão vazias.")
    logger.info("🔐 Credenciais carregadas a partir dos arquivos.")
except Exception as e:
    logger.exception("❌ Erro ao ler credenciais: %s", e)
    try:
        atualizar_contador("erro")
    except Exception:
        pass
    remover_lock()
    sys.exit(1)

# ============================================================
# FUNÇÕES DE DOWNLOAD / DETECÇÃO ROBUSTA
# ============================================================
def listar_arquivos(dir_path):
    try:
        return [f for f in os.listdir(dir_path)]
    except FileNotFoundError:
        return []

def encontra_arquivo_com_prefixo(dir_path, prefixo):
    arquivos = listar_arquivos(dir_path)
    candidatos = [
        os.path.join(dir_path, f)
        for f in arquivos
        if f.lower().endswith(".csv")
        and f.startswith(prefixo)
        and not f.startswith("~$")
    ]
    if not candidatos:
        return None
    candidatos.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return candidatos[0]

def existe_crdownload_com_prefixo(dir_path, prefixo):
    arquivos = listar_arquivos(dir_path)
    for f in arquivos:
        if f.endswith(".crdownload") and prefixo.lower() in f.lower():
            return True
    return False

def wait_for_csv_ready(download_dir, prefixo=CSV_PREFIX, max_wait=MAX_WAIT_SECONDS):
    """
    Espera até que um arquivo CSV que comece com `prefixo` esteja pronto:
    - não exista .crdownload correspondente
    - possua tamanho estável por STABLE_CHECKS verificações
    Retorna caminho completo do CSV pronto ou lança TimeoutError.
    """
    logger.info("⏳ Aguardando CSV (%s) na pasta %s (timeout %ds)...", prefixo, download_dir, max_wait)
    start = time.time()
    last_log = 0

    while True:
        elapsed = time.time() - start
        if elapsed > max_wait:
            logger.error("⏱️ Timeout (%ds) aguardando CSV.", max_wait)
            raise TimeoutError(f"Timeout aguardando CSV ({prefixo}).")

        # se houver .crdownload com o prefixo, só esperar
        if existe_crdownload_com_prefixo(download_dir, prefixo):
            if time.time() - last_log > 10:
                logger.debug("Encontrado .crdownload com prefixo '%s' — aguardando término...", prefixo)
                last_log = time.time()
            time.sleep(POLL_INTERVAL)
            continue

        # tenta encontrar CSV finalizado
        candidate = encontra_arquivo_com_prefixo(download_dir, prefixo)
        if candidate:
            # checar estabilidade de tamanho
            stable = 0
            prev = -1
            logger.debug("Arquivo candidato encontrado: %s", os.path.basename(candidate))
            while stable < STABLE_CHECKS:
                try:
                    cur = os.path.getsize(candidate)
                except FileNotFoundError:
                    logger.warning("Arquivo candidato desapareceu durante verificação: %s", candidate)
                    break  # volta ao loop externo
                logger.debug("Tamanho atual: %d (prev: %d)", cur, prev)
                if cur == prev:
                    stable += 1
                else:
                    stable = 0
                    prev = cur
                time.sleep(POLL_INTERVAL)
                if time.time() - start > max_wait:
                    logger.error("⏱️ Timeout durante verificação de estabilidade do arquivo.")
                    raise TimeoutError("Timeout durante verificação de estabilidade do arquivo.")
            else:
                logger.info("✅ CSV pronto e estável: %s (%d bytes)", os.path.basename(candidate), os.path.getsize(candidate))
                return candidate

        # caso exista um CSV já anteriormente baixado (sem .crdownload) e sem precisar checar estabilidade
        if int(elapsed) % 10 == 0 and time.time() - last_log > 5:
            logger.debug("Aguardando aparecimento do CSV... (elapsed %ds)", int(elapsed))
            last_log = time.time()

        time.sleep(POLL_INTERVAL)

# ============================================================
# LIMPEZA DE PROCESSOS (apenas filhos do driver criado)
# ============================================================
def kill_driver_children(driver_pid):
    """
    Termina apenas os processos filhos do processo do chromedriver / chrome criado pelo script.
    """
    try:
        if not driver_pid:
            return
        parent = psutil.Process(driver_pid)
        children = parent.children(recursive=True)
        for child in children:
            try:
                name = (child.name() or "").lower()
                logger.debug("Terminando processo filho: PID=%s name=%s", child.pid, name)
                child.terminate()
            except Exception:
                logger.exception("Erro terminando processo filho.")
        # aguardar finalização
        gone, alive = psutil.wait_procs(children, timeout=5)
        for p in alive:
            try:
                logger.debug("Forçando kill: PID=%s", p.pid)
                p.kill()
            except Exception:
                pass
    except psutil.NoSuchProcess:
        pass
    except Exception:
        logger.exception("Erro ao tentar finalizar processos filhos do driver.")

# ============================================================
# FUNÇÃO PARA PROCURAR RELATÓRIO 3.11
# ============================================================
def procurar_relatorio_311(driver, wait, tentativa=1):
    """Tenta localizar o campo de pesquisa e digitar '3.11'. Se falhar, recarrega e tenta de novo."""
    try:
        logger.info("🔍 Tentando localizar campo de pesquisa (tentativa %d)...", tentativa)
        campo_busca = wait.until(
            EC.visibility_of_element_located((By.XPATH, '//input[@aria-label="Procurar conteúdo"]'))
        )
        time.sleep(1)
        try:
            campo_busca.clear()
        except Exception:
            campo_busca = driver.find_element(By.XPATH, '//input[@aria-label="Procurar conteúdo"]')
        campo_busca.send_keys("3.11")
        campo_busca.send_keys(Keys.ENTER)
        logger.info("✅ Pesquisa '3.11' enviada com sucesso")
        return True

    except (TimeoutException, NoSuchElementException, StaleElementReferenceException, WebDriverException) as e:
        logger.warning("⚠️ Erro ao digitar '3.11' (tentativa %d): %s", tentativa, e)
        if tentativa == 1:
            logger.info("🔁 Recarregando página e tentando novamente...")
            try:
                driver.refresh()
                wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            except Exception:
                pass
            time.sleep(3)
            return procurar_relatorio_311(driver, wait, tentativa + 1)
        else:
            logger.error("❌ Falha ao tentar digitar '3.11' duas vezes.")
            raise

# ============================================================
# MAIN FLOW
# ============================================================
def main():
    temp_profile = None
    driver = None
    wait = None
    driver_pid = None

    try:
        # criar profile temporário para isolar o Chrome deste script
        temp_profile = tempfile.mkdtemp(prefix="script311_chrome_profile_")
        logger.debug("Profile temporário criado: %s", temp_profile)

        # configurar Chrome options
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-infobars")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--log-level=3")
        chrome_options.add_argument(f"--user-data-dir={temp_profile}")
        # chrome_options.add_argument("--headless=new")  # descomente se quiser headless

        # download directory (compatível com qualquer usuário)
        prefs = {
            "download.default_directory": DOWNLOADS_DIR,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)

        # iniciar webdriver (assume chromedriver no PATH / driver compatível já disponível)
        service = Service()
        driver = webdriver.Chrome(service=service, options=chrome_options)
        wait = WebDriverWait(driver, 45)
        try:
            driver_pid = driver.service.process.pid if driver.service else None
        except Exception:
            driver_pid = None

        logger.info("✅ Navegador iniciado (PID service: %s)", driver_pid)

        # ============================================================
        # LOGIN E ACESSO DIRETO AO RELATÓRIO
        # ============================================================
        link_csv = "https://viavp-sci.sce.manh.com/bi/?perspective=authoring&id=i79E326D8D72B45F795E0897FCE0606F6&objRef=i79E326D8D72B45F795E0897FCE0606F6&action=run&format=CSV&cmPropStr=%7B%22id%22%3A%22i79E326D8D72B45F795E0897FCE0606F6%22%2C%22type%22%3A%22report%22%2C%22defaultName%22%3A%223.11%20-%20Status%20Wave%20%2B%20oLPN%22%2C%22permissions%22%3A%5B%22execute%22%2C%22read%22%2C%22traverse%22%5D%7D"
        driver.get(link_csv)
        logger.info("✅ Página CSV acessada diretamente")

        # Dropdown namespace (se existir)
        try:
            namespace_button = wait.until(EC.element_to_be_clickable((By.ID, "downshift-0-toggle-button")))
            namespace_button.click()
            namespace_option = wait.until(EC.element_to_be_clickable((By.ID, "downshift-0-item-0")))
            namespace_option.click()
            logger.debug("Namespace selecionado (se aplicável).")
        except Exception:
            logger.debug("Namespace dropdown não apareceu ou não é necessário.")

        # Login Microsoft/SSO
        email_input = wait.until(EC.presence_of_element_located((By.ID, "i0116")))
        email_input.send_keys(matricula)
        email_input.send_keys(Keys.ENTER)

        senha_input = wait.until(EC.presence_of_element_located((By.ID, "i0118")))
        senha_input.send_keys(senha)
        entrar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@id='idSIButton9' and @value='Entrar']")))
        entrar_btn.click()
        logger.info("✅ Botão 'Entrar' clicado")

        try:
            confirm_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@id='idSIButton9' and @value='Sim']")))
            confirm_btn.click()
            logger.info("✅ Botão 'Sim' clicado")
        except Exception:
            logger.debug("Botão 'Sim' não apareceu, continuando...")

        # Procurar relatório e entrar
        procurar_relatorio_311(driver, wait)
        folha_link = wait.until(EC.element_to_be_clickable((By.XPATH, '//a[.//div[@aria-label="3.11 - Status Wave + oLPN"]]')))
        driver.execute_script("arguments[0].scrollIntoView(true);", folha_link)
        folha_link.click()
        logger.info("✅ Link do relatório clicado")

        wait.until(EC.frame_to_be_available_and_switch_to_it(0))
        logger.info("✅ Entrou no iframe do relatório")

        # Selecionar unidade
        xpath_dropdown = "//select[contains(@id, '_ValueComboBox')]"
        select_element = wait.until(EC.presence_of_element_located((By.XPATH, xpath_dropdown)))
        Select(select_element).select_by_value("1200")
        logger.info("✅ Unidade 1200 selecionada")
        time.sleep(1)

                # ============================================================
        # INSERIR DATA COM NOVA REGRA
        # ============================================================
        hoje = datetime.today()

        # Nova lógica:
        # Segunda → última sexta (hoje - 3 dias)
        # Terça a Sexta → 2 dias atrás
        # Sábado → 2 dias atrás
        # Domingo → 3 dias atrás
        if hoje.weekday() == 0:
            # Segunda-feira → última sexta
            data_alvo = hoje - timedelta(days=3)

        elif hoje.weekday() in [1, 2, 3, 4]:
            # Terça a Sexta → 2 dias atrás
            data_alvo = hoje - timedelta(days=2)

        elif hoje.weekday() == 5:
            # Sábado → 2 dias atrás
            data_alvo = hoje - timedelta(days=2)

        else:
            # Domingo → 3 dias atrás
            data_alvo = hoje - timedelta(days=3)

        data_formatada = data_alvo.strftime("%d/%m/%Y")
        logger.info("📅 Data selecionada automaticamente: %s", data_formatada)

        try:
            input_data = wait.until(EC.presence_of_element_located((
                By.XPATH,
                "//input[contains(@id,'DateTextBox') or contains(@placeholder,'Data') or contains(@title,'Data')]"
            )))
            input_data.clear()
            input_data.send_keys(data_formatada)
            input_data.send_keys(Keys.ENTER)
            logger.info("✅ Data %s inserida", data_formatada)
        except Exception as e:
            logger.warning("❌ Erro ao inserir a data: %s", e)

        time.sleep(2)


        # ============================================================
        # SELECIONAR BOXES
        # ============================================================
        boxes = [
            "S11 - TRANSF. LOJA VIA DEPOSITO BOA",
            "S12 - TRANSF.LOJA VIA DEPOSITO QEB",
            "S13 - ABASTECIMENTO DE LOJA BOA",
            "S14 - ABASTECIMENTO DE LOJA QEB",
            "S46 - ABASTECIMENTO RETIRA LOJA",
            "S48 - ABASTECIMENTO CEL RJ",
            "S53 - TRANSFERENCIA ENTRE CDS"
        ]

        for box in boxes:
            try:
                opcao = wait.until(EC.element_to_be_clickable((
                    By.XPATH,
                    f'//span[@class="clsListItemLabel" and normalize-space(text())="{box}"]'
                )))
                opcao.click()
                logger.info("✅ Box '%s' selecionado", box)
                time.sleep(0.5)
            except Exception as e:
                logger.warning("❌ Erro ao selecionar box '%s': %s", box, e)

        try:
            concluir_btn = wait.until(EC.element_to_be_clickable((
                By.XPATH, "//button[@specname='promptButton' and contains(text(),'Concluir')]"
            )))
            concluir_btn.click()
            logger.info("✅ Botão 'Concluir' clicado – download do CSV deve iniciar")
        except Exception as e:
            logger.error("❌ Erro ao clicar em 'Concluir': %s", e)

        # ============================================================
        # AGUARDAR O CSV (90s) E MOVER PARA PASTA DESTINO
        # ============================================================
        os.makedirs(PASTA_DESTINO, exist_ok=True)

        caminho_csv = wait_for_csv_ready(DOWNLOADS_DIR, prefixo=CSV_PREFIX, max_wait=MAX_WAIT_SECONDS)
        destino_final = os.path.join(PASTA_DESTINO, "Base.csv")

        if os.path.exists(destino_final):
            try:
                os.remove(destino_final)
                logger.info("🗑️ CSV antigo deletado: %s", destino_final)
            except Exception as e:
                logger.warning("⚠️ Não foi possível deletar CSV antigo: %s", e)

        shutil.move(caminho_csv, destino_final)
        logger.info("✅ CSV movido para Base.csv em: %s", destino_final)

        dados = atualizar_contador("sucesso")
        logger.info("✅ Execução concluída com sucesso (Hoje: %d sucesso(s) | %d erro(s)).", dados['sucesso'], dados['erro'])

    except Exception as exc:
        logger.exception("❌ Ocorreu um erro durante a execução: %s", exc)
        try:
            dados = atualizar_contador("erro")
            logger.info("❌ Execução com erro (Hoje: %d sucesso(s) | %d erro(s)).", dados['sucesso'], dados['erro'])
        except Exception:
            pass
    finally:
        # fechamento seguro do driver criado por este script
        try:
            if driver:
                try:
                    logger.info("🧹 Finalizando WebDriver (quit)...")
                    driver.quit()
                except Exception:
                    logger.exception("Erro ao dar quit no driver.")
                try:
                    if driver_pid:
                        logger.debug("Tentando finalizar processos filhos do driver PID %s", driver_pid)
                        kill_driver_children(driver_pid)
                except Exception:
                    logger.exception("Erro tentando finalizar processos filhos.")
        except Exception:
            logger.exception("Erro durante rotina finally de fechamento do driver.")
        try:
            if temp_profile and os.path.isdir(temp_profile):
                shutil.rmtree(temp_profile, ignore_errors=True)
                logger.debug("Profile temporário removido: %s", temp_profile)
        except Exception:
            logger.exception("Erro removendo profile temporário.")
        remover_lock()


if __name__ == "__main__":
    main()
