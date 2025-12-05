# ============================================================
# Automação: 5.03 - Produtividade de Packing - Packed por hora
# Versão: reforçada (login robusto, timeout global, download wait)
# Logs adicionados na pasta de logs
# ============================================================

import os
import time
import tempfile
import shutil
import threading
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

# -----------------------------
# CONFIGURAÇÕES
# -----------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOADS_DIR = os.path.join(BASE_DIR, "downloads")
LOGS_DIR = os.path.join(BASE_DIR, "logs")
os.makedirs(DOWNLOADS_DIR, exist_ok=True)
os.makedirs(LOGS_DIR, exist_ok=True)

CSV_PREFIX = "5.03 - Produtividade de Packing - Packed por hora"

GLOBAL_TIMEOUT_SECONDS = 15 * 60
STEP_WAIT = 30
DOWNLOAD_TIMEOUT = 3 * 60
STABLE_CHECKS = 2

TEMP_PROFILE = tempfile.mkdtemp(prefix="script503_profile_")
_global_timeout_hit = False

log_path = os.path.join(LOGS_DIR, f"log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt")

def log(msg):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    full_msg = f"[{timestamp}] {msg}"
    print(full_msg)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(full_msg + "\n")

# -----------------------------
# FUNÇÕES UTEIS
# -----------------------------
def _global_timeout_watcher(seconds, driver_ref):
    global _global_timeout_hit
    time.sleep(seconds)
    _global_timeout_hit = True
    log("⛔ Global timeout atingido. Encerrando automação por segurança...")
    try:
        if driver_ref:
            driver_ref.quit()
    except:
        pass
    try:
        shutil.rmtree(TEMP_PROFILE, ignore_errors=True)
    except:
        pass
    os._exit(1)

def clicar_seguro(el):
    try:
        el.click()
        return True
    except:
        try:
            driver.execute_script("arguments[0].click();", el)
            return True
        except:
            return False

def limpar_pasta_downloads():
    log("🧹 Limpando pasta de downloads (início)...")
    for filename in os.listdir(DOWNLOADS_DIR):
        full = os.path.join(DOWNLOADS_DIR, filename)
        try:
            if os.path.isfile(full) or os.path.islink(full):
                os.unlink(full)
            elif os.path.isdir(full):
                shutil.rmtree(full)
            log(f"   - removido: {filename}")
        except Exception as e:
            log(f"   ! erro removendo {filename}: {e}")
    log("📂 Pasta de downloads limpa.\n")

def aguardar_arquivo_pronto(prefix, timeout_seconds):
    log(f"📥 Aguardando arquivo começando com '{prefix}' (timeout {timeout_seconds}s)...")
    start = time.time()
    while True:
        if _global_timeout_hit:
            raise RuntimeError("Global timeout atingido.")

        files = os.listdir(DOWNLOADS_DIR)
        candidatos = [f for f in files if f.lower().startswith(prefix.lower()) and f.lower().endswith((".csv", ".xlsx", ".txt"))]

        if candidatos:
            candidatos.sort(key=lambda n: os.path.getmtime(os.path.join(DOWNLOADS_DIR, n)), reverse=True)
            cand = candidatos[0]
            caminho = os.path.join(DOWNLOADS_DIR, cand)
            stable_count = 0
            prev = -1
            while stable_count < STABLE_CHECKS:
                try:
                    cur = os.path.getsize(caminho)
                except:
                    break
                if cur == prev:
                    stable_count += 1
                else:
                    stable_count = 0
                    prev = cur
                if time.time() - start > timeout_seconds:
                    break
                time.sleep(1)
            if stable_count >= STABLE_CHECKS and not caminho.endswith(".crdownload"):
                log(f"✅ Arquivo pronto: {cand} ({os.path.getsize(caminho)} bytes)")
                return caminho

        if time.time() - start > timeout_seconds:
            break
        time.sleep(1)
    return None

# -----------------------------
# INICIALIZAÇÃO PRINCIPAL
# -----------------------------
driver = None
try:
    chrome_options = Options()
    chrome_options.add_argument(f"--user-data-dir={TEMP_PROFILE}")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOADS_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    })
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, STEP_WAIT)

    threading.Thread(target=_global_timeout_watcher, args=(GLOBAL_TIMEOUT_SECONDS, driver), daemon=True).start()
    limpar_pasta_downloads()

    url = "https://viavp-sci.sce.manh.com/bi/?perspective=authoring&factoryMode=true&id=i3E96C0044C974F60ABC41F536196806B&objRef=i3E96C0044C974F60ABC41F536196806B&action=run&format=CSV&cmPropStr=%7B%22id%22%3A%22i3E96C0044C974F60ABC41F536196806B%22%2C%22type%22%3A%22report%22%2C%22defaultName%22%3A%225.03%20-%20Produtividade%20de%20Packing%20-%20Packed%20por%20hora%22%2C%22permissions%22%3A%5B%22execute%22%2C%22read%22%2C%22traverse%22%5D%7D"
    driver.get(url)
    time.sleep(4)

    log("✅ Página carregada, iniciando login e seleção de parâmetros...")

    # ----- selecionar tipo de login (dropdown) -----
    try:
        dd = wait.until(EC.element_to_be_clickable((By.ID, "downshift-0-toggle-button")))
        driver.execute_script("arguments[0].scrollIntoView(true);", dd)
        clicar_seguro(dd)
        time.sleep(0.5)
        item = wait.until(EC.element_to_be_clickable((By.ID, "downshift-0-item-2")))
        clicar_seguro(item)
        log("✅ Tipo de login selecionado.")
    except Exception as e:
        log(f"⚠ Falha ao selecionar tipo de login (seguindo mesmo assim): {e}")

    # ----- clicar VIAV Users -----
    try:
        viav = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'kc-social-provider-name') and contains(text(),'VIAV Users')]")))
        driver.execute_script("arguments[0].scrollIntoView(true);", viav)
        clicar_seguro(viav)
        log("✅ 'VIAV Users' clicado.")
    except Exception as e:
        log(f"❌ Erro ao clicar VIAV Users: {e}")
        raise

    # ----- carregar credenciais -----
    CRED_FOLDER = os.path.join(os.environ.get("USERPROFILE"), "Documents", "Usuario")
    MAT_FILE = os.path.join(CRED_FOLDER, "Matricula.txt")
    SENHA_FILE = os.path.join(CRED_FOLDER, "Senha.txt")
    with open(MAT_FILE, "r", encoding="utf-8") as f:
        matricula = f.read().strip()
    with open(SENHA_FILE, "r", encoding="utf-8") as f:
        senha = f.read().strip()

    # ---------------------------
    # LOGIN MICROSOFT REFORÇADO
    # ---------------------------
    def enviar_email_reforcado(matricula_val):
        for attempt in range(5):
            try:
                email_in = wait.until(EC.presence_of_element_located((By.ID, "i0116")))
                email_in.clear()
                email_in.send_keys(matricula_val)
                email_in.send_keys(Keys.ENTER)
                log("📨 Matrícula enviada.")
                return True
            except Exception as e:
                log(f"⚠ Tentativa {attempt+1}/5 enviar matrícula falhou: {e}")
                time.sleep(1.5)
        return False

    def enviar_senha_reforcado(senha_val):
        for attempt in range(5):
            try:
                pwd = wait.until(EC.presence_of_element_located((By.ID, "i0118")))
                pwd.clear()
                pwd.send_keys(senha_val)
                log("🔐 Senha digitada.")
                try:
                    btn = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
                    clicar_seguro(btn)
                except:
                    try:
                        btn2 = driver.find_element(By.XPATH, "//input[@type='submit' and (@value='Entrar' or @value='Sign in' or @value='Sign In')]")
                        clicar_seguro(btn2)
                    except:
                        pass
                log("➡ Tentativa de envio da senha executada.")
                return True
            except Exception as e:
                log(f"⚠ Tentativa {attempt+1}/5 enviar senha falhou: {e}")
                time.sleep(1.5)
        return False

    if not enviar_email_reforcado(matricula):
        raise RuntimeError("Não foi possível enviar matrícula depois de 5 tentativas.")
    if not enviar_senha_reforcado(senha):
        raise RuntimeError("Não foi possível enviar senha depois de 5 tentativas.")

    # tentar clicar SIM
    sim_clicked = False
    for attempt in range(10):
        try:
            sim = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
            clicar_seguro(sim)
            sim_clicked = True
            log("✅ Botão 'Sim' clicado.")
            break
        except:
            try:
                alt = driver.find_element(By.XPATH, "//input[@type='submit' and (translate(@value,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')='SIM' or @value='Yes')]")
                clicar_seguro(alt)
                sim_clicked = True
                log("✅ Botão 'Sim' (alternativo) clicado.")
                break
            except:
                time.sleep(1)
    if not sim_clicked:
        log("ℹ Botão 'Sim' não apareceu — seguindo fluxo.")

    # ---------------------------
    # Selecionar Unidade 1200 (iframe)
    # ---------------------------
    log("--- Selecionando unidade 1200 (Jundiaí) ---")
    try:
        wait.until(EC.frame_to_be_available_and_switch_to_it(1))
        time.sleep(1)
        sel = wait.until(EC.presence_of_element_located((By.XPATH, "//select[contains(@id,'_ValueComboBox')]")))
        Select(sel).select_by_value("1200")
        log("✅ Unidade 1200 selecionada.")
    except Exception as e:
        log(f"⚠ Erro ao selecionar unidade 1200: {e} — tentando outros iframes")
        try:
            driver.switch_to.default_content()
            frames = driver.find_elements(By.TAG_NAME, "iframe")
            found = False
            for idx, fr in enumerate(frames):
                try:
                    driver.switch_to.frame(fr)
                    sel = driver.find_element(By.XPATH, "//select[contains(@id,'_ValueComboBox')]")
                    Select(sel).select_by_value("1200")
                    log(f"✅ Unidade 1200 selecionada no iframe {idx}")
                    found = True
                    break
                except:
                    driver.switch_to.default_content()
                    continue
            if not found:
                raise RuntimeError("Não localizei o select de unidade em nenhum iframe.")
        except Exception as ex:
            log(f"❌ Fallback falhou: {ex}")
            raise
    finally:
        driver.switch_to.default_content()

    # ---------------------------
    # Selecionar BOXES
    # ---------------------------
    log("--- Selecionando BOXES ---")
    boxes = [
        "S11 - TRANSF. LOJA VIA DEPOSITO BOA",
        "S12 - TRANSF.LOJA VIA DEPOSITO QEB",
        "S13 - ABASTECIMENTO DE LOJA BOA",
        "S14 - ABASTECIMENTO DE LOJA QEB",
        "S46 - ABASTECIMENTO RETIRA LOJA",
        "S48 - ABASTECIMENTO CEL RJ",
        "S53 - TRANSFERENCIA ENTRE CDS"
    ]

    try:
        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it(1))
    except:
        driver.switch_to.default_content()
        frames = driver.find_elements(By.TAG_NAME, "iframe")
        for idx, fr in enumerate(frames):
            try:
                driver.switch_to.frame(fr)
                driver.find_element(By.XPATH, "//span[@class='clsListItemLabel']")
                log(f"📌 Itens localizados no iframe {idx}")
                break
            except:
                driver.switch_to.default_content()
                continue

    for box in boxes:
        try:
            opcao = wait.until(EC.element_to_be_clickable((
                By.XPATH,
                f'//span[@class="clsListItemLabel" and normalize-space(text())="{box}"]'
            )))
            driver.execute_script("arguments[0].scrollIntoView(true);", opcao)
            clicar_seguro(opcao)
            log(f"✅ Box selecionado: {box}")
            time.sleep(0.3)
        except Exception as e:
            log(f"❌ Erro ao selecionar box '{box}': {e}")

    try:
        concluir_btn = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//button[@specname='promptButton' and contains(text(),'Concluir')]"
        )))
        clicar_seguro(concluir_btn)
        log("🎯 Botão 'Concluir' clicado – download deve iniciar...")
    except Exception as e:
        log(f"❌ Erro ao clicar em Concluir: {e}")

    driver.switch_to.default_content()

    # ---------------------------
    # Aguardar download
    # ---------------------------
    arquivo_pronto = aguardar_arquivo_pronto(CSV_PREFIX, DOWNLOAD_TIMEOUT)
    if arquivo_pronto:
        log(f"🎉 Download finalizado: {arquivo_pronto}")
    else:
        log("⛔ Timeout: não detectei arquivo finalizado dentro do tempo permitido.")

except Exception as main_exc:
    log(f"❌ Erro crítico na automação: {main_exc}")

finally:
    log("🔒 Finalizando — fechando apenas o Chrome que o script abriu...")
    try:
        if driver:
            driver.quit()
    except:
        pass
    try:
        shutil.rmtree(TEMP_PROFILE, ignore_errors=True)
    except:
        pass
    log("✅ Finalizado. Profile temporário removido.")
