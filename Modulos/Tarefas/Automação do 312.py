# ============================================================
# Automação: 3.12 - Tarefas na Onda - Tabela
# Versão: reforçada (login robusto, timeout global, download wait)
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

# -----------------------------
# CONFIG
# -----------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOADS_DIR = os.path.join(BASE_DIR, "downloads")
os.makedirs(DOWNLOADS_DIR, exist_ok=True)

# Prefix do arquivo (conforme você informou)
CSV_PREFIX = "3.12 - Tarefas na Onda - Tabela"

# Timeout gerais
GLOBAL_TIMEOUT_SECONDS = 15 * 60   # 15 minutos (fail-safe global)
STEP_WAIT = 30                     # WebDriverWait padrão
DOWNLOAD_TIMEOUT = 3 * 60          # 3 minutos para o download finalizar
STABLE_CHECKS = 2                  # vezes que o tamanho precisa permanecer igual

# Temp profile do Chrome (o script só fecha o Chrome que ele criou)
TEMP_PROFILE = tempfile.mkdtemp(prefix="script312_profile_")

# Controle para a thread de timeout global
_global_timeout_hit = False

# -----------------------------
# FUNÇÃO: trava global (fecha tudo se passar do tempo)
# -----------------------------
def _global_timeout_watcher(seconds, driver_ref):
    global _global_timeout_hit
    time.sleep(seconds)
    _global_timeout_hit = True
    print("\n⛔ Global timeout atingido. Encerrando automação por segurança...")
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

# -----------------------------
# FUNÇÃO: clique seguro (normal -> js fallback)
# -----------------------------
def clicar_seguro(el):
    try:
        el.click()
        return True
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", el)
            return True
        except Exception:
            return False

# -----------------------------
# FUNÇÃO: limpa pasta downloads somente no inicio
# -----------------------------
def limpar_pasta_downloads():
    print("🧹 Limpando pasta de downloads (início)...")
    for filename in os.listdir(DOWNLOADS_DIR):
        full = os.path.join(DOWNLOADS_DIR, filename)
        try:
            if os.path.isfile(full) or os.path.islink(full):
                os.unlink(full)
            elif os.path.isdir(full):
                shutil.rmtree(full)
            print(f"   - removido: {filename}")
        except Exception as e:
            print(f"   ! erro removendo {filename}: {e}")
    print("📂 Pasta de downloads limpa.\n")

# -----------------------------
# FUNÇÃO: encontrar arquivo finalizado (prefixo + sem .crdownload)
# Retorna o caminho completo do arquivo quando estiver pronto.
# -----------------------------
def aguardar_arquivo_pronto(prefix, timeout_seconds):
    print(f"📥 Aguardando arquivo começando com '{prefix}' (timeout {timeout_seconds}s)...")
    start = time.time()
    last_sizes = {}

    while True:
        # check global timeout thread flag
        if _global_timeout_hit:
            raise RuntimeError("Global timeout atingido.")

        # listar arquivos
        files = os.listdir(DOWNLOADS_DIR)
        # procurar candidatos que comecem com prefix e terminem .csv (ou .xlsx/.txt)
        candidatos = []
        for f in files:
            low = f.lower()
            if low.startswith(prefix.lower()) and (low.endswith(".csv") or low.endswith(".xlsx") or low.endswith(".txt")):
                candidatos.append(f)
        if candidatos:
            # pegar o mais recente
            candidatos.sort(key=lambda n: os.path.getmtime(os.path.join(DOWNLOADS_DIR, n)), reverse=True)
            cand = candidatos[0]
            caminho = os.path.join(DOWNLOADS_DIR, cand)
            # checar se existe .crdownload correspondente (in progress)
            crdownload_present = any(x.endswith(".crdownload") and x.startswith(cand) for x in files)
            if crdownload_present:
                # ainda baixando
                pass
            else:
                # checar estabilidade de tamanho
                stable_count = 0
                prev = -1
                while stable_count < STABLE_CHECKS:
                    try:
                        cur = os.path.getsize(caminho)
                    except Exception:
                        break  # arquivo sumiu; volta ao loop externo
                    if cur == prev:
                        stable_count += 1
                    else:
                        stable_count = 0
                        prev = cur
                    if time.time() - start > timeout_seconds:
                        break
                    time.sleep(1)
                if stable_count >= STABLE_CHECKS:
                    print(f"✅ Arquivo pronto: {cand} ({os.path.getsize(caminho)} bytes)")
                    return caminho
        # timeout check
        if time.time() - start > timeout_seconds:
            break
        time.sleep(1)
    return None

# -----------------------------
# MAIN
# -----------------------------
driver = None
try:
    # iniciar driver
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

    # iniciar watcher global (fecha tudo se passar do tempo)
    threading.Thread(target=_global_timeout_watcher, args=(GLOBAL_TIMEOUT_SECONDS, driver), daemon=True).start()

    # limpar pasta downloads apenas no inicio
    limpar_pasta_downloads()

    # abrir link direto
    url = "https://viavp-sci.sce.manh.com/bi/?perspective=authoring&id=i86FCDA079A6C403994BD3B3C65C230FF&objRef=i86FCDA079A6C403994BD3B3C65C230FF&action=run&format=CSV"
    driver.get(url)
    time.sleep(4)

    # ----- selecionar tipo de login (dropdown) -----
    try:
        dd = wait.until(EC.element_to_be_clickable((By.ID, "downshift-0-toggle-button")))
        driver.execute_script("arguments[0].scrollIntoView(true);", dd)
        clicar_seguro(dd)
        time.sleep(0.5)
        item = wait.until(EC.element_to_be_clickable((By.ID, "downshift-0-item-2")))
        clicar_seguro(item)
        print("✅ Tipo de login selecionado.")
    except Exception as e:
        print(f"⚠ Não foi possível selecionar tipo de login (continuando): {e}")

    # ----- clicar VIAV Users -----
    try:
        viav = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'kc-social-provider-name') and contains(text(),'VIAV Users')]")))
        driver.execute_script("arguments[0].scrollIntoView(true);", viav)
        clicar_seguro(viav)
        print("✅ 'VIAV Users' clicado.")
    except Exception as e:
        print(f"❌ Erro ao clicar VIAV Users: {e}")
        raise

    # ----- carregar credenciais
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
                print("📨 Matrícula enviada.")
                return True
            except Exception as e:
                print(f"⚠ Tentativa {attempt+1}/5 enviar matrícula falhou: {e}")
                time.sleep(1.5)
        return False

    def enviar_senha_reforcado(senha_val):
        for attempt in range(5):
            try:
                pwd = wait.until(EC.presence_of_element_located((By.ID, "i0118")))
                pwd.clear()
                pwd.send_keys(senha_val)
                print("🔐 Senha digitada.")
                # tentar clicar Entrar (pode haver mais de um botão com mesmo id)
                try:
                    btn = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
                    clicar_seguro(btn)
                except:
                    # fallback por xpath
                    try:
                        btn2 = driver.find_element(By.XPATH, "//input[@type='submit' and (@value='Entrar' or @value='Sign in' or @value='Sign In')]")
                        clicar_seguro(btn2)
                    except:
                        pass
                print("➡ Tentativa de envio da senha executada.")
                return True
            except Exception as e:
                print(f"⚠ Tentativa {attempt+1}/5 enviar senha falhou: {e}")
                time.sleep(1.5)
        return False

    if not enviar_email_reforcado(matricula):
        raise RuntimeError("Não foi possível enviar matrícula depois de 5 tentativas.")

    if not enviar_senha_reforcado(senha):
        raise RuntimeError("Não foi possível enviar senha depois de 5 tentativas.")

    # tentar clicar SIM (confirmação) com várias estratégias
    sim_clicked = False
    for attempt in range(10):
        try:
            sim = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
            clicar_seguro(sim)
            sim_clicked = True
            print("✅ Botão 'Sim' clicado.")
            break
        except:
            # tentar outros seletores possíveis
            try:
                alt = driver.find_element(By.XPATH, "//input[@type='submit' and (translate(@value,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')='SIM' or @value='Yes')]")
                clicar_seguro(alt)
                sim_clicked = True
                print("✅ Botão 'Sim' (alternativo) clicado.")
                break
            except:
                time.sleep(1)
    if not sim_clicked:
        print("ℹ Botão 'Sim' não apareceu — seguindo fluxo.")

    # ---------------------------
    # Selecionar Unidade 1200 (entrar no iframe correto)
    # ---------------------------
    print("\n--- Selecionando unidade 1200 (Jundiaí) ---")
    try:
        # esperar e entrar no iframe que contém os filtros; normalmente é iframe index 1
        wait.until(EC.frame_to_be_available_and_switch_to_it(1))
        time.sleep(1)
        sel = wait.until(EC.presence_of_element_located((By.XPATH, "//select[contains(@id,'_ValueComboBox')]")))
        Select(sel).select_by_value("1200")
        print("✅ Unidade 1200 selecionada.")
    except Exception as e:
        print(f"❌ Erro ao selecionar unidade 1200: {e}")
        # tenta procurar em outros iframes (fallback)
        try:
            driver.switch_to.default_content()
            found = False
            frames = driver.find_elements(By.TAG_NAME, "iframe")
            for idx, fr in enumerate(frames):
                try:
                    driver.switch_to.frame(fr)
                    sel = driver.find_element(By.XPATH, "//select[contains(@id,'_ValueComboBox')]")
                    Select(sel).select_by_value("1200")
                    print(f"✅ Unidade 1200 selecionada no iframe {idx}")
                    found = True
                    break
                except:
                    driver.switch_to.default_content()
                    continue
            if not found:
                print("⚠ Não localizei o select de unidade em nenhum iframe.")
        except:
            pass
    finally:
        driver.switch_to.default_content()

    # ---------------------------
    # Selecionar opção 10 (procura em todos os iframes)
    # ---------------------------
    print("\n--- Selecionando opção 10 - Abastecimento Lojas - Pick por BOX ---")
    xpath_opcao_10 = "//span[contains(text(),'10 - Abastecimento Lojas - Pick por BOX')]"
    iframe_opcao10 = None
    frames = driver.find_elements(By.TAG_NAME, "iframe")
    for idx, fr in enumerate(frames):
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(fr)
            el = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, xpath_opcao_10)))
            clicar_seguro(el)
            iframe_opcao10 = fr
            print(f"✅ Opção 10 selecionada no iframe {idx}")
            break
        except:
            continue
    driver.switch_to.default_content()
    if iframe_opcao10 is None:
        raise RuntimeError("Não foi possível localizar a opção 10 em nenhum iframe.")

    # ---------------------------
    # Clicar no botão Concluir (no iframe onde a opção estava)
    # ---------------------------
    print("\n--- Clicando no botão 'Concluir' ---")
    try:
        driver.switch_to.frame(iframe_opcao10)
        btn = wait.until(EC.element_to_be_clickable((By.ID, "dv100")))
        driver.execute_script("arguments[0].scrollIntoView(true);", btn)
        clicar_seguro(btn)
        print("✅ Botão 'Concluir' clicado.")
    except Exception as e:
        print(f"❌ Erro ao clicar em 'Concluir': {e}")
    finally:
        driver.switch_to.default_content()

    # ---------------------------
    # Aguardar download REAL (arquivo pronto, sem .crdownload)
    # ---------------------------
    arquivo_pronto = aguardar_arquivo_pronto(CSV_PREFIX, DOWNLOAD_TIMEOUT)
    if arquivo_pronto:
        print(f"\n🎉 Download finalizado: {arquivo_pronto}")
    else:
        print("\n⛔ Timeout: não detectei arquivo finalizado dentro do tempo permitido.")

    # NÃO APAGAR NADA daqui pra frente; o usuário pediu para manter o arquivo na pasta

except Exception as main_exc:
    print(f"\n❌ Erro crítico na automação: {main_exc}")

finally:
    print("\n🔒 Finalizando — fechando apenas o Chrome que o script abriu...")
    try:
        if driver:
            driver.quit()
    except:
        pass
    try:
        shutil.rmtree(TEMP_PROFILE, ignore_errors=True)
    except:
        pass
    print("✅ Finalizado. Profile temporário removido.")
