#!/usr/bin/env python3
"""
Agente Baixador de XML de NF-e   (roda no SEU computador)

Deixe esta janela aberta. Quando você clicar em "Executar" na página do servidor, o Agente:
  1. recebe a lista de chaves da sua tarefa
  2. abre o Google Chrome AQUI, no seu computador (se aparecer a verificação anti-robô,
     resolva você mesmo na janela do Chrome)
  3. baixa cada XML e o envia para a pasta do servidor
  4. informa o andamento, que aparece na página
"""
import glob
import json
import os
import random
import shutil
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path

BASE = Path(__file__).resolve().parent

try:
    from selenium import webdriver
    from selenium.common.exceptions import (ElementClickInterceptedException,
                                            ElementNotInteractableException)
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait
except ImportError:
    print("\nFalta o pacote 'selenium'. Instale com:  pip install selenium")
    print("(Se você recebeu o Agente em .zip com a pasta 'python', use o iniciar-agente.bat.)\n")
    input("Pressione Enter para sair...")
    sys.exit(1)

XP_BUSCAR = "//button[normalize-space(.)='Imprimir DANFE']"
XP_BAIXAR_XML = "//button[normalize-space(.)='Baixar XML']"


# --------------------------------------------------------------------------- #
# Comunicação com o servidor
# --------------------------------------------------------------------------- #
def ler_servidor():
    """Endereço do servidor: variável NFE_SERVIDOR, arquivo servidor.txt ou perguntado (e salvo)."""
    s = os.environ.get("NFE_SERVIDOR", "").strip()
    arq = BASE / "servidor.txt"
    if not s and arq.exists():
        s = arq.read_text(encoding="utf-8").strip()
    if not s:
        print("Informe o endereço do servidor (ex.: 192.168.1.25:8765)")
        s = input("Servidor: ").strip()
        if s:
            arq.write_text(s + "\n", encoding="utf-8")
    s = s.replace("http://", "").rstrip("/")
    if ":" not in s:
        s += ":8765"
    return s


SERVIDOR = ""


class ErroServidor(Exception):
    pass


def http(metodo, caminho, dados=None, bruto=None, timeout=30):
    url = f"http://{SERVIDOR}{caminho}"
    headers = {}
    corpo = None
    if dados is not None:
        corpo = json.dumps(dados).encode("utf-8")
        headers["Content-Type"] = "application/json"
    elif bruto is not None:
        corpo = bruto
        headers["Content-Type"] = "application/xml"
    req = urllib.request.Request(url, data=corpo, headers=headers, method=metodo)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8") or "{}")
    except urllib.error.HTTPError as e:
        try:
            msg = json.loads(e.read().decode("utf-8")).get("erro", "")
        except Exception:  # noqa: BLE001
            msg = ""
        raise ErroServidor(f"HTTP {e.code} {msg}".strip())
    except (urllib.error.URLError, OSError, TimeoutError) as e:
        raise ErroServidor(f"sem conexão com o servidor: {e}")


# --------------------------------------------------------------------------- #
# Relato do andamento (um segundo fio envia "sinais de vida" a cada poucos segundos)
# --------------------------------------------------------------------------- #
class Relato:
    def __init__(self, job_id):
        self.job_id = job_id
        self.lock = threading.Lock()
        self.logs = []
        self.proc = 0
        self.baixadas = 0
        self.falhas = []
        self.atual = ""
        self.parar = threading.Event()
        self.falhas_conexao = 0
        self.encerrado = False

    def log(self, msg, nivel="info"):
        print(f"  {time.strftime('%H:%M:%S')} {msg}")
        with self.lock:
            self.logs.append([nivel, msg])

    def enviar(self, fim=None):
        with self.lock:
            enviados = list(self.logs)
            payload = {"job": self.job_id, "logs": enviados, "proc": self.proc,
                       "baixadas": self.baixadas, "falhas": self.falhas,
                       "atual": self.atual, "fim": fim}
            try:
                r = http("POST", "/api/agente/relatorio", dados=payload, timeout=20)
            except ErroServidor:
                self.falhas_conexao += 1
                if self.falhas_conexao >= 40:  # ~2 min sem falar com o servidor
                    self.parar.set()
                return
            self.falhas_conexao = 0
            del self.logs[:len(enviados)]
            if r.get("parar"):
                self.parar.set()

    def batimento(self):
        while not self.encerrado:
            time.sleep(3)
            if not self.encerrado:
                self.enviar()


# --------------------------------------------------------------------------- #
# Chrome
# --------------------------------------------------------------------------- #
def criar_driver(pasta_download):
    opts = webdriver.ChromeOptions()
    opts.add_experimental_option("prefs", {
        "download.default_directory": pasta_download,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    })
    # Variáveis só para testes automatizados. No uso normal o Chrome fica VISÍVEL
    # (o site costuma bloquear o modo invisível).
    if os.environ.get("NFE_CHROME_BINARY"):
        opts.binary_location = os.environ["NFE_CHROME_BINARY"]
    if os.environ.get("NFE_HEADLESS") == "1":
        opts.add_argument("--headless=new")
        opts.add_argument("--no-sandbox")
    if os.environ.get("NFE_CHROMEDRIVER"):
        from selenium.webdriver.chrome.service import Service
        return webdriver.Chrome(service=Service(os.environ["NFE_CHROMEDRIVER"]), options=opts)
    try:
        return webdriver.Chrome(options=opts)
    except Exception as e1:  # noqa: BLE001
        try:  # Selenium antigo: tenta o webdriver_manager, se existir
            from selenium.webdriver.chrome.service import Service
            from webdriver_manager.chrome import ChromeDriverManager
            return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
        except ImportError:
            raise e1


def dormir(rel, seg):
    fim = time.time() + seg
    while time.time() < fim and not rel.parar.is_set():
        time.sleep(0.3)


def enviar_xml(job_id, chave, caminho):
    with open(caminho, "rb") as f:
        dados = f.read()
    ultimo = None
    for tentativa in range(4):
        try:
            http("POST", f"/api/agente/xml?job={job_id}&chave={chave}", bruto=dados, timeout=60)
            return
        except ErroServidor as e:
            ultimo = e
            if "HTTP 4" in str(e):  # erro do pedido: repetir não adianta
                break
            time.sleep(2 * (tentativa + 1))
    raise ErroServidor(f"não consegui enviar o XML ao servidor ({ultimo})")


def executar(tarefa):
    job_id, itens, cfg = tarefa["job"], tarefa["itens"], tarefa["cfg"]
    site = tarefa["site_url"]
    pausa = (cfg["pausa_min"], cfg["pausa_max"])
    pausa_longa = (cfg["pausa_longa_min"], cfg["pausa_longa_max"])
    lote, tmax = int(cfg["lote_pausa_longa"]), cfg["tempo_max_download"]

    rel = Relato(job_id)
    tmp = tempfile.mkdtemp(prefix="nfe_agente_")

    def xmls():
        return set(glob.glob(os.path.join(tmp, "*.xml")))

    def aguardar_download(antes):
        fim = time.time() + tmax
        while time.time() < fim:
            novos = xmls() - antes
            if novos and not glob.glob(os.path.join(tmp, "*.crdownload")):
                return novos.pop()
            time.sleep(0.5)
        return None

    def fechar_popup(driver):
        try:
            for b in driver.find_elements(By.CSS_SELECTOR, "button.swal-button--cancel"):
                if b.is_displayed():
                    driver.execute_script("arguments[0].click();", b)
                    time.sleep(0.8)
        except Exception:  # noqa: BLE001
            pass

    def clicar(driver, el):
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(random.uniform(0.8, 1.5))
        try:
            el.click()
        except (ElementClickInterceptedException, ElementNotInteractableException):
            driver.execute_script("arguments[0].click();", el)

    threading.Thread(target=rel.batimento, daemon=True).start()
    fim_status = "concluido"
    driver = None
    try:
        rel.log(f"Tarefa recebida: {len(itens)} nota(s). Abrindo o Google Chrome neste computador…")
        rel.enviar()
        try:
            driver = criar_driver(tmp)
        except Exception as e:  # noqa: BLE001
            det = (str(e).strip().splitlines() or [""])[0][:200]
            rel.log("Não foi possível abrir o Chrome neste computador. Verifique se o Google Chrome está "
                    f"instalado e se há internet (o driver é baixado na 1ª vez). Detalhe: {det}", "erro")
            fim_status = "erro"
            return

        falhas_seguidas = 0
        for pos, (num, chave) in enumerate(itens, 1):
            if rel.parar.is_set():
                fim_status = "parado"
                break
            rel.atual = chave
            rel.log(f"[{num}] Processando chave: {chave}")
            try:
                driver.get(site)
                wait = WebDriverWait(driver, 30)
                fechar_popup(driver)
                campo = wait.until(EC.element_to_be_clickable((By.ID, "accessKey")))
                campo.clear()
                campo.send_keys(chave)
                clicar(driver, wait.until(EC.element_to_be_clickable((By.XPATH, XP_BUSCAR))))
                botao_xml = wait.until(EC.visibility_of_element_located((By.XPATH, XP_BAIXAR_XML)))
                antes = xmls()
                clicar(driver, botao_xml)
                novo = aguardar_download(antes)
                if novo is None:
                    raise TimeoutError("o arquivo XML não apareceu na pasta")
                enviar_xml(job_id, chave, novo)
                os.remove(novo)
                rel.baixadas += 1
                falhas_seguidas = 0
                rel.log(f"   OK: {chave}.xml enviado ao servidor", "ok")
            except Exception as e:  # noqa: BLE001
                msg = (str(e).strip().splitlines() or [""])[0][:160]
                rel.falhas.append({"linha": num, "chave": chave, "erro": f"{type(e).__name__}: {msg}"})
                rel.log(f"   ERRO na chave {chave}: {type(e).__name__}: {msg}", "erro")
                falhas_seguidas += 1
                for velho in xmls():  # apaga sobras de um download que falhou
                    try:
                        os.remove(velho)
                    except OSError:
                        pass
                try:  # a janela do Chrome foi fechada?
                    if not driver.window_handles:
                        raise RuntimeError
                except Exception:  # noqa: BLE001
                    rel.log("A janela do Chrome foi fechada. Execução interrompida.", "erro")
                    rel.proc = pos
                    fim_status = "erro"
                    break
                if falhas_seguidas >= 5 and falhas_seguidas % 5 == 0:
                    rel.log(f"{falhas_seguidas} falhas seguidas — o site pode estar bloqueando ou pedindo a "
                            "verificação anti-robô: resolva na janela do Chrome.", "aviso")
            rel.proc = pos
            rel.enviar()

            if pos >= len(itens) or rel.parar.is_set():
                continue
            if rel.baixadas and lote and rel.baixadas % lote == 0:
                p = random.uniform(*pausa_longa)
                rel.log(f"   (pausa longa de {p:.0f}s)")
                dormir(rel, p)
            else:
                dormir(rel, random.uniform(*pausa))
        if rel.parar.is_set() and fim_status == "concluido" and rel.proc < len(itens):
            fim_status = "parado"
    except Exception as e:  # noqa: BLE001
        rel.log(f"Erro inesperado no Agente: {type(e).__name__}: {e}", "erro")
        fim_status = "erro"
    finally:
        rel.encerrado = True
        if driver is not None:
            try:
                driver.quit()
            except Exception:  # noqa: BLE001
                pass
        shutil.rmtree(tmp, ignore_errors=True)
        rel.atual = ""
        for _ in range(3):  # entrega o relatório final (com algumas tentativas)
            rel.enviar(fim=fim_status)
            if not rel.logs and rel.falhas_conexao == 0:
                break
            time.sleep(2)
        print(f"  Tarefa encerrada: {fim_status} — {rel.baixadas} XML(s) baixado(s).\n")


def main():
    global SERVIDOR
    print("=" * 62)
    print(" Agente Baixador de XML de NF-e")
    SERVIDOR = ler_servidor()
    print(f" Servidor: http://{SERVIDOR}/")
    print(" Deixe esta janela ABERTA. Clique em Executar na página do servidor.")
    print(" Para sair: feche esta janela.")
    print("=" * 62)
    avisou = False
    while True:
        try:
            tarefa = http("GET", "/api/agente/proximo", timeout=15)
            avisou = False
        except ErroServidor as e:
            if not avisou:
                print(f"  {time.strftime('%H:%M:%S')} Servidor indisponível ({e}). Tentando de novo…")
                avisou = True
            time.sleep(5)
            continue
        if tarefa.get("job"):
            executar(tarefa)
        else:
            time.sleep(2)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
