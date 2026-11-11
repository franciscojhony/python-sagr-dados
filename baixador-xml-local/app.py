#!/usr/bin/env python3
"""
Baixador de XML de NF-e  (consultadanfe.com)

Aplicação web LOCAL: roda no seu computador e é usada pelo navegador.
  1. Escolha a planilha (.xlsx) com as chaves de acesso
  2. Informe a linha inicial e final
  3. Escolha a pasta onde salvar os XMLs
  4. Clique em Executar

Por baixo, o programa abre uma janela real do Google Chrome (Selenium) e repete
o que você faria à mão: cola a chave, clica em "Imprimir DANFE" e depois em
"Baixar XML".
"""
import glob
import os
import random
import subprocess
import sys
import tempfile
import threading
import time
import uuid
import webbrowser
from pathlib import Path

try:
    import pandas as pd
    import openpyxl  # noqa: F401
    import selenium  # noqa: F401
    from flask import Flask, jsonify, request, send_from_directory
except ImportError as _e:
    print(f"\nFalta um pacote Python: {_e.name}\nInstale com:  pip install {_e.name}\n")
    input("Pressione Enter para sair...")
    sys.exit(1)

BASE = Path(__file__).resolve().parent
SITE_URL = os.environ.get("NFE_SITE_URL", "https://consultadanfe.com/")
HOST = "127.0.0.1"  # somente este computador consegue acessar
PORT = int(os.environ.get("NFE_PORT", "8765"))

XP_BUSCAR = "//button[normalize-space(.)='Imprimir DANFE']"
XP_BAIXAR_XML = "//button[normalize-space(.)='Baixar XML']"

app = Flask(__name__, static_folder=str(BASE / "static"), static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = 300 * 1024 * 1024
UPLOAD_DIR = Path(tempfile.gettempdir()) / "baixador_nfe_uploads"
UPLOAD_DIR.mkdir(exist_ok=True)


# --------------------------------------------------------------------------- #
# Estado global (um único trabalho por vez)
# --------------------------------------------------------------------------- #
class Planilha:
    nome = ""
    df = None


class Job:
    def __init__(self):
        self.lock = threading.Lock()
        self.reset()

    def reset(self):
        self.status = "ocioso"  # ocioso | rodando | parando | concluido | parado | erro
        self.logs = []
        self.total = 0
        self.processadas = 0
        self.baixadas = 0
        self.existentes = 0
        self.invalidas = 0
        self.falhas = []
        self.atual = ""
        self.pasta = ""
        self.simulado = False
        self.t0 = None
        self.t1 = None
        self.stop = threading.Event()
        self.logado = False
        self.fim_ok = False


PLANILHA = Planilha()
JOB = Job()   # download dos XMLs
IMP = Job()   # importação no SISGI


def log(msg, nivel="info"):
    with JOB.lock:
        if len(JOB.logs) < 50000:
            JOB.logs.append({"t": time.strftime("%H:%M:%S"), "n": nivel, "m": msg})


def so_digitos(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    return "".join(ch for ch in str(v) if ch.isdigit())


# --------------------------------------------------------------------------- #
# Planilha
# --------------------------------------------------------------------------- #
def ler_planilha(caminho: Path) -> pd.DataFrame:
    if caminho.suffix.lower() == ".csv":
        for enc in ("utf-8-sig", "latin-1"):
            try:
                return pd.read_csv(caminho, dtype=str, sep=None, engine="python", encoding=enc)
            except UnicodeDecodeError:
                continue
    return pd.read_excel(caminho, dtype=str, engine="openpyxl")


def detectar_coluna(df: pd.DataFrame) -> str:
    cols = [str(c) for c in df.columns]
    if "CHAVE_ACESSO_NOTA" in cols:
        return "CHAVE_ACESSO_NOTA"
    melhor, melhor_pct = None, 0.0
    for c in df.columns:
        amostra = df[c].dropna().head(200)
        if amostra.empty:
            continue
        pct = sum(len(so_digitos(v)) == 44 for v in amostra) / len(amostra)
        if pct > melhor_pct:
            melhor, melhor_pct = str(c), pct
    if melhor and melhor_pct >= 0.5:
        return melhor
    for c in cols:
        if "chave" in c.lower():
            return c
    return cols[0]


def info_coluna(nome: str) -> dict:
    serie = PLANILHA.df[nome]
    validas = sum(len(so_digitos(v)) == 44 for v in serie)
    return {"total": int(len(serie)), "validas": int(validas)}


@app.post("/api/planilha")
def api_planilha():
    arq = request.files.get("arquivo")
    if not arq or not arq.filename:
        return jsonify(erro="Nenhum arquivo enviado."), 400
    ext = Path(arq.filename).suffix.lower()
    if ext not in (".xlsx", ".xlsm", ".csv"):
        return jsonify(erro="Formato não suportado. Use .xlsx, .xlsm ou .csv."), 400
    tmp = UPLOAD_DIR / f"{uuid.uuid4().hex}{ext}"
    arq.save(tmp)
    try:
        df = ler_planilha(tmp)
    except Exception as e:  # noqa: BLE001
        return jsonify(erro=f"Não consegui ler a planilha: {e}"), 400
    finally:
        try:
            tmp.unlink()
        except OSError:
            pass
    if df.empty or len(df.columns) == 0:
        return jsonify(erro="A planilha está vazia."), 400
    df.columns = [str(c) for c in df.columns]
    PLANILHA.nome, PLANILHA.df = arq.filename, df
    sugerida = detectar_coluna(df)
    return jsonify(nome=arq.filename, colunas=list(df.columns), sugerida=sugerida,
                   **info_coluna(sugerida))


@app.get("/api/coluna")
def api_coluna():
    nome = request.args.get("nome", "")
    if PLANILHA.df is None or nome not in PLANILHA.df.columns:
        return jsonify(erro="Coluna não encontrada."), 400
    return jsonify(info_coluna(nome))


# --------------------------------------------------------------------------- #
# Seletor de pastas (o navegador não revela caminhos, então listamos no servidor)
# --------------------------------------------------------------------------- #
def atalhos():
    home = Path.home()
    out = [("Início", str(home))]
    for n in ("Documents", "Documentos", "Downloads", "Desktop", "Área de Trabalho"):
        p = home / n
        if p.is_dir():
            out.append((n, str(p)))
    if os.name == "nt":
        import string
        for letra in string.ascii_uppercase:
            d = f"{letra}:\\"
            if os.path.exists(d):
                out.append((d, d))
    else:
        out.append(("/", "/"))
    return [{"nome": n, "caminho": c} for n, c in out]


@app.get("/api/pastas")
def api_pastas():
    p = request.args.get("path") or str(Path.home())
    path = Path(p).expanduser()
    while not path.is_dir() and path != path.parent:
        path = path.parent
    try:
        dirs = sorted(
            (d.name for d in path.iterdir()
             if d.is_dir() and not d.name.startswith(".")),
            key=str.lower)
    except (PermissionError, OSError):
        dirs = []
    pai = str(path.parent) if path.parent != path else None
    return jsonify(path=str(path), pai=pai, pastas=dirs, atalhos=atalhos())


@app.post("/api/pastas/nova")
def api_pasta_nova():
    d = request.get_json(force=True)
    nome = (d.get("nome") or "").strip()
    if not nome or any(c in nome for c in '\\/:*?"<>|'):
        return jsonify(erro="Nome de pasta inválido."), 400
    novo = Path(d.get("path", "")) / nome
    try:
        novo.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        return jsonify(erro=f"Não foi possível criar a pasta: {e}"), 400
    return jsonify(path=str(novo))


@app.post("/api/abrir-pasta")
def api_abrir_pasta():
    pasta = IMP.pasta if request.args.get("de") == "sisgi" else JOB.pasta
    if not pasta or not os.path.isdir(pasta):
        return jsonify(erro="Pasta não encontrada."), 400
    try:
        if sys.platform.startswith("win"):
            os.startfile(pasta)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", pasta])
        else:
            subprocess.Popen(["xdg-open", pasta])
    except Exception as e:  # noqa: BLE001
        return jsonify(erro=str(e)), 500
    return jsonify(ok=True)


# --------------------------------------------------------------------------- #
# Execução
# --------------------------------------------------------------------------- #
def dormir(seg):
    """Dorme em fatias curtas para responder rápido ao botão Parar."""
    fim = time.time() + seg
    while time.time() < fim and not JOB.stop.is_set():
        time.sleep(0.3)


def criar_driver(pasta, ignorar_alertas=False):
    from selenium import webdriver

    opts = webdriver.ChromeOptions()
    opts.add_experimental_option("prefs", {
        "download.default_directory": pasta,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    })
    # Variáveis só para testes automatizados; o uso normal NÃO usa modo invisível,
    # pois o site costuma bloquear.
    if ignorar_alertas:  # não aborta se aparecer alert() no site
        opts.set_capability("unhandledPromptBehavior", "ignore")
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
        try:  # Selenium antigo / sem acesso ao Selenium Manager: usa webdriver_manager
            from selenium.webdriver.chrome.service import Service
            from webdriver_manager.chrome import ChromeDriverManager
            return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
        except ImportError:
            raise e1


def executar(cfg, itens):
    """itens = [(numero_da_linha, valor_bruto)]"""
    from selenium.common.exceptions import (ElementClickInterceptedException,
                                            ElementNotInteractableException)
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait

    pasta = cfg["pasta"]
    pausa = (cfg["pausa_min"], cfg["pausa_max"])
    pausa_longa = (cfg["pausa_longa_min"], cfg["pausa_longa_max"])
    lote = cfg["lote_pausa_longa"]
    tmax = cfg["tempo_max_download"]

    def xmls():
        return set(glob.glob(os.path.join(pasta, "*.xml")))

    def aguardar_download(antes):
        fim = time.time() + tmax
        while time.time() < fim:
            novos = xmls() - antes
            if novos and not glob.glob(os.path.join(pasta, "*.crdownload")):
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

    log("Abrindo o Google Chrome… (não feche a janela que vai aparecer)")
    try:
        driver = criar_driver(pasta)
    except Exception as e:  # noqa: BLE001
        det = (str(e).strip().splitlines() or [""])[0][:200]
        log("Não foi possível abrir o Chrome. Verifique se o Google Chrome está instalado "
            "e atualizado e se há internet (o driver é baixado na 1ª vez). "
            f"Detalhe: {det}", "erro")
        JOB.logado = True
        raise

    falhas_seguidas = 0
    try:
        for num, bruto in itens:
            if JOB.stop.is_set():
                break
            chave = so_digitos(bruto)
            JOB.atual = chave
            if len(chave) != 44:
                JOB.invalidas += 1
                JOB.processadas += 1
                log(f"[{num}] Pulando: chave inválida ({chave or 'vazia'}).", "aviso")
                continue
            destino = os.path.join(pasta, f"{chave}.xml")
            if os.path.exists(destino):
                JOB.existentes += 1
                JOB.processadas += 1
                log(f"[{num}] Já baixada, pulando: {chave}")
                continue

            log(f"[{num}] Processando chave: {chave}")
            try:
                driver.get(SITE_URL)
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
                os.replace(novo, destino)
                JOB.baixadas += 1
                falhas_seguidas = 0
                log(f"   OK: {chave}.xml", "ok")
            except Exception as e:  # noqa: BLE001
                msg = (str(e).strip().splitlines() or [""])[0][:160]
                JOB.falhas.append({"linha": num, "chave": chave, "erro": f"{type(e).__name__}: {msg}"})
                log(f"   ERRO na chave {chave}: {type(e).__name__}: {msg}", "erro")
                falhas_seguidas += 1
                try:  # a janela foi fechada?
                    _ = driver.window_handles
                    if not driver.window_handles:
                        raise RuntimeError
                except Exception:  # noqa: BLE001
                    log("A janela do Chrome foi fechada. Execução interrompida.", "erro")
                    JOB.processadas += 1
                    break
                if falhas_seguidas >= 5 and falhas_seguidas % 5 == 0:
                    log(f"{falhas_seguidas} falhas seguidas — o site pode estar bloqueando "
                        "ou exigindo a verificação anti-robô na janela do Chrome.", "aviso")
            JOB.processadas += 1

            if JOB.stop.is_set():
                break
            if JOB.baixadas and lote and JOB.baixadas % lote == 0 and JOB.processadas < JOB.total:
                p = random.uniform(*pausa_longa)
                log(f"   (pausa longa de {p:.0f}s)")
                dormir(p)
            else:
                dormir(random.uniform(*pausa))
    finally:
        try:
            driver.quit()
        except Exception:  # noqa: BLE001
            pass


def simular(itens, pasta):
    novas = 0
    for num, bruto in itens:
        if JOB.stop.is_set():
            break
        chave = so_digitos(bruto)
        JOB.processadas += 1
        if len(chave) != 44:
            JOB.invalidas += 1
            log(f"[{num}] Chave inválida ({chave or 'vazia'}).", "aviso")
        elif os.path.exists(os.path.join(pasta, f"{chave}.xml")):
            JOB.existentes += 1
            if JOB.existentes <= 20:
                log(f"[{num}] Já baixada: {chave}")
        else:
            novas += 1
            if novas <= 20:
                log(f"[{num}] Seria baixada: {chave}")
    if novas > 20:
        log(f"… e mais {novas - 20} chave(s) a baixar.")
    log(f"Simulação: {novas} a baixar, {JOB.existentes} já existem, {JOB.invalidas} inválida(s).", "ok")
    JOB.baixadas = 0


def trabalho(cfg, itens):
    JOB.t0 = time.time()
    try:
        if cfg["simular"]:
            log("MODO SIMULAÇÃO: nada será baixado e o Chrome não será aberto.", "aviso")
            simular(itens, cfg["pasta"])
        else:
            executar(cfg, itens)
        JOB.status = "parado" if JOB.stop.is_set() else "concluido"
        if not cfg["simular"]:
            log(f"Processo finalizado! {JOB.baixadas} XML(s) baixado(s), "
                f"{len(JOB.falhas)} falha(s). Pasta: {cfg['pasta']}",
                "ok" if not JOB.falhas else "aviso")
            if JOB.falhas:
                log("Dica: é só executar de novo com o mesmo intervalo — as chaves já "
                    "baixadas são puladas e só as que falharam são tentadas outra vez.")
    except Exception as e:  # noqa: BLE001
        JOB.status = "erro"
        if not getattr(JOB, "logado", False):
            log(f"Erro inesperado: {type(e).__name__}: {(str(e).strip().splitlines() or [''])[0][:200]}", "erro")
    finally:
        JOB.t1 = time.time()
        JOB.atual = ""


def _num(d, chave, padrao, minimo=0.0):
    try:
        v = float(d.get(chave, padrao))
    except (TypeError, ValueError):
        v = padrao
    return max(v, minimo)


@app.post("/api/iniciar")
def api_iniciar():
    if JOB.status in ("rodando", "parando") or IMP.status in ("rodando", "parando"):
        return jsonify(erro="Já existe uma execução em andamento."), 409
    d = request.get_json(force=True)
    if PLANILHA.df is None:
        return jsonify(erro="Escolha a planilha primeiro."), 400
    coluna = d.get("coluna")
    if coluna not in PLANILHA.df.columns:
        return jsonify(erro="Escolha a coluna das chaves de acesso."), 400
    try:
        inicio = int(d.get("inicio") or 1)
        fim_raw = d.get("fim")
        fim = int(fim_raw) if fim_raw not in (None, "") else len(PLANILHA.df)
    except ValueError:
        return jsonify(erro="Linha inicial/final inválida."), 400
    total_linhas = len(PLANILHA.df)
    if inicio < 1 or inicio > total_linhas:
        return jsonify(erro=f"A linha inicial deve estar entre 1 e {total_linhas}."), 400
    if fim < inicio:
        return jsonify(erro="A linha final deve ser maior ou igual à inicial."), 400
    fim = min(fim, total_linhas)
    pasta = (d.get("pasta") or "").strip()
    if not pasta:
        return jsonify(erro="Escolha a pasta onde salvar os XMLs."), 400
    pasta = os.path.abspath(os.path.expanduser(pasta))
    try:
        os.makedirs(pasta, exist_ok=True)
        teste = os.path.join(pasta, ".teste_escrita")
        with open(teste, "w") as f:
            f.write("ok")
        os.remove(teste)
    except OSError as e:
        return jsonify(erro=f"Sem permissão para gravar nessa pasta: {e}"), 400

    o = d.get("opcoes") or {}
    cfg = {
        "pasta": pasta,
        "simular": bool(d.get("simular")),
        "pausa_min": _num(o, "pausa_min", 5),
        "pausa_max": _num(o, "pausa_max", 9),
        "lote_pausa_longa": int(_num(o, "lote_pausa_longa", 10)),
        "pausa_longa_min": _num(o, "pausa_longa_min", 30),
        "pausa_longa_max": _num(o, "pausa_longa_max", 60),
        "tempo_max_download": _num(o, "tempo_max_download", 40, 5),
    }
    cfg["pausa_max"] = max(cfg["pausa_max"], cfg["pausa_min"])
    cfg["pausa_longa_max"] = max(cfg["pausa_longa_max"], cfg["pausa_longa_min"])

    serie = PLANILHA.df[coluna]
    itens = [(i + 1, serie.iloc[i]) for i in range(inicio - 1, fim)]

    JOB.reset()
    JOB.status = "rodando"
    JOB.total = len(itens)
    JOB.pasta = pasta
    JOB.simulado = cfg["simular"]
    log(f"Planilha: {PLANILHA.nome} | coluna: {coluna} | total: {total_linhas} linhas")
    log(f"Processando linhas {inicio} a {fim} ({len(itens)} chaves) → {pasta}")
    if not cfg["simular"]:
        media = (cfg["pausa_min"] + cfg["pausa_max"]) / 2 + 8
        extra = (len(itens) // max(cfg["lote_pausa_longa"], 1)) * \
                (cfg["pausa_longa_min"] + cfg["pausa_longa_max"]) / 2 if cfg["lote_pausa_longa"] else 0
        log(f"Tempo estimado: cerca de {int((len(itens) * media + extra) / 60) + 1} min.")
    threading.Thread(target=trabalho, args=(cfg, itens), daemon=True).start()
    return jsonify(ok=True)


@app.post("/api/parar")
def api_parar():
    if JOB.status == "rodando":
        JOB.status = "parando"
        JOB.stop.set()
        log("Parando… vou terminar a nota atual e fechar o navegador.", "aviso")
    return jsonify(ok=True)


@app.get("/api/status")
def api_status():
    desde = int(request.args.get("desde", 0))
    t0 = JOB.t0
    fim = JOB.t1 or time.time()
    with JOB.lock:
        logs = JOB.logs[desde:]
        n = len(JOB.logs)
    return jsonify(
        status=JOB.status, total=JOB.total, processadas=JOB.processadas,
        baixadas=JOB.baixadas, existentes=JOB.existentes, invalidas=JOB.invalidas,
        falhas=JOB.falhas, atual=JOB.atual, pasta=JOB.pasta, simulado=JOB.simulado,
        decorrido=int(fim - t0) if t0 else 0, logs=logs, n=n)


@app.post("/api/encerrar")
def api_encerrar():
    if JOB.status in ("rodando", "parando") or IMP.status in ("rodando", "parando"):
        return jsonify(erro="Pare a execução antes de encerrar."), 409
    threading.Timer(0.5, lambda: os._exit(0)).start()
    return jsonify(ok=True)



# --------------------------------------------------------------------------- #
# Etapa 4 — Importar os XMLs no SISGI
# --------------------------------------------------------------------------- #
import xml.etree.ElementTree as ET  # noqa: E402
from collections import defaultdict  # noqa: E402

URL_SISGI_PADRAO = os.environ.get("NFE_SISGI_URL", "http://sisgi.hmg.seduc.mt.gov.br/despesas_lista")
ARQUIVO_LOG_SISGI = "sisgi_importados.txt"  # guardado dentro da pasta dos XMLs


def logi(msg, nivel="info"):
    with IMP.lock:
        if len(IMP.logs) < 50000:
            IMP.logs.append({"t": time.strftime("%H:%M:%S"), "n": nivel, "m": msg})


def dormir_i(seg):
    fim = time.time() + seg
    while time.time() < fim and not IMP.stop.is_set():
        time.sleep(0.3)


def mes_emissao(caminho):
    """AAAA-MM da data de emissão (dhEmi/dEmi) — o SISGI só aceita notas do mês de competência."""
    for _, el in ET.iterparse(caminho, events=("end",)):
        tag = el.tag.split("}")[-1]
        if tag in ("dhEmi", "dEmi") and el.text:
            return el.text.strip()[:7]
    return None


def planejar_importacao(cfg):
    """Devolve (por_mes, sem_data, ja_enviados:int) conforme as configurações."""
    pasta = cfg["pasta"]
    enviados = set()
    log_path = os.path.join(pasta, ARQUIVO_LOG_SISGI)
    if not cfg["reenviar"] and os.path.exists(log_path):
        with open(log_path, encoding="utf-8") as f:
            enviados = {l.strip() for l in f if l.strip()}
    por_mes, sem_data = defaultdict(list), []
    for arq in sorted(glob.glob(os.path.join(pasta, "*.xml"))):
        if os.path.basename(arq) in enviados:
            continue
        try:
            mes = mes_emissao(arq)
        except ET.ParseError:
            mes = None
        if not mes:
            sem_data.append(arq)
            continue
        if cfg["mes_filtro"] and mes != cfg["mes_filtro"]:
            continue
        por_mes[mes].append(arq)
    if cfg["limite"]:
        restante = cfg["limite"]
        for mes in sorted(por_mes):
            por_mes[mes] = por_mes[mes][:max(restante, 0)]
            restante -= len(por_mes[mes])
    return por_mes, sem_data, len(enviados)


def executar_importacao(cfg, por_mes):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait

    pasta, url, lote_n = cfg["pasta"], cfg["url"], cfg["lote"]
    log_path = os.path.join(pasta, ARQUIVO_LOG_SISGI)

    logi("Abrindo o Google Chrome… (não feche a janela que vai aparecer)")
    try:
        driver = criar_driver(pasta, ignorar_alertas=True)
    except Exception as e:  # noqa: BLE001
        det = (str(e).strip().splitlines() or [""])[0][:200]
        logi(f"Não foi possível abrir o Chrome. Detalhe: {det}", "erro")
        IMP.logado = True
        raise
    try:
        driver.maximize_window()
    except Exception:  # noqa: BLE001
        pass

    def mensagens_da_pagina():
        sel = ".alert, .toast, .swal2-popup, .swal-modal, .mensagem, .flash, [role=alert]"
        out = []
        for e in driver.find_elements(By.CSS_SELECTOR, sel):
            try:
                if e.is_displayed() and e.text.strip():
                    out.append(e.text.strip())
            except Exception:  # noqa: BLE001
                pass
        return out

    try:
        driver.get(url)
        logi(">>> FAÇA O LOGIN na janela do Chrome. O programa continua sozinho quando a lista de despesas aparecer.", "aviso")
        vistos = set()

        def login_ok(d):
            if IMP.stop.is_set():
                return True
            try:
                return bool(d.find_elements(By.ID, "btnAbrirImportarXml"))
            except Exception:  # noqa: BLE001  (ex.: alert aberto na tela)
                try:
                    txt = d.switch_to.alert.text
                    if txt not in vistos:
                        vistos.add(txt)
                        logi(f"(alerta do site: {txt!r} — clique OK na janela)", "aviso")
                except Exception:  # noqa: BLE001
                    pass
                return False

        try:
            WebDriverWait(driver, cfg["tempo_login"], poll_frequency=1).until(login_ok)
        except Exception:  # noqa: BLE001
            logi("Tempo de login esgotado (a lista de despesas não apareceu).", "erro")
            IMP.logado = True
            raise
        if IMP.stop.is_set():
            return
        logi("Login detectado. Iniciando importação…", "ok")

        def abrir_modal():
            wait = WebDriverWait(driver, 20)
            btn = wait.until(EC.element_to_be_clickable((By.ID, "btnAbrirImportarXml")))
            driver.execute_script("arguments[0].click();", btn)
            wait.until(EC.visibility_of_element_located((By.ID, "modal_arquivos_xml")))

        for mes in sorted(por_mes):
            arquivos = por_mes[mes]
            for i in range(0, len(arquivos), lote_n):
                if IMP.stop.is_set():
                    return
                lote = arquivos[i:i + lote_n]
                logi(f"[{mes}] lote {i // lote_n + 1}: {len(lote)} arquivo(s)")
                IMP.atual = f"{mes} · lote {i // lote_n + 1}"
                try:
                    driver.get(url)
                    abrir_modal()
                    comp = driver.find_element(By.ID, "modal_competencia")
                    driver.execute_script(
                        "arguments[0].value = arguments[1];"
                        "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
                        "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", comp, mes)
                    driver.find_element(By.ID, "modal_arquivos_xml").send_keys("\n".join(lote))
                    antigo = driver.find_element(By.TAG_NAME, "html")
                    driver.execute_script("arguments[0].click();",
                                          driver.find_element(By.ID, "btnConfirmarImportarXml"))

                    # o SISGI responde com alert() e/ou recarrega a página: lê, confirma e segue
                    fim = time.time() + cfg["tempo_importar"]
                    ultimo_alerta, respondeu = None, False
                    while time.time() < fim and not IMP.stop.is_set():
                        try:
                            al = driver.switch_to.alert
                            txt = al.text
                            al.accept()
                            logi("SISGI: " + txt.replace("\n", " | "))
                            ultimo_alerta = time.time()
                            respondeu = True
                            time.sleep(1.5)
                            continue
                        except Exception:  # noqa: BLE001
                            pass
                        try:
                            antigo.is_enabled()  # lança erro se a página já recarregou
                        except Exception as ex:  # noqa: BLE001
                            if "stale" in type(ex).__name__.lower():
                                respondeu = True
                                time.sleep(1.5)
                                break
                        if ultimo_alerta and time.time() - ultimo_alerta > 5:
                            break
                        time.sleep(0.5)
                    else:
                        if not IMP.stop.is_set():
                            logi("⚠ sem resposta no tempo esperado; verifique na tela do Chrome.", "aviso")
                    try:
                        for m in mensagens_da_pagina():
                            logi("SISGI: " + m.replace("\n", " | "))
                    except Exception:  # noqa: BLE001
                        pass
                    with open(log_path, "a", encoding="utf-8") as f:
                        for a in lote:
                            f.write(os.path.basename(a) + "\n")
                    IMP.baixadas += len(lote)
                except Exception as e:  # noqa: BLE001
                    msg = (str(e).strip().splitlines() or [""])[0][:160]
                    for a in lote:
                        IMP.falhas.append({"linha": mes, "chave": os.path.basename(a),
                                           "erro": f"{type(e).__name__}: {msg}"})
                    logi(f"ERRO no lote {mes}/{i // lote_n + 1}: {type(e).__name__}: {msg}", "erro")
                IMP.processadas += len(lote)
                dormir_i(2)
    finally:
        # Avisa na tela do Chrome e só fecha quando você clicar em OK (como no notebook)
        try:
            if not IMP.stop.is_set() and IMP.baixadas:
                IMP.fim_ok = True
                driver.execute_script(
                    "setTimeout(function(){alert(arguments[0]);}, 100);",
                    f"Importação finalizada!\n\n{IMP.baixadas} arquivo(s) enviado(s) ao SISGI.\n\n"
                    "Clique em OK para fechar o navegador.")
                logi("Importação finalizada. Clique em OK na janela do Chrome para fechá-la.", "ok")
                fim = time.time() + 600
                apareceu = False
                while time.time() < fim and not IMP.stop.is_set():
                    try:
                        driver.switch_to.alert.text
                        apareceu = True
                    except Exception:  # noqa: BLE001
                        if apareceu:
                            break
                    time.sleep(0.5)
        except Exception:  # noqa: BLE001
            pass
        try:
            driver.quit()
        except Exception:  # noqa: BLE001
            pass


def trabalho_importacao(cfg):
    IMP.t0 = time.time()
    try:
        por_mes, sem_data, ja = planejar_importacao(cfg)
        total = sum(len(v) for v in por_mes.values())
        IMP.total = total
        IMP.existentes = ja
        IMP.invalidas = len(sem_data)
        logi(f"Pasta: {cfg['pasta']}")
        logi(f"XMLs a importar: {total} | já enviados antes: {ja} | ilegíveis/sem data: {len(sem_data)}")
        for mes in sorted(por_mes):
            logi(f"   {mes}: {len(por_mes[mes])} arquivo(s)")
        for a in sem_data:
            logi(f"⚠ sem data de emissão: {os.path.basename(a)}", "aviso")
        if total == 0:
            logi("Nada a enviar.", "aviso")
        elif cfg["simular"]:
            logi("MODO SIMULAÇÃO: nenhum arquivo foi enviado e o Chrome não foi aberto.", "aviso")
        else:
            executar_importacao(cfg, por_mes)
            logi(f"Finalizado. {IMP.baixadas} arquivo(s) enviado(s) ao SISGI"
                 + (f", {len(IMP.falhas)} com erro." if IMP.falhas else "."),
                 "ok" if not IMP.falhas else "aviso")
        IMP.status = "parado" if (IMP.stop.is_set() and not IMP.fim_ok) else "concluido"
    except Exception as e:  # noqa: BLE001
        IMP.status = "erro"
        if not IMP.logado:
            logi(f"Erro inesperado: {type(e).__name__}: {(str(e).strip().splitlines() or [''])[0][:200]}", "erro")
    finally:
        IMP.t1 = time.time()
        IMP.atual = ""


def _status_json(job, desde):
    t0 = job.t0
    fim = job.t1 or time.time()
    with job.lock:
        logs = job.logs[desde:]
        n = len(job.logs)
    return jsonify(
        status=job.status, total=job.total, processadas=job.processadas,
        baixadas=job.baixadas, existentes=job.existentes, invalidas=job.invalidas,
        falhas=job.falhas, atual=job.atual, pasta=job.pasta, simulado=job.simulado,
        decorrido=int(fim - t0) if t0 else 0, logs=logs, n=n)


@app.get("/api/sisgi/padrao")
def api_sisgi_padrao():
    return jsonify(url=URL_SISGI_PADRAO)


@app.post("/api/sisgi/iniciar")
def api_sisgi_iniciar():
    if JOB.status in ("rodando", "parando") or IMP.status in ("rodando", "parando"):
        return jsonify(erro="Já existe uma execução em andamento."), 409
    d = request.get_json(force=True)
    pasta = os.path.abspath(os.path.expanduser((d.get("pasta") or "").strip() or "."))
    if not (d.get("pasta") or "").strip():
        return jsonify(erro="Escolha a pasta onde estão os XMLs."), 400
    if not os.path.isdir(pasta):
        return jsonify(erro="Pasta não encontrada."), 400
    if not glob.glob(os.path.join(pasta, "*.xml")):
        return jsonify(erro="Não há arquivos .xml nessa pasta."), 400
    url = (d.get("url") or "").strip()
    if not url.lower().startswith(("http://", "https://")):
        return jsonify(erro="Informe o endereço do SISGI começando com http:// ou https://"), 400
    mes = (d.get("mes") or "").strip() or None
    if mes and not (len(mes) == 7 and mes[4] == "-" and mes[:4].isdigit() and mes[5:].isdigit()):
        return jsonify(erro="Mês de competência inválido (use AAAA-MM)."), 400
    try:
        limite = int(d.get("limite")) if str(d.get("limite") or "").strip() else None
        lote = int(d.get("lote") or 20)
    except ValueError:
        return jsonify(erro="Limite/lote inválido."), 400
    if limite is not None and limite < 1:
        return jsonify(erro="O limite de arquivos deve ser 1 ou mais."), 400
    cfg = {"pasta": pasta, "url": url, "mes_filtro": mes, "limite": limite,
           "lote": min(max(lote, 1), 20),  # limite do sistema: 20 arquivos por vez
           "simular": bool(d.get("simular")), "reenviar": bool(d.get("reenviar")),
           "tempo_login": 300, "tempo_importar": 180}
    IMP.reset()
    IMP.status = "rodando"
    IMP.pasta = pasta
    IMP.simulado = cfg["simular"]
    threading.Thread(target=trabalho_importacao, args=(cfg,), daemon=True).start()
    return jsonify(ok=True)


@app.post("/api/sisgi/parar")
def api_sisgi_parar():
    if IMP.status == "rodando":
        IMP.status = "parando"
        IMP.stop.set()
        logi("Parando… vou fechar o navegador.", "aviso")
    return jsonify(ok=True)


@app.get("/api/sisgi/status")
def api_sisgi_status():
    return _status_json(IMP, int(request.args.get("desde", 0)))


@app.get("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


def main():
    url = f"http://{HOST}:{PORT}/"
    print("=" * 60)
    print(" Baixador de XML de NF-e")
    print(f" Abra no navegador: {url}")
    print(" Para encerrar: feche esta janela (ou use o botão 'Encerrar' na página).")
    print("=" * 60)
    if os.environ.get("NFE_NO_BROWSER") != "1":
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    app.run(host=HOST, port=PORT, debug=False, threaded=True)


if __name__ == "__main__":
    main()
