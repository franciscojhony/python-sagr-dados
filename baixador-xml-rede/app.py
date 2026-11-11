#!/usr/bin/env python3
"""
Baixador de XML de NF-e  (consultadanfe.com)  —  SERVIDOR + AGENTES

Arquitetura:
  * SERVIDOR (este programa): guarda a planilha de cada pessoa, distribui as chaves,
    recebe os XMLs e mostra o andamento no navegador. NÃO abre o Chrome.
  * AGENTE (agente/agente.py): programinha que roda no computador de CADA usuário,
    abre o Chrome ali (a verificação anti-robô é resolvida pelo próprio usuário),
    baixa os XMLs e os envia para a pasta do servidor.

Fluxo de uso:
  1. O usuário abre o navegador em http://IP-DO-SERVIDOR:8765, escolhe a planilha e as linhas
  2. Clica em Executar
  3. O Agente (aberto no computador dele) detecta a tarefa, abre o Chrome e baixa
  4. Os XMLs chegam na pasta do servidor, e o andamento aparece na página
"""
import glob
import io
import os
import socket
import subprocess
import sys
import tempfile
import threading
import time
import uuid
import webbrowser
import zipfile
from pathlib import Path

try:
    import pandas as pd
    import openpyxl  # noqa: F401
    from flask import Flask, g, jsonify, request, send_file, send_from_directory
except ImportError as _e:
    print(f"\nFalta um pacote Python: {_e.name}\nInstale com:  pip install {_e.name}\n")
    input("Pressione Enter para sair...")
    sys.exit(1)

BASE = Path(__file__).resolve().parent

# =========================================================================== #
#  CONFIGURAÇÕES — o que normalmente você vai querer mudar está AQUI
# =========================================================================== #

# Quantas pessoas podem baixar ao mesmo tempo (cada uma usa o Agente + 1 Chrome no PRÓPRIO computador).
# Quem tentar iniciar além desse limite recebe um aviso para aguardar uma vaga.
MAX_USUARIOS = int(os.environ.get("NFE_MAX_USUARIOS", "5"))

# Pasta NO SERVIDOR onde TODOS os XMLs serão salvos (a mesma para todos).
# Exemplos:  r"D:\NotasFiscais\xmls"   ou   r"C:\git-repositorio\baixador-xml-nfe\xmls_baixados"
#PASTA_XML = os.environ.get("NFE_PASTA_XML", str(BASE / "xmls_baixados"))
PASTA_XML = os.environ.get("NFE_PASTA_XML", r"C:\Users\89821491120\xmls_baixados")

# True  = outros computadores da mesma rede também acessam (http://IP-DO-SERVIDOR:8765)
# False = somente este computador acessa (http://127.0.0.1:8765)   ← uso local, mais seguro
ACESSO_PELA_REDE = True

# False = todos usam a PASTA_XML acima, e ninguém navega nas pastas do servidor (recomendado).
# True  = cada pessoa pode escolher qualquer pasta do servidor (menos seguro numa rede).
PERMITIR_ESCOLHER_PASTA = False

# Segundos esperando o Agente se conectar depois do "Executar" (depois disso a tarefa é cancelada).
TEMPO_ESPERA_AGENTE = int(os.environ.get("NFE_TEMPO_ESPERA_AGENTE", "600"))
# Segundos sem notícias do Agente durante uma tarefa (computador desligou, Agente fechado...).
TEMPO_SEM_SINAL = int(os.environ.get("NFE_TEMPO_SEM_SINAL", "180"))

# Sessões sem atividade (e sem download em andamento) são esquecidas após esse tempo.
HORAS_SESSAO_OCIOSA = 8

# =========================================================================== #

SITE_URL = os.environ.get("NFE_SITE_URL", "https://consultadanfe.com/")
HOST = "0.0.0.0" if ACESSO_PELA_REDE else "127.0.0.1"
PORT = int(os.environ.get("NFE_PORT", "8765"))

app = Flask(__name__, static_folder=str(BASE / "static"), static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = 300 * 1024 * 1024
UPLOAD_DIR = Path(tempfile.gettempdir()) / "baixador_nfe_uploads"
UPLOAD_DIR.mkdir(exist_ok=True)


# --------------------------------------------------------------------------- #
# Estado: UMA sessão por pessoa (identificada por um cookie no navegador)
# --------------------------------------------------------------------------- #
class Planilha:
    def __init__(self):
        self.nome = ""
        self.df = None
        self.fp = None  # "impressão digital": identifica a MESMA planilha em sessões diferentes


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
        self.simulado = False
        self.inicio = 0
        self.fim = 0
        self.agente = False       # um Agente já assumiu a tarefa?
        self.ultimo_sinal = 0.0   # último contato do Agente
        self.base_proc = 0        # linhas resolvidas no servidor (inválidas / já existentes)
        self.itens = []
        self.cfg = None
        self.t0 = None
        self.t1 = None
        self.stop = threading.Event()
        self.logado = False


class Sessao:
    def __init__(self, sid):
        self.sid = sid
        self.planilha = Planilha()
        self.job = Job()
        self.nome = ""
        self.ip = ""
        self.ultimo = time.time()

    @property
    def rotulo(self):
        return self.nome or self.ip or "alguém"

    @property
    def ativa(self):  # usando um Chrome agora?
        return self.job.status in ("rodando", "parando") and not self.job.simulado


SESSOES = {}
LOCK = threading.RLock()
AGENTES_VISTOS = {}  # IP do computador -> hora do último contato do Agente


def log(job, msg, nivel="info"):
    with job.lock:
        if len(job.logs) < 50000:
            job.logs.append({"t": time.strftime("%H:%M:%S"), "n": nivel, "m": msg})


def so_digitos(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    return "".join(ch for ch in str(v) if ch.isdigit())


def ips_deste_computador():
    ips = {"127.0.0.1", "::1"}
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None):
            ips.add(info[4][0])
    except OSError:
        pass
    return ips


IPS_LOCAIS = ips_deste_computador()


def ident(ip):
    """Identifica o computador: todos os IPs do próprio servidor contam como um só."""
    return "servidor" if ip in IPS_LOCAIS else ip


def eh_servidor():
    """True se a requisição vem do próprio computador do servidor."""
    return request.remote_addr in IPS_LOCAIS


def so_servidor(f):
    """Decorador: a rota só funciona no próprio servidor (protege quem acessa pela rede)."""
    from functools import wraps

    @wraps(f)
    def interno(*a, **kw):
        if not eh_servidor():
            return jsonify(erro="Esta ação só pode ser feita no computador servidor."), 403
        return f(*a, **kw)
    return interno


@app.before_request
def identificar_sessao():
    if request.path.startswith(("/static/", "/api/agente/", "/agente.zip")):
        return  # chamadas do Agente não têm sessão de navegador
    sid = request.cookies.get("nfe_sid")
    with LOCK:
        s = SESSOES.get(sid) if sid else None
        if s is None:
            sid = uuid.uuid4().hex
            s = SESSOES[sid] = Sessao(sid)
            g.novo_sid = sid
        s.ultimo = time.time()
        s.ip = request.remote_addr or ""
        g.sessao = s
        limite = time.time() - HORAS_SESSAO_OCIOSA * 3600
        for k in [k for k, v in SESSOES.items() if v.ultimo < limite and not v.ativa]:
            SESSOES.pop(k, None)


@app.after_request
def gravar_cookie(resp):
    sid = getattr(g, "novo_sid", None)
    if sid:
        resp.set_cookie("nfe_sid", sid, max_age=7 * 24 * 3600, httponly=True, samesite="Lax")
    return resp


# --------------------------------------------------------------------------- #
# Planilha (cada pessoa envia a sua, do próprio computador)
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


def info_coluna(df, nome: str) -> dict:
    serie = df[nome]
    validas = sum(len(so_digitos(v)) == 44 for v in serie)
    return {"total": int(len(serie)), "validas": int(validas)}


@app.post("/api/planilha")
def api_planilha():
    sess = g.sessao
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
    sess.planilha.nome, sess.planilha.df = arq.filename, df
    sess.planilha.fp = (arq.filename, len(df), tuple(str(x) for x in df.iloc[0].tolist()))
    sugerida = detectar_coluna(df)
    return jsonify(nome=arq.filename, colunas=list(df.columns), sugerida=sugerida,
                   **info_coluna(df, sugerida))


@app.get("/api/coluna")
def api_coluna():
    df = g.sessao.planilha.df
    nome = request.args.get("nome", "")
    if df is None or nome not in df.columns:
        return jsonify(erro="Coluna não encontrada."), 400
    return jsonify(info_coluna(df, nome))


# --------------------------------------------------------------------------- #
# Configuração exibida na tela
# --------------------------------------------------------------------------- #
@app.get("/api/config")
def api_config():
    return jsonify(max_usuarios=MAX_USUARIOS, pasta=PASTA_XML, escolher_pasta=PERMITIR_ESCOLHER_PASTA,
                   servidor=eh_servidor(), nome=g.sessao.nome)


# --------------------------------------------------------------------------- #
# Seletor de pastas — só existe se PERMITIR_ESCOLHER_PASTA = True
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


def exige_escolher_pasta():
    if not PERMITIR_ESCOLHER_PASTA:
        return jsonify(erro="A escolha de pasta está desativada: os XMLs vão para a pasta do servidor."), 403
    return None


@app.get("/api/pastas")
def api_pastas():
    bloqueio = exige_escolher_pasta()
    if bloqueio:
        return bloqueio
    p = request.args.get("path") or str(Path.home())
    path = Path(p).expanduser()
    while not path.is_dir() and path != path.parent:
        path = path.parent
    try:
        dirs = sorted((d.name for d in path.iterdir() if d.is_dir() and not d.name.startswith(".")),
                      key=str.lower)
    except (PermissionError, OSError):
        dirs = []
    pai = str(path.parent) if path.parent != path else None
    return jsonify(path=str(path), pai=pai, pastas=dirs, atalhos=atalhos())


@app.post("/api/pastas/nova")
def api_pasta_nova():
    bloqueio = exige_escolher_pasta()
    if bloqueio:
        return bloqueio
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
@so_servidor
def api_abrir_pasta():
    pasta = os.path.abspath(PASTA_XML)
    if not os.path.isdir(pasta):
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
# Execução: simulação roda aqui; o download de verdade é feito pelo Agente
# --------------------------------------------------------------------------- #
def simular(job, itens, pasta):
    novas = 0
    for num, bruto in itens:
        if job.stop.is_set():
            break
        chave = so_digitos(bruto)
        job.processadas += 1
        if len(chave) != 44:
            job.invalidas += 1
            log(job, f"[{num}] Chave inválida ({chave or 'vazia'}).", "aviso")
        elif os.path.exists(os.path.join(pasta, f"{chave}.xml")):
            job.existentes += 1
            if job.existentes <= 20:
                log(job, f"[{num}] Já baixada: {chave}")
        else:
            novas += 1
            if novas <= 20:
                log(job, f"[{num}] Seria baixada: {chave}")
    if novas > 20:
        log(job, f"… e mais {novas - 20} chave(s) a baixar.")
    log(job, f"Simulação: {novas} a baixar, {job.existentes} já existem, {job.invalidas} inválida(s).", "ok")
    job.baixadas = 0


def trabalho_simulacao(job, cfg, itens):
    job.t0 = time.time()
    try:
        log(job, "MODO SIMULAÇÃO: nada será baixado e o Agente não será acionado.", "aviso")
        simular(job, itens, cfg["pasta"])
        job.status = "concluido"
    except Exception as e:  # noqa: BLE001
        job.status = "erro"
        log(job, f"Erro inesperado: {type(e).__name__}: {e}", "erro")
    finally:
        job.t1 = time.time()
        job.atual = ""


def _num(d, chave, padrao, minimo=0.0):
    try:
        v = float(d.get(chave, padrao))
    except (TypeError, ValueError):
        v = padrao
    return max(v, minimo)


@app.post("/api/iniciar")
def api_iniciar():
    sess = g.sessao
    job = sess.job
    if job.status in ("rodando", "parando"):
        return jsonify(erro="Você já tem uma execução em andamento."), 409
    d = request.get_json(force=True)
    df = sess.planilha.df
    if df is None:
        return jsonify(erro="Escolha a planilha primeiro."), 400
    coluna = d.get("coluna")
    if coluna not in df.columns:
        return jsonify(erro="Escolha a coluna das chaves de acesso."), 400
    try:
        inicio = int(d.get("inicio") or 1)
        fim_raw = d.get("fim")
        fim = int(fim_raw) if fim_raw not in (None, "") else len(df)
    except ValueError:
        return jsonify(erro="Linha inicial/final inválida."), 400
    total_linhas = len(df)
    if inicio < 1 or inicio > total_linhas:
        return jsonify(erro=f"A linha inicial deve estar entre 1 e {total_linhas}."), 400
    if fim < inicio:
        return jsonify(erro="A linha final deve ser maior ou igual à inicial."), 400
    fim = min(fim, total_linhas)

    # Pasta de destino: a do servidor (padrão) ou a escolhida, se isso estiver liberado.
    pasta = (d.get("pasta") or "").strip() if PERMITIR_ESCOLHER_PASTA else ""
    pasta = os.path.abspath(os.path.expanduser(pasta or PASTA_XML))
    try:
        os.makedirs(pasta, exist_ok=True)
        teste = os.path.join(pasta, f".teste_escrita_{sess.sid[:8]}")
        with open(teste, "w") as f:
            f.write("ok")
        os.remove(teste)
    except OSError as e:
        return jsonify(erro=f"O servidor não consegue gravar na pasta de destino: {e}"), 400

    sess.nome = (d.get("nome") or "").strip()[:40]
    simular_ = bool(d.get("simular"))
    o = d.get("opcoes") or {}
    cfg = {
        "pasta": pasta,
        "simular": simular_,
        "pausa_min": _num(o, "pausa_min", 5),
        "pausa_max": _num(o, "pausa_max", 9),
        "lote_pausa_longa": int(_num(o, "lote_pausa_longa", 10)),
        "pausa_longa_min": _num(o, "pausa_longa_min", 30),
        "pausa_longa_max": _num(o, "pausa_longa_max", 60),
        "tempo_max_download": _num(o, "tempo_max_download", 40, 5),
    }
    cfg["pausa_max"] = max(cfg["pausa_max"], cfg["pausa_min"])
    cfg["pausa_longa_max"] = max(cfg["pausa_longa_max"], cfg["pausa_longa_min"])

    serie = df[coluna]
    itens = [(i + 1, serie.iloc[i]) for i in range(inicio - 1, fim)]

    with LOCK:  # verificações e reserva da vaga são feitas "de uma vez só" (sem corrida)
        if not simular_:
            outros = [s for s in SESSOES.values() if s is not sess and s.ativa]
            if len(outros) >= MAX_USUARIOS:
                return jsonify(erro=f"O servidor já está com {MAX_USUARIOS} downloads em andamento "
                                    "(limite de usuários simultâneos). Aguarde alguém terminar e tente de novo."), 429
            # Evita duas pessoas baixarem as MESMAS linhas da MESMA planilha ao mesmo tempo.
            for s in outros:
                j = s.job
                if s.planilha.fp == sess.planilha.fp and inicio <= j.fim and fim >= j.inicio:
                    return jsonify(erro=f"As linhas {max(inicio, j.inicio)} a {min(fim, j.fim)} já estão sendo "
                                        f"baixadas por {s.rotulo} (linhas {j.inicio} a {j.fim}). "
                                        "Escolha outro intervalo."), 409
        job.reset()
        job.status = "rodando"
        job.total = len(itens)
        job.simulado = simular_
        job.inicio, job.fim = inicio, fim
        job.t0 = time.time()
        job.itens, job.cfg = itens, cfg

    log(job, f"Planilha: {sess.planilha.nome} | coluna: {coluna} | total: {total_linhas} linhas")
    log(job, f"Linhas {inicio} a {fim} ({len(itens)} chaves) → pasta do servidor: {pasta}")
    if simular_:
        threading.Thread(target=trabalho_simulacao, args=(job, cfg, itens), daemon=True).start()
    else:
        log(job, "Aguardando o Agente neste computador… abra o \"Agente Baixador\" (iniciar-agente.bat). "
                 "O Chrome vai abrir no SEU computador.", "aviso")
    return jsonify(ok=True)


@app.post("/api/parar")
def api_parar():
    job = g.sessao.job
    if job.status == "rodando":
        if not job.agente:  # ninguém assumiu a tarefa ainda: cancela direto
            job.status = "parado"
            job.t1 = time.time()
            log(job, "Tarefa cancelada antes de o Agente conectar.", "aviso")
        else:
            job.status = "parando"
            job.stop.set()
            log(job, "Parando… o Agente termina a nota atual e fecha o Chrome.", "aviso")
    return jsonify(ok=True)


@app.get("/api/status")
def api_status():
    sess = g.sessao
    job = sess.job
    desde = int(request.args.get("desde", 0))
    fim = job.t1 or time.time()
    with job.lock:
        logs = job.logs[desde:]
        n = len(job.logs)
    with LOCK:
        ativos = [s for s in SESSOES.values() if s.ativa]
        outros = [dict(nome=s.rotulo, inicio=s.job.inicio, fim=s.job.fim, processadas=s.job.processadas,
                       total=s.job.total, status=s.job.status) for s in ativos if s is not sess]
    return jsonify(
        status=job.status, total=job.total, processadas=job.processadas,
        baixadas=job.baixadas, existentes=job.existentes, invalidas=job.invalidas,
        falhas=job.falhas, atual=job.atual, simulado=job.simulado,
        inicio=job.inicio, fim=job.fim,
        agente=("conectado" if job.agente else "aguardando") if job.status in ("rodando", "parando") and not job.simulado else "",
        agente_online=(time.time() - AGENTES_VISTOS.get(ident(request.remote_addr), 0) < 10),
        decorrido=int(fim - job.t0) if job.t0 else 0, logs=logs, n=n,
        outros=outros, vagas=max(MAX_USUARIOS - len(ativos), 0), max_usuarios=MAX_USUARIOS)


# --------------------------------------------------------------------------- #
# API usada pelo AGENTE (programa que roda no computador do usuário)
# --------------------------------------------------------------------------- #
def _preparar_para_agente(sess):
    """Separa o que o Agente precisa baixar: descarta chaves inválidas e XMLs que já existem."""
    job, cfg = sess.job, sess.job.cfg
    pend, inv, exi = [], 0, 0
    for num, bruto in job.itens:
        chave = so_digitos(bruto)
        if len(chave) != 44:
            inv += 1
            log(job, f"[{num}] Pulando: chave inválida ({chave or 'vazia'}).", "aviso")
        elif os.path.exists(os.path.join(cfg["pasta"], f"{chave}.xml")):
            exi += 1
            if exi <= 20:
                log(job, f"[{num}] Já baixada, pulando: {chave}")
        else:
            pend.append([num, chave])
    if exi > 20:
        log(job, f"… e mais {exi - 20} já baixada(s).")
    job.invalidas, job.existentes = inv, exi
    job.base_proc = inv + exi
    job.processadas = job.base_proc
    return pend


@app.get("/api/agente/proximo")
def agente_proximo():
    ip = ident(request.remote_addr)
    AGENTES_VISTOS[ip] = time.time()
    with LOCK:
        for s in SESSOES.values():
            j = s.job
            if ident(s.ip) == ip and j.status == "rodando" and not j.simulado and not j.agente:
                pend = _preparar_para_agente(s)
                j.agente = True
                j.ultimo_sinal = time.time()
                log(j, f"Agente conectado ({ip if ip != 'servidor' else 'servidor'}): {len(pend)} nota(s) a baixar.", "ok")
                if not pend:
                    j.status = "concluido"
                    j.t1 = time.time()
                    log(j, "Nada a baixar: todas as chaves já existem ou são inválidas.", "aviso")
                    return jsonify(job=None)
                c = j.cfg
                return jsonify(job=s.sid, itens=pend, site_url=SITE_URL,
                               cfg={k: c[k] for k in ("pausa_min", "pausa_max", "lote_pausa_longa",
                                                     "pausa_longa_min", "pausa_longa_max", "tempo_max_download")})
    return jsonify(job=None)


@app.post("/api/agente/relatorio")
def agente_relatorio():
    d = request.get_json(force=True) or {}
    AGENTES_VISTOS[ident(request.remote_addr)] = time.time()
    with LOCK:
        s = SESSOES.get(d.get("job"))
        if not s or not s.job.agente or s.job.status not in ("rodando", "parando"):
            return jsonify(parar=True)  # tarefa não existe mais / foi encerrada: o Agente deve parar
        j = s.job
        j.ultimo_sinal = time.time()
        for item in d.get("logs") or []:
            try:
                log(j, str(item[1])[:500], item[0] if item[0] in ("info", "ok", "aviso", "erro") else "info")
            except (IndexError, TypeError):
                pass
        try:
            j.processadas = j.base_proc + int(d.get("proc", 0))
            j.baixadas = int(d.get("baixadas", 0))
        except (TypeError, ValueError):
            pass
        if isinstance(d.get("falhas"), list):
            j.falhas = [{"linha": f.get("linha"), "chave": str(f.get("chave", ""))[:44],
                         "erro": str(f.get("erro", ""))[:300]} for f in d["falhas"][:5000] if isinstance(f, dict)]
        j.atual = str(d.get("atual", ""))[:44]
        fim = d.get("fim")
        if fim in ("concluido", "parado", "erro"):
            j.status = fim
            j.t1 = time.time()
            j.atual = ""
            log(j, f"Processo finalizado! {j.baixadas} XML(s) baixado(s), {len(j.falhas)} falha(s). "
                   f"Pasta no servidor: {j.cfg['pasta']}", "ok" if not j.falhas else "aviso")
            if j.falhas:
                log(j, "Dica: é só executar de novo com o mesmo intervalo — as chaves já baixadas "
                       "são puladas e só as que falharam são tentadas outra vez.")
        return jsonify(parar=j.stop.is_set())


@app.post("/api/agente/xml")
def agente_xml():
    """Recebe o conteúdo de um XML baixado pelo Agente e grava na pasta do servidor."""
    s = SESSOES.get(request.args.get("job", ""))
    if not s or not s.job.agente or s.job.status not in ("rodando", "parando"):
        return jsonify(erro="Tarefa não encontrada."), 404
    AGENTES_VISTOS[ident(request.remote_addr)] = time.time()
    chave = request.args.get("chave", "")
    if len(chave) != 44 or not chave.isdigit():
        return jsonify(erro="Chave inválida."), 400
    dados = request.get_data()
    if not dados or len(dados) > 5 * 1024 * 1024:
        return jsonify(erro="Arquivo vazio ou grande demais."), 400
    if not dados[:2000].lstrip(b"\xef\xbb\xbf \r\n\t").startswith(b"<"):
        return jsonify(erro="O arquivo recebido não parece um XML."), 400
    pasta = s.job.cfg["pasta"]
    destino = os.path.join(pasta, f"{chave}.xml")
    tmp = os.path.join(pasta, f".{chave}.{uuid.uuid4().hex[:6]}.part")
    try:
        with open(tmp, "wb") as f:
            f.write(dados)
        os.replace(tmp, destino)  # gravação atômica: nunca fica um XML pela metade
    except OSError as e:
        return jsonify(erro=f"O servidor não conseguiu gravar: {e}"), 500
    return jsonify(ok=True)


def vigia():
    """Cancela tarefas cujo Agente nunca apareceu ou parou de responder."""
    while True:
        time.sleep(5)
        agora = time.time()
        with LOCK:
            for s in SESSOES.values():
                j = s.job
                if j.status not in ("rodando", "parando") or j.simulado:
                    continue
                if not j.agente and agora - j.t0 > TEMPO_ESPERA_AGENTE:
                    j.status, j.t1 = "erro", agora
                    log(j, f"O Agente não se conectou em {TEMPO_ESPERA_AGENTE // 60} min. Abra o Agente neste "
                           "computador e clique em Executar de novo.", "erro")
                elif j.agente and agora - j.ultimo_sinal > TEMPO_SEM_SINAL:
                    j.status, j.t1 = "erro", agora
                    log(j, "Perdi o contato com o Agente (computador desligado ou Agente fechado). "
                           "Abra o Agente e execute de novo o mesmo intervalo: o que já foi baixado é pulado.", "erro")


@app.get("/agente.zip")
def agente_zip():
    """Gera o pacote do Agente já com o endereço deste servidor preenchido."""
    host = request.host
    if host.split(":")[0] in ("127.0.0.1", "localhost", "[::1]"):
        ips = meus_ips()
        if ips:
            host = f"{ips[0]}:{PORT}"
    buf = io.BytesIO()
    pasta_ag = BASE / "agente"
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(pasta_ag.glob("*")):
            if f.is_file():
                z.write(f, f"agente-baixador/{f.name}")
        z.writestr("agente-baixador/servidor.txt", host + "\r\n")
        py = BASE / "agente-portatil" / "python"  # Python embutido (veja preparar-agente.bat)
        if py.is_dir():
            for f in py.rglob("*"):
                if f.is_file():
                    z.write(f, "agente-baixador/python/" + f.relative_to(py).as_posix())
    buf.seek(0)
    return send_file(buf, mimetype="application/zip", as_attachment=True, download_name="agente-baixador.zip")


@app.post("/api/encerrar")
@so_servidor
def api_encerrar():
    if any(s.ativa for s in SESSOES.values()):
        return jsonify(erro="Há downloads em andamento. Aguarde ou peça para pararem."), 409
    threading.Timer(0.5, lambda: os._exit(0)).start()
    return jsonify(ok=True)


@app.get("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


def meus_ips():
    """IPs deste computador na rede (para mostrar na janela preta)."""
    ips = set()
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            if not info[4][0].startswith("127."):
                ips.add(info[4][0])
    except OSError:
        pass
    return sorted(ips)


def main():
    # O navegador DESTE computador sempre abre em 127.0.0.1.
    # (Não use HOST aqui: "0.0.0.0" serve para escutar, não para ser aberto no navegador.)
    url = f"http://127.0.0.1:{PORT}/"
    print("=" * 66)
    print(" Baixador de XML de NF-e — servidor (os Chromes abrem nos computadores dos usuários)")
    print(f" Neste computador:  {url}")
    if ACESSO_PELA_REDE:
        for ip in meus_ips():
            print(f" Outros computadores: http://{ip}:{PORT}/")
        print(" ATENÇÃO: sem senha. Use só em rede confiável.")
    print(f" Usuários simultâneos: até {MAX_USUARIOS}")
    print(f" XMLs salvos em: {os.path.abspath(PASTA_XML)}")
    print(f" Pacote do Agente para os usuários: http://<IP>:{PORT}/agente.zip")
    print(" Para encerrar: feche esta janela (ou use o botão 'Encerrar' na página, no servidor).")
    print("=" * 66)
    threading.Thread(target=vigia, daemon=True).start()
    if os.environ.get("NFE_NO_BROWSER") != "1":
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    app.run(host=HOST, port=PORT, debug=False, threaded=True)


if __name__ == "__main__":
    main()
