# === LANCADOR (auto-atualizacao) ===
# O PyInstaller so analisa ESTE arquivo, NAO o .py baixado em runtime.
# Por isso, espelhamos aqui TODOS os imports de terceiros da automacao
# (com submodulos) para forcar o empacotamento das dependencias.

# --- imports espelhados de issec_taxa_biopsia.py (terceiros) ---
# Nenhum destes e usado aqui de proposito: existem so para o PyInstaller enxergar
# as dependencias. O "noqa: F401" evita que um linter os remova por engano.
import pandas  # noqa: F401  leitura/escrita de .xlsx (read_excel / to_excel)
import openpyxl  # noqa: F401  engine implicita do pandas para .xlsx
import requests  # noqa: F401  POST no Supabase

import tkinter  # noqa: F401  stdlib, mas precisa do import p/ o PyInstaller incluir o Tk
from tkinter import simpledialog, messagebox  # noqa: F401

# selenium: submodulos dinamicos + binario do Selenium Manager
# (espelhar a raiz NAO arrasta os submodulos; precisa listar cada um)
from selenium import webdriver  # noqa: F401
from selenium.webdriver.chrome.options import Options  # noqa: F401
from selenium.webdriver.common.by import By  # noqa: F401
from selenium.webdriver.common.keys import Keys  # noqa: F401
from selenium.webdriver.support.ui import WebDriverWait  # noqa: F401
from selenium.webdriver.support import expected_conditions as EC  # noqa: F401
from selenium.common.exceptions import (  # noqa: F401
    StaleElementReferenceException,
    TimeoutException,
)

# --- nucleo do lancador (stdlib) ---
import sys, os, urllib.request, runpy, tempfile, configparser
from datetime import datetime

# Bumpar a cada .exe distribuido: e o unico jeito de saber qual build uma
# maquina tem sem inspecionar o binario.
VERSAO_LANCADOR = "2026-08-31 repo robo-taxas-biopsia"

# Ancora o cwd na pasta do .exe (onde estao config.ini / acesso.xlsx / logs),
# nao no diretorio temporario de onde o script baixado roda.
if getattr(sys, "frozen", False):
    os.chdir(os.path.dirname(sys.executable))
else:
    os.chdir(os.path.dirname(os.path.abspath(__file__)))


def registrar(msg):
    """Grava em logs/lancador.log ao lado do .exe.

    O console fecha junto com o robo, entao um print nao sobrevive. Foi
    justamente por isso que maquinas ficaram meses congeladas numa versao
    antiga sem ninguem perceber: a falha de download nao deixava rastro.
    """
    print(f"[lancador] {msg}")
    try:
        os.makedirs("logs", exist_ok=True)
        carimbo = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(os.path.join("logs", "lancador.log"), "a", encoding="utf-8") as f:
            f.write(f"{carimbo} - {msg}\n")
    except Exception:
        # Log e diagnostico: nunca pode impedir o robo de rodar.
        pass


registrar(f"=== Lancador Taxas Biopsia {VERSAO_LANCADOR} ===")

# TLS corporativo: usa o cofre de certificados do Windows (CA do orgao instalada
# via GPO). Sem isso, redes que interceptam HTTPS quebram a verificacao e o
# download falha sempre - a maquina congela na versao que estiver em cache.
try:
    import truststore
    truststore.inject_into_ssl()
except Exception as e:
    registrar(f"truststore indisponivel ({e}) - verificacao TLS padrao")

OWNER = "tcidrack"
REPO = "robo-taxas-biopsia"
ARQ = "issec_taxa_biopsia.py"

# Token NAO embutido no .exe: vem do config.ini ao lado do executavel.
# Assim, ao expirar, troca-se so o arquivo (sem rebuildar).
_cfg = configparser.ConfigParser()
_cfg.read(os.path.join(os.getcwd(), "config.ini"), encoding="utf-8")
TOKEN = _cfg.get("github", "token", fallback="").strip()

URL = f"https://api.github.com/repos/{OWNER}/{REPO}/contents/{ARQ}"
CACHE = os.path.join(tempfile.gettempdir(), ARQ)

baixou = False
motivo_falha = ""

if TOKEN:
    req = urllib.request.Request(URL, headers={
        "Authorization": f"Bearer {TOKEN}",
        "Accept": "application/vnd.github.raw",
        "User-Agent": "lancador",
    })
    try:
        # Verificacao TLS via truststore (cofre de certificados do Windows). Se
        # falhar, NAO baixa por canal inseguro: propaga o erro e cai no cache.
        dados = urllib.request.urlopen(req, timeout=10).read()
        with open(CACHE, "wb") as f:
            f.write(dados)
        baixou = True
        registrar(f"versao mais recente baixada do GitHub ({len(dados)} bytes)")
    except Exception as e:
        motivo_falha = str(e)
        registrar(f"FALHA no download ({e}) - tentando cache local")
else:
    motivo_falha = "token do GitHub ausente no config.ini"
    registrar(f"{motivo_falha} - tentando cache local")

if not os.path.exists(CACHE):
    registrar("sem download e sem cache - encerrando")
    messagebox.showerror(
        "Lancador Taxas Biopsia",
        "Nao foi possivel baixar o robo e nao ha copia local.\n\n"
        f"Motivo: {motivo_falha}\n\n"
        "Verifique o token no config.ini e a conexao, depois tente de novo."
    )
    sys.exit(1)

if not baixou:
    # Avisa que vai rodar codigo possivelmente desatualizado, em vez de deixar
    # o usuario acreditando que esta na versao mais recente.
    idade = datetime.fromtimestamp(os.path.getmtime(CACHE)).strftime("%d/%m/%Y %H:%M")
    registrar(f"executando cache local de {idade}")
    messagebox.showwarning(
        "Lancador Taxas Biopsia",
        "Nao foi possivel buscar a versao mais recente no GitHub.\n\n"
        f"Motivo: {motivo_falha}\n"
        f"Sera usada a copia local de {idade}, que pode estar desatualizada.\n\n"
        "Se o robo apresentar problemas ja corrigidos, avise o suporte."
    )

runpy.run_path(CACHE, run_name="__main__")
