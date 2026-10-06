import os
import sys
import time
import logging
import pandas as pd
from datetime import datetime
import requests
from collections import Counter
import tkinter as tk
from tkinter import simpledialog, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    InvalidSessionIdException,
    NoSuchWindowException,
)


root = tk.Tk()
root.withdraw()
numero = simpledialog.askstring("Processo", "Digite o número do processo:")

if not numero:
    print("Processo não informado. Encerrando...")
    exit()

# Procedimentos radiologicos que pagam o KIT de agulha. Todos apontam para a
# MESMA taxa - diferente do robo Oftalmo, onde cada procedimento tem a sua.
# Codigos em ASCII puro: o material de origem veio com homoglifos gregos
# (U+0391 no lugar do "A") e a comparacao de string nunca casaria.
MAPEAMENTO_PROCED_TAXA = {
    "3213004A": "70016941",  # BIOP.PERC.ORIENT.POR CT,US/RX - TIREOIDE
    "3213004B": "70016941",  # BIOP.PERC.ORIENT.POR CT,US/RX - MAMA
    "3213004C": "70016941",  # BIOP.PERC.ORIENT.POR CT,US/RX - CORE BIOPSIA
    "3213004D": "70016941",  # BIOP.PERC.ORIENT.POR CT,US/RX - MARCACAO
}

# Colunas da tabela de itens da guia (#table_itens_guia).
IDX_COL_PROCEDIMENTO = 1
IDX_COL_QTD_PROCEDIMENTO = 6

# Colunas de #taxas_table, confirmadas em 28/08/2026 pelo log de producao:
#   0=ordem  1=codigo+descricao  2=Qtd Cobrada  3=Qtd
#   4,5,6=valores  7=(-)  8=Editar/Excluir
# A primeira versao lia a coluna 6 achando que era a quantidade; e o valor total
# ("174,80"), que nao e digito - toda guia caia em "quantidade ilegivel".
IDX_COL_QTD_COB_TAXA = 2
IDX_COL_QTD_TAXA = 3

# Botao de editar da linha de taxa: a celula traz o texto literal "Editar".
# Casar pelo texto e mais confiavel do que chutar o nome da classe.
XPATH_BOTAO_EDITAR_TAXA = (
    ".//a[normalize-space()='Editar'] | .//button[normalize-space()='Editar'] "
    "| .//*[contains(@class,'editar')]"
)

# Fallback dos campos de quantidade do modal de edicao, caso a descoberta por
# sufixo (_descobrir_campos_qtd_taxa) nao ache nada.
ID_EDIT_TAXA_QT = "edit_item_taxas_QT_SERV_HOSP"
ID_EDIT_TAXA_QT_C = "edit_item_taxas_QT_SERV_HOSP_C"

# Botao de salvar, na ordem em que sao tentados. O primeiro e o mesmo que
# adicionar_taxa usa e que ja funciona em producao no modal de taxa.
SELETORES_BOTAO_SALVAR = [
    (By.XPATH, '//*[@id="modal_novo_item"]/div/div/div[3]/button[2]'),
    (By.CLASS_NAME, "btn_2_novo_item"),
]

# Bumpar a cada versao publicada: e o unico jeito de saber, olhando um log de
# usuario, qual codigo aquela maquina executou.
VERSAO = "2026-10-06 firefox-reserva"

# Entrada do SISWEB: /sisweb/ redireciona para a tela de login atual. Sem sessao,
# /sisweb/principal/index.php devolve HTTP 500 ("Pagina nao encontrada").
URL_LOGIN = "http://autoriza.issec.ce.gov.br/sisweb/"

CAMINHO_RAIZ = os.getcwd()

CAMINHO_LOG = os.path.join(CAMINHO_RAIZ, "logs")
os.makedirs(CAMINHO_LOG, exist_ok=True)

CAMINHO_RELATORIO = os.path.join(CAMINHO_RAIZ, "relatorios")
os.makedirs(CAMINHO_RELATORIO, exist_ok=True)

CAMINHO_LOG = os.path.join(
    CAMINHO_LOG, f"biopsia_{datetime.today().strftime('%Y_%m')}.log"
)

logging.basicConfig(
    filename=CAMINHO_LOG,
    level=logging.INFO,
    # O PID separa execucoes simultaneas: sem ele as linhas de dois processos se
    # misturam no mesmo arquivo e o log passa a mentir sobre o que esta rodando.
    format="%(asctime)s - [%(process)d] - %(levelname)s - %(message)s",
    encoding="utf-8"
)

# "origem" distingue de onde o codigo veio: sob o lancador aponta para o cache
# em %TEMP% (baixado do GitHub); numa copia solta, para o .py que a pessoa abriu.
_ORIGEM = globals().get("__file__", "<desconhecida>")
logging.info(f"=== Taxas Biopsia {VERSAO} | origem={_ORIGEM} ===")


def _logar_excecao_nao_tratada(tipo, valor, tb):
    """Grava no log qualquer excecao que derrube o script.

    O codigo roda quase todo no nivel do modulo e o logging so tem handler de
    arquivo, entao sem isto uma falha (ex.: tela de login fora do ar) mataria
    o processo deixando o log em branco.
    """
    logging.critical("Falha fatal nao tratada", exc_info=(tipo, valor, tb))
    sys.__excepthook__(tipo, valor, tb)


sys.excepthook = _logar_excecao_nao_tratada

relatorio_guias = []
contador_kits_inseridos = 0
contador_guias_taxadas = 0
contador_guias_corrigidas = 0

CAMINHO_ACESSO = os.path.join(CAMINHO_RAIZ, "acesso.xlsx")
if not os.path.exists(CAMINHO_ACESSO):
    pd.DataFrame([{"usuario": "SEU_USUARIO", "senha": "SUA_SENHA"}]).to_excel(
        CAMINHO_ACESSO, index=False
    )

df_usuario = pd.read_excel(CAMINHO_ACESSO, dtype=str)
usuario_sistema = df_usuario.iloc[0]["usuario"]
senha_sistema = df_usuario.iloc[0]["senha"]

# Credenciais: _secrets.py ao lado do exe (opcional) sobrepõe os valores embutidos
import importlib.util

def _carregar_secrets():
    caminho = os.path.join(CAMINHO_RAIZ, "_secrets.py")
    if not os.path.exists(caminho):
        return None
    spec = importlib.util.spec_from_file_location("_secrets", caminho)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo

_SECRETS = _carregar_secrets()

SUPABASE_URL = getattr(_SECRETS, "SUPABASE_URL", "") or "https://azjkpuulzmtyubosolck.supabase.co/rest/v1/execucoes_taxas_biopsia"
SUPABASE_APIKEY = getattr(_SECRETS, "SUPABASE_APIKEY", "") or "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImF6amtwdXVsem10eXVib3NvbGNrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ5NTE2MTksImV4cCI6MjA5MDUyNzYxOX0.vGis0gxHcuvXX6PGYaW72WnObF9sgaYRI3UejiUAc3g"

options = Options()
options.add_argument("--start-maximized")

# O SISWEB atende http e https, mas a Capa de Processo (CodeIgniter, dentro do
# iframe) referencia CSS/JS por http:// e monta os AJAX POST de
# load_content_view() sobre JS_BASE_URL = 'http://...'. Se o Chrome promover a
# pagina para https, o CSS/JS vira Mixed Content e os POST falham ("Erro, ao
# carregar permissoes!", capa vazia). Por isso:
# - a flag impede o Chrome de promover HTTP para HTTPS por conta propria;
# - se ainda assim cair em https (ex.: ISSEC forcar 301, como em 05/10/2026),
#   o Mixed Content e liberado e o script abaixo alinha JS_BASE_URL ao
#   protocolo da pagina antes dela usa-lo. Em http ele nao altera nada.
options.add_argument(
    "--disable-features=HttpsUpgrades,HttpsFirstBalancedMode,HttpsFirstModeV2"
)
options.add_argument("--allow-running-insecure-content")

SCRIPT_JS_BASE_URL_HTTPS = """
(function () { var v; try { Object.defineProperty(window, 'JS_BASE_URL', {
  get: function () { return v; },
  set: function (x) { v = (typeof x === 'string') ? x.replace(/^https?:/, location.protocol) : x; },
  configurable: false }); } catch (e) {} })();
"""


def _iniciar_chrome(max_tentativas=3):
    ultimo_erro = None
    for tentativa in range(1, max_tentativas + 1):
        try:
            logging.info(f"Iniciando Chrome - tentativa {tentativa}/{max_tentativas}")
            nav = webdriver.Chrome(options=options)
            nav.execute_cdp_cmd(
                "Page.addScriptToEvaluateOnNewDocument",
                {"source": SCRIPT_JS_BASE_URL_HTTPS},
            )
            logging.info("Chrome iniciado com sucesso")
            return nav
        except Exception as e:
            ultimo_erro = e
            logging.warning(f"Falha ao iniciar Chrome (tentativa {tentativa}): {e}")
            time.sleep(5)
    logging.error(
        f"Não foi possível iniciar o Chrome após {max_tentativas} tentativas: {ultimo_erro}"
    )
    raise ultimo_erro


def _abrir_firefox():
    # O Firefox também tem HTTPS-First; desligado para o SISWEB ficar em http.
    opcoes = FirefoxOptions()
    opcoes.set_preference("dom.security.https_only_mode", False)
    opcoes.set_preference("dom.security.https_first", False)
    opcoes.set_preference("dom.security.https_first_schemeless", False)
    nav = webdriver.Firefox(options=opcoes)
    nav.maximize_window()
    return nav


def iniciar_navegador():
    """Abre o Chrome na tela de login; se ele não subir ou cair em https (onde a Capa
    de Processo quebra por mixed content), usa o Firefox em http como reserva.
    ROBO_NAVEGADOR=firefox na máquina pula direto para o Firefox."""
    erro_chrome = None
    if os.environ.get("ROBO_NAVEGADOR", "").strip().lower() != "firefox":
        nav = None
        try:
            nav = _iniciar_chrome()
            nav.get(URL_LOGIN)
            if not nav.current_url.lower().startswith("https://"):
                return nav
            logging.warning(f"Chrome abriu em https ({nav.current_url}); usando o Firefox em http")
        except Exception as e:
            erro_chrome = e
            logging.warning(f"Chrome falhou ({e}); usando o Firefox em http")
        if nav is not None:
            try:
                nav.quit()
            except Exception:
                pass

    try:
        nav = _abrir_firefox()
        nav.get(URL_LOGIN)
        logging.info(f"Navegador: Firefox ({nav.current_url})")
        return nav
    except Exception:
        logging.exception("Firefox também falhou")
        if erro_chrome is not None:
            raise erro_chrome
        raise


navegador = iniciar_navegador()
wait = WebDriverWait(navegador, 30)


def _estado_navegador():
    """Descreve url/titulo sem nunca lancar.

    So e chamado de dentro de handlers de erro. Se a sessao do Chrome morreu,
    tocar o driver lanca de novo e mascara a excecao original - foi assim que
    um "invalid session id" impediu FalhaLeituraGuias de ser levantada.
    """
    try:
        return f"url={navegador.current_url} titulo={navegador.title!r}"
    except Exception:
        return "url=<sessao indisponivel> titulo=<sessao indisponivel>"


def fazer_login():
    logging.info(f"Abrindo tela de login: {URL_LOGIN}")
    navegador.get(URL_LOGIN)

    try:
        campo_login = wait.until(EC.presence_of_element_located((By.ID, "login")))
    except TimeoutException:
        # Sem isto, o find_element direto estoura no nivel do modulo e o
        # script morre sem deixar nada no arquivo de log.
        logging.exception(f"Campo de login nao encontrado - {_estado_navegador()}")
        messagebox.showerror(
            "Taxas Biópsia",
            "Nao foi possivel abrir a tela de login do SISWEB.\n\n"
            f"Endereco: {URL_LOGIN}\n"
            "O sistema pode estar fora do ar ou ter mudado de endereco.\n"
            "Avise o suporte e envie a pasta logs."
        )
        raise

    campo_login.send_keys(usuario_sistema)
    navegador.find_element(By.ID, "senha").send_keys(senha_sistema, Keys.ENTER)
    logging.info("Login enviado")


fazer_login()


def aguardar_loading_fade(tempo_maximo=30):
    try:
        WebDriverWait(navegador, tempo_maximo).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, "loading_fade"))
        )
    except TimeoutException:
        print("⚠️ Timeout aguardando loader - continuando")
    except:
        pass


def click_seguro(localizador):
    aguardar_loading_fade()
    ultimo_erro = None
    for _ in range(3):
        try:
            el = wait.until(EC.element_to_be_clickable(localizador))
            navegador.execute_script("arguments[0].scrollIntoView(true);", el)
            el.click()
            return True
        except Exception as e:
            ultimo_erro = e
            time.sleep(0.5)

    # Sem este log o robo falha calado e segue como se tivesse clicado - a tela
    # nao avanca e o log nao diz nada.
    logging.warning(f"Nao foi possivel clicar em {localizador}: {ultimo_erro}")
    return False


def abrir_aba_guia():
    click_seguro((By.XPATH, '//*[@id="guia_painel"]/div[3]/ul/li[1]/a'))


def abrir_aba_taxas():
    click_seguro((By.XPATH, '//*[@id="guia_painel"]/div[3]/ul/li[3]/a'))


def fechar_popups_ok():
    """Fecha modais de consistencia/confirmacao que o SISWEB abre ao editar.

    Alterar quantidade dispara validacoes do sistema; se o modal ficar aberto,
    o clique seguinte cai no overlay e o robo trava sem erro visivel.
    """
    seletores_ok = [
        (By.ID, "ok_botao"),
        (By.XPATH, "//*[@id='modal_consistencia_msg']//button"),
        (By.XPATH, "//*[@id='modal_msg']//button"),
        (By.XPATH, "//button[contains(text(), 'OK') or contains(text(), 'Ok')]"),
    ]

    for seletor in seletores_ok:
        try:
            botao = WebDriverWait(navegador, 2).until(
                EC.element_to_be_clickable(seletor)
            )
            navegador.execute_script("arguments[0].click();", botao)
            time.sleep(0.5)
            print("✅ Popup/consistência fechado com OK")
            logging.info("Popup/consistência fechado com OK")
        except:
            pass


def preencher_campo(id_campo, valor):
    aguardar_loading_fade()
    campo = wait.until(EC.visibility_of_element_located((By.ID, id_campo)))
    campo.clear()
    campo.send_keys(str(valor))
    campo.send_keys(Keys.TAB)
    campo.send_keys(Keys.ENTER)
    navegador.execute_script("arguments[0].blur();", campo)
    return campo


def obter_itens_biopsia_guia():
    """Itens de biopsia da guia atual, com a quantidade de cada linha.

    Diferente do robo Oftalmo, que le so o texto do procedimento: aqui a
    quantidade do KIT e a SOMA das quantidades, entao a coluna Qtd faz parte da
    regra de negocio e nao pode ser ignorada.
    """
    wait.until(
        EC.invisibility_of_element_located((By.ID, "table_itens_guia_processing"))
    )

    itens = []
    linhas = navegador.find_elements(By.XPATH, '//*[@id="table_itens_guia"]/tbody/tr')

    if len(linhas) == 1 and "Nenhum registro encontrado" in linhas[0].text:
        return itens

    for linha in linhas:
        cols = linha.find_elements(By.TAG_NAME, "td")

        if len(cols) <= IDX_COL_QTD_PROCEDIMENTO:
            continue

        procedimento = cols[IDX_COL_PROCEDIMENTO].text.strip().upper()
        codigo = procedimento.split("-")[0].strip()

        if codigo not in MAPEAMENTO_PROCED_TAXA:
            continue

        qtd_txt = cols[IDX_COL_QTD_PROCEDIMENTO].text.strip()

        if qtd_txt.isdigit() and int(qtd_txt) > 0:
            qtd = int(qtd_txt)
        else:
            # Assumir 1 e o menor erro possivel: pular a linha perderia um KIT
            # devido, e inventar um numero maior cobraria a mais.
            qtd = 1
            logging.warning(
                f"Quantidade ilegivel ('{qtd_txt}') em {codigo} - assumindo 1"
            )

        itens.append({"codigo": codigo, "procedimento": procedimento, "qtd": qtd})

    return itens


def obter_taxa_existente(codigo_taxa):
    """Localiza a taxa na aba Taxas da guia atual.

    Devolve {"qtd": int|None, "qtd_cob": int|None, "cols": [...]} se existir,
    None se nao existir. "qtd" None significa que a linha existe mas a quantidade
    nao pode ser lida - nesse caso o chamador nao deve alterar nada.

    Levanta excecao em falha de leitura: quem chama decide. O loop trata como
    indeterminado e nao mexe na guia, igual ao comportamento conservador do
    robo Oftalmo.
    """
    abrir_aba_taxas()
    time.sleep(1)

    wait.until(EC.presence_of_element_located((By.ID, "taxas_table")))

    print("[TAXA] ⏳ Aguardando carregamento da tabela...")

    try:
        wait.until(
            EC.invisibility_of_element_located((By.ID, "taxas_table_processing"))
        )
    except:
        print("[TAXA] ⚠️ Timeout esperando 'Processando...' sumir")

    wait.until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='taxas_table']/tbody/tr"))
    )

    linhas = navegador.find_elements(By.XPATH, "//*[@id='taxas_table']/tbody/tr")

    print(f"[TAXA] 🔎 Verificando {codigo_taxa}")

    for linha in linhas:
        texto = linha.text.strip()

        if not texto or "Nenhum registro encontrado" in texto:
            continue

        if codigo_taxa not in texto:
            continue

        cols = [c.text.strip() for c in linha.find_elements(By.TAG_NAME, "td")]

        # Loga a linha inteira: e por aqui que se confere/ajusta IDX_COL_QTD_TAXA
        # sem precisar inspecionar a tela manualmente.
        logging.info(f"Taxa {codigo_taxa} encontrada - colunas={cols}")

        def _ler_coluna(indice):
            if len(cols) > indice and cols[indice].isdigit():
                return int(cols[indice])
            logging.warning(
                f"Quantidade da taxa {codigo_taxa} ilegivel na coluna "
                f"{indice} - colunas={cols}"
            )
            return None

        qtd = _ler_coluna(IDX_COL_QTD_TAXA)
        qtd_cob = _ler_coluna(IDX_COL_QTD_COB_TAXA)

        print(f"[TAXA] ⚠️ Já existe: {codigo_taxa} (qtd={qtd} / cobrada={qtd_cob})")
        return {"qtd": qtd, "qtd_cob": qtd_cob, "cols": cols}

    print(f"[TAXA] ➕ Não encontrada: {codigo_taxa}")
    return None


def adicionar_taxa(codigo_taxa, quantidade):
    abrir_aba_taxas()
    time.sleep(1)

    print(f"[TAXA] ➕ Inserindo {codigo_taxa} (qtd {quantidade})")
    click_seguro((By.XPATH, '//*[@id="guia_novo_item_taxas"]'))
    time.sleep(1)

    click_seguro((By.ID, "novo_item_taxas_CD_SERV_HOSP_chosen"))
    time.sleep(1)

    campo = wait.until(
        EC.visibility_of_element_located(
            (By.XPATH, '//*[@id="novo_item_taxas_CD_SERV_HOSP_chosen"]//input')
        )
    )

    navegador.execute_script("arguments[0].value = '';", campo)
    campo.send_keys(codigo_taxa)
    time.sleep(1)
    campo.send_keys(Keys.ENTER)

    qt_cob = navegador.find_element(By.ID, "novo_item_taxas_QT_SERV_HOSP_C")
    qt_cob.clear()
    qt_cob.send_keys(str(quantidade))

    qt = navegador.find_element(By.ID, "novo_item_taxas_QT_SERV_HOSP")
    qt.clear()
    qt.send_keys(str(quantidade), Keys.ENTER)

    print("[TAXA] → Abrindo modal")
    click_seguro((By.XPATH, '//*[@id="modal_novo_item"]/div/div/div[3]/button[2]'))
    time.sleep(1)

    fechar_popups_ok()
    print(f"[TAXA] ✅ Inserida: {codigo_taxa} (qtd {quantidade})")
    logging.info(f"Taxa {codigo_taxa} inserida com quantidade {quantidade}")


def _localizar_linha_taxa(codigo_taxa):
    abrir_aba_taxas()
    try:
        wait.until(
            EC.invisibility_of_element_located((By.ID, "taxas_table_processing"))
        )
    except:
        pass

    for linha in navegador.find_elements(By.XPATH, "//*[@id='taxas_table']/tbody/tr"):
        if codigo_taxa in linha.text:
            return linha
    return None


_campos_qtd_logados = False


def _descobrir_campos_qtd_taxa():
    """Ids dos campos Qtd/Qtd Cobrada do modal de taxa que esta aberto.

    O SISWEB pode nomea-los 'novo_item_taxas_*' ou 'edit_item_taxas_*' - o modal
    de taxa e o mesmo 'modal_novo_item' nos dois casos. Ambos terminam em
    QT_SERV_HOSP / QT_SERV_HOSP_C, entao casar pelo sufixo dispensa adivinhar.
    O filtro por offsetParent descarta o que estiver no DOM porem oculto.
    """
    global _campos_qtd_logados

    ids = navegador.execute_script(
        "return Array.from(document.querySelectorAll('input[id]'))"
        ".filter(e => e.offsetParent).map(e => e.id);"
    ) or []

    # endswith("QT_SERV_HOSP") nao casa com "..._C", entao os dois saem distintos.
    id_qt = next((i for i in ids if i.endswith("QT_SERV_HOSP")), ID_EDIT_TAXA_QT)
    id_qt_c = next((i for i in ids if i.endswith("QT_SERV_HOSP_C")), ID_EDIT_TAXA_QT_C)

    if not _campos_qtd_logados:
        logging.info(
            f"Campos de quantidade do modal de taxa: qtd='{id_qt}' "
            f"cobrada='{id_qt_c}' (inputs visiveis: {ids})"
        )
        _campos_qtd_logados = True

    return id_qt, id_qt_c


def click_seguro_qualquer(seletores):
    """Clica no primeiro seletor da lista que responder.

    O botao de salvar do modal de taxa nao esta confirmado: tenta o XPath que
    adicionar_taxa ja usa em producao e, se nao houver, a classe do modal de
    item de procedimento.
    """
    for seletor in seletores:
        try:
            el = WebDriverWait(navegador, 5).until(
                EC.element_to_be_clickable(seletor)
            )
            navegador.execute_script("arguments[0].click();", el)
            return True
        except Exception:
            continue

    logging.warning(f"Nenhum dos seletores respondeu: {seletores}")
    return False


def corrigir_qtd_taxa(codigo_taxa, quantidade):
    """Ajusta a quantidade de uma taxa ja lancada na guia atual.

    Capacidade que o robo Oftalmo nao tem: la cada procedimento tem taxa propria
    e quantidade fixa 1, entao "ja existe" sempre significava "nada a fazer".
    """
    linha = _localizar_linha_taxa(codigo_taxa)

    if linha is None:
        logging.warning(f"Linha da taxa {codigo_taxa} sumiu antes da edicao")
        return False

    try:
        botao_editar = linha.find_element(By.XPATH, XPATH_BOTAO_EDITAR_TAXA)
    except Exception:
        # Logar o HTML da linha e o unico jeito de descobrir o seletor real sem
        # inspecionar a tela na mao - foi assim que as colunas foram resolvidas.
        try:
            html = linha.get_attribute("outerHTML")
        except Exception:
            html = "<indisponivel>"
        logging.error(
            f"Botao de editar da taxa {codigo_taxa} nao encontrado. "
            f"HTML da linha: {html}"
        )
        return False

    aguardar_loading_fade()
    navegador.execute_script("arguments[0].click();", botao_editar)
    time.sleep(1)

    id_qt, id_qt_c = _descobrir_campos_qtd_taxa()

    try:
        preencher_campo(id_qt_c, quantidade)
        fechar_popups_ok()

        preencher_campo(id_qt, quantidade)
        fechar_popups_ok()
    except TimeoutException:
        logging.error(
            f"Campos de quantidade da taxa nao encontrados ('{id_qt_c}' / '{id_qt}')"
        )
        return False

    # Validacao antes de salvar: o SISWEB reverte campo em silencio quando a
    # consistencia reclama, e salvar assim gravaria a quantidade errada.
    for tentativa in range(1, 4):
        val_qt = navegador.find_element(By.ID, id_qt).get_attribute("value").strip()
        val_qt_c = navegador.find_element(By.ID, id_qt_c).get_attribute("value").strip()

        print(f"[TAXA] 🔎 Validação {tentativa}: Qt={val_qt} | QtCob={val_qt_c}")

        if val_qt == str(quantidade) and val_qt_c == str(quantidade):
            break

        print("[TAXA] ⚠️ Corrigindo campos...")
        preencher_campo(id_qt_c, quantidade)
        preencher_campo(id_qt, quantidade)
        time.sleep(1)
    else:
        logging.warning(
            f"Taxa {codigo_taxa}: campos nao aceitaram a quantidade "
            f"{quantidade} - nao salvando"
        )
        return False

    fechar_popups_ok()

    print("[TAXA] 💾 Salvando...")
    aguardar_loading_fade()

    if not click_seguro_qualquer(SELETORES_BOTAO_SALVAR):
        logging.error(f"Taxa {codigo_taxa}: nenhum botao de salvar respondeu")
        return False

    time.sleep(2)

    fechar_popups_ok()

    # Revalida na tabela: so assim se sabe que o SISWEB persistiu a alteracao.
    existente = obter_taxa_existente(codigo_taxa)

    if existente and existente["qtd"] == existente["qtd_cob"] == quantidade:
        logging.info(f"Taxa {codigo_taxa} corrigida para quantidade {quantidade}")
        return True

    logging.warning(
        f"Taxa {codigo_taxa}: apos salvar a tabela mostra {existente} "
        f"(esperado {quantidade})"
    )
    return False


class FalhaLeituraGuias(Exception):
    """Contador de guias ilegivel - navegacao quebrada, nao da para continuar."""


def obter_totais_guias(max_tentativas=3):
    # Devolver (1, 1) em silencio faria o loop encerrar achando que ha uma unica
    # guia e o Supabase registraria "1 guias / 0 taxas" como sucesso.
    # As tentativas cobrem o contador ainda em carregamento; esgotadas elas, a
    # pagina esta realmente quebrada e insistir nao ajuda.
    ultimo_erro = None
    for tentativa in range(1, max_tentativas + 1):
        try:
            texto = navegador.find_element(By.ID, "map_holder").text
            partes = texto.replace("Total:", "").replace("de", "").split()
            return int(partes[0]), int(partes[1])
        except Exception as e:
            ultimo_erro = e
            logging.warning(
                f"Falha ao ler total de guias (tentativa {tentativa}/{max_tentativas}): {e}"
            )
            aguardar_loading_fade()
            time.sleep(1)

    raise FalhaLeituraGuias(
        f"Nao foi possivel ler o total de guias em #map_holder apos "
        f"{max_tentativas} tentativas ({_estado_navegador()})"
    ) from ultimo_erro


# Deste ponto ate a primeira guia o log e o unico rastro: sem ele, uma falha
# aqui deixaria o arquivo parado em "Login enviado" sem dizer onde parou.
logging.info("Abrindo menu Contas Medicas")
click_seguro((By.XPATH, "//a[contains(., 'Contas Médicas')]"))

logging.info("Abrindo Capa De Processo")
click_seguro((By.XPATH, "//a[contains(., 'Capa De Processo')]"))

logging.info("Aguardando iframe da Capa de Processo")
iframe = wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
navegador.switch_to.frame(iframe)
logging.info(f"Dentro do iframe - {_estado_navegador()}")

campo_num = wait.until(EC.element_to_be_clickable((By.ID, "numero_processo")))
campo_num.clear()
campo_num.send_keys(numero, Keys.ENTER)
logging.info(f"Processo {numero} pesquisado")

click_seguro((By.ID, "editar_guias"))
logging.info("Guias abertas - iniciando processamento")


guias_erro_count = {}
# Inicializados aqui porque o handler de erro do loop referencia ambos: se a
# excecao ocorresse antes da linha que os define, o proprio handler quebraria.
atual, total = "?", 0

while True:
    try:
        abrir_aba_guia()

        atual, total = obter_totais_guias()

        print(f"\n📄 ===== GUIA {atual}/{total} =====")
        logging.info(f"[GUIA {atual}] Início processamento")

        itens = obter_itens_biopsia_guia()

        if not itens:
            print(f"[GUIA {atual}] ℹ️ Nenhum procedimento de biópsia")
            relatorio_guias.append({
                "guia": atual,
                "status": "Nenhuma necessária",
                "procedimentos": "",
                "qtd_procedimentos": 0,
                "codigo_taxa": "",
                "qtd_taxa": "",
            })
        else:
            # A quantidade do KIT e a SOMA das quantidades - uma linha com Qtd 2
            # vale 2 KITs, nao 1.
            qtd_esperada = sum(item["qtd"] for item in itens)
            codigo_taxa = MAPEAMENTO_PROCED_TAXA[itens[0]["codigo"]]
            descricao = " | ".join(item["procedimento"] for item in itens)

            print(
                f"[GUIA {atual}] 🎯 {len(itens)} item(ns) → "
                f"taxa {codigo_taxa} qtd {qtd_esperada}"
            )
            logging.info(
                f"[GUIA {atual}] {len(itens)} item(ns) de biopsia, "
                f"qtd esperada do KIT = {qtd_esperada}"
            )

            try:
                existente = obter_taxa_existente(codigo_taxa)
            except Exception as e:
                # Conservador de proposito: sem saber o estado da guia, alterar
                # e pior do que nao alterar.
                print(f"[GUIA {atual}] ❌ Erro ao verificar taxa {codigo_taxa}")
                logging.warning(f"Falha ao verificar taxa {codigo_taxa}: {e}")
                relatorio_guias.append({
                    "guia": atual,
                    "status": "Verificação falhou - não alterada",
                    "procedimentos": descricao,
                    "qtd_procedimentos": qtd_esperada,
                    "codigo_taxa": codigo_taxa,
                    "qtd_taxa": "",
                })
                existente = "erro"

            if existente == "erro":
                pass

            elif existente is None:
                adicionar_taxa(codigo_taxa, qtd_esperada)
                contador_kits_inseridos += qtd_esperada
                contador_guias_taxadas += 1
                relatorio_guias.append({
                    "guia": atual,
                    "status": "Inserida",
                    "procedimentos": descricao,
                    "qtd_procedimentos": qtd_esperada,
                    "codigo_taxa": codigo_taxa,
                    "qtd_taxa": qtd_esperada,
                })

            elif existente["qtd"] is None or existente["qtd_cob"] is None:
                relatorio_guias.append({
                    "guia": atual,
                    "status": "Já existente - quantidade ilegível",
                    "procedimentos": descricao,
                    "qtd_procedimentos": qtd_esperada,
                    "codigo_taxa": codigo_taxa,
                    "qtd_taxa": "",
                })

            # Exige que Qtd e Qtd Cobrada batam com o esperado: uma taxa lancada
            # a mao com 2/1 esta errada mesmo que a Qtd sozinha pareca certa.
            elif existente["qtd"] == existente["qtd_cob"] == qtd_esperada:
                print(f"[GUIA {atual}] ✅ Taxa já correta (qtd {qtd_esperada})")
                relatorio_guias.append({
                    "guia": atual,
                    "status": "Já existente (qtd correta)",
                    "procedimentos": descricao,
                    "qtd_procedimentos": qtd_esperada,
                    "codigo_taxa": codigo_taxa,
                    "qtd_taxa": existente["qtd"],
                })

            else:
                qtd_antiga = f"{existente['qtd']}/{existente['qtd_cob']}"
                print(f"[GUIA {atual}] ✏️ Corrigindo qtd {qtd_antiga} → {qtd_esperada}")

                if corrigir_qtd_taxa(codigo_taxa, qtd_esperada):
                    contador_guias_corrigidas += 1
                    relatorio_guias.append({
                        "guia": atual,
                        "status": f"Quantidade corrigida {qtd_antiga} → {qtd_esperada}",
                        "procedimentos": descricao,
                        "qtd_procedimentos": qtd_esperada,
                        "codigo_taxa": codigo_taxa,
                        "qtd_taxa": qtd_esperada,
                    })
                else:
                    relatorio_guias.append({
                        "guia": atual,
                        "status": (
                            f"Falha ao corrigir quantidade "
                            f"({qtd_antiga} → {qtd_esperada})"
                        ),
                        "procedimentos": descricao,
                        "qtd_procedimentos": qtd_esperada,
                        "codigo_taxa": codigo_taxa,
                        "qtd_taxa": qtd_antiga,
                    })

    except (InvalidSessionIdException, NoSuchWindowException):
        # Navegador fechado/derrubado: toda iteracao seguinte repetiria o mesmo
        # ciclo sem chance de sucesso.
        print("❌ Navegador foi fechado — abortando")
        logging.exception("Sessao do Chrome perdida - abortando execucao")
        raise

    except FalhaLeituraGuias:
        # Fatal: insistir so geraria loop infinito sobre uma pagina quebrada.
        print("❌ Nao foi possivel ler a lista de guias — abortando")
        logging.exception("Falha ao ler total de guias - abortando execucao")
        raise

    except Exception as e:
        print(f"[GUIA {atual}] ❌ ERRO NÃO TRATADO: {e}")
        logging.error(f"[GUIA {atual}] Erro geral: {e}", exc_info=True)

        for seletor in [
            (By.XPATH, '//*[@id="modal_novo_item"]/div/div/div[3]/button[1]'),
            (By.ID, "ok_botao"),
            (By.XPATH, "//button[contains(text(), 'OK') or contains(text(), 'Ok')]"),
        ]:
            try:
                el = WebDriverWait(navegador, 2).until(
                    EC.element_to_be_clickable(seletor)
                )
                navegador.execute_script("arguments[0].click();", el)
                time.sleep(0.5)
            except:
                pass

        aguardar_loading_fade()
        click_seguro((By.ID, "next"))
        time.sleep(1)

        guias_erro_count[atual] = guias_erro_count.get(atual, 0) + 1
        if guias_erro_count[atual] >= 5:
            print(
                f"[GUIA {atual}] ⛔ Pulando após "
                f"{guias_erro_count[atual]} erros consecutivos"
            )
            logging.warning(
                f"Guia {atual} pulada apos {guias_erro_count[atual]} erros consecutivos"
            )

        continue

    if atual >= total:
        break

    click_seguro((By.ID, "next"))
    time.sleep(1)

caminho_planilha = os.path.join(
    CAMINHO_RELATORIO,
    f"Relatorio_Taxas_Biopsia_Processo_{numero}.xlsx"
)

df_relatorio = pd.DataFrame(relatorio_guias)
df_relatorio = df_relatorio[
    ["guia", "status", "procedimentos", "qtd_procedimentos", "codigo_taxa", "qtd_taxa"]
]
df_relatorio.to_excel(caminho_planilha, index=False)

navegador.quit()


def registrar_supabase(numero_processo, qtd_guias, qtd_taxas, qtd_guias_taxadas,
                       qtd_corrigidas, codigos_taxas):
    payload = {
        "numero_processo": numero_processo,
        "qtd_guias": qtd_guias,
        "qtd_taxas": qtd_taxas,
        "qtd_guias_taxadas": qtd_guias_taxadas,
        "qtd_corrigidas": qtd_corrigidas,
        "codigos_taxas": codigos_taxas,
    }
    headers = {
        "apikey": SUPABASE_APIKEY,
        "Content-Type": "application/json"
    }
    try:
        r = requests.post(SUPABASE_URL, json=payload, headers=headers, timeout=30)
    except requests.RequestException:
        # O trabalho no ISSEC ja foi concluido e a planilha ja esta salva:
        # uma falha de rede aqui nao pode derrubar o script no ultimo passo.
        logging.exception("Falha de rede ao registrar no Supabase")
        return False

    if r.status_code == 201:
        logging.info(
            f"Registrado no Supabase: {numero_processo} / "
            f"{qtd_guias} guias / {qtd_taxas} kits"
        )
    else:
        logging.warning(f"Erro ao registrar no Supabase: {r.status_code} - {r.text}")
    return r.status_code == 201


# Conta KITs por codigo de taxa. Soma a quantidade (nao o numero de guias), para
# que o total bata com contador_kits_inseridos. O formato de dicionario
# acompanha o do robo Oftalmo, para os dashboards lerem os dois do mesmo jeito.
contagem_taxas = Counter()
for r in relatorio_guias:
    if r.get("status") == "Inserida":
        contagem_taxas[r["codigo_taxa"]] += r["qtd_taxa"]
contagem_taxas = dict(contagem_taxas)

registrar_supabase(
    numero,
    total,
    contador_kits_inseridos,
    contador_guias_taxadas,
    contador_guias_corrigidas,
    contagem_taxas,
)

janela = tk.Tk()
janela.withdraw()

messagebox.showinfo(
    "Automação Finalizada",
    f"Processo: {numero}\n"
    f"Guias com taxa inserida: {contador_guias_taxadas}\n"
    f"Total de KITs lançados: {contador_kits_inseridos}\n"
    f"Quantidades corrigidas: {contador_guias_corrigidas}\n\n"
    f"Arquivo gerado:\n{caminho_planilha}"
)

janela.destroy()
