#ONDE PAREI
#PRECISO AJUSTAR NA PARTE EM QUE FECHA A PLANILHA (ALGUM ERRO RELACIONADO AO SHAREPOINT) E DEPOIS MANDAR O EMAIL COM ELA
import subprocess
import sys
import tempfile
import os
import psutil
import datetime
import pandas as pd
from pywinauto import Desktop
import win32com.client as win32
import pythoncom
import win32com.client
import time, threading, unicodedata, win32con, win32api, win32gui, win32process
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# Variáveis de usuário e senha do SAP (lidas de arquivo .env na mesma pasta do .py/.exe)
def _base_dir():
    # Quando empacotado com PyInstaller, o .exe fica em dist/; o .env deve ficar ao lado dele.
    # Quando rodando como .py, o .env fica ao lado do script.
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _load_dotenv(path):
    """Carrega KEY=VALUE de um .env simples (sem dependências externas)."""
    if not os.path.exists(path):
        return
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key, value = key.strip(), value.strip()
            # Remove aspas simples/duplas ao redor do valor, se houver
            if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
                value = value[1:-1]
            os.environ.setdefault(key.strip(), value)


_load_dotenv(os.path.join(_base_dir(), ".env"))


def _require_env(name):
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(
            f"Variável '{name}' não encontrada. Crie um arquivo '.env' na mesma pasta "
            f"do script/executável ({_base_dir()}) com:\n{name}=seu_valor"
        )
    return value


sap_user = _require_env("SAP_USER")
sap_password = _require_env("SAP_PASSWORD")
# Mandante (client) e opcional; usado para preencher a tela de logon
# quando a conexao nao traz o mandante pre-definido (ex.: SAP_CLIENT=300)
sap_client = os.environ.get("SAP_CLIENT", "").strip()

TARGET_YEAR_TEXT = "2026"

# Pasta de saida dos .txt extraidos do SAP (o SAP NAO cria a pasta - o script cria)
OUTPUT_DIR = os.environ.get("SAP_OUTPUT_DIR", r"C:\Automat").strip()

# Diagnostico: grava <OUTPUT_DIR>\_tree_<segmento>.txt com TODOS os nos lidos
# da arvore "Variacao: Exercicio" (key + texto). Use DUMP_TREE=1 no .env para ligar.
DUMP_TREE = os.environ.get("DUMP_TREE", "").strip().lower() in ("1", "true", "yes", "sim", "on")

# Piloto: limita quantos segmentos rodar (vazio/0 = todos os 25).
_seg_limit = os.environ.get("SEGMENT_LIMIT", "").strip()
SEGMENT_LIMIT = int(_seg_limit) if _seg_limit.isdigit() else 0

segments_to_extract = [
    ("77699AAA", "77699ZZZ"),
    ("77698AAA", "77698ZZZ"),
    ("78220AAA", "78220ZZZ"),
    ("77855AAA", "77855ZZZ"),
    ("65745AAA", "65745ZZZ"),
    ("73834AAA", "73834ZZZ"),
    ("76541AAA", "76541ZZZ"),
    ("74085AAA", "74085ZZZ"),
    ("78484AAA", "78484ZZZ"),
    ("79683AAA", "79683ZZZ"),
    ("79535AAA", "79535ZZZ"),
    ("79982AAA", "79982ZZZ"),
    ("79819AAA", "79819ZZZ"),
    ("71575AAA", "71575ZZZ"),
    ("76710AAA", "76710ZZZ"),
    ("77682AAA", "77682ZZZ"),
    ("77194AAA", "77194ZZZ"),
    ("77696AAA", "77696ZZZ"),
    ("78586AAA", "78586ZZZ"),
    ("78587AAA", "78587ZZZ"),
    ("100386AAA", "100386ZZZ"),
    ("77605AAA", "77605ZZZ"),
    ("67972AAA", "67972ZZZ"),
    ("79501AAA", "79501ZZZ"),
]


# ---------------------------------------------------------------------------
# CONEXAO COM O SAP
#
# O metodo antigo (application.OpenConnection via pywin32) falha nesta maquina
# com o erro 605 "The 'Sapgui Component' could not be instantiated." (o
# saplogon.exe nao consegue criar o processo do componente grafico).
# Por isso a obtencao da sessao agora acontece em cascata, das formas que
# funcionam neste ambiente, sem depender exclusivamente do OpenConnection:
#   1) reaproveitar a sessao que ja estiver aberta (nao usa OpenConnection);
#   2) OpenConnection via VBScript/cscript (late binding nativo do SAP GUI);
#   3) abrir a conexao com sapshcut.exe (mesmo caminho do atalho do desktop,
#      que nao passa por COM) e apenas anexar a sessao criada.
# ---------------------------------------------------------------------------

SAP_CONNECTION_NAME = "001 - SAP S/4 HANA PRODUÇÃO CLARO - CLIQUE AQUI"  # nome exato no SAP Logon (SAPUILandscape.xml)
SAP_CONNECTION_STRING = "/H/172.25.73.15/S/3200"  # do SAPUILandscape (server=172.25.73.15:3200) - sem acentos
SAP_SYSTEM_ID = "S4P"                             # systemid do SAPUILandscape
SAP_LANGUAGE = "PT"
SAP_LOGON_PATHS = [
    r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe",
    r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",
]
SAPSHCUT_EXE = r"C:\Program Files\SAP\FrontEnd\SAPgui\sapshcut.exe"


def garantir_saplogon_unico():
    """Mantem no maximo 1 SAP Logon aberto (2 instancias quebram o OpenConnection)."""
    procs = [
        p for p in psutil.process_iter(["name", "create_time"])
        if (p.info.get("name") or "").lower() == "saplogon.exe"
    ]
    if not procs:
        exe = next((p for p in SAP_LOGON_PATHS if os.path.exists(p)), "saplogon.exe")
        print(f"Abrindo o SAP Logon: {exe}")
        subprocess.Popen([exe], shell=False)
        time.sleep(5)
        return
    if len(procs) > 1:
        procs.sort(key=lambda p: p.info.get("create_time") or 0)
        print(f"{len(procs)} SAP Logon abertos - mantendo o mais antigo (PID {procs[0].pid}).")
        for p in procs[1:]:
            try:
                p.kill()
            except Exception:
                pass
    else:
        print("SAP Logon ja esta aberto - reutilizando.")

def obter_engine(retries=30):
    """Devolve o GuiApplication (engine de scripting), com retry."""
    for i in range(retries):
        try:
            return win32com.client.GetObject("SAPGUI").GetScriptingEngine
        except Exception:
            if i == 0:
                print("Aguardando o SAP GUI disponibilizar o scripting...")
            time.sleep(1)
    raise RuntimeError(
        "Scripting do SAP nao disponivel. Abra o SAP Logon e confirme em "
        "Opcoes > Acessibilidade e Scripting que 'Ativar Scripting' esta marcado."
    )


def sessao_existente(application):
    """Primeira sessao de qualquer conexao ja aberta (sem chamar OpenConnection)."""
    try:
        for conn in application.Children:
            try:
                if conn.Children.Count > 0:
                    return conn.Children(0)
            except Exception:
                continue
    except Exception:
        pass
    return None


def aguardar_sessao(application, timeout=60):
    """Espera ate 'timeout' segundos por uma sessao disponivel."""
    limite = time.time() + timeout
    while time.time() < limite:
        sess = sessao_existente(application)
        if sess is not None:
            return sess
        time.sleep(2)
    return None

def openconn_via_vbs(spec, encoding="cp1252", timeout=45):
    """Tenta OpenConnection pelo VBScript (mesmo caminho late binding do script).

    'encoding' cp1252 e necessario quando 'spec' tem acentos (nome da conexao),
    pois o cscript le arquivos .vbs em ANSI.
    """
    code = fr"""
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
On Error Resume Next
Set connection = application.OpenConnection("{spec}", True)
If Err.Number <> 0 Then
   WScript.Echo "FALHA_OPENCONNECTION " & Err.Number & ": " & Err.Description
   WScript.Quit 3
End If
On Error GoTo 0
"""
    return "FALHA_OPENCONNECTION" not in run_vbs(
        code, timeout=timeout, capture=True, encoding=encoding
    )


def abrir_via_sapshcut():
    """Abre a conexao com sapshcut.exe (nao usa COM - contorna o erro 605).

    A senha NAO vai na linha de comando: a janela abre na tela de login e o
    login continua sendo feito pelo VBScript de login do script.
    Cliente (mandante) opcional via variavel SAP_CLIENT no .env
    """
    if not os.path.exists(SAPSHCUT_EXE):
        print(f"sapshcut.exe nao encontrado em {SAPSHCUT_EXE}.")
        return False
    args = [SAPSHCUT_EXE, f"-system={SAP_SYSTEM_ID}"]
    client = os.environ.get("SAP_CLIENT", "").strip()
    if client:
        args.append(f"-client={client}")
    args.append(f"-language={SAP_LANGUAGE}")
    print("Tentando abrir a conexao via sapshcut.exe:", " ".join(args))
    subprocess.Popen(args, shell=False)
    return True

def obter_sessao():
    """Devolve uma GuiSession pronta para uso, tentando em cascata."""
    garantir_saplogon_unico()
    application = obter_engine()
    try:
        print(f"Conexoes abertas: {application.Connections.Count}")
    except Exception:
        pass

    # 1) Sessao ja aberta (caminho mais confiavel neste ambiente)
    sessao = aguardar_sessao(application, timeout=5)
    if sessao is not None:
        try:
            print(f"Reaproveitando sessao existente: {sessao.Info.SystemName}/"
                  f"{sessao.Info.Client} (usuario {sessao.Info.User or 'nao logado'})")
        except Exception:
            print("Reaproveitando sessao existente.")
        try:
            sessao.findById("wnd[0]").maximize()
        except Exception:
            pass
        return sessao

    # 2) OpenConnection via VBScript: por descricao e por connection string
    for spec, enc in ((SAP_CONNECTION_NAME, "cp1252"), (SAP_CONNECTION_STRING, "utf-8")):
        if openconn_via_vbs(spec, encoding=enc):
            sessao = aguardar_sessao(application, timeout=45)
            if sessao is not None:
                print(f"Conectado via OpenConnection ({spec}).")
                try:
                    sessao.findById("wnd[0]").maximize()
                except Exception:
                    pass
                return sessao
        print(f"OpenConnection nao funcionou para: {spec}")

    # 3) sapshcut.exe (abre sem COM) + anexar a sessao criada
    if abrir_via_sapshcut():
        sessao = aguardar_sessao(application, timeout=90)
        if sessao is not None:
            print("Conectado via sapshcut.exe.")
            try:
                sessao.findById("wnd[0]").maximize()
            except Exception:
                pass
            return sessao

    raise RuntimeError(
        "Nao foi possivel abrir/reaproveitar uma sessao do SAP.\n"
        "Contorno manual: abra o SAP Logon, conecte-se em "
        f"'{SAP_CONNECTION_NAME}' e rode o script novamente - ele vai "
        "reaproveitar a sessao que ja estiver aberta."
    )

# ---------------------------------------------------------------------------
# WATCHDOG DO POPUP "SEGURANCA SAPGUI"
#
# Quando o script manda o SAP GUI gravar arquivo no PC (download para
# C:\Automat\), o SAP mostra um dialogo NATIVO do Windows com o check
# 'Memorizar aqui' e os botoes 'Permitir'/'Nao permitir'. Esse dialogo NAO
# esta na arvore de scripting (session.findById nao o encontra), por isso e
# tratado por API Win32, em uma thread paralela - enquanto o cscript esta
# bloqueado esperando o SAP responder.
#
# Nao interfere na extracao: nao usa COM, nao mexe na sessao, nao move o
# mouse nem rouba o foco (usa BM_SETCHECK/BM_CLICK direto no controle).
#
# Variaveis no .env (opcionais):
#   WATCHDOG_SEGURANCA=1|0|log   (padrao 1)  log = so registra, nao clica
#   MEMORIZAR_SEGURANCA=1|0      (padrao 1)  marca o check 'Memorizar aqui'
# ---------------------------------------------------------------------------

WATCHDOG_SEGURANCA = os.environ.get("WATCHDOG_SEGURANCA", "1").strip().lower()
MEMORIZAR_SEGURANCA = os.environ.get("MEMORIZAR_SEGURANCA", "1").strip() != "0"
WATCHDOG_INTERVALO = 0.5

_BS_TIPO = 0x000F              # mascara do tipo do botao (BS_*)
_BS_CHECKBOX = (2, 3, 5, 6)    # BS_CHECKBOX/BS_AUTOCHECKBOX/BS_3STATE/BS_AUTO3STATE
_BS_PUSHBUTTON = (0, 1)        # BS_PUSHBUTTON/BS_DEFPUSHBUTTON
_BM_SETCHECK = 0x00F1
_BST_CHECKED = 1
_BM_CLICK = 0x00F5
_PALAVRAS_PERMITIR = ("permitir", "allow", "zulassen", "erlauben", "aceitar")
_PALAVRAS_RECUSAR = ("não permitir", "nao permitir", "deny", "recusar", "rejeitar",
                     "nicht zulassen", "abbrechen", "cancelar")
_HINT_TITULO = ("seguran", "security", "sicherheit", "seguridad")

_watchdog_stats = {"popups": 0}


def _pids_sap_gui():
    """PIDs dos processos do SAP GUI (saplogon.exe / sapgui.exe)."""
    try:
        return {
            p.pid for p in psutil.process_iter(["name"])
            if (p.info.get("name") or "").lower() in ("saplogon.exe", "sapgui.exe")
        }
    except Exception:
        return set()


def _controles_do_dialogo(hwnd):
    """[(hwnd, classe, texto, style)] de todos os controles filhos."""
    itens = []

    def _cb(h, _):
        try:
            itens.append((
                h,
                win32gui.GetClassName(h),
                (win32gui.GetWindowText(h) or "").strip(),
                win32gui.GetWindowLong(h, win32con.GWL_STYLE),
            ))
        except Exception:
            pass
        return True

    try:
        win32gui.EnumChildWindows(hwnd, _cb, None)
    except Exception:
        pass
    return itens


def tratar_popups_seguranca(pids=None, classe="#32770"):
    """Responde os dialogos do SAP GUI ('Seguranca SAPGUI').

    Marca 'Memorizar aqui' (se MEMORIZAR_SEGURANCA) e clica em 'Permitir'.
    Devolve quantos dialogos foram respondidos (0 = nenhum na tela).

    Os parametros 'pids'/'classe' existem para teste (default = SAP GUI real).
    """
    if pids is None:
        pids = _pids_sap_gui()
    if not pids:
        return 0

    dialogos = []

    def _cb(hwnd, _):
        try:
            if win32gui.GetClassName(hwnd) == classe and win32gui.IsWindowVisible(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid in pids:
                    dialogos.append((hwnd, pid))
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:
        return 0

    respondidos = 0
    for dlg, pid in dialogos:
        try:
            titulo = (win32gui.GetWindowText(dlg) or "").strip()
            controles = _controles_do_dialogo(dlg)
            checks = [h for h, c, _t, st in controles
                      if c.lower() == "button" and (st & _BS_TIPO) in _BS_CHECKBOX]
            botoes = [(h, t) for h, c, t, st in controles
                      if c.lower() == "button" and (st & _BS_TIPO) in _BS_PUSHBUTTON and t]

            # so age em dialogo com cara de seguranca (independe do idioma)
            if not checks and not any(k in titulo.lower() for k in _HINT_TITULO):
                continue

            if WATCHDOG_SEGURANCA == "log":
                print(f"  [watchdog/auditoria] dialogo do SAP GUI: {titulo!r} (pid {pid})")
                for h, c, t, st in controles:
                    print(f"     hwnd={h} classe={c!r} texto={t!r} tipo={st & _BS_TIPO}")
                continue

            # botao de permitir: por texto; se nao achar, o unico que nao e recusa
            alvo = None
            for h, t in botoes:
                tl = t.lower()
                if any(p in tl for p in _PALAVRAS_PERMITIR) and \
                        not any(r in tl for r in _PALAVRAS_RECUSAR):
                    alvo = (h, t)
                    break
            if alvo is None:
                restantes = [(h, t) for h, t in botoes
                             if not any(r in t.lower() for r in _PALAVRAS_RECUSAR)]
                if len(restantes) == 1:
                    alvo = restantes[0]
            if alvo is None:
                print(f"  [watchdog] dialogo {titulo!r} nao reconhecido - aguardando clique manual.")
                continue

            if MEMORIZAR_SEGURANCA:
                for h in checks:
                    try:
                        win32gui.SendMessage(h, _BM_SETCHECK, _BST_CHECKED, 0)
                    except Exception:
                        pass
            win32gui.SendMessage(alvo[0], _BM_CLICK, 0, 0)
            respondidos += 1
            _watchdog_stats["popups"] += 1
            print(f"  [watchdog] popup {titulo!r}: 'Memorizar aqui' marcado e "
                  f"'{alvo[1]}' clicado (total no run: {_watchdog_stats['popups']}).")
        except Exception:
            continue
    return respondidos


def iniciar_watchdog_seguranca():
    """Sobe a thread paralela que responde aos popups durante todo o fluxo."""
    if WATCHDOG_SEGURANCA in ("0", "off", "nao", "no", "false"):
        print("Watchdog do popup de seguranca: DESATIVADO (.env WATCHDOG_SEGURANCA=0).")
        return None
    if WATCHDOG_SEGURANCA == "log":
        print("Watchdog do popup de seguranca: modo AUDITORIA (so registra, nao clica).")

    def _loop():
        while True:
            try:
                tratar_popups_seguranca()
            except Exception:
                pass
            time.sleep(WATCHDOG_INTERVALO)

    t = threading.Thread(target=_loop, name="watchdog_seguranca", daemon=True)
    t.start()
    print("Watchdog do popup de seguranca: ATIVO (thread paralela).")
    return t


def run_vbs(code, timeout=30, capture=False, encoding="utf-8"):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".vbs") as tmp:
        tmp.write(code.encode(encoding, errors="replace"))
        tmp_path = tmp.name

    try:
        if capture:
            proc = subprocess.run(
                ["cscript.exe", "//Nologo", tmp_path],
                timeout=timeout,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            saida = (proc.stdout or "") + (proc.stderr or "")
            if saida.strip():
                print(saida.strip())
            return saida
        subprocess.run(
            ["cscript.exe", "//Nologo", tmp_path],
            timeout=timeout,  # Tempo máximo de execução em segundos
            check=True
        )
        return ""
    except subprocess.TimeoutExpired:
        print(f"O script VBS ultrapassou {timeout} segundos e foi encerrado.")
        return ""
    finally:
        os.remove(tmp_path)


def generate_vbs_code(segment_start, segment_end, year_text=TARGET_YEAR_TEXT):
    # Caminho de saida com barra final (o SAP espera "C:\Automat\")
    output_dir = OUTPUT_DIR if OUTPUT_DIR.endswith("\\") else OUTPUT_DIR + "\\"

    # Bloco opcional que grava o dump da arvore lida (diagnostico DUMP_TREE=1)
    dump_block = ""
    if DUMP_TREE:
        dump_block = (
            "On Error Resume Next\n"
            'Set fso = CreateObject("Scripting.FileSystemObject")\n'
            f'Set f = fso.CreateTextFile("{output_dir}_tree_{segment_start}_{segment_end}.txt", True)\n'
            "f.Write dumpLines\n"
            "f.Close\n"
            "On Error GoTo 0\n"
        )

    vbs_script = fr"""
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "S_ALR_87013019"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/txt$6-KOKRS").text = "BR01"
session.findById("wnd[0]/usr/ctxt_6ORDGRP-LOW").text = "{segment_start}"
session.findById("wnd[0]/usr/ctxt_6ORDGRP-HIGH").text = "{segment_end}"
session.findById("wnd[0]/usr/ctxt_6ORDGRP-HIGH").setFocus
session.findById("wnd[0]/usr/ctxt_6ORDGRP-HIGH").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 500
' Seleciona a pasta do ano pelo TEXTO (ex: "2026 2026") e nao pelo ID fixo,
' pois o ID (000002/000003/000004/...) muda de posicao conforme o projeto.
On Error Resume Next
Set shell = session.findById("wnd[0]/shellcont/shell/shellcont[2]/shell")
' Expande a pasta raiz para revelar os anos filhos (ignora erro se nao for pasta)
shell.expandNode "000001"
WScript.Sleep 300
' Varre TODOS os nos da arvore (GetAllNodeKeys) e le o texto de cada um
targetKey = ""
targetText = ""
dumpLines = ""
Set allKeys = shell.GetAllNodeKeys
For i = 0 To allKeys.Count - 1
    nodeKey = allKeys.ElementAt(i)
    nodeText = shell.GetNodeTextByKey(nodeKey)
    dumpLines = dumpLines & nodeKey & vbTab & nodeText & vbCrLf
    If targetKey = "" And InStr(1, nodeText, "{year_text}", 1) > 0 Then
        targetKey = nodeKey
        targetText = nodeText
    End If
Next
On Error GoTo 0
{dump_block}If targetKey = "" Then
    Err.Raise 9999, "SAP_EXTRACT", "Pasta do ano {year_text} nao encontrada na arvore Variacao: Exercicio (segmento {segment_start}-{segment_end})."
End If
shell.selectedNode = targetKey
WScript.Sleep 200

session.findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").select
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "{output_dir}"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "{segment_start}_{segment_end}.txt"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]/tbar[0]/btn[11]").press
"""
    return vbs_script


if __name__ == "__main__":
    # Garante a pasta de saida ANTES de extrair (o SAP nao cria a pasta sozinho)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print(f"Pasta de saida: {OUTPUT_DIR}")
    if DUMP_TREE:
        print("Diagnostico DUMP_TREE=1: a arvore lida de cada segmento sera gravada "
              f"em {OUTPUT_DIR}\\_tree_<segmento>.txt")
    # Executar
    iniciar_watchdog_seguranca()
    session = obter_sessao()
    data_um = (datetime.datetime.now() - datetime.timedelta(days=3)).strftime("%d.%m.%Y")
    data_2 = datetime.datetime.now().strftime("%d.%m.%Y")

    # VBScript para login único
    login_vbs_code = fr"""
    If Not IsObject(application) Then
       Set SapGuiAuto  = GetObject("SAPGUI")
       Set application = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(connection) Then
       Set connection = application.Children(0)
    End If
    If Not IsObject(session) Then
       Set session    = connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session,     "on"
       WScript.ConnectObject application, "on"
    End If
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "{sap_user}"
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "{sap_password}"
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 13
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    """
    # Login somente se a sessao ainda nao estiver autenticada
    try:
        usuario_logado = (session.Info.User or "").strip()
    except Exception:
        usuario_logado = ""
    if usuario_logado:
        print(f"Sessao ja autenticada como {usuario_logado} - pulando login.")
    else:
        # Preenche o mandante (client) quando a tela de logon vem sem ele
        if sap_client:
            client_vbs_code = fr"""
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(session) Then
   Set session = application.Children(0).Children(0)
End If
On Error Resume Next
session.findById("wnd[0]/usr/ctxtRSYST-MANDT").text = "{sap_client}"
On Error GoTo 0
"""
            run_vbs(client_vbs_code, timeout=15)
        else:
            print("SAP_CLIENT nao definido no .env - se a tela de logon pedir o mandante, "
                  "preencha a mao ou adicione SAP_CLIENT=300 no .env")
        run_vbs(login_vbs_code.format(sap_user=sap_user, sap_password=sap_password), timeout=30)

    segmentos = segments_to_extract[:SEGMENT_LIMIT] if SEGMENT_LIMIT > 0 else segments_to_extract
    if SEGMENT_LIMIT > 0:
        print(f"Piloto SEGMENT_LIMIT={SEGMENT_LIMIT}: rodando apenas {len(segmentos)} de "
              f"{len(segments_to_extract)} segmentos.")
    for start, end in segmentos:
        print(f"Extraindo dados para os segmentos: {start} - {end}, Ano: {TARGET_YEAR_TEXT}")
        current_vbs_code = generate_vbs_code(start, end, TARGET_YEAR_TEXT)
        # capture=True: mostra/loga o erro do VBScript sem derrubar o programa inteiro
        run_vbs(current_vbs_code, timeout=60, capture=True)
        # Comando para resetar a tela do SAP após a extração
        reset_vbs_code = fr"""
    If Not IsObject(application) Then
       Set SapGuiAuto  = GetObject("SAPGUI")
       Set application = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(connection) Then
       Set connection = application.Children(0)
    End If
    If Not IsObject(session) Then
       Set session    = connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session,     "on"
       WScript.ConnectObject application, "on"
    End If
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n S_ALR_87013019"
    session.findById("wnd[0]").sendVKey 0

    """
        run_vbs(reset_vbs_code, timeout=10) # Timeout menor para reset, pois é uma operação rápida
