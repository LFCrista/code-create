# automatizador_prompts_docs.py – 2025-06-26-g
"""
► Procura prompts marcados com !*! em um Google Docs.
► Para cada prompt:
    1. Envia ao ChatGPT (aba em modo-debug já aberta).
    2. Espera a NOVA resposta terminar.
    3. Cola a resposta DUAS linhas abaixo do prompt
       (Arial 11 pt, sem bold/itálico).
    4. Verifica se a resposta está realmente abaixo do prompt;
       só então envia o próximo.
    5. Mantém uma fila no Google Sheets com status:
       Pendente • Em Processo • Concluído • Erro.
"""

import os
import re
import sys
import time
import asyncio
import subprocess
import tkinter as tk
from tkinter import filedialog, simpledialog, scrolledtext, messagebox
import webbrowser

from playwright.sync_api import sync_playwright
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from bs4 import BeautifulSoup

# ---------- CONFIG ---------------------------------------------------
CHROME_PATH              = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
CHROME_USER_DATA_DIR     = r"C:\temp\chrome"
CHROME_REMOTE_DEBUG_PORT = 9222
CHROME_DEBUG_URL         = f"http://localhost:{CHROME_REMOTE_DEBUG_PORT}"

SCOPES = [
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

if sys.platform.startswith("win"):
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

# ---------- GOOGLE DOCS ---------------------------------------------
def autenticar_google_docs():
    if os.path.exists("token.json"):
        try:
            return Credentials.from_authorized_user_file("token.json", SCOPES)
        except Exception:
            os.remove("token.json")

    cred = filedialog.askopenfilename(
        title="Selecione credentials.json",
        filetypes=[("JSON", "*.json")]
    )
    if not cred:
        return None

    flow  = InstalledAppFlow.from_client_secrets_file(cred, SCOPES)
    creds = flow.run_local_server(port=0)
    open("token.json", "w").write(creds.to_json())
    return creds


def extrair_document_id(url: str | None):
    m = re.search(r"/d/([A-Za-z0-9\-_]+)", url or "")
    return m.group(1) if m else None

# ---------- GOOGLE SHEETS -------------------------------------------
def criar_planilha_fila(sheets_svc):
    planilha = sheets_svc.spreadsheets().create(body={
        "properties": {"title": "Fila de Prompts"},
        "sheets": [{"properties": {"title": "Fila"}}]
    }).execute()

    sheet_id = planilha["spreadsheetId"]
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}"

    # Cabeçalhos
    sheets_svc.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range="Fila!A1:D1",
        valueInputOption="RAW",
        body={"values": [["Prompt", "Status", "Timestamp", "Observação"]]}
    ).execute()

    webbrowser.open_new_tab(url)
    return sheet_id, url


def registrar_prompts_iniciais(sheets_svc, sheet_id, lista_prompts):
    """
    Recebe lista [(end_idx, prompt_txt), …] e grava todos na planilha como Pendente.
    """
    linhas = [[txt, "Pendente",
               time.strftime("%Y-%m-%d %H:%M:%S"), ""]
              for _, txt in lista_prompts]

    if linhas:
        sheets_svc.spreadsheets().values().append(
            spreadsheetId=sheet_id,
            range="Fila!A:D",
            valueInputOption="RAW",
            body={"values": linhas}
        ).execute()


def atualizar_status(sheets_svc, sheet_id,
                     prompt_txt, novo_status, observacao: str = ""):
    valores = sheets_svc.spreadsheets().values().get(
        spreadsheetId=sheet_id, range="Fila!A2:D").execute().get("values", [])

    for i, row in enumerate(valores):
        if row and row[0].strip() == prompt_txt:
            sheets_svc.spreadsheets().values().update(
                spreadsheetId=sheet_id,
                range=f"Fila!B{i+2}:D{i+2}",
                valueInputOption="RAW",
                body={"values": [[
                    novo_status,
                    time.strftime("%Y-%m-%d %H:%M:%S"),
                    observacao
                ]]}
            ).execute()
            break

# ---------- PLAYWRIGHT HELPERS --------------------------------------
def _stop(page):
    return page.locator("button:has(svg[aria-label='Stop generating'])").count()

def _stream(page):
    return page.locator(".result-streaming, .animate-spin").count()

def _composer(page):
    ta = page.locator("textarea")
    if ta.count():
        try:
            return bool(ta.input_value().strip())
        except Exception:
            pass
    return False

def aguardar_pronto(page, stab: float = 4.0, buf: float = 1.5):
    while _stop(page):
        time.sleep(0.3)

    last_cnt  = page.locator("[data-message-author-role='assistant']").count()
    last_html = page.locator("[data-message-author-role='assistant']").nth(-1) \
        .locator(".markdown").inner_html() if last_cnt else ""

    fase_buf, t0 = False, time.time()
    while True:
        busy = _stop(page) or _stream(page) or _composer(page)
        cnt  = page.locator("[data-message-author-role='assistant']").count()
        html = page.locator("[data-message-author-role='assistant']").nth(-1) \
            .locator(".markdown").inner_html() if cnt else ""

        mudou = busy or cnt != last_cnt or html != last_html
        if mudou:
            fase_buf, t0 = False, time.time()
            last_cnt, last_html = cnt, html
        else:
            now = time.time()
            if not fase_buf and now - t0 >= buf:
                fase_buf, t0 = True, now
            elif fase_buf and now - t0 >= stab:
                return
        time.sleep(10)

def digitar_prompt(page, prompt: str):
    for sel in ("textarea", "div[role='textbox']"):
        try:
            page.wait_for_selector(sel, timeout=2000)
            box = page.locator(sel).first
            if sel == "textarea":
                box.evaluate("n=>n.value=''")
                box.fill(prompt)
            else:
                box.click()
                box.evaluate("n=>n.innerText=''")
                box.type(prompt)
            page.keyboard.press("Enter")
            return
        except Exception:
            pass
    page.keyboard.type(prompt)
    page.keyboard.press("Enter")

def esperar_html_estavel(locator, segundos: float = 3.0, dt: float = 0.4):
    tam = len(locator.inner_html(timeout=0))
    t0  = time.time()
    while True:
        time.sleep(dt)
        novo = len(locator.inner_html(timeout=0))
        if novo != tam:
            tam, t0 = novo, time.time()
        elif time.time() - t0 >= segundos:
            return

# ---------- OBTÉM RESPOSTA COMPLETA ---------------------------------
def obter_resposta(page, prompt_txt: str) -> str:
    prev_cnt = page.locator("[data-message-author-role='assistant']").count()
    digitar_prompt(page, prompt_txt)

    while page.locator("[data-message-author-role='assistant']").count() <= prev_cnt:
        time.sleep(0.3)

    aguardar_pronto(page)

    bolha = page.locator("[data-message-author-role='assistant']").nth(-1)
    md    = bolha.locator(".markdown")
    md.wait_for(state="attached", timeout=60_000)

    while True:
        html = md.inner_html(timeout=0).strip()
        texto = html_para_texto(html).strip()
        if len(texto) > 5:
            return html
        time.sleep(0.3)

# ---------- CONVERSORES ---------------------------------------------
def html_para_texto(html: str) -> str:
    return BeautifulSoup(html, "html.parser").get_text("\n")

# ---------- INSERE E VERIFICA NO DOCS -------------------------------
def inserir_resposta(svc, doc_id: str, insert_at: int, texto_puro: str) -> int:
    """
    Insere a resposta logo após o prompt, com no máximo 1 linha de espaço.
    • Move o ponto de inserção 1 caractere para trás (antes do '\n' do prompt).
    • Adiciona APENAS 1 quebra de linha após o prompt e 1 no final da resposta.
    • Fonte Arial 11 pt, cor preta, sem bold/itálico.
    """
    posicao = max(1, insert_at - 1)          # ← antes do '\n' do prompt
    bloco   = "\n" + texto_puro       # ← 1 antes, 1 depois
    tam     = len(bloco)

    svc.documents().batchUpdate(
        documentId=doc_id,
        body={
            "requests": [
                {"insertText": {
                    "location": {"index": posicao},
                    "text": bloco
                }},
                {"updateTextStyle": {
                    "range": {"startIndex": posicao, "endIndex": posicao + tam},
                    "textStyle": {
                        "weightedFontFamily": {"fontFamily": "Arial"},
                        "fontSize": {"magnitude": 11, "unit": "PT"},
                        "bold": False,
                        "italic": False,
                        "foregroundColor": {
                            "color": {"rgbColor": {"red": 0, "green": 0, "blue": 0}}
                        }
                    },
                    "fields": ("weightedFontFamily,fontSize,"
                               "bold,italic,foregroundColor")
                }}
            ]
        }
    ).execute()

    return tam



def doc_para_texto(svc, doc_id: str) -> str:
    partes = []
    for el in svc.documents().get(documentId=doc_id).execute()["body"]["content"]:
        for e in el.get("paragraph", {}).get("elements", []):
            partes.append(e.get("textRun", {}).get("content", ""))
    return "".join(partes)

def verificar_insercao(svc, doc_id: str, prompt: str, resposta: str) -> bool:
    plano = doc_para_texto(svc, doc_id)
    i_p = plano.find(prompt.strip())
    i_r = plano.find(resposta.strip())
    return i_p != -1 and i_r != -1 and i_r > i_p

# ---------- INTERFACE DE CHROME DEBUG -------------------------------
def abrir_debug():
    subprocess.Popen([
        CHROME_PATH,
        f"--remote-debugging-port={CHROME_REMOTE_DEBUG_PORT}",
        f"--user-data-dir={CHROME_USER_DATA_DIR}"
    ])

# ---------- PROCESSAMENTO PRINCIPAL ---------------------------------
def processar():
    link_gpt = simpledialog.askstring("GPT",  "Cole o link da sala GPT:")
    link_doc = simpledialog.askstring("Docs", "Cole o link do Google Docs:")
    if not (link_gpt and link_doc):
        return

    doc_id = extrair_document_id(link_doc)
    if not doc_id:
        messagebox.showerror("Erro", "ID do documento inválido.")
        return

    texto_log.delete("1.0", tk.END)
    texto_log.insert(tk.END, "🔑 Autenticando Google Docs e Sheets…\n")
    janela.update()

    creds = autenticar_google_docs()
    if not creds:
        return
    svc_docs   = build("docs",   "v1", credentials=creds)
    svc_sheets = build("sheets", "v4", credentials=creds)

    texto_log.insert(tk.END, "📄 Criando planilha da fila…\n")
    janela.update()
    sheet_id, sheet_url = criar_planilha_fila(svc_sheets)
    texto_log.insert(tk.END, f"📎 Fila: {sheet_url}\n\n")
    janela.update()

    texto_log.insert(tk.END, "🕸️ Conectando ao Chrome Debug…\n")
    janela.update()

    with sync_playwright() as p:
        ctx  = p.chromium.connect_over_cdp(CHROME_DEBUG_URL).contexts[0]
        page = ctx.new_page()
        page.goto(link_gpt)
        time.sleep(5)

        # ---- coleta prompts:
        body = svc_docs.documents().get(documentId=doc_id).execute()["body"]["content"]
        prompts, desloc = [], 0
        for el in body:
            elems = el.get("paragraph", {}).get("elements", [])
            if elems:
                txt = elems[0].get("textRun", {}).get("content", "").strip()
                if txt.startswith("!*!"):
                    prompts.append((el["endIndex"], txt[3:].strip()))

        if not prompts:
            messagebox.showinfo("Aviso", "Nenhum !*! encontrado.")
            return

        # grava todos como Pendente
        registrar_prompts_iniciais(svc_sheets, sheet_id, prompts)

        texto_log.insert(tk.END, f"{len(prompts)} prompt(s) listado(s) na fila.\n\n")
        janela.update()

        for i, (end_idx, prompt_txt) in enumerate(prompts):
            is_last = i == len(prompts) - 1
            if is_last:
                end_idx -= 1  # evita erro ao inserir no final do doc

            texto_log.insert(tk.END, f"→ {prompt_txt[:60]}…\n")
            janela.update()

            atualizar_status(svc_sheets, sheet_id, prompt_txt, "Em Processo")

            html_resp  = obter_resposta(page, prompt_txt)
            texto_resp = html_para_texto(html_resp).strip()
            if not texto_resp:
                atualizar_status(svc_sheets, sheet_id, prompt_txt, "Erro", "Resposta vazia")
                texto_log.insert(tk.END, "⚠ Resposta vazia\n")
                continue

            tent, ok = 0, False
            while tent < 3 and not ok:
                if tent:
                    time.sleep(2)
                delta = inserir_resposta(svc_docs, doc_id, end_idx + desloc, texto_resp)
                ok    = verificar_insercao(svc_docs, doc_id, prompt_txt, texto_resp)
                tent += 1

            if ok:
                desloc += delta
                atualizar_status(svc_sheets, sheet_id, prompt_txt, "Concluído")
                texto_log.insert(tk.END, "✓ OK\n")
            else:
                atualizar_status(svc_sheets, sheet_id, prompt_txt, "Erro", "Falha ao inserir")
                texto_log.insert(tk.END, "✗ Falha ao inserir — abortando.\n")
                break
            janela.update()

        messagebox.showinfo("Concluído", "Todos os prompts processados!")

# ---------- INTERFACE TK --------------------------------------------
janela = tk.Tk()
janela.title("Automatizador GPT → Docs")
janela.geometry("780x580")

top = tk.Frame(janela)
top.pack(pady=8)

tk.Button(top, text="Abrir Chrome Debug", command=abrir_debug).pack(side=tk.LEFT, padx=6)
tk.Button(top, text="Processar Docs",    command=processar   ).pack(side=tk.LEFT, padx=6)

texto_log = scrolledtext.ScrolledText(janela, width=100, height=28)
texto_log.pack(pady=10)

janela.mainloop()
