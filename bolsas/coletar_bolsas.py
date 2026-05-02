"""
Monitor de Bolsas de Pesquisa
==============================
Coleta oportunidades de: CNPq, CAPES, FAPESP, Fulbright, DAAD e feeds RSS.
Salva em Excel no OneDrive (ou pasta local).

Dependências:
    pip install requests beautifulsoup4 openpyxl feedparser lxml

Agendamento: Windows Task Scheduler (ver instrucoes em LEIAME.txt)
"""

import requests
from bs4 import BeautifulSoup
import feedparser
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime, date
import os
import json
import logging
import subprocess
from datetime import date, datetime

import requests
from bs4 import BeautifulSoup
import openpyxl

# ─── CONFIG ─────────────────────────────────────────────

PASTA_SAIDA = r"C:\Users\lucas\OneDrive\Doutorado\Scripts\bolsas"

ARQUIVO_EXCEL = os.path.join(PASTA_SAIDA, "bolsas_pesquisa.xlsx")
ARQUIVO_IDS   = os.path.join(PASTA_SAIDA, "ids_vistos.json")

os.makedirs(PASTA_SAIDA, exist_ok=True)

# ─── LOG ────────────────────────────────────────────────

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

# ─── RESET (NOVO) ───────────────────────────────────────

def resetar_excel():
    """Apaga Excel antigo e cria novo do zero"""
    if os.path.exists(ARQUIVO_EXCEL):
        os.remove(ARQUIVO_EXCEL)
        log.info("Excel antigo removido")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bolsas"

    ws.append([
        "titulo","fonte","pais","area",
        "nivel","prazo","link","data_coleta","ativa"
    ])

    wb.save(ARQUIVO_EXCEL)
    log.info("Excel recriado (zerado)")

# ─── COLETAS ────────────────────────────────────────────

def coletar_cnpq():
    url = "https://www.gov.br/cnpq/pt-br/acesso-a-informacao/bolsas-e-auxilios"
    bolsas = []

    try:
        r = requests.get(url)
        soup = BeautifulSoup(r.text, "html.parser")

        for link in soup.select("a")[:20]:
            titulo = link.text.strip()

            if len(titulo) < 10:
                continue

            bolsas.append({
                "titulo": titulo,
                "fonte": "CNPq",
                "pais": "Brasil",
                "area": "Pesquisa",
                "nivel": "Diversos",
                "prazo": "Consultar site",
                "link": link.get("href"),
                "data_coleta": str(date.today()),
                "ativa": "Sim"
            })

    except Exception as e:
        log.warning(f"CNPq erro: {e}")

    return bolsas


def coletar_capes():
    url = "https://www.gov.br/capes/pt-br/acesso-a-informacao/acoes-e-programas/bolsas"
    bolsas = []

    try:
        r = requests.get(url)
        soup = BeautifulSoup(r.text, "html.parser")

        for link in soup.select("a")[:20]:
            titulo = link.text.strip()

            if len(titulo) < 10:
                continue

            bolsas.append({
                "titulo": titulo,
                "fonte": "CAPES",
                "pais": "Brasil",
                "area": "Educação",
                "nivel": "Diversos",
                "prazo": "Consultar site",
                "link": link.get("href"),
                "data_coleta": str(date.today()),
                "ativa": "Sim"
            })

    except Exception as e:
        log.warning(f"CAPES erro: {e}")

    return bolsas


def coletar_fapesp():
    url = "https://fapesp.br/oportunidades"
    bolsas = []

    try:
        r = requests.get(url)
        soup = BeautifulSoup(r.text, "html.parser")

        for item in soup.select("a")[:20]:
            titulo = item.text.strip()

            if len(titulo) < 10:
                continue

            bolsas.append({
                "titulo": titulo,
                "fonte": "FAPESP",
                "pais": "Brasil",
                "area": "Pesquisa",
                "nivel": "Diversos",
                "prazo": "Consultar site",
                "link": "https://fapesp.br" + item.get("href", ""),
                "data_coleta": str(date.today()),
                "ativa": "Sim"
            })

    except Exception as e:
        log.warning(f"FAPESP erro: {e}")

    return bolsas


def coletar_fulbright():
    url = "https://fulbright.org.br/bolsas/"
    bolsas = []

    try:
        r = requests.get(url)
        soup = BeautifulSoup(r.text, "html.parser")

        for item in soup.select("h2")[:10]:
            titulo = item.text.strip()

            bolsas.append({
                "titulo": titulo,
                "fonte": "Fulbright",
                "pais": "Internacional",
                "area": "Diversos",
                "nivel": "Diversos",
                "prazo": "Consultar site",
                "link": url,
                "data_coleta": str(date.today()),
                "ativa": "Sim"
            })

    except Exception as e:
        log.warning(f"Fulbright erro: {e}")

    return bolsas


def coletar_daad():
    url = "https://www.daad.de/en/study-and-research-in-germany/scholarships/"
    bolsas = []

    try:
        r = requests.get(url)
        soup = BeautifulSoup(r.text, "html.parser")

        for item in soup.select("a")[:20]:
            titulo = item.text.strip()

            if len(titulo) < 15:
                continue

            bolsas.append({
                "titulo": titulo,
                "fonte": "DAAD",
                "pais": "Alemanha",
                "area": "Diversos",
                "nivel": "Diversos",
                "prazo": "Consultar site",
                "link": item.get("href"),
                "data_coleta": str(date.today()),
                "ativa": "Sim"
            })

    except Exception as e:
        log.warning(f"DAAD erro: {e}")

    return bolsas


def coletar_horizon():
    return [{
        "titulo": "Chamadas Horizon Europe (Portal Oficial)",
        "fonte": "Horizon Europe",
        "pais": "Europa",
        "area": "Pesquisa",
        "nivel": "Doutorado/Pós-doc",
        "prazo": "Variável",
        "link": "https://ec.europa.eu/info/funding-tenders/opportunities/portal/",
        "data_coleta": str(date.today()),
        "ativa": "Sim"
    }]

# ─── COLETOR PRINCIPAL ──────────────────────────────────

def coletar_bolsas():
    bolsas = []
    bolsas += coletar_cnpq()
    bolsas += coletar_capes()
    bolsas += coletar_fapesp()
    bolsas += coletar_fulbright()
    bolsas += coletar_daad()
    bolsas += coletar_horizon()

    log.info(f"{len(bolsas)} bolsas coletadas")
    return bolsas

# ─── EXCEL ──────────────────────────────────────────────

def salvar_excel(bolsas):
    wb = openpyxl.load_workbook(ARQUIVO_EXCEL)
    ws = wb.active

    for b in bolsas:
        ws.append(list(b.values()))

    wb.save(ARQUIVO_EXCEL)
    log.info("Excel preenchido")

# ─── HTML ───────────────────────────────────────────────

def gerar_html():
    wb = openpyxl.load_workbook(ARQUIVO_EXCEL)
    ws = wb.active

    linhas = ""

    for row in ws.iter_rows(min_row=2, values_only=True):
        linhas += f"""
        <tr>
          <td>{row[0]}</td>
          <td>{row[1]}</td>
          <td>{row[2]}</td>
          <td>{row[4]}</td>
          <td>{row[5]}</td>
          <td><a href="{row[6]}" target="_blank">Abrir</a></td>
        </tr>
        """

    html = f"""
<html>
<body>
<h1>Bolsas de Pesquisa</h1>
<p>{date.today()}</p>

<table border="1">
<tr><th>Título</th><th>Fonte</th><th>País</th><th>Nível</th><th>Prazo</th><th>Link</th></tr>
{linhas}
</table>

</body>
</html>
"""

    with open(os.path.join(PASTA_SAIDA, "index.html"), "w", encoding="utf-8") as f:
        f.write(html)

    log.info("HTML gerado")

# ─── GITHUB ─────────────────────────────────────────────

def publicar():
    try:
        subprocess.run(["git", "-C", PASTA_SAIDA, "add", "."], check=True)
        subprocess.run(["git", "-C", PASTA_SAIDA, "commit", "-m",
                        f"update {date.today()}"], check=False)
        subprocess.run(["git", "-C", PASTA_SAIDA, "pull", "--rebase"], check=True)
        subprocess.run(["git", "-C", PASTA_SAIDA, "push"], check=True)
        log.info("GitHub atualizado")
    except Exception as e:
        log.warning(f"Erro no git: {e}")

# ─── MAIN ───────────────────────────────────────────────

def main():
    log.info("Iniciando processo")

    resetar_excel()

    bolsas = coletar_bolsas()

    salvar_excel(bolsas)

    gerar_html()

    publicar()

    log.info("Processo finalizado")

if __name__ == "__main__":
    main()