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
import openpyxl
import feedparser

# ─── CONFIG ─────────────────────────────────────────────

PASTA_SAIDA = r"C:\Users\lucas\OneDrive\Doutorado\Scripts\bolsas"

ARQUIVO_EXCEL = os.path.join(PASTA_SAIDA, "bolsas_pesquisa.xlsx")
ARQUIVO_IDS   = os.path.join(PASTA_SAIDA, "ids_vistos.json")

os.makedirs(PASTA_SAIDA, exist_ok=True)

# ─── LOG ────────────────────────────────────────────────

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

# ─── AUX ────────────────────────────────────────────────

def gerar_id(fonte, titulo):
    return f"{fonte}::{titulo[:80].lower()}"

def carregar_ids():
    if os.path.exists(ARQUIVO_IDS):
        with open(ARQUIVO_IDS, "r", encoding="utf-8") as f:
            return set(json.load(f))
    return set()

def salvar_ids(ids):
    with open(ARQUIVO_IDS, "w", encoding="utf-8") as f:
        json.dump(list(ids), f, indent=2, ensure_ascii=False)

# ─── COLETA REAL (RSS CORRIGIDO) ────────────────────────

def coletar_rss():
    feeds = [
        ("ScholarshipPortal", "https://www.scholarshipportal.com/rss/scholarships"),
        ("FindAPhD", "https://www.findaphd.com/rss"),
        ("AcademicPositions", "https://academicpositions.com/rss"),
    ]

    bolsas = []

    for nome, url in feeds:
        try:
            feed = feedparser.parse(url)

            if not feed.entries:
                log.warning(f"Feed vazio: {nome}")
                continue

            for entry in feed.entries[:15]:
                bolsas.append({
                    "titulo": entry.get("title", "Sem título"),
                    "fonte": nome,
                    "pais": "Internacional",
                    "area": "Não informado",
                    "nivel": "Diversos",
                    "prazo": "Não informado",
                    "link": entry.get("link", ""),
                    "data_coleta": str(date.today()),
                    "ativa": "Sim"
                })

        except Exception as e:
            log.warning(f"Erro no feed {nome}: {e}")

    log.info(f"{len(bolsas)} bolsas coletadas via RSS")
    return bolsas

# ─── EXCEL ──────────────────────────────────────────────

COLUNAS = [
    "titulo","fonte","pais","area",
    "nivel","prazo","link","data_coleta","ativa"
]

def salvar_excel(bolsas, ids):
    if not os.path.exists(ARQUIVO_EXCEL):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Bolsas"
        ws.append(COLUNAS)
        wb.save(ARQUIVO_EXCEL)
        log.info("Excel criado")

    wb = openpyxl.load_workbook(ARQUIVO_EXCEL)
    ws = wb.active

    adicionadas = 0

    for b in bolsas:
        bid = gerar_id(b["fonte"], b["titulo"])

        if bid in ids:
            continue

        ws.append([
            b["titulo"], b["fonte"], b["pais"], b["area"],
            b["nivel"], b["prazo"], b["link"],
            b["data_coleta"], b["ativa"]
        ])

        ids.add(bid)
        adicionadas += 1

    wb.save(ARQUIVO_EXCEL)
    log.info(f"{adicionadas} novas bolsas adicionadas no Excel")

# ─── HTML ───────────────────────────────────────────────

def gerar_html():
    if not os.path.exists(ARQUIVO_EXCEL):
        log.error("Excel não existe — HTML não pode ser gerado")
        return

    wb = openpyxl.load_workbook(ARQUIVO_EXCEL)
    ws = wb.active

    linhas = ""

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue

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

    if not linhas:
        linhas = "<tr><td colspan='6'>Nenhuma bolsa encontrada</td></tr>"

    html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Bolsas</title>
<style>
body{{font-family:sans-serif;max-width:1100px;margin:auto}}
table{{width:100%;border-collapse:collapse}}
th{{background:#1F3864;color:white}}
td,th{{padding:8px;border:1px solid #ddd}}
</style>
</head>
<body>

<h1>Bolsas de Pesquisa</h1>
<p>Atualizado em {date.today()}</p>

<input type="text" id="busca" placeholder="Buscar..." onkeyup="filtrar()">

<table id="tabela">
<thead>
<tr>
<th>Título</th><th>Fonte</th><th>País</th>
<th>Nível</th><th>Prazo</th><th>Link</th>
</tr>
</thead>

<tbody>
{linhas}
</tbody>
</table>

<script>
function filtrar(){{
  let input = document.getElementById("busca").value.toLowerCase();
  document.querySelectorAll("#tabela tbody tr").forEach(tr => {{
    tr.style.display = tr.innerText.toLowerCase().includes(input) ? "" : "none";
  }});
}}
</script>

</body>
</html>
"""

    caminho = os.path.join(PASTA_SAIDA, "index.html")

    with open(caminho, "w", encoding="utf-8") as f:
        f.write(html)

    log.info("HTML gerado com sucesso")

# ─── GITHUB (CORRIGIDO) ─────────────────────────────────

def publicar():
    try:
        subprocess.run(["git", "-C", PASTA_SAIDA, "add", "."], check=True)

        subprocess.run([
            "git", "-C", PASTA_SAIDA,
            "commit", "-m", f"update {date.today()}"
        ], check=False)

        subprocess.run(["git", "-C", PASTA_SAIDA, "pull", "--rebase"], check=True)

        subprocess.run(["git", "-C", PASTA_SAIDA, "push"], check=True)

        log.info("GitHub atualizado")

    except Exception as e:
        log.warning(f"Erro no git: {e}")

# ─── MAIN ───────────────────────────────────────────────

def main():
    log.info("Iniciando processo")

    ids = carregar_ids()

    bolsas = coletar_rss()

    salvar_excel(bolsas, ids)
    salvar_ids(ids)

    gerar_html()
    publicar()

    log.info("Processo finalizado")

if __name__ == "__main__":
    main()