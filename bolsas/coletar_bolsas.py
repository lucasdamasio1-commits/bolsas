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
import time
import subprocess

# ─── CONFIGURAÇÃO ─────────────────────────────────────────────

PASTA_SAIDA = r"C:\Users\lucas\OneDrive\Doutorado\Scripts\bolsas"

ARQUIVO_EXCEL = os.path.join(PASTA_SAIDA, "bolsas_pesquisa.xlsx")
ARQUIVO_LOG   = os.path.join(PASTA_SAIDA, "coleta.log")
ARQUIVO_IDS   = os.path.join(PASTA_SAIDA, "ids_vistos.json")

os.makedirs(PASTA_SAIDA, exist_ok=True)

# ─── LOG ─────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.FileHandler(ARQUIVO_LOG, encoding="utf-8"),
        logging.StreamHandler()
    ]
)
log = logging.getLogger(__name__)

# ─── AUXILIARES ──────────────────────────────────────────────

def carregar_ids_vistos():
    if os.path.exists(ARQUIVO_IDS):
        with open(ARQUIVO_IDS, "r", encoding="utf-8") as f:
            return set(json.load(f))
    return set()

def salvar_ids_vistos(ids):
    with open(ARQUIVO_IDS, "w", encoding="utf-8") as f:
        json.dump(list(ids), f, indent=2, ensure_ascii=False)

def gerar_id(fonte, titulo):
    return f"{fonte}::{titulo[:80].lower()}"

# ─── EXCEL ───────────────────────────────────────────────────

COLUNAS = ["titulo","fonte","pais","area","nivel","prazo","link","data_coleta","ativa"]

def salvar_excel(bolsas, ids_vistos):
    if os.path.exists(ARQUIVO_EXCEL):
        wb = openpyxl.load_workbook(ARQUIVO_EXCEL)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(COLUNAS)

    adicionadas = 0

    for b in bolsas:
        bid = gerar_id(b["fonte"], b["titulo"])
        if bid in ids_vistos:
            continue

        ws.append([
            b["titulo"], b["fonte"], b["pais"], b["area"],
            b["nivel"], b["prazo"], b["link"],
            b["data_coleta"], b["ativa"]
        ])

        ids_vistos.add(bid)
        adicionadas += 1

    wb.save(ARQUIVO_EXCEL)
    return adicionadas

# ─── HTML (AGORA LENDO DO EXCEL) ─────────────────────────────

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
          <td><a href="{row[6]}" target="_blank">Acessar</a></td>
        </tr>
        """

    html = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Bolsas</title>
<style>
body{{font-family:sans-serif}}
table{{width:100%;border-collapse:collapse}}
th{{background:#1F3864;color:#fff}}
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
  let rows = document.querySelectorAll("#tabela tbody tr");

  rows.forEach(r => {{
    r.style.display = r.innerText.toLowerCase().includes(input) ? "" : "none";
  }});
}}
</script>

</body>
</html>
"""

    caminho = os.path.join(PASTA_SAIDA, "index.html")
    with open(caminho, "w", encoding="utf-8") as f:
        f.write(html)

    log.info("HTML atualizado")

# ─── GITHUB AUTOMÁTICO ───────────────────────────────────────

def publicar_github():
    pasta = PASTA_SAIDA

    try:
        subprocess.run(["git", "-C", pasta, "add", "."], check=True)
        subprocess.run(["git", "-C", pasta, "commit", "-m",
                        f"Atualização {date.today()}"], check=True)
        subprocess.run(["git", "-C", pasta, "push"], check=True)
        log.info("GitHub atualizado com sucesso")
    except Exception as e:
        log.warning(f"Erro no git: {e}")

# ─── MOCK COLETA (mantive simples) ───────────────────────────

def coletar_mock():
    return [{
        "titulo": "Bolsa Exemplo",
        "fonte": "Teste",
        "pais": "Brasil",
        "area": "Administração",
        "nivel": "Mestrado",
        "prazo": "2026-12-01",
        "link": "https://exemplo.com",
        "data_coleta": str(date.today()),
        "ativa": "Sim"
    }]

# ─── MAIN ────────────────────────────────────────────────────

def main():
    log.info("Iniciando coleta")

    ids_vistos = carregar_ids_vistos()

    bolsas = coletar_mock()  # substitua pelos seus coletores reais

    adicionadas = salvar_excel(bolsas, ids_vistos)

    salvar_ids_vistos(ids_vistos)

    gerar_html()

    publicar_github()

    log.info(f"{adicionadas} novas bolsas adicionadas")

if __name__ == "__main__":
    main()