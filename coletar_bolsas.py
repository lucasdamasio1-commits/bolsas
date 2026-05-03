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
import logging
import subprocess
import re
from datetime import date
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
import openpyxl

# ─── CONFIG ─────────────────────────────────────────────

PASTA_SAIDA = r"C:\Users\lucas\OneDrive\Doutorado\Scripts\bolsas"

ARQUIVO_EXCEL = os.path.join(PASTA_SAIDA, "bolsas_pesquisa.xlsx")

os.makedirs(PASTA_SAIDA, exist_ok=True)

# ─── LOG ────────────────────────────────────────────────

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

# ─── PALAVRAS-CHAVE (PT + EN) ───────────────────────────

PALAVRAS_CHAVE = [
    # PT
    "bolsa", "edital", "pesquisa", "doutorado", "mestrado",
    "pós-doutorado", "financiamento", "auxílio",

    # EN
    "scholarship", "fellowship", "grant", "funding",
    "phd", "doctoral", "master", "research",
    "call for applications", "open call"
]

EXCLUIR = [
    "ir para", "menu", "navegação", "mapa",
    "login", "home", "início", "imposto",
    "acesso", "institucional", "organograma",
    "privacy", "terms", "cookies"
]

# ─── FUNÇÕES INTELIGENTES ──────────────────────────────

def detectar_nivel(texto):
    if "phd" in texto or "doutorado" in texto:
        return "Doutorado"
    elif "master" in texto or "mestrado" in texto:
        return "Mestrado"
    elif "postdoc" in texto or "pós" in texto:
        return "Pós-doc"
    else:
        return "Pesquisa / Geral"


def link_valido(link, base):
    if not link:
        return None

    if link.startswith("http"):
        return link

    if link.startswith("/"):
        return base + link

    return None


def extrair_links_filtrados(url, fonte, pais):
    bolsas = []

    try:
        r = requests.get(url, timeout=10)
        soup = BeautifulSoup(r.text, "html.parser")

        for a in soup.find_all("a"):
            titulo = a.get_text(strip=True)
            link = a.get("href")

            if not titulo or not link:
                continue

            titulo_lower = titulo.lower()

            # filtro positivo
            if not any(p in titulo_lower for p in PALAVRAS_CHAVE):
                continue

            # filtro negativo
            if any(e in titulo_lower for e in EXCLUIR):
                continue

            link = link_valido(link, url)

            if not link:
                continue

            bolsas.append({
                "titulo": titulo,
                "fonte": fonte,
                "pais": pais,
                "area": "Pesquisa",
                "nivel": detectar_nivel(titulo_lower),
                "prazo": "Consultar edital",
                "link": link,
                "data_coleta": str(date.today()),
                "ativa": "Sim"
            })

    except Exception as e:
        log.warning(f"{fonte} erro: {e}")

    return bolsas

# ─── COLETA POR FONTE ──────────────────────────────────

def coletar_euraxess_it_fr_es():
    """
    Busca contratos, financiamentos e oportunidades de hosting no EURAXESS
    para Italia, Franca e Espanha, mantendo apenas resultados relacionados a
    marketing, administracao/gestao e inteligencia artificial.
    """
    bolsas = []
    url_base = "https://euraxess.ec.europa.eu/jobs/search"
    paises = {"Italia": "781", "Franca": "793", "Espanha": "788"}
    tipos = {
        "job_offer": "Contrato / Vaga",
        "funding": "Financiamento",
        "hosting": "Hosting Offer",
    }
    termos = re.compile(
        r"\b("
        r"artificial intelligence|machine learning|deep learning|data science|"
        r"big data|analytics|algorithm|algorithms|intelig[eê]ncia artificial|"
        r"marketing|consumer|brand|branding|advertising|communication|"
        r"management|business administration|administration|organizational|"
        r"governance|leadership|strategy|innovation|entrepreneurship|"
        r"gest[aã]o|administra[cç][aã]o|inova[cç][aã]o|empreendedorismo"
        r")\b",
        flags=re.I,
    )
    headers = {"User-Agent": "Mozilla/5.0"}
    vistos = set()
    sessao = requests.Session()
    sessao.headers.update(headers)

    def limpar_texto(elemento):
        return re.sub(r"\s+", " ", elemento.get_text(" ", strip=True)).strip()

    def extrair_prazo(texto):
        m = re.search(r"Application Deadline:\s*(.*?)(?:\s+Work Locations:|$)", texto, flags=re.I)
        return m.group(1).strip() if m else "Consultar edital"

    def pagina_filtrada(id_pais, tipo):
        inicial = sessao.get(url_base, timeout=20)
        inicial.raise_for_status()
        soup_inicial = BeautifulSoup(inicial.text, "html.parser")
        form = soup_inicial.find("form", {"id": "oe-list-pages-facets-form"})
        dados = {}
        if form:
            for inp in form.find_all("input"):
                nome = inp.get("name")
                if nome:
                    dados[nome] = inp.get("value", "")
        dados["job_country[]"] = id_pais
        dados["offer_type[]"] = tipo
        resposta = sessao.post(url_base, data=dados, timeout=20)
        resposta.raise_for_status()
        return BeautifulSoup(resposta.text, "html.parser")

    for pais, id_pais in paises.items():
        for tipo, rotulo_tipo in tipos.items():
            try:
                soup = pagina_filtrada(id_pais, tipo)
            except Exception as e:
                log.warning(f"EURAXESS {pais} erro: {e}")
                continue

            for item in soup.find_all("article"):
                a = item.select_one('h3 a[href*="/jobs/"]')
                if not a:
                    continue

                titulo = limpar_texto(a)
                texto = limpar_texto(item)
                termo = termos.search(f"{titulo} {texto}")
                if not titulo or not termo:
                    continue

                link = urljoin("https://euraxess.ec.europa.eu", a.get("href", ""))
                chave = f"{pais.lower()}::{tipo}::{link.rstrip('/').lower()}"
                if chave in vistos:
                    continue

                bolsas.append({
                    "titulo": titulo,
                    "fonte": "EURAXESS",
                    "pais": pais,
                    "area": f"{rotulo_tipo} - {termo.group(1).title()}",
                    "nivel": detectar_nivel(f"{titulo.lower()} {texto.lower()}"),
                    "prazo": extrair_prazo(texto),
                    "link": link,
                    "data_coleta": str(date.today()),
                    "ativa": "Sim"
                })
                vistos.add(chave)

    log.info(f"EURAXESS Italia/Franca/Espanha: {len(bolsas)} oportunidades filtradas")
    return bolsas


def coletar_bolsas():
    bolsas = []

    bolsas += extrair_links_filtrados(
        "https://www.gov.br/cnpq/pt-br/acesso-a-informacao/bolsas-e-auxilios",
        "CNPq",
        "Brasil"
    )

    bolsas += extrair_links_filtrados(
        "https://www.gov.br/capes/pt-br/acesso-a-informacao/acoes-e-programas/bolsas",
        "CAPES",
        "Brasil"
    )

    bolsas += extrair_links_filtrados(
        "https://fapesp.br/oportunidades",
        "FAPESP",
        "Brasil"
    )

    bolsas += extrair_links_filtrados(
        "https://fulbright.org.br/bolsas/",
        "Fulbright",
        "Internacional"
    )

    bolsas += extrair_links_filtrados(
        "https://www.daad.de/en/study-and-research-in-germany/scholarships/",
        "DAAD",
        "Alemanha"
    )

    bolsas += coletar_euraxess_it_fr_es()

    # Horizon (manual — melhor abordagem)
    bolsas.append({
        "titulo": "Horizon Europe – Funding & Tenders Portal",
        "fonte": "Horizon Europe",
        "pais": "Europa",
        "area": "Pesquisa",
        "nivel": "Doutorado/Pós-doc",
        "prazo": "Variável",
        "link": "https://ec.europa.eu/info/funding-tenders/opportunities/portal/",
        "data_coleta": str(date.today()),
        "ativa": "Sim"
    })

    log.info(f"{len(bolsas)} bolsas coletadas (filtradas)")
    return bolsas

# ─── EXCEL (ZERADO SEMPRE) ─────────────────────────────

def salvar_excel(bolsas):
    wb = openpyxl.Workbook()
    ws = wb.active

    ws.append([
        "titulo","fonte","pais","area",
        "nivel","prazo","link","data_coleta","ativa"
    ])

    for b in bolsas:
        ws.append(list(b.values()))

    wb.save(ARQUIVO_EXCEL)

    log.info("Excel recriado com dados novos")

# ─── HTML ──────────────────────────────────────────────

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
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Bolsas</title>

<style>
body{{font-family:Arial;max-width:1100px;margin:auto;padding:20px}}
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

    with open(os.path.join(PASTA_SAIDA, "index.html"), "w", encoding="utf-8") as f:
        f.write(html)

    log.info("HTML atualizado")

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

    bolsas = coletar_bolsas()

    salvar_excel(bolsas)

    gerar_html()

    publicar()

    log.info("Processo finalizado")

if __name__ == "__main__":
    main()
