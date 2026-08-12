from __future__ import annotations

import hashlib
import json
import math
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import unicodedata
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path
from typing import Any
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
import streamlit as st

APP_NAME = "Oráculo Cultural"
BASE_DIR = Path(__file__).resolve().parent
SCRAPERS_DIR = BASE_DIR / "scrapers"
DATA_DIR = BASE_DIR / "data"
COLETAS_DIR = BASE_DIR / "coletas"
UPLOAD_DIR = BASE_DIR / "imagens_upload"
DB_PATH = DATA_DIR / "oraculo_cultural.db"
USD_CACHE_PATH = DATA_DIR / "usd_brl_cache.json"
SEM_INFO = "Sem informação"

DATA_DIR.mkdir(exist_ok=True)
COLETAS_DIR.mkdir(exist_ok=True)
UPLOAD_DIR.mkdir(exist_ok=True)

FUSO_BRASILIA = ZoneInfo("America/Sao_Paulo")


def agora_brasilia() -> datetime:
    return datetime.now(FUSO_BRASILIA)


def agora_brasilia_iso() -> str:
    return agora_brasilia().isoformat(timespec="seconds")


def _converter_para_brasilia(valor: Any):
    if valor is None or limpar_texto(valor) == SEM_INFO:
        return None
    try:
        dt = pd.to_datetime(valor, errors="coerce")
        if pd.isna(dt):
            return None
        if getattr(dt, "tzinfo", None) is None:
            # Timestamps antigos do Streamlit Cloud foram gravados como UTC sem fuso.
            dt = dt.tz_localize("UTC")
        return dt.tz_convert("America/Sao_Paulo")
    except Exception:
        return None


def formatar_data_brasilia(valor: Any) -> str:
    dt = _converter_para_brasilia(valor)
    if dt is None:
        return SEM_INFO
    return dt.strftime("%d/%m/%Y %H:%M:%S")


def em_streamlit_cloud() -> bool:
    return bool(os.environ.get("STREAMLIT_SHARING") or os.environ.get("STREAMLIT_SERVER_PORT") or os.environ.get("HOSTNAME"))


def garantir_playwright_chromium() -> None:
    """Garante o Chromium do Playwright no deploy em nuvem.

    ArtSoul usa requests e funciona sem navegador. Blombo, Gagosian e Saatchi
    precisam do Chromium. Em Streamlit Cloud, o browser pode não existir depois
    do build, então instalamos uma vez e criamos um marcador local.
    """
    flag = BASE_DIR / ".playwright_chromium_instalado"
    if flag.exists():
        return
    try:
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=300,
            check=False,
        )
        flag.write_text(datetime.now().isoformat(timespec="seconds"), encoding="utf-8")
    except Exception:
        pass


# Executa no início para o deploy já nascer preparado. Se falhar, o erro real
# ainda aparecerá no log de coleta do scraper correspondente.
garantir_playwright_chromium()

SITES = {
    "ArtSoul": {
        "script": "artsoulWS_sem_miniaturas.py",
        "xlsx": "artes_ArtSoul.xlsx",
        "header": 0,
        "runner": "artsoul_cli",
        "internacional": False,
    },
    "Blombo": {
        "script": "blomboWS_sem_miniaturas.py",
        "xlsx": "artes_BLOMBO.xlsx",
        "header": 1,
        "runner": "legacy_playwright",
        "internacional": False,
    },
    "Gagosian": {
        "script": "gagosianWS_sem_miniaturas.py",
        "xlsx": "artes_Gagosian.xlsx",
        "header": 1,
        "runner": "legacy_playwright",
        "internacional": True,
    },
    "Saatchi Art": {
        "script": "saatchiWS_sem_miniaturas.py",
        "xlsx": "artes_SAATCHIART.xlsx",
        "header": 1,
        "runner": "saatchi_cli",
        "internacional": True,
    },
    # ── Artistas específicos ────────────────────────────────────────────────
    "Cícero Dias (Blombo)": {
        "script": "cicero_dias_blomboWS.py",
        "xlsx": "artes_CiceroDias_Blombo.xlsx",
        "header": 0,
        "runner": "playwright_cli",
        "internacional": False,
    },
    "Milan Dusek (Arte Galeria)": {
        "script": "milan_dusek_artegalariaWS.py",
        "xlsx": "artes_MilanDusek_ArteGaleria.xlsx",
        "header": 0,
        "runner": "artsoul_cli",
        "internacional": False,
    },
    "Rubem Grilo (Arremate Arte)": {
        "script": "rubem_grilo_arremateWS.py",
        "xlsx": "artes_RubemGrilo_Arremate.xlsx",
        "header": 0,
        "runner": "playwright_cli",
        "internacional": False,
    },

    "Leilões BR": {
        "script": "leiloesbrWS_sem_miniaturas.py",
        "xlsx": "artes_LeiloesBR.xlsx",
        "header": 0,
        "runner": "playwright_cli",
        "internacional": False,
    },
}

TECNICAS_PADRONIZADAS = [
    "Pintura",
    "Pintura acrílica",
    "Pintura a óleo",
    "Aquarela",
    "Guache",
    "Têmpera",
    "Encáustica",
    "Afresco",
    "Técnica mista",
    "Desenho",
    "Grafite",
    "Carvão",
    "Pastel seco",
    "Pastel oleoso",
    "Lápis de cor",
    "Nanquim / tinta",
    "Caneta / marcador",
    "Colagem",
    "Assemblage",
    "Serigrafia",
    "Litografia",
    "Xilogravura",
    "Linogravura",
    "Gravura em metal",
    "Água-forte",
    "Água-tinta",
    "Ponta-seca",
    "Monotipia",
    "Gravura",
    "Impressão fine art",
    "Giclée",
    "Impressão digital",
    "Fotografia",
    "Arte digital",
    "Arte generativa",
    "Spray / aerosol",
    "Escultura",
    "Cerâmica",
    "Porcelana",
    "Vidro",
    "Metal",
    "Bronze",
    "Madeira",
    "Mármore / pedra",
    "Resina",
    "Têxtil",
    "Tapeçaria",
    "Bordado",
    "Instalação",
    "Objeto",
    "Performance / vídeo",
    "Outros",
    "Sem informação",
]
UI_TO_DB = {
    "Nome da obra": "nome_obra",
    "Título": "nome_obra",
    "Autor": "autor",
    "Preço": "preco",
    "Preço BRL": "preco",
    "Dimensões": "dimensoes",
    "Técnica": "tecnica",
    "Técnica original": "tecnica_original",
    "Ano da obra": "ano_obra",
    "Ano": "ano_obra",
    "Descrição": "descricao",
    "Link da obra": "link_obra",
    "Link da imagem da obra": "link_imagem",
}


def limpar_texto(valor: Any) -> str:
    if valor is None:
        return SEM_INFO
    try:
        if isinstance(valor, float) and math.isnan(valor):
            return SEM_INFO
    except Exception:
        pass
    texto = str(valor).replace("\xa0", " ").replace("\ufeff", " ").strip()
    texto = re.sub(r"\s+", " ", texto)
    if not texto or texto.lower() in {"nan", "none", "null", "false", "true"}:
        return SEM_INFO
    return texto


def sem_info(valor: Any) -> bool:
    return limpar_texto(valor) == SEM_INFO


def chave_coluna(col: str) -> str:
    base = str(col).strip().lower()
    base = "".join(ch for ch in unicodedata.normalize("NFKD", base) if not unicodedata.combining(ch))
    base = re.sub(r"[^a-z0-9]+", "_", base).strip("_")
    return base


def parse_preco(valor: Any) -> float | None:
    if valor is None:
        return None
    if isinstance(valor, (int, float)) and not pd.isna(valor):
        return float(valor)
    texto = limpar_texto(valor)
    if texto == SEM_INFO:
        return None
    texto = re.sub(r"[^\d,.-]", "", texto)
    if not texto:
        return None
    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "." in texto and re.fullmatch(r"\d{1,3}(?:\.\d{3})+", texto):
        texto = texto.replace(".", "")
    try:
        return float(texto)
    except Exception:
        return None


def normalizar_ano(valor: Any) -> str:
    texto = limpar_texto(valor)
    if texto == SEM_INFO:
        return SEM_INFO
    m = re.search(r"\b(18\d{2}|19\d{2}|20\d{2})\b", texto)
    if not m:
        return SEM_INFO
    ano = int(m.group(1))
    if 1800 <= ano <= datetime.now().year + 2:
        return str(ano)
    return SEM_INFO


def padronizar_dimensoes(valor: Any) -> str:
    texto = limpar_texto(valor)
    if texto == SEM_INFO:
        return SEM_INFO
    nums = re.findall(r"\d+(?:[.,]\d+)?", texto)
    if len(nums) < 2:
        return SEM_INFO
    floats = []
    for n in nums[:3]:
        try:
            floats.append(float(n.replace(",", ".")))
        except Exception:
            pass
    if len(floats) < 2:
        return SEM_INFO
    if any(n == 0 for n in floats) or any(1500 <= n <= datetime.now().year + 2 for n in floats):
        return SEM_INFO
    def fmt(n: float) -> str:
        return str(int(n)) if n.is_integer() else f"{n:.2f}".rstrip("0").rstrip(".")
    return " x ".join(fmt(n) for n in floats) + " cm"


def normalizar_tecnica(valor: Any) -> str:
    texto = limpar_texto(valor)
    if texto == SEM_INFO:
        return SEM_INFO
    t = texto.lower()
    t = "".join(ch for ch in unicodedata.normalize("NFKD", t) if not unicodedata.combining(ch))

    # Técnicas combinadas ou descrições com múltiplos materiais artísticos.
    grupos = {
        "oleo": ["oil", "oleo"],
        "acrilica": ["acrylic", "acrilica", "acrilico"],
        "aquarela": ["watercolor", "watercolour", "aquarela"],
        "guache": ["gouache", "guache"],
        "pastel": ["pastel"],
        "nanquim": ["india ink", "nanquim", "ink", "tinta"],
        "grafite": ["graphite", "grafite", "lapis", "pencil"],
        "carvao": ["charcoal", "carvao"],
        "colagem": ["collage", "colagem"],
        "spray": ["spray", "aerosol", "aerossol"],
    }
    materiais = [nome for nome, termos in grupos.items() if any(term in t for term in termos)]
    if any(x in t for x in ["mixed media", "mixed technique", "tecnica mista", "tecnicas mistas", "mista", "mixed"]):
        return "Técnica mista"
    if len(set(materiais)) >= 2:
        return "Técnica mista"

    # Gravuras e impressões: termos específicos antes de termos genéricos.
    if any(x in t for x in ["woodcut", "xilogravura", "xilografia", "woodblock"]):
        return "Xilogravura"
    if any(x in t for x in ["linocut", "linogravura", "linoleogravura"]):
        return "Linogravura"
    if any(x in t for x in ["etching", "agua-forte", "agua forte", "aguaforte", "acid etching"]):
        return "Água-forte"
    if any(x in t for x in ["aquatint", "agua-tinta", "agua tinta", "aguatinta"]):
        return "Água-tinta"
    if any(x in t for x in ["drypoint", "ponta-seca", "ponta seca", "pontaseca"]):
        return "Ponta-seca"
    if any(x in t for x in ["monotype", "monoprint", "monotipia"]):
        return "Monotipia"
    if any(x in t for x in ["engraving", "intaglio", "calcogravura", "calcografia", "gravura em metal", "gravura sobre metal", "gravura metal", "metal engraving", "metalgravura", "burin", "buril"]):
        return "Gravura em metal"
    if any(x in t for x in ["screenprint", "screen print", "serigraph", "silkscreen", "serigrafia"]):
        return "Serigrafia"
    if any(x in t for x in ["lithograph", "lithography", "litografia"]):
        return "Litografia"
    if any(x in t for x in ["giclee", "giclée"]):
        return "Giclée"
    if any(x in t for x in ["fine art print", "archival print", "print on paper"]):
        return "Impressão fine art"
    if any(x in t for x in ["digital print", "impressao digital", "impressão digital"]):
        return "Impressão digital"
    if any(x in t for x in ["offset", "off-set"]):
        return "Impressão digital"
    if any(x in t for x in ["printmaking", "gravura", "print"]):
        return "Gravura"

    # Pintura e desenho.
    if any(x in t for x in ["acrylic", "acrilica", "acrilico"]):
        return "Pintura acrílica"
    if any(x in t for x in ["oil", "oleo"]):
        return "Pintura a óleo"
    if any(x in t for x in ["watercolor", "watercolour", "aquarela"]):
        return "Aquarela"
    if any(x in t for x in ["gouache", "guache"]):
        return "Guache"
    if any(x in t for x in ["tempera", "têmpera"]):
        return "Têmpera"
    if any(x in t for x in ["encaustic", "encaustica", "encáustica"]):
        return "Encáustica"
    if any(x in t for x in ["fresco", "afresco"]):
        return "Afresco"
    if any(x in t for x in ["pastel oil", "oil pastel", "pastel oleoso"]):
        return "Pastel oleoso"
    if any(x in t for x in ["soft pastel", "dry pastel", "pastel seco", "pastel"]):
        return "Pastel seco"
    if any(x in t for x in ["graphite", "grafite"]):
        return "Grafite"
    if any(x in t for x in ["charcoal", "carvao", "carvão"]):
        return "Carvão"
    if any(x in t for x in ["colored pencil", "colour pencil", "lapis de cor", "lápis de cor"]):
        return "Lápis de cor"
    if any(x in t for x in ["india ink", "nanquim", "ink"]):
        return "Nanquim / tinta"
    if any(x in t for x in ["marker", "marcador", "caneta", "pen"]):
        return "Caneta / marcador"
    if any(x in t for x in ["drawing", "desenho"]):
        return "Desenho"
    if any(x in t for x in ["painting", "pintura"]):
        return "Pintura"

    # Outras categorias artísticas e materiais.
    if any(x in t for x in ["collage", "colagem"]):
        return "Colagem"
    if any(x in t for x in ["assemblage", "assemblagem"]):
        return "Assemblage"
    if any(x in t for x in ["photography", "photograph", "fotografia", "photo"]):
        return "Fotografia"
    if any(x in t for x in ["generative", "arte generativa", "ai art"]):
        return "Arte generativa"
    if any(x in t for x in ["digital", "new media", "arte digital"]):
        return "Arte digital"
    if any(x in t for x in ["spray", "aerosol", "aerossol"]):
        return "Spray / aerosol"
    if any(x in t for x in ["ceramic", "ceramics", "ceramica"]):
        return "Cerâmica"
    if any(x in t for x in ["porcelain", "porcelana"]):
        return "Porcelana"
    if any(x in t for x in ["glass", "vidro"]):
        return "Vidro"
    if any(x in t for x in ["bronze"]):
        return "Bronze"
    if any(x in t for x in ["metal", "steel", "iron", "aluminum", "aluminium", "copper", "brass"]):
        return "Metal"
    if any(x in t for x in ["wood", "madeira"]):
        return "Madeira"
    if any(x in t for x in ["marble", "stone", "marmore", "mármore", "pedra"]):
        return "Mármore / pedra"
    if any(x in t for x in ["resin", "resina"]):
        return "Resina"
    if any(x in t for x in ["tapestry", "tapecaria", "tapeçaria"]):
        return "Tapeçaria"
    if any(x in t for x in ["embroidery", "bordado"]):
        return "Bordado"
    if any(x in t for x in ["textile", "fabric", "fiber", "fibre", "tecido", "textil", "têxtil"]):
        return "Têxtil"
    if any(x in t for x in ["installation", "instalacao", "instalação"]):
        return "Instalação"
    if any(x in t for x in ["object", "objeto"]):
        return "Objeto"
    if any(x in t for x in ["performance", "video", "vídeo"]):
        return "Performance / vídeo"
    if any(x in t for x in ["sculpture", "escultura"]):
        return "Escultura"

    return "Outros"

def nome_valido(valor: Any) -> bool:
    texto = limpar_texto(valor)
    if texto == SEM_INFO:
        return False
    low = texto.lower()
    if any(x in low for x in [".com", "r$", "us$", "$", "false", "true", "saatchi art", "artsoul"]):
        return False
    return len(texto) >= 2


def corrigir_nome_obra(nome: Any, link: Any = None) -> str:
    texto = limpar_texto(nome)
    if nome_valido(texto):
        return texto
    link = limpar_texto(link)
    if link != SEM_INFO:
        partes = [p for p in re.split(r"[/#?]", link) if p]
        for p in reversed(partes):
            if p.lower() in {"view", "print", "obras", "art", "paintings"} or p.isdigit():
                continue
            p = re.sub(r"^(Painting|Photography|Drawing|Sculpture|Print|Mixed-Media)-", "", p, flags=re.I)
            p = p.replace("-", " ").replace("_", " ").strip()
            if nome_valido(p):
                return p[:1].upper() + p[1:]
    return SEM_INFO


def corrigir_autor(autor: Any, nome_obra: Any = None) -> str:
    texto = limpar_texto(autor)
    if texto == SEM_INFO:
        return SEM_INFO
    texto = re.split(r"\b(?:Date|Data|Ano|Producer|Medium|Dimensions|Preço|Price)\s*:", texto, maxsplit=1, flags=re.I)[0]
    texto = re.sub(r",\s*(?:Brazil|Brasil|United States|USA|Ukraine|France|Italy|Spain|Portugal|Germany|United Kingdom|Canada|Japan|China|Australia)\.?$", "", texto, flags=re.I)
    texto = limpar_texto(texto)
    low = texto.lower()
    if low in {"artist", "artista", "saatchi art", "false", "true"} or ".com" in low:
        return SEM_INFO
    if limpar_texto(nome_obra).lower() == low:
        return SEM_INFO
    if len(texto) > 90:
        return SEM_INFO
    return texto


def texto_busca(valor: Any) -> str:
    """Texto sem acento e sem pontuação para comparação robusta."""
    texto = limpar_texto(valor)
    if texto == SEM_INFO:
        return ""
    texto = "".join(ch for ch in unicodedata.normalize("NFKD", texto.lower()) if not unicodedata.combining(ch))
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


ARTISTAS_ESPECIFICOS = [
    {
        "autor": "Cícero Dias",
        "nomes": ["cícero dias", "cicero dias"],
        "sites": ["blombo"],
        "ruidos": ["blombo"],
        "tecnica_padrao": None,
    },
    {
        "autor": "Milan Dusek",
        "nomes": ["milan dusek"],
        "sites": ["arte galeria", "artegaleria"],
        "ruidos": ["arte galeria", "artegaleria"],
        "tecnica_padrao": "Gravura em metal",
    },
    {
        "autor": "Rubem Grilo",
        "nomes": ["rubem grilo"],
        "sites": ["arremate arte", "arremate"],
        "ruidos": ["arremate arte", "arremate"],
        # As páginas desse recorte são de gravuras; se o site não entregar
        # o campo técnico, evita gravar "Sem informação" no acervo.
        "tecnica_padrao": "Gravura",
    },
]


def contexto_linha(row: pd.Series) -> str:
    partes: list[str] = []
    for col, valor in row.items():
        texto = limpar_texto(valor)
        if texto != SEM_INFO:
            partes.append(f"{col}: {texto}")
    return "\n".join(partes)


def link_http(valor: Any) -> bool:
    texto = limpar_texto(valor)
    return texto.startswith(("http://", "https://"))


@st.cache_data(show_spinner=False, ttl=60 * 60 * 12)
def obter_texto_pagina_obra(link: str) -> str:
    """Busca texto visível da página da obra para completar campos ausentes.

    Usado apenas como reforço para os scrapers específicos. Se a página falhar,
    a limpeza continua com os dados da planilha.
    """
    if not link_http(link):
        return ""
    try:
        resp = requests.get(
            link,
            timeout=12,
            headers={"User-Agent": "Mozilla/5.0 (compatible; OraculoCultural/1.0)"},
        )
        resp.raise_for_status()
        ctype = resp.headers.get("content-type", "")
        if "html" not in ctype.lower() and "text" not in ctype.lower():
            return ""
        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup(["script", "style", "noscript", "svg"]):
            tag.decompose()
        extras: list[str] = []
        if soup.title and soup.title.get_text(strip=True):
            extras.append(f"Título da página: {soup.title.get_text(' ', strip=True)}")
        for meta in soup.find_all("meta"):
            nome = (meta.get("property") or meta.get("name") or "").lower()
            if nome in {"og:title", "twitter:title", "description", "og:description", "twitter:description"}:
                conteudo = meta.get("content")
                if conteudo:
                    extras.append(conteudo)
        texto = soup.get_text("\n", strip=True)
        linhas = [limpar_texto(x) for x in ("\n".join(extras) + "\n" + texto).splitlines()]
        linhas = [x for x in linhas if x != SEM_INFO]
        return "\n".join(linhas)[:30000]
    except Exception:
        return ""


def detectar_artista_especifico(site: str | None, autor: Any, contexto: Any) -> dict[str, Any] | None:
    base = texto_busca(" ".join([limpar_texto(site), limpar_texto(autor), limpar_texto(contexto)]))
    for cfg in ARTISTAS_ESPECIFICOS:
        nomes = [texto_busca(n) for n in cfg["nomes"]]
        sites = [texto_busca(s) for s in cfg["sites"]]
        if any(n and n in base for n in nomes) and (not sites or any(s and s in base for s in sites)):
            return cfg
        # Quando o usuário seleciona o scraper específico, o nome do site no app
        # já é uma evidência suficiente para aplicar a limpeza daquele artista.
        site_norm = texto_busca(site)
        if any(n and n in site_norm for n in nomes):
            return cfg
    return None


ROTULOS_CAMPOS = r"(?:nome(?:\s+da\s+obra)?|t[ií]tulo|titulo|autor|artista|pre[cç]o|valor|t[eé]cnica|tecnica|medium|m[eé]dia|material|dimens[õo]es|dimensoes|medidas|ano|data|descri[cç][aã]o|link|imagem)"


def extrair_campo_por_rotulo(texto: Any, rotulos: list[str], limite: int = 180) -> str:
    base = limpar_texto(texto)
    if base == SEM_INFO:
        return SEM_INFO
    rot = "|".join(rotulos)
    padrao = re.compile(
        rf"(?:^|[\n\r|;•])\s*(?:{rot})\s*[:\-–—]\s*(.{{1,{limite}}}?)(?=\s*(?:[\n\r|;•]|{ROTULOS_CAMPOS}\s*[:\-–—])|$)",
        flags=re.I,
    )
    m = padrao.search(base)
    if not m:
        return SEM_INFO
    valor = re.sub(r"\s+", " ", m.group(1)).strip(" -–—|;•")
    return limpar_texto(valor)


def extrair_dimensoes_contexto(texto: Any) -> str:
    base = limpar_texto(texto)
    if base == SEM_INFO:
        return SEM_INFO
    rotulada = extrair_campo_por_rotulo(base, [r"dimens[õo]es", r"dimensoes", r"medidas", r"dimensions", r"size"], limite=90)
    for fonte in [rotulada, base]:
        if limpar_texto(fonte) == SEM_INFO:
            continue
        m = re.search(
            r"(\d+(?:[,.]\d+)?)\s*(?:x|×|por)\s*(\d+(?:[,.]\d+)?)(?:\s*(?:x|×|por)\s*(\d+(?:[,.]\d+)?))?\s*(?:cm|cent[ií]metros?)?",
            str(fonte),
            flags=re.I,
        )
        if m:
            partes = [g for g in m.groups() if g]
            return padronizar_dimensoes(" x ".join(partes))
    return SEM_INFO


def extrair_ano_contexto(texto: Any) -> str:
    base = limpar_texto(texto)
    if base == SEM_INFO:
        return SEM_INFO
    rotulado = extrair_campo_por_rotulo(base, [r"ano", r"data", r"year", r"date"], limite=60)
    ano = normalizar_ano(rotulado)
    if ano != SEM_INFO:
        return ano
    for padrao in [
        r"\b(?:datad[oa]|dated|assinado\s+e\s+datado|ass\.?[^.]{0,35}datad[oa]|c\.|circa)\D{0,25}(18\d{2}|19\d{2}|20\d{2})\b",
        r"\b(18\d{2}|19\d{2}|20\d{2})\b\s*(?:\([^)]*\))?\s*(?:ass\.|datad[oa]|dated)",
    ]:
        m = re.search(padrao, base, flags=re.I)
        if m:
            return normalizar_ano(m.group(1))
    return SEM_INFO


def tecnica_util(valor: Any) -> bool:
    return limpar_texto(valor) not in {SEM_INFO, "Outros"}


def resolver_tecnica(row: pd.Series, contexto: str, cfg: dict[str, Any] | None) -> tuple[str, str]:
    rotulada = extrair_campo_por_rotulo(
        contexto,
        [r"t[eé]cnica", r"tecnica", r"medium", r"m[eé]dia", r"material(?:\s+e\s+t[eé]cnica)?", r"materials?"],
        limite=140,
    )
    fontes = [
        row.get("Técnica original"),
        row.get("Técnica"),
        rotulada,
        row.get("Nome da obra"),
        row.get("Descrição"),
        row.get("Link da obra"),
        contexto,
    ]
    for fonte in fontes:
        texto = limpar_texto(fonte)
        if texto == SEM_INFO:
            continue
        tecnica = normalizar_tecnica(texto)
        if tecnica_util(tecnica):
            # Se a técnica veio de título/link contaminado, salva a forma limpa
            # como técnica original para não levar o título inteiro para o banco.
            original = texto if fonte in [row.get("Técnica original"), row.get("Técnica"), rotulada] else tecnica
            return limpar_texto(original), tecnica
    if cfg and cfg.get("tecnica_padrao"):
        return cfg["tecnica_padrao"], cfg["tecnica_padrao"]
    return SEM_INFO, SEM_INFO


def parte_parece_tecnica(texto: Any) -> bool:
    parte = limpar_texto(texto)
    if parte == SEM_INFO:
        return False
    return tecnica_util(normalizar_tecnica(parte)) and len(texto_busca(parte).split()) <= 7


def cfg_eh_cicero_blombo(cfg: dict[str, Any] | None) -> bool:
    return bool(cfg and texto_busca(cfg.get("autor")) == "cicero dias")


def titulo_biografico_cicero(texto: Any) -> bool:
    """Detecta o título famoso citado na biografia do artista, não no produto.

    Nas páginas da Blombo, a biografia de Cícero Dias repete "Eu Vi o Mundo...
    Ele Começava no Recife". Esse texto estava sendo capturado como título
    de todas as obras. Para os itens vendidos no site, o título correto deve vir
    do produto/link, não desse trecho biográfico.
    """
    norm = texto_busca(texto)
    if not norm:
        return False
    return (
        "eu vi o mundo" in norm and "recife" in norm
    ) or "guernica brasileira" in norm


def limpar_pontuacao_titulo(texto: Any) -> str:
    t = limpar_texto(texto)
    if t == SEM_INFO:
        return SEM_INFO
    t = re.sub(r"\s+", " ", t)
    t = re.sub(r"^[\s'\"“”‘’]+|[\s'\"“”‘’.]+$", "", t)
    t = re.sub(r"\s+([,.;:])", r"\1", t)
    return limpar_texto(t)


def titulo_invalido_ou_contaminado(texto: Any, cfg: dict[str, Any] | None, autor: Any, site: Any) -> bool:
    t = limpar_pontuacao_titulo(texto)
    if t == SEM_INFO:
        return True
    norm = texto_busca(t)
    if not norm:
        return True
    if cfg_eh_cicero_blombo(cfg) and titulo_biografico_cicero(t):
        return True
    if any(x in norm for x in ["http", "https", "www", "com br", "r ", "preco", "price", "dimensoes", "dimensions", "tecnica", "medium"]):
        return True
    if norm in {"http", "https", "www", "com", "br", "obra", "produto", "sem", "sem informacao", "sem titulo", "untitled", "artista", "artist"}:
        return True
    if parte_parece_tecnica(t):
        return True
    tokens_site = ["arte galeria", "artegaleria", "arremate arte", "arremate", "blombo", "artsoul", "saatchi art", "gagosian"]
    if any(texto_busca(s) == norm or texto_busca(s) in norm for s in tokens_site) and len(norm.split()) <= 4:
        return True
    autor_norm = texto_busca(autor)
    if autor_norm and (norm == autor_norm or (autor_norm in norm and len(norm.split()) <= len(autor_norm.split()) + 3)):
        return True
    if cfg:
        nomes = [texto_busca(n) for n in cfg["nomes"]]
        ruidos = [texto_busca(r) for r in cfg.get("ruidos", [])]
        if any(n and (norm == n or (n in norm and len(norm.split()) <= len(n.split()) + 3)) for n in nomes):
            return True
        if any(r and (norm == r or (r in norm and len(norm.split()) <= len(r.split()) + 2)) for r in ruidos):
            return True
    return not nome_valido(t)


def remover_ruidos_titulo(parte: str, cfg: dict[str, Any] | None, autor: Any, site: Any) -> str:
    t = limpar_pontuacao_titulo(parte)
    if t == SEM_INFO:
        return SEM_INFO
    termos = [limpar_texto(autor), limpar_texto(site), "Arte Galeria", "ArteGaleria", "Arremate Arte", "Arremate", "Blombo"]
    if cfg:
        termos.extend(cfg.get("nomes", []))
        termos.extend(cfg.get("ruidos", []))
    for termo in sorted({x for x in termos if limpar_texto(x) != SEM_INFO}, key=len, reverse=True):
        t = re.sub(rf"\b{re.escape(termo)}\b", " ", t, flags=re.I)
    t = re.sub(r"^\s*(?:comprar\s+)?obra\s+arte\s+", "", t, flags=re.I)
    t = re.sub(r"\s+(?:original\s+do\s+artista|obra\s+de|obra\s+do\s+artista)\b.*$", "", t, flags=re.I)
    t = re.sub(r"\s+obra\s*$", "", t, flags=re.I)
    t = re.sub(r"\b(?:t[eé]cnica|ano|id\s+blombo)\b.*$", "", t, flags=re.I)
    t = re.sub(r"\b(?:por|by)\b\s*$", "", t, flags=re.I)
    return limpar_pontuacao_titulo(t)


def titulo_a_partir_link(link: Any, cfg: dict[str, Any] | None, autor: Any, site: Any) -> str:
    link_txt = limpar_texto(link)
    if link_txt == SEM_INFO:
        return SEM_INFO
    partes = [p for p in re.split(r"[/#?]", link_txt) if p]
    for p in reversed(partes):
        if texto_busca(p) in {"http", "https", "www"} or "." in p:
            continue
        p = re.sub(r"\.[a-z0-9]{2,5}$", "", p, flags=re.I)
        p = re.sub(r"^(Painting|Photography|Drawing|Sculpture|Print|Mixed-Media)-", "", p, flags=re.I)
        p = p.replace("-", " ").replace("_", " ")
        p = remover_ruidos_titulo(p, cfg, autor, site)
        if not titulo_invalido_ou_contaminado(p, cfg, autor, site):
            return p[:1].upper() + p[1:]
    return SEM_INFO


def extrair_titulo_blombo_cicero(texto: Any, cfg: dict[str, Any] | None, autor: Any, site: Any) -> str:
    """Extrai o título real do produto Cícero Dias na Blombo.

    A Blombo coloca uma biografia longa do artista na página. Essa biografia
    contém a obra histórica "Eu Vi o Mundo... Ele Começava no Recife", que não
    deve ser usada como título dos demais produtos.
    """
    if not cfg_eh_cicero_blombo(cfg):
        return SEM_INFO
    base = limpar_texto(texto)
    if base == SEM_INFO:
        return SEM_INFO

    padroes = [
        r"comprar\s+obra\s+arte\s+(.{2,180}?)\s+original\s+do\s+artista\s+c[ií]cero\s+dias",
        r"image\s*:\s*(.{2,160}?)(?=\s*(?:image\s*:|c[ií]cero\s+dias|r\$|comprar|quantidade|eu\s+quero|fale\s+com|\*|$))",
        r"obras\s+gravuras\s+e\s+m[uú]ltiplos\s+(.{2,160}?)(?=\s*image\s*:)",
    ]
    for padrao in padroes:
        for m in re.finditer(padrao, base, flags=re.I):
            cand = remover_ruidos_titulo(m.group(1), cfg, autor, site)
            if not titulo_invalido_ou_contaminado(cand, cfg, autor, site):
                return cand

    linhas = [limpar_texto(x) for x in str(base).splitlines()]
    linhas = [x for x in linhas if x != SEM_INFO]
    for i, linha in enumerate(linhas):
        if texto_busca(linha) == "cicero dias":
            for cand in linhas[i + 1 : i + 8]:
                if re.search(r"^(?:r\$|quantidade|eu quero|fale com|comprar|image\s*:)", cand, flags=re.I):
                    continue
                if normalizar_tecnica(cand) != "Outros" or padronizar_dimensoes(cand) != SEM_INFO or normalizar_ano(cand) != SEM_INFO:
                    continue
                cand = remover_ruidos_titulo(cand, cfg, autor, site)
                if not titulo_invalido_ou_contaminado(cand, cfg, autor, site):
                    return cand
    return SEM_INFO


def extrair_titulo_de_texto(texto: Any, cfg: dict[str, Any] | None, autor: Any, site: Any) -> str:
    base = limpar_texto(texto)
    if base == SEM_INFO:
        return SEM_INFO

    titulo_cicero = extrair_titulo_blombo_cicero(base, cfg, autor, site)
    if titulo_cicero != SEM_INFO:
        return titulo_cicero

    # Para Cícero Dias/Blombo, não usa títulos entre aspas nem metatítulos
    # genéricos da biografia, pois eles repetem o painel histórico do artista.
    if cfg_eh_cicero_blombo(cfg):
        return SEM_INFO

    # Primeiro busca títulos entre aspas, padrão muito comum em páginas de leilão.
    for padrao in [r"[“\"]([^“”\"]{2,140})[”\"]", r"[‘']([^‘’']{2,140})[’']"]:
        for m in re.finditer(padrao, base):
            cand = limpar_pontuacao_titulo(m.group(1))
            if not titulo_invalido_ou_contaminado(cand, cfg, autor, site):
                return cand

    # Depois tenta padrões de título em metadados/HTML.
    rotulado = extrair_campo_por_rotulo(base, [r"t[ií]tulo(?:\s+da\s+p[aá]gina)?", r"titulo", r"title"], limite=160)
    if rotulado != SEM_INFO:
        cand = remover_ruidos_titulo(rotulado, cfg, autor, site)
        if not titulo_invalido_ou_contaminado(cand, cfg, autor, site):
            return cand
    return SEM_INFO



def cfg_eh_rubem_arremate(cfg: dict[str, Any] | None) -> bool:
    return bool(cfg and texto_busca(cfg.get("autor")) == "rubem grilo")


TITULOS_NAVEGACAO_ARREMATE = {
    "contas",
    "leiloes",
    "lotes",
    "categorias",
    "casas",
    "casas de leilao",
    "blog",
    "login",
    "cadastrar",
    "home",
    "termo de qualidade",
    "regulamento",
    "termo de uso",
    "privacidade",
}


def imagem_logo_ou_generica(valor: Any) -> bool:
    texto = limpar_texto(valor)
    if texto == SEM_INFO:
        return True
    norm = texto_busca(texto)
    baixo = texto.lower()
    if baixo.endswith(".svg"):
        return True
    return any(
        ruido in norm
        for ruido in [
            "images logo",
            "footer logo",
            "logo svg",
            "arrematearte com br images logo",
            "arrematearte com br images footer logo",
            "placeholder",
            "sem imagem",
            "no image",
            "favicon",
            "bandeira",
            "flag",
        ]
    )


def registro_lixo_rubem_arremate(nome: Any, link: Any, imagem: Any, site: Any, autor: Any) -> bool:
    base = texto_busca(" ".join([limpar_texto(site), limpar_texto(autor)]))
    if "rubem grilo" not in base and "arremate" not in base:
        return False
    nome_norm = texto_busca(nome)
    link_norm = texto_busca(link)
    if nome_norm in TITULOS_NAVEGACAO_ARREMATE:
        return True
    if "lotes disponiveis" in nome_norm or "categorias" in nome_norm and "lote" in nome_norm:
        return True
    if "artistas rubem grilo" in link_norm and imagem_logo_ou_generica(imagem):
        return True
    return False


def normalizar_url_imagem(src: Any, base_url: Any) -> str:
    texto = limpar_texto(src)
    if texto == SEM_INFO:
        return SEM_INFO
    texto = texto.strip().strip('"\'')
    if texto.startswith("//"):
        texto = "https:" + texto
    elif texto.startswith(("/", "./", "../")):
        texto = urljoin(limpar_texto(base_url), texto)
    if not texto.startswith(("http://", "https://")):
        return SEM_INFO
    return texto


def urls_de_srcset(srcset: Any, base_url: Any) -> list[str]:
    texto = limpar_texto(srcset)
    if texto == SEM_INFO:
        return []
    urls: list[str] = []
    for parte in texto.split(","):
        src = parte.strip().split(" ")[0]
        url = normalizar_url_imagem(src, base_url)
        if url != SEM_INFO:
            urls.append(url)
    return urls


def imagem_candidata_util(url: Any, contexto: str = "") -> bool:
    u = normalizar_url_imagem(url, contexto or "https://www.arrematearte.com.br")
    if u == SEM_INFO or imagem_logo_ou_generica(u):
        return False
    baixo = u.lower()
    if any(x in baixo for x in ["/images/logo", "/footer/", "sprite", "icon", "avatar", "user.svg"]):
        return False
    if re.search(r"\.(?:jpg|jpeg|png|webp)(?:\?|$)", baixo):
        return True
    return any(x in baixo for x in ["cloudfront", "amazonaws", "storage", "uploads", "cdn", "lote", "lot", "image"])


def escolher_melhor_imagem(candidatos: list[dict[str, Any]], contexto: str = "") -> str:
    melhores: list[tuple[int, str]] = []
    ctx = texto_busca(contexto)
    for cand in candidatos:
        url = normalizar_url_imagem(cand.get("url"), cand.get("base") or contexto)
        if not imagem_candidata_util(url, contexto):
            continue
        alt = texto_busca(" ".join([limpar_texto(cand.get("alt")), limpar_texto(cand.get("text"))]))
        try:
            w = int(float(cand.get("w") or 0))
            h = int(float(cand.get("h") or 0))
        except Exception:
            w = h = 0
        score = 0
        if "rubem" in alt or "grilo" in alt or "rubem" in texto_busca(url):
            score += 40
        if any(x in alt for x in ["lote", "obra", "xilogravura", "gravura", "serigrafia", "nanquim"]):
            score += 25
        if any(x in texto_busca(url) for x in ["lote", "obra", "auction", "uploads", "storage", "cdn"]):
            score += 20
        if w >= 250 and h >= 250:
            score += 18
        elif w >= 120 and h >= 120:
            score += 8
        if ctx and alt and any(palavra in alt for palavra in ctx.split()[:8] if len(palavra) > 3):
            score += 10
        melhores.append((score, url))
    if not melhores:
        return SEM_INFO
    melhores.sort(reverse=True)
    return melhores[0][1]


@st.cache_data(show_spinner=False, ttl=60 * 60 * 12)
def extrair_imagem_estatica_pagina(link: str, contexto: str = "") -> str:
    if not link_http(link):
        return SEM_INFO
    try:
        resp = requests.get(
            link,
            timeout=12,
            headers={"User-Agent": "Mozilla/5.0 (compatible; OraculoCultural/1.0)"},
        )
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        candidatos: list[dict[str, Any]] = []
        for meta in soup.find_all("meta"):
            nome = (meta.get("property") or meta.get("name") or "").lower()
            if nome in {"og:image", "twitter:image", "og:image:secure_url"}:
                candidatos.append({"url": meta.get("content"), "alt": nome, "base": link})
        for tag in soup.find_all(["img", "source"]):
            texto_pai = ""
            try:
                pai = tag.find_parent(["article", "section", "li", "div", "a"])
                texto_pai = pai.get_text(" ", strip=True)[:500] if pai else ""
            except Exception:
                pass
            for attr in ["src", "data-src", "data-lazy", "data-original", "data-ng-src", "ng-src", "data-zoom-image"]:
                if tag.get(attr):
                    candidatos.append({
                        "url": tag.get(attr),
                        "alt": " ".join([tag.get("alt") or "", tag.get("title") or ""]),
                        "text": texto_pai,
                        "base": link,
                    })
            for attr in ["srcset", "data-srcset"]:
                for url in urls_de_srcset(tag.get(attr), link):
                    candidatos.append({"url": url, "alt": tag.get("alt") or "", "text": texto_pai, "base": link})
        for m in re.finditer(r"https?://[^\s'\"<>]+\.(?:jpg|jpeg|png|webp)(?:\?[^\s'\"<>]*)?", resp.text, flags=re.I):
            candidatos.append({"url": m.group(0), "base": link})
        return escolher_melhor_imagem(candidatos, contexto)
    except Exception:
        return SEM_INFO


@st.cache_data(show_spinner=False, ttl=60 * 60 * 12)
def extrair_imagem_renderizada_pagina(link: str, contexto: str = "") -> str:
    if not link_http(link):
        return SEM_INFO
    try:
        from playwright.sync_api import sync_playwright
    except Exception:
        return SEM_INFO
    try:
        garantir_playwright_chromium()
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page(
                viewport={"width": 1365, "height": 1000},
                user_agent="Mozilla/5.0 (compatible; OraculoCultural/1.0)",
            )
            page.goto(link, wait_until="domcontentloaded", timeout=45000)
            try:
                page.wait_for_load_state("networkidle", timeout=12000)
            except Exception:
                pass
            page.wait_for_timeout(2500)
            dados = page.evaluate(
                """
                () => {
                  const arr = [];
                  const push = (url, alt, text, w, h) => { if (url) arr.push({url, alt, text, w, h, base: location.href}); };
                  for (const img of Array.from(document.images || [])) {
                    const box = img.getBoundingClientRect();
                    const parent = img.closest('article, section, li, a, div');
                    push(img.currentSrc || img.src || img.getAttribute('data-src') || img.getAttribute('data-lazy'),
                         `${img.alt || ''} ${img.title || ''}`,
                         parent ? (parent.innerText || '').slice(0, 700) : '',
                         img.naturalWidth || Math.round(box.width) || 0,
                         img.naturalHeight || Math.round(box.height) || 0);
                  }
                  for (const el of Array.from(document.querySelectorAll('*'))) {
                    const bg = getComputedStyle(el).backgroundImage || '';
                    const m = bg.match(/url\(["']?([^"')]+)["']?\)/);
                    if (m) push(m[1], el.getAttribute('aria-label') || '', (el.innerText || '').slice(0, 700), el.clientWidth || 0, el.clientHeight || 0);
                  }
                  return arr;
                }
                """
            )
            browser.close()
        if isinstance(dados, list):
            return escolher_melhor_imagem(dados, contexto)
    except Exception:
        return SEM_INFO
    return SEM_INFO


def corrigir_imagem_artista(row: pd.Series, cfg: dict[str, Any] | None, contexto: str) -> str:
    img = limpar_texto(row.get("Link da imagem da obra"))
    if not cfg_eh_rubem_arremate(cfg):
        return img
    if img != SEM_INFO and not imagem_logo_ou_generica(img):
        return img
    link = limpar_texto(row.get("Link da obra"))
    if not link_http(link):
        return SEM_INFO
    contexto_img = " ".join([
        limpar_texto(row.get("Nome da obra")),
        limpar_texto(row.get("Autor")),
        contexto,
    ])[:1500]
    nova = extrair_imagem_estatica_pagina(link, contexto_img)
    if nova != SEM_INFO:
        return nova
    nova = extrair_imagem_renderizada_pagina(link, contexto_img)
    return nova if nova != SEM_INFO else SEM_INFO


def corrigir_titulo_artista(nome: Any, link: Any, autor: Any, site: str | None, contexto: str, cfg: dict[str, Any] | None) -> str:
    bruto = limpar_texto(nome)
    candidatos: list[str] = []

    rotulado = extrair_campo_por_rotulo(contexto, [r"nome(?:\s+da\s+obra)?", r"t[ií]tulo", r"titulo", r"title"], limite=160)
    if rotulado != SEM_INFO and rotulado != bruto:
        candidatos.append(rotulado)
    if bruto != SEM_INFO:
        candidatos.append(bruto)

    for cand in candidatos:
        cand = re.sub(r"^(?:nome(?:\s+da\s+obra)?|t[ií]tulo|titulo|title)\s*[:\-–—]\s*", "", cand, flags=re.I)
        # Primeiro tenta pedaços separados por barra vertical, bullets e quebras.
        partes = re.split(r"\s*(?:\||·|•|\n|\r)\s*", cand)
        # Depois tenta hífen/dash com espaço, comum em "Artista - Galeria".
        if len(partes) == 1:
            partes = re.split(r"\s+[\-–—]\s+", cand)
        for parte in partes:
            limpo = remover_ruidos_titulo(parte, cfg, autor, site)
            if not titulo_invalido_ou_contaminado(limpo, cfg, autor, site):
                return limpo
        # Se a string inteira, depois de remover artista/site, virou um título bom.
        limpo = remover_ruidos_titulo(cand, cfg, autor, site)
        if not titulo_invalido_ou_contaminado(limpo, cfg, autor, site):
            return limpo

    link_titulo = titulo_a_partir_link(link, cfg, autor, site)
    if link_titulo != SEM_INFO:
        return link_titulo

    # Para os scrapers específicos, é melhor mostrar ausência real de título do
    # que preencher o campo com técnica, artista ou galeria.
    return "Sem título" if cfg else SEM_INFO


def adaptar_colunas_entrada(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    lookup: dict[str, str] = {}
    for col in out.columns:
        lookup.setdefault(chave_coluna(col), col)
        lookup.setdefault(str(col).strip().lower(), col)
    prioridades = {
        "Nome da obra": ["nome_da_obra", "nome_obra", "nome da obra", "titulo", "título"],
        "Autor": ["autor", "artista"],
        "Preço": ["preco_brl", "preço brl", "preço_brl", "preco_numero", "preco", "preço", "valor"],
        "Dimensões": ["dimensoes", "dimensões", "medidas"],
        "Técnica": ["tecnica", "técnica"],
        "Técnica original": ["tecnica_original", "técnica original", "tecnica", "técnica"],
        "Ano da obra": ["ano_da_obra", "ano_obra", "ano"],
        "Descrição": ["descricao", "descrição"],
        "Link da obra": ["url_da_obra", "link_da_obra", "link_obra", "link da obra"],
        "Link da imagem da obra": ["imagem_da_obra", "url_imagem", "link_imagem", "link_da_imagem_da_obra", "link da imagem da obra"],
    }
    for destino, candidatos in prioridades.items():
        origem = None
        for c in candidatos:
            origem = lookup.get(c) or lookup.get(chave_coluna(c))
            if origem is not None:
                break
        if origem is not None:
            if destino not in out.columns:
                out[destino] = out[origem]
            else:
                out[destino] = out[destino].where(~out[destino].apply(sem_info), out[origem])
    if "Técnica original" not in out.columns and "Técnica" in out.columns:
        out["Técnica original"] = out["Técnica"]
    return out



def link_leiloesbr_lote(valor: Any) -> bool:
    """Aceita página individual de lote do LeilõesBR ou página direta peca.asp.

    O LeilõesBR muitas vezes aponta para o lote por
    abre_catalogo.asp?t=1|site_do_leiloeiro|leilao|id_da_peca. O scraper novo
    usa esse t=1 para descobrir a página direta do leiloeiro, geralmente
    /peca.asp?ID=..., onde os dados saem mais limpos. Por isso, a validação do
    app aceita os dois formatos e continua rejeitando busca/catálogo/menu.
    """
    link = limpar_texto(valor)
    if link == SEM_INFO:
        return False
    low = texto_busca(link)
    if "busca_andamento.asp" in low or "t=2" in low or "catalogo.asp" in low:
        return False

    # Formato LeilõesBR: abre_catalogo.asp?t=1|base|leilao|peca
    if "leiloesbr.com.br" in low and "abre_catalogo.asp" in low and "t=1" in low:
        try:
            from urllib.parse import unquote, urlsplit, parse_qsl
            query = urlsplit(link).query
            t = ""
            for k, v in parse_qsl(query, keep_blank_values=True):
                if k.lower() == "t":
                    t = unquote(v)
                    break
            if not t:
                m = re.search(r"[?&]t=([^&]+)", link, flags=re.I)
                t = unquote(m.group(1)) if m else ""
            partes = [p for p in t.split("|") if p]
            return len(partes) >= 4 and partes[0] == "1"
        except Exception:
            return "t=1" in low and low.count("|") >= 3

    # Formato direto dos sites de leiloeiros parceiros: /peca.asp?ID=123...
    if "peca.asp" in low and re.search(r"(?:\?|&)(?:id|cod|codigo|item|lote)=\d+", low, flags=re.I):
        return True

    return False


def imagem_lixo_leiloesbr(valor: Any) -> bool:
    img = limpar_texto(valor)
    if img == SEM_INFO:
        return False
    low = texto_busca(img)
    return any(x in low for x in [
        "logo", "placeholder", "banner", "icon", "icone", "avatar", "default",
        "sem_imagem", "sem-imagem", "no_image", "no-image", "spacer", "blank",
        "loading", "loader", "sprite", "pixel", "whatsapp", "facebook", "instagram",
        "youtube", "twitter", "botao", "button", "menu",
    ])


def registro_lixo_leiloesbr(row: Any, site: Any) -> bool:
    if "leilões br" not in texto_busca(site) and "leiloes br" not in texto_busca(site):
        return False
    link = limpar_texto(row.get("Link da obra") if hasattr(row, "get") else SEM_INFO)
    if not link_leiloesbr_lote(link):
        return True
    nome = limpar_texto(row.get("Nome da obra") if hasattr(row, "get") else SEM_INFO)
    low_nome = texto_busca(nome)
    if low_nome in {
        "leiloes", "leiloes br", "catalogo", "catalogo de pecas", "pinturas e gravuras",
        "resultado da busca", "meus lances", "meus favoritos", "informacoes", "pagamento frete",
    }:
        return True
    return False

def preparar_dataframe_obras(df: pd.DataFrame, site: str | None = None) -> pd.DataFrame:
    df = adaptar_colunas_entrada(df)
    for col in ["Nome da obra", "Autor", "Ano da obra", "Técnica", "Técnica original", "Dimensões", "Preço", "Descrição", "Link da obra", "Link da imagem da obra"]:
        if col not in df.columns:
            df[col] = SEM_INFO

    registros: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        contexto_base = contexto_linha(row)
        if registro_lixo_leiloesbr(row, site):
            continue
        cfg = detectar_artista_especifico(site, row.get("Autor"), contexto_base)
        if cfg_eh_rubem_arremate(cfg) and registro_lixo_rubem_arremate(
            row.get("Nome da obra"), row.get("Link da obra"), row.get("Link da imagem da obra"), site, row.get("Autor")
        ):
            continue
        texto_pagina = ""
        if cfg and link_http(row.get("Link da obra")):
            titulo_problematico = titulo_invalido_ou_contaminado(row.get("Nome da obra"), cfg, row.get("Autor"), site)
            campos_incompletos = any(sem_info(row.get(c)) for c in ["Técnica", "Técnica original", "Dimensões", "Ano da obra"])
            if titulo_problematico or campos_incompletos:
                texto_pagina = obter_texto_pagina_obra(limpar_texto(row.get("Link da obra")))
        contexto = contexto_base + ("\n" + texto_pagina if texto_pagina else "")

        autor = corrigir_autor(row.get("Autor"), row.get("Nome da obra"))
        if cfg and (autor == SEM_INFO or titulo_invalido_ou_contaminado(autor, cfg, autor, site)):
            autor = cfg["autor"]
        elif cfg and texto_busca(cfg["autor"]) in texto_busca(contexto) and texto_busca(autor) not in [texto_busca(cfg["autor"]), ""]:
            # Evita que galeria/site ou título contaminado ocupem o campo autor.
            if any(r in texto_busca(autor) for r in [texto_busca(x) for x in cfg.get("ruidos", [])]):
                autor = cfg["autor"]

        tecnica_original, tecnica = resolver_tecnica(row, contexto, cfg)

        dimensoes_pagina = extrair_dimensoes_contexto(texto_pagina) if texto_pagina else SEM_INFO
        dimensoes = dimensoes_pagina if dimensoes_pagina != SEM_INFO else padronizar_dimensoes(row.get("Dimensões"))
        if dimensoes == SEM_INFO:
            dimensoes = extrair_dimensoes_contexto(contexto)

        ano_pagina = extrair_ano_contexto(texto_pagina) if texto_pagina else SEM_INFO
        ano = ano_pagina if ano_pagina != SEM_INFO else normalizar_ano(row.get("Ano da obra"))
        if ano == SEM_INFO:
            ano = extrair_ano_contexto(contexto)

        nome_extraido = extrair_titulo_de_texto(texto_pagina, cfg, autor, site) if texto_pagina else SEM_INFO
        nome_base = nome_extraido if nome_extraido != SEM_INFO else row.get("Nome da obra")
        nome = corrigir_titulo_artista(nome_base, row.get("Link da obra"), autor, site, contexto, cfg)
        imagem_final = corrigir_imagem_artista(row, cfg, contexto)
        if ("leilões br" in texto_busca(site) or "leiloes br" in texto_busca(site)) and imagem_lixo_leiloesbr(imagem_final):
            imagem_final = SEM_INFO

        if cfg_eh_rubem_arremate(cfg) and registro_lixo_rubem_arremate(nome, row.get("Link da obra"), imagem_final, site, autor):
            continue

        registros.append({
            "Nome da obra": nome,
            "Autor": autor,
            "Preço": parse_preco(row.get("Preço")),
            "Dimensões": dimensoes,
            "Técnica original": tecnica_original,
            "Técnica": tecnica,
            "Ano da obra": ano,
            "Descrição": limpar_texto(row.get("Descrição")),
            "Link da obra": limpar_texto(row.get("Link da obra")),
            "Link da imagem da obra": imagem_final,
        })
    return pd.DataFrame(registros)


def _ler_cache_dolar() -> dict:
    try:
        if USD_CACHE_PATH.exists():
            return json.loads(USD_CACHE_PATH.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}

def _salvar_cache_dolar(valor: float) -> None:
    try:
        USD_CACHE_PATH.write_text(
            json.dumps(
                {"rate": float(valor), "updated_at": agora_brasilia_iso()},
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
    except Exception:
        pass

def atualizar_cotacao_dolar() -> float:
    headers = {"User-Agent": "Mozilla/5.0"}
    fontes = [
        ("https://economia.awesomeapi.com.br/json/last/USD-BRL", lambda d: d.get("USDBRL", {}).get("bid")),
        ("https://economia.awesomeapi.com.br/last/USD-BRL", lambda d: d.get("USDBRL", {}).get("bid")),
        ("https://api.frankfurter.app/latest?from=USD&to=BRL", lambda d: (d.get("rates") or {}).get("BRL")),
        ("https://open.er-api.com/v6/latest/USD", lambda d: (d.get("rates") or {}).get("BRL")),
    ]
    for url, getter in fontes:
        try:
            resp = requests.get(url, timeout=10, headers=headers)
            resp.raise_for_status()
            valor = parse_preco(getter(resp.json()))
            if valor and valor > 0:
                _salvar_cache_dolar(valor)
                return valor
        except Exception:
            continue
    cache = _ler_cache_dolar()
    valor = parse_preco(cache.get("rate"))
    if valor:
        # Mesmo quando a API falha, o botão de atualização registra a tentativa usando a cotação em cache.
        _salvar_cache_dolar(valor)
        return valor
    _salvar_cache_dolar(5.0)
    return 5.0

@st.cache_data(show_spinner=False, ttl=60 * 60)
def obter_cotacao_dolar() -> float:
    cache = _ler_cache_dolar()
    valor = parse_preco(cache.get("rate"))
    if valor:
        return valor
    return atualizar_cotacao_dolar()

def ultima_atualizacao_dolar() -> str:
    cache = _ler_cache_dolar()
    return formatar_data_brasilia(cache.get("updated_at"))

def conectar() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def inicializar_banco() -> None:
    with conectar() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS obras (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome_obra TEXT,
                autor TEXT,
                preco REAL,
                dimensoes TEXT,
                tecnica TEXT,
                tecnica_original TEXT,
                ano_obra TEXT,
                descricao TEXT,
                link_obra TEXT UNIQUE,
                link_imagem TEXT,
                site TEXT,
                origem TEXT,
                cotacao_dolar REAL,
                coletado_em TEXT,
                atualizado_em TEXT
            )
            """
        )
        conn.commit()


def dataframe_para_banco(df: pd.DataFrame, site: str, origem: str, cotacao_dolar: float | None) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    df_norm = preparar_dataframe_obras(df, site=site).rename(columns=UI_TO_DB)
    agora = agora_brasilia_iso()
    df_norm["site"] = site
    df_norm["origem"] = origem
    df_norm["cotacao_dolar"] = cotacao_dolar
    df_norm["coletado_em"] = agora
    df_norm["atualizado_em"] = agora
    cols = ["nome_obra", "autor", "preco", "dimensoes", "tecnica", "tecnica_original", "ano_obra", "descricao", "link_obra", "link_imagem", "site", "origem", "cotacao_dolar", "coletado_em", "atualizado_em"]
    for c in cols:
        if c not in df_norm.columns:
            df_norm[c] = None
    df_norm = df_norm[cols]
    df_norm = df_norm[df_norm["nome_obra"].apply(nome_valido)]
    if "link_obra" in df_norm.columns:
        df_norm = df_norm.drop_duplicates(subset=["link_obra"], keep="first")
    return df_norm


def inserir_obras(df_db: pd.DataFrame) -> int:
    if df_db.empty:
        return 0
    inseridas = 0
    with conectar() as conn:
        for _, row in df_db.iterrows():
            dados = row.to_dict()
            try:
                conn.execute(
                    """
                    INSERT INTO obras (nome_obra, autor, preco, dimensoes, tecnica, tecnica_original, ano_obra, descricao, link_obra, link_imagem, site, origem, cotacao_dolar, coletado_em, atualizado_em)
                    VALUES (:nome_obra, :autor, :preco, :dimensoes, :tecnica, :tecnica_original, :ano_obra, :descricao, :link_obra, :link_imagem, :site, :origem, :cotacao_dolar, :coletado_em, :atualizado_em)
                    ON CONFLICT(link_obra) DO UPDATE SET
                        nome_obra=excluded.nome_obra,
                        autor=excluded.autor,
                        preco=excluded.preco,
                        dimensoes=excluded.dimensoes,
                        tecnica=excluded.tecnica,
                        tecnica_original=excluded.tecnica_original,
                        ano_obra=excluded.ano_obra,
                        descricao=excluded.descricao,
                        link_imagem=excluded.link_imagem,
                        site=excluded.site,
                        origem=excluded.origem,
                        cotacao_dolar=excluded.cotacao_dolar,
                        atualizado_em=excluded.atualizado_em
                    """,
                    dados,
                )
                inseridas += 1
            except Exception:
                pass
        conn.commit()
    carregar_acervo.clear()
    return inseridas


@st.cache_data(show_spinner=False)
def carregar_acervo() -> pd.DataFrame:
    inicializar_banco()
    with conectar() as conn:
        df = pd.read_sql_query("SELECT * FROM obras ORDER BY id DESC", conn)
    return df


def numero(v: Any) -> str:
    try:
        return f"{int(v):,}".replace(",", ".")
    except Exception:
        return "0"


def dinheiro(v: Any) -> str:
    try:
        if pd.isna(v):
            return SEM_INFO
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return SEM_INFO


def percentual(v: Any) -> str:
    try:
        return f"{float(v):.1f}%".replace(".", ",")
    except Exception:
        return "0,0%"


def patch_legacy_script(src: Path, max_obras: int, headless: bool, destino: Path) -> None:
    texto = src.read_text(encoding="utf-8")
    limite = max_obras if max_obras > 0 else 999999
    texto = re.sub(r"max_links\s*=\s*\d+", f"max_links = {limite}", texto)
    texto = re.sub(r"^\s*import\s+pyautogui\s*$", "", texto, flags=re.MULTILINE)
    texto = re.sub(r"chromium\.launch\(headless\s*=\s*False\)", f"chromium.launch(headless={str(headless)})", texto)
    texto = re.sub(r"chromium\.launch\(headless\s*=\s*True\)", f"chromium.launch(headless={str(headless)})", texto)
    texto = re.sub(r"chromium\.launch\(headless=False\)", f"chromium.launch(headless={str(headless)})", texto)
    texto = re.sub(r"chromium\.launch\(headless=True\)", f"chromium.launch(headless={str(headless)})", texto)
    texto = texto.replace('pyautogui.alert("Sistema Executado")', 'print("Sistema Executado")')
    texto = texto.replace("pyautogui.alert('Sistema Executado')", 'print("Sistema Executado")')
    destino.write_text(texto, encoding="utf-8")



def _extrair_preco_contexto(texto: Any) -> str:
    base = limpar_texto(texto)
    if base == SEM_INFO:
        return SEM_INFO
    m = re.search(r"R\$\s*\d[\d.]*,\d{2}|R\$\s*\d[\d.]*", base)
    return m.group(0) if m else SEM_INFO


def _coletar_cards_rubem_renderizado(page) -> list[dict[str, Any]]:
    cards = []
    try:
        for _ in range(8):
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            page.wait_for_timeout(900)
        cards = page.evaluate(
            """
            () => {
              const out = [];
              const seen = new Set();
              const bestImg = (root) => {
                const img = root ? root.querySelector('img') : null;
                return img ? (img.currentSrc || img.src || img.getAttribute('data-src') || img.getAttribute('data-lazy') || '') : '';
              };
              for (const a of Array.from(document.querySelectorAll('a[href*="/lote/"]'))) {
                const href = a.href || '';
                if (!href || seen.has(href)) continue;
                seen.add(href);
                const card = a.closest('article, section, li, [class*="card"], [class*="lot"], [class*="lote"], [class*="item"], div') || a;
                out.push({href, text: (card.innerText || a.innerText || '').slice(0, 2500), img: bestImg(card)});
              }
              return out;
            }
            """
        )
    except Exception:
        return []
    if not isinstance(cards, list):
        return []
    limpos: list[dict[str, Any]] = []
    vistos: set[str] = set()
    for card in cards:
        href = limpar_texto(card.get("href"))
        if href == SEM_INFO or href in vistos:
            continue
        vistos.add(href)
        if "/lote/" not in href:
            continue
        limpos.append(card)
    return limpos


def _extrair_dados_lote_rubem(page, link: str, texto_card: str = "", imagem_card: str = "") -> dict[str, Any] | None:
    try:
        page.goto(link, wait_until="domcontentloaded", timeout=45000)
        try:
            page.wait_for_load_state("networkidle", timeout=12000)
        except Exception:
            pass
        page.wait_for_timeout(1800)
        dados = page.evaluate(
            """
            () => {
              const clean = (s) => (s || '').replace(/\s+/g, ' ').trim();
              const headings = Array.from(document.querySelectorAll('h1,h2,h3,[class*="title"],[class*="titulo"]')).map(e => clean(e.innerText)).filter(Boolean).slice(0, 20);
              const images = [];
              const push = (url, alt, text, w, h) => { if (url) images.push({url, alt, text, w, h, base: location.href}); };
              for (const img of Array.from(document.images || [])) {
                const box = img.getBoundingClientRect();
                const parent = img.closest('article, section, li, a, div');
                push(img.currentSrc || img.src || img.getAttribute('data-src') || img.getAttribute('data-lazy'),
                     `${img.alt || ''} ${img.title || ''}`,
                     parent ? clean(parent.innerText).slice(0, 800) : '',
                     img.naturalWidth || Math.round(box.width) || 0,
                     img.naturalHeight || Math.round(box.height) || 0);
              }
              return {title: document.title || '', headings, text: clean(document.body ? document.body.innerText : '').slice(0, 9000), images};
            }
            """
        )
    except Exception:
        return None
    if not isinstance(dados, dict):
        return None
    contexto = "\n".join([
        limpar_texto(dados.get("title")),
        "\n".join(dados.get("headings") or []),
        limpar_texto(texto_card),
        limpar_texto(dados.get("text")),
    ])
    cfg = ARTISTAS_ESPECIFICOS[2]
    autor = "Rubem Grilo"
    site = "Rubem Grilo (Arremate Arte)"
    titulo = extrair_titulo_de_texto(contexto, cfg, autor, site)
    if titulo == SEM_INFO:
        titulo = titulo_a_partir_link(link, cfg, autor, site)
    if titulo == SEM_INFO:
        titulo = "Sem título"
    img = imagem_card if imagem_candidata_util(imagem_card, link) else SEM_INFO
    if img == SEM_INFO:
        img = escolher_melhor_imagem(dados.get("images") or [], contexto)
    if registro_lixo_rubem_arremate(titulo, link, img, site, autor):
        return None
    return {
        "Nome da obra": titulo,
        "Autor": autor,
        "Preço": _extrair_preco_contexto(contexto),
        "Técnica": contexto,
        "Técnica original": contexto,
        "Dimensões": extrair_dimensoes_contexto(contexto),
        "Ano da obra": extrair_ano_contexto(contexto),
        "Descrição": limpar_texto(contexto)[:1200],
        "Link da obra": link,
        "Link da imagem da obra": img,
    }


def executar_scraper_rubem_grilo_arremate(max_obras: int, headless: bool) -> tuple[bool, str, pd.DataFrame]:
    """Coleta Rubem Grilo diretamente com Playwright para não importar logo/menu.

    O scraper antigo do Arremate Arte capturava elementos de navegação e o logo
    do site. Esta rotina busca apenas links de lote renderizados e extrai a
    imagem real do lote.
    """
    garantir_playwright_chromium()
    limite = max_obras if max_obras > 0 else 999999
    url_artista = "https://www.arrematearte.com.br/artistas/rubem-grilo-1946"
    registros: list[dict[str, Any]] = []
    logs: list[str] = []
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=bool(headless))
            page = browser.new_page(
                viewport={"width": 1365, "height": 1100},
                user_agent="Mozilla/5.0 (compatible; OraculoCultural/1.0)",
            )
            page.goto(url_artista, wait_until="domcontentloaded", timeout=60000)
            try:
                page.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                pass
            page.wait_for_timeout(2500)
            cards = _coletar_cards_rubem_renderizado(page)
            logs.append(f"Links de lote encontrados: {len(cards)}")
            for card in cards:
                if len(registros) >= limite:
                    break
                dados = _extrair_dados_lote_rubem(
                    page,
                    limpar_texto(card.get("href")),
                    limpar_texto(card.get("text")),
                    normalizar_url_imagem(card.get("img"), url_artista),
                )
                if dados:
                    registros.append(dados)
            browser.close()
    except Exception as exc:
        return False, f"Erro no scraper Rubem Grilo/Arremate: {exc}", pd.DataFrame()

    df = pd.DataFrame(registros)
    if not df.empty:
        df = df.drop_duplicates(subset=["Link da obra"], keep="first")
        try:
            df.to_excel(COLETAS_DIR / "ultima_coleta_rubem_grilo_arremate_arte.xlsx", index=False)
        except Exception:
            pass
        return True, "\n".join(logs + [f"Obras válidas coletadas: {len(df)}"]), df
    return False, "\n".join(logs + ["Nenhuma obra válida foi encontrada. O sistema não importou menus/logos como obras."]), pd.DataFrame()


def executar_scraper(site: str, max_obras: int, headless: bool) -> tuple[bool, str, pd.DataFrame]:
    if site == "Rubem Grilo (Arremate Arte)":
        return executar_scraper_rubem_grilo_arremate(max_obras, headless)

    config = SITES[site]
    src = SCRAPERS_DIR / config["script"]
    workdir = COLETAS_DIR / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{chave_coluna(site)}"
    workdir.mkdir(parents=True, exist_ok=True)
    runner = config["runner"]
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["MAX_OBRAS"] = str(max_obras)
    if runner in {"saatchi_cli", "legacy_playwright", "playwright_cli"}:
        garantir_playwright_chromium()

    if runner == "saatchi_cli":
        target = workdir / config["script"]
        shutil.copy(src, target)
        limite = max_obras if max_obras > 0 else 999999
        cmd = [sys.executable, str(target), "--max", str(limite), "--paginas", "999", "--headless", "true" if headless else "false", "--saida", config["xlsx"]]
    elif runner == "artsoul_cli":
        target = workdir / config["script"]
        shutil.copy(src, target)
        cmd = [sys.executable, str(target), "--max-obras", str(max_obras), "--max-pages", "0", "--output", config["xlsx"]]
    elif runner == "playwright_cli":
        target = workdir / config["script"]
        shutil.copy(src, target)
        cmd = [
            sys.executable, str(target),
            "--max-obras", str(max_obras),
            "--output", config["xlsx"],
            "--headless", "true" if headless else "false",
        ]
    else:
        target = workdir / config["script"]
        patch_legacy_script(src, max_obras, headless, target)
        cmd = [sys.executable, str(target)]
    try:
        proc = subprocess.run(cmd, cwd=workdir, env=env, capture_output=True, text=True, timeout=60 * 60)
        log = (proc.stdout or "") + "\n" + (proc.stderr or "")
        xlsx = workdir / config["xlsx"]
        if not xlsx.exists():
            return False, f"Arquivo de saída não encontrado.\n\n{log[-4000:]}", pd.DataFrame()
        df = pd.read_excel(xlsx, header=int(config.get("header", 1)))
        if not any(str(c).strip().lower() in {"título", "titulo", "nome da obra", "nome_da_obra"} for c in df.columns):
            try:
                df0 = pd.read_excel(xlsx, header=0)
                if any(str(c).strip().lower() in {"título", "titulo", "nome da obra", "nome_da_obra"} for c in df0.columns):
                    df = df0
            except Exception:
                pass
        df.to_excel(COLETAS_DIR / f"ultima_coleta_{chave_coluna(site)}.xlsx", index=False)
        return proc.returncode == 0, log[-5000:], df
    except subprocess.TimeoutExpired:
        return False, "Tempo máximo excedido durante a coleta.", pd.DataFrame()
    except Exception as exc:
        return False, f"Erro ao executar scraper: {exc}", pd.DataFrame()


def salvar_upload_imagem(uploaded) -> str:
    if uploaded is None:
        return SEM_INFO
    ext = Path(uploaded.name).suffix.lower() or ".jpg"
    digest = hashlib.sha1(uploaded.getvalue()).hexdigest()[:16]
    destino = UPLOAD_DIR / f"obra_{digest}{ext}"
    destino.write_bytes(uploaded.getvalue())
    return str(destino)


def imagem_para_exibir(valor: Any) -> str | None:
    """Resolve imagem para st.image.

    Aceita URL externa, caminho absoluto, caminho relativo do projeto
    ou valor vazio. Essa função estava sendo usada no Acervo e
    Comparativo, mas faltou entrar na versão anterior.
    """
    texto = limpar_texto(valor)
    if texto == SEM_INFO or imagem_logo_ou_generica(texto):
        return None
    if texto.startswith(("http://", "https://")):
        return texto

    caminho = Path(texto)
    if caminho.is_absolute() and caminho.exists():
        return str(caminho)

    relativo_base = BASE_DIR / texto
    if relativo_base.exists():
        return str(relativo_base)

    relativo_upload = UPLOAD_DIR / texto
    if relativo_upload.exists():
        return str(relativo_upload)

    return None


def sanear_acervo() -> int:
    alterados = 0
    agora = agora_brasilia_iso()
    ids_excluir: list[int] = []
    with conectar() as conn:
        rows = conn.execute("SELECT * FROM obras").fetchall()
        for row in rows:
            if registro_lixo_rubem_arremate(row["nome_obra"], row["link_obra"], row["link_imagem"], row["site"], row["autor"]):
                ids_excluir.append(int(row["id"]))
                continue
            bruto = pd.DataFrame([{
                "Nome da obra": row["nome_obra"],
                "Autor": row["autor"],
                "Preço": row["preco"],
                "Dimensões": row["dimensoes"],
                "Técnica": row["tecnica"],
                "Técnica original": row["tecnica_original"] or row["tecnica"],
                "Ano da obra": row["ano_obra"],
                "Descrição": row["descricao"],
                "Link da obra": row["link_obra"],
                "Link da imagem da obra": row["link_imagem"],
            }])
            saneado = preparar_dataframe_obras(bruto, site=row["site"])
            if saneado.empty:
                continue
            item = saneado.iloc[0]
            novo_nome = item["Nome da obra"]
            novo_autor = item["Autor"]
            tecnica_original = item["Técnica original"]
            nova_tecnica = item["Técnica"]
            novo_ano = item["Ano da obra"]
            nova_dim = item["Dimensões"]
            nova_imagem = item["Link da imagem da obra"]
            if any([
                novo_nome != row["nome_obra"],
                novo_autor != row["autor"],
                nova_tecnica != row["tecnica"],
                tecnica_original != row["tecnica_original"],
                novo_ano != row["ano_obra"],
                nova_dim != row["dimensoes"],
                nova_imagem != row["link_imagem"],
            ]):
                conn.execute(
                    "UPDATE obras SET nome_obra=?, autor=?, tecnica=?, tecnica_original=?, ano_obra=?, dimensoes=?, link_imagem=?, atualizado_em=? WHERE id=?",
                    (novo_nome, novo_autor, nova_tecnica, tecnica_original, novo_ano, nova_dim, nova_imagem, agora, row["id"]),
                )
                alterados += 1
        if ids_excluir:
            placeholders = ",".join(["?"] * len(ids_excluir))
            cur = conn.execute(f"DELETE FROM obras WHERE id IN ({placeholders})", tuple(ids_excluir))
            alterados += cur.rowcount or 0
        conn.commit()
    if alterados:
        carregar_acervo.clear()
    return alterados


def excluir_por_site(site: str | None = None) -> int:
    with conectar() as conn:
        if site:
            cur = conn.execute("DELETE FROM obras WHERE site = ?", (site,))
        else:
            cur = conn.execute("DELETE FROM obras")
        conn.commit()
        total = cur.rowcount or 0
    carregar_acervo.clear()
    return total


def excluir_obras_por_ids(ids: list[int]) -> int:
    if not ids:
        return 0
    placeholders = ",".join(["?"] * len(ids))
    with conectar() as conn:
        cur = conn.execute(f"DELETE FROM obras WHERE id IN ({placeholders})", tuple(ids))
        conn.commit()
        total = cur.rowcount or 0
    carregar_acervo.clear()
    return total


def calcular_area(dim: str) -> float | None:
    nums = [float(x.replace(",", ".")) for x in re.findall(r"\d+(?:[.,]\d+)?", str(dim))]
    if len(nums) >= 2:
        return nums[0] * nums[1]
    return None


def similaridade(obra: pd.Series, comp: pd.Series) -> tuple[float, dict[str, float]]:
    pontos = {"Autor": 0.0, "Técnica": 0.0, "Dimensões": 0.0, "Ano": 0.0, "Preço": 0.0}

    if limpar_texto(obra.get("autor")).lower() == limpar_texto(comp.get("autor")).lower() and limpar_texto(obra.get("autor")) != SEM_INFO:
        pontos["Autor"] = 40

    if limpar_texto(obra.get("tecnica")) == limpar_texto(comp.get("tecnica")) and limpar_texto(obra.get("tecnica")) != SEM_INFO:
        pontos["Técnica"] = 20

    area1, area2 = calcular_area(obra.get("dimensoes", "")), calcular_area(comp.get("dimensoes", ""))
    if area1 and area2:
        ratio = min(area1, area2) / max(area1, area2)
        pontos["Dimensões"] = 15 * ratio

    a1, a2 = normalizar_ano(obra.get("ano_obra")), normalizar_ano(comp.get("ano_obra"))
    if a1 != SEM_INFO and a2 != SEM_INFO:
        dif_ano = abs(int(a1) - int(a2))
        pontos["Ano"] = max(0, 15 * (1 - dif_ano / 50))

    p1, p2 = obra.get("preco"), comp.get("preco")
    try:
        if pd.notna(p1) and pd.notna(p2) and float(p1) > 0 and float(p2) > 0:
            dif = abs(math.log(float(p1)) - math.log(float(p2)))
            pontos["Preço"] = max(0, 10 * (1 - dif / 2.5))
    except Exception:
        pass

    total = round(sum(pontos.values()), 2)
    return total, pontos


def classificar_nivel(score: float) -> str:
    if score >= 75:
        return "Nível 1"
    if score >= 50:
        return "Nível 2"
    return "Nível 3"


def diagnostico_comparacao(obra: pd.Series, comp: pd.Series, pontos: dict[str, float]) -> dict[str, Any]:
    area_ref = calcular_area(obra.get("dimensoes", ""))
    area_comp = calcular_area(comp.get("dimensoes", ""))
    if area_ref and area_comp:
        razao_area = min(area_ref, area_comp) / max(area_ref, area_comp)
        area_txt = f"{razao_area * 100:.1f}%".replace(".", ",")
    else:
        area_txt = SEM_INFO
    try:
        p_ref = float(obra.get("preco"))
        p_comp = float(comp.get("preco"))
        dif_preco = abs(p_ref - p_comp) / max(p_ref, p_comp) if p_ref > 0 and p_comp > 0 else None
        preco_txt = f"{dif_preco * 100:.1f}%".replace(".", ",") if dif_preco is not None else SEM_INFO
    except Exception:
        preco_txt = SEM_INFO
    a_ref = normalizar_ano(obra.get("ano_obra"))
    a_comp = normalizar_ano(comp.get("ano_obra"))
    dif_ano_txt = SEM_INFO if a_ref == SEM_INFO or a_comp == SEM_INFO else f"{abs(int(a_ref) - int(a_comp))} ano(s)"
    return {
        "Autor igual": pontos.get("Autor", 0) > 0,
        "Técnica igual": pontos.get("Técnica", 0) > 0,
        "Proximidade área": area_txt,
        "Diferença ano": dif_ano_txt,
        "Diferença preço": preco_txt,
    }



def aba_inicio(df: pd.DataFrame) -> None:
    st.title("🔮 Oráculo Cultural")
    st.markdown("Coleta obras culturais e organiza comparáveis para análise contábil a valor justo.")

    ultima_coleta_obras = SEM_INFO
    if not df.empty and "coletado_em" in df.columns and df["coletado_em"].notna().any():
        try:
            datas = [_converter_para_brasilia(v) for v in df["coletado_em"].dropna().tolist()]
            datas = [d for d in datas if d is not None]
            if datas:
                ultima_coleta_obras = max(datas).strftime("%d/%m/%Y %H:%M:%S")
        except Exception:
            pass

    cbtn, _, _, _ = st.columns([1.2, 1, 1, 1])
    if cbtn.button("Atualizar cotação do dólar"):
        atualizar_cotacao_dolar()
        obter_cotacao_dolar.clear()
        st.success("Cotação do dólar atualizada com sucesso.")
        st.rerun()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Obras cadastradas", numero(len(df)))
    c2.metric("Preço médio", dinheiro(df["preco"].dropna().mean()) if not df.empty and "preco" in df and df["preco"].notna().any() else SEM_INFO)
    c3.metric("Sites", numero(df["site"].nunique()) if not df.empty and "site" in df else "0")
    c4.metric("Cotação USD-BRL", f"R$ {obter_cotacao_dolar():.4f}".replace(".", ","))

    i1, i2 = st.columns(2)
    i1.info(f"**Última coleta de obras (horário de Brasília):** {ultima_coleta_obras}")
    i2.info(f"**Última atualização da cotação do dólar (horário de Brasília):** {ultima_atualizacao_dolar()}")

    st.subheader("Passos")
    st.markdown(
        '''
        1. Use **Coleta** para executar os scrapers, importar planilhas ou cadastrar obras manualmente.  
        2. Confira imagens e campos centrais em **Acervo**.  
        3. Use **Comparativo** para selecionar uma obra de referência e encontrar comparáveis por similaridade.  
        4. Consulte **Informações** para entender valor justo, heritage assets e os limites da metodologia.  
        5. Classifique a evidência em **Nível 1**, **Nível 2** ou **Nível 3** conforme a proximidade dos comparáveis.
        '''
    )


def aba_coleta(df: pd.DataFrame) -> None:
    st.header("Coleta")
    st.subheader("Coleta automática")
    st.caption("Use 0 para tentar coletar o máximo disponível. Também é possível executar os scrapers específicos e o Leilões BR com validação de lotes/imagens.")
    with st.form("form_scrapers"):
        sites = st.multiselect("Sites", list(SITES.keys()), default=[])
        max_obras = st.number_input("Quantidade máxima por site (0 = sem limite)", 0, 5000, 100, 10)
        headless = st.checkbox("Executar navegador em segundo plano", value=True)
        executar = st.form_submit_button("Executar coleta")
    if executar:
        if not sites:
            st.warning("Selecione pelo menos um site.")
        else:
            cot = obter_cotacao_dolar()
            total_salvas = 0
            for site in sites:
                with st.status(f"Coletando {site}...", expanded=True) as status:
                    ok, log, df_coleta = executar_scraper(site, int(max_obras), bool(headless))
                    st.text(log[-3000:] if log else "Sem log.")
                    if not df_coleta.empty:
                        df_db = dataframe_para_banco(df_coleta, site, "scraper", cot)
                        inseridas = inserir_obras(df_db)
                        total_salvas += inseridas
                        status.update(label=f"{site}: {inseridas} obras salvas", state="complete")
                    else:
                        status.update(label=f"{site}: nenhuma obra importada", state="error" if not ok else "complete")
            if total_salvas > 0:
                st.session_state["coleta_sucesso"] = f"Obras coletadas com sucesso. {total_salvas} obra(s) foram salvas."
            st.rerun()

    if st.session_state.get("coleta_sucesso"):
        st.success(st.session_state.pop("coleta_sucesso"))

    st.divider()
    st.subheader("Importar planilha")
    arq = st.file_uploader("Planilha Excel ou CSV", type=["xlsx", "xls", "csv"])
    site_plan = st.selectbox("Site/origem da planilha", list(SITES.keys()) + ["Manual", "Outro"])
    if arq and st.button("Importar planilha"):
        if arq.name.lower().endswith(".csv"):
            df_imp = pd.read_csv(arq)
        else:
            df_imp = pd.read_excel(arq)
        inseridas = inserir_obras(dataframe_para_banco(df_imp, site_plan, "importação", obter_cotacao_dolar()))
        st.success(f"{inseridas} obra(s) importada(s).")
        st.rerun()

    st.divider()
    st.subheader("Cadastro manual")
    with st.form("manual"):
        nome = st.text_input("Nome da obra")
        autor = st.text_input("Autor")
        tipo_preco = st.radio(
            "Preço da obra",
            ["Adicionar preço", "Preço sem informação"],
            horizontal=True,
            index=0,
            help="Escolha 'Adicionar preço' quando souber o valor da obra ou 'Preço sem informação' quando o valor ainda não estiver disponível.",
        )
        preco_sem_info = tipo_preco == "Preço sem informação"
        preco = None
        if not preco_sem_info:
            preco = st.number_input("Preço em R$", min_value=0.0, step=100.0)
        tecnica = st.selectbox("Técnica", TECNICAS_PADRONIZADAS, index=TECNICAS_PADRONIZADAS.index("Sem informação"))
        dimensoes = st.text_input("Dimensões", placeholder="Ex.: 80 x 120 cm")
        ano = st.text_input("Ano da obra")
        link = st.text_input("Link da obra")
        link_img = st.text_input("Link da imagem da obra")
        img_upload = st.file_uploader("Ou envie a imagem da obra", type=["png", "jpg", "jpeg", "webp"])
        descricao = st.text_area("Descrição")
        salvar = st.form_submit_button("Salvar obra manual")
    if salvar:
        imagem_local = salvar_upload_imagem(img_upload)
        imagem_final = link_img if link_img.strip() else imagem_local
        link_final = link.strip() or f"manual://{hashlib.sha1((nome+autor+str(datetime.now())).encode()).hexdigest()}"
        df_manual = pd.DataFrame([{
            "Nome da obra": nome, "Autor": autor, "Preço": SEM_INFO if preco_sem_info else preco, "Técnica": tecnica,
            "Técnica original": tecnica, "Dimensões": dimensoes, "Ano da obra": ano,
            "Descrição": descricao, "Link da obra": link_final, "Link da imagem da obra": imagem_final,
        }])
        inseridas = inserir_obras(dataframe_para_banco(df_manual, "Manual", "manual", obter_cotacao_dolar()))
        st.success(f"{inseridas} obra(s) salva(s).")
        st.rerun()

    st.divider()
    st.subheader("Limpeza e correção")
    a, b = st.columns([1.5, 1], gap="large")
    with a:
        st.markdown("**Apagar registros do acervo local**")
        site_del = st.selectbox("Apagar por site", ["Todos"] + list(SITES.keys()) + ["Manual"])
        st.caption("Você pode apagar um site específico ou toda a base local.")
    with b:
        st.markdown("&nbsp;", unsafe_allow_html=True)
        if st.button("Apagar dados selecionados", use_container_width=True):
            total = excluir_por_site(None if site_del == "Todos" else site_del)
            st.success(f"{total} registro(s) apagado(s).")
            st.rerun()
        if st.button("Corrigir/sanear acervo já salvo", use_container_width=True):
            total = sanear_acervo()
            st.success(f"{total} registro(s) corrigido(s).")
            st.rerun()

    st.markdown("**Excluir obras específicas**")
    if df.empty:
        st.caption("Nenhuma obra cadastrada para exclusão individual.")
    else:
        opcoes_obras = {
            f"#{int(row['id'])} · {limpar_texto(row.get('nome_obra'))} · {limpar_texto(row.get('autor'))} · {limpar_texto(row.get('site'))}": int(row["id"])
            for _, row in df.sort_values("id", ascending=False).iterrows()
        }
        obras_para_excluir = st.multiselect(
            "Selecione uma ou mais obras para excluir",
            list(opcoes_obras.keys()),
            help="Esta opção remove apenas as obras selecionadas, mantendo as demais opções de limpeza por site e saneamento.",
        )
        if st.button("Excluir obras específicas", use_container_width=True, disabled=not obras_para_excluir):
            ids = [opcoes_obras[item] for item in obras_para_excluir]
            total = excluir_obras_por_ids(ids)
            st.success(f"{total} obra(s) específica(s) excluída(s).")
            st.rerun()

def aba_acervo(df: pd.DataFrame) -> None:
    st.header("Acervo")
    if df.empty:
        st.info("Nenhuma obra cadastrada ainda.")
        return
    c1, c2, c3, c4 = st.columns(4)
    sites = c1.multiselect("Sites", sorted(df["site"].dropna().unique().tolist()), default=[])
    tecnicas = c2.multiselect("Técnicas", sorted(df["tecnica"].dropna().unique().tolist()), default=[])
    busca = c3.text_input("Buscar por nome/autor")
    ordem = c4.selectbox("Ordenar", ["Mais recentes", "Preço maior", "Preço menor", "Nome"])
    filtrado = df.copy()
    if sites:
        filtrado = filtrado[filtrado["site"].isin(sites)]
    if tecnicas:
        filtrado = filtrado[filtrado["tecnica"].isin(tecnicas)]
    if busca:
        b = busca.lower()
        filtrado = filtrado[filtrado.apply(lambda r: b in str(r.get("nome_obra", "")).lower() or b in str(r.get("autor", "")).lower(), axis=1)]
    if ordem == "Preço maior":
        filtrado = filtrado.sort_values("preco", ascending=False, na_position="last")
    elif ordem == "Preço menor":
        filtrado = filtrado.sort_values("preco", ascending=True, na_position="last")
    elif ordem == "Nome":
        filtrado = filtrado.sort_values("nome_obra")
    st.caption(f"{len(filtrado)} obra(s) exibida(s)")
    cols = st.columns(3)
    for idx, (_, row) in enumerate(filtrado.iterrows()):
        with cols[idx % 3]:
            img = imagem_para_exibir(row.get("link_imagem"))
            if img:
                st.image(img, use_container_width=True)
            st.markdown(f"**{limpar_texto(row.get('nome_obra'))}**")
            st.caption(f"{limpar_texto(row.get('autor'))} · {limpar_texto(row.get('site'))}")
            st.write(f"**Preço:** {dinheiro(row.get('preco'))}")
            st.write(f"**Técnica:** {limpar_texto(row.get('tecnica'))}")
            st.write(f"**Dimensões:** {limpar_texto(row.get('dimensoes'))}")
            st.write(f"**Ano:** {limpar_texto(row.get('ano_obra'))}")
            link_obra = limpar_texto(row.get("link_obra"))
            if link_obra != SEM_INFO and link_obra.startswith(("http://", "https://")):
                st.link_button("Abrir obra no site", link_obra, use_container_width=True)


def montar_resultados_comparativo(df: pd.DataFrame, ref: pd.Series) -> pd.DataFrame:
    comps = []
    for _, row in df[df["id"] != ref["id"]].iterrows():
        score, pontos = similaridade(ref, row)
        diag = diagnostico_comparacao(ref, row, pontos)
        comps.append({
            "id": row["id"],
            "Nível": classificar_nivel(score),
            "Similaridade (%)": score,
            "Nome da obra": row["nome_obra"],
            "Autor": row["autor"],
            "Site": row["site"],
            "Preço": dinheiro(row["preco"]),
            "Preço numérico": row.get("preco"),
            "Técnica": row["tecnica"],
            "Dimensões": row["dimensoes"],
            "Ano": row["ano_obra"],
            "Imagem": row.get("link_imagem"),
            "Link": row.get("link_obra"),
            "Pts Autor": round(pontos["Autor"], 2),
            "Pts Técnica": round(pontos["Técnica"], 2),
            "Pts Dimensões": round(pontos["Dimensões"], 2),
            "Pts Ano": round(pontos["Ano"], 2),
            "Pts Preço": round(pontos["Preço"], 2),
            **diag,
        })
    if not comps:
        return pd.DataFrame()
    return pd.DataFrame(comps).sort_values("Similaridade (%)", ascending=False)



def aba_comparativo(df: pd.DataFrame) -> None:
    st.header("Comparativo")
    if df.empty or len(df) < 2:
        st.info("Cadastre ao menos duas obras para comparar.")
        return

    with st.expander("Metodologia de similaridade e níveis de comparabilidade", expanded=False):
        st.markdown(
            """
            O comparativo utiliza uma lógica de **abordagem de mercado**: a obra selecionada é confrontada com outras obras do acervo para verificar a força da evidência comparável. A pontuação total varia de **0 a 100 pontos**. A distribuição abaixo constitui uma **calibração teórico-metodológica inicial**: ela traduz uma hierarquia de relevância dos atributos, mas **não corresponde a coeficientes econométricos estimados**.

            | Critério | Peso | Como é usado |
            |---|---:|---|
            | **Autoria / proximidade reputacional** | **40 pts** | Critério principal. Na versão atual, a coincidência de autoria é utilizada como indicador operacional de proximidade reputacional, sem ser tratada como medida completa da reputação do artista. |
            | **Técnica** | **20 pts** | Compara o meio artístico e a materialidade da obra. Técnicas iguais aumentam a similaridade, mas permanecem subordinadas à autoria na hierarquia metodológica. |
            | **Dimensões** | **15 pts** | Compara a escala física pela área aproximada. Quanto menor o desvio proporcional entre as dimensões, maior a pontuação. |
            | **Ano / fase de produção** | **15 pts** | Compara a proximidade temporal da produção, funcionando como aproximação operacional de períodos ou fases produtivas semelhantes. |
            | **Preço** | **10 pts** | Critério complementar. Compara a ordem de grandeza dos preços quando disponíveis, sem transformar o preço observado em valor justo nem fazê-lo dominar a seleção dos comparáveis. |

            **Leitura dos níveis de similaridade:**
            - **Nível 1 — alta similaridade:** de **75 a 100 pontos**. Pela própria ponderação, exige coincidência de autoria, pois uma obra de outro artista pode alcançar no máximo 60 pontos.
            - **Nível 2 — similaridade intermediária:** de **50 a 74,99 pontos**. Reúne evidências parcialmente semelhantes e requer maior julgamento técnico.
            - **Nível 3 — baixa similaridade:** abaixo de **50 pontos**. Deve ser tratado como evidência complementar, com maior necessidade de documentação e análise especializada.

            A pontuação **não determina o valor justo da obra**. Ela apenas organiza e hierarquiza evidências comparáveis para apoiar o julgamento técnico.
            """
        )

    nomes = df.apply(lambda r: f"#{r['id']} · {r['nome_obra']} · {r['autor']}", axis=1).tolist()
    escolha = st.selectbox("Obra de referência", nomes)
    id_ref = int(re.match(r"#(\d+)", escolha).group(1))
    ref = df[df["id"] == id_ref].iloc[0]

    st.subheader("Obra selecionada")
    cimg, cinfo = st.columns([1.1, 1.9], gap="large")
    with cimg:
        img = imagem_para_exibir(ref.get("link_imagem"))
        if img:
            st.image(img, use_container_width=True)
    with cinfo:
        st.markdown(f"## {limpar_texto(ref.get('nome_obra'))}")
        st.caption(f"{limpar_texto(ref.get('autor'))} · {limpar_texto(ref.get('site'))}")
        st.markdown(
            '''
            <style>
            .oc-meta-grid {display:grid; grid-template-columns:repeat(4,minmax(100px,1fr)); gap:12px; margin-top:10px; margin-bottom:14px;}
            .oc-meta-card {background:rgba(255,255,255,0.02); border:1px solid rgba(255,255,255,0.08); border-radius:12px; padding:12px 14px;}
            .oc-meta-label {font-size:0.86rem; opacity:.8; margin-bottom:4px;}
            .oc-meta-value {font-size:0.95rem; font-weight:600; line-height:1.25; word-break:break-word;}
            </style>
            ''',
            unsafe_allow_html=True,
        )
        html = f"""
        <div class="oc-meta-grid">
            <div class="oc-meta-card"><div class="oc-meta-label">Preço</div><div class="oc-meta-value">{dinheiro(ref.get('preco'))}</div></div>
            <div class="oc-meta-card"><div class="oc-meta-label">Técnica</div><div class="oc-meta-value">{limpar_texto(ref.get('tecnica'))}</div></div>
            <div class="oc-meta-card"><div class="oc-meta-label">Dimensões</div><div class="oc-meta-value">{limpar_texto(ref.get('dimensoes'))}</div></div>
            <div class="oc-meta-card"><div class="oc-meta-label">Ano</div><div class="oc-meta-value">{limpar_texto(ref.get('ano_obra'))}</div></div>
        </div>
        """
        st.markdown(html, unsafe_allow_html=True)
        if limpar_texto(ref.get("link_obra")) != SEM_INFO:
            st.link_button("Abrir obra original", str(ref.get("link_obra")))

    f1, f2, f3, f4 = st.columns(4)
    nivel_escolha = f1.multiselect("Níveis", ["Nível 1", "Nível 2", "Nível 3"], default=["Nível 1", "Nível 2", "Nível 3"])
    minimo = f2.slider("Similaridade mínima", 0, 100, 0, 1)
    mesmo_site = f3.checkbox("Somente mesmo site", value=False)
    mesmo_autor = f4.checkbox("Somente mesmo autor", value=False)

    comps = []
    for _, row in df[df["id"] != id_ref].iterrows():
        score, pontos = similaridade(ref, row)
        nivel = classificar_nivel(score)
        if nivel not in nivel_escolha or score < minimo:
            continue
        if mesmo_site and limpar_texto(row.get("site")) != limpar_texto(ref.get("site")):
            continue
        if mesmo_autor and limpar_texto(row.get("autor")) != limpar_texto(ref.get("autor")):
            continue
        comps.append({
            "Nível": nivel,
            "Similaridade": round(score, 2),
            "Nome da obra": row["nome_obra"],
            "Autor": row["autor"],
            "Site": row["site"],
            "Preço": dinheiro(row["preco"]),
            "Técnica": row["tecnica"],
            "Dimensões": row["dimensoes"],
            "Ano": row["ano_obra"],
            "Pts Autor": round(pontos["Autor"], 2),
            "Pts Técnica": round(pontos["Técnica"], 2),
            "Pts Dimensões": round(pontos["Dimensões"], 2),
            "Pts Ano": round(pontos["Ano"], 2),
            "Pts Preço": round(pontos["Preço"], 2),
            "_img": row.get("link_imagem"),
            "_link": row.get("link_obra"),
        })

    if not comps:
        st.warning("Nenhuma obra similar encontrada com os filtros atuais.")
        return

    res = pd.DataFrame(comps).sort_values("Similaridade", ascending=False).reset_index(drop=True)
    st.dataframe(
        res.drop(columns=["_img", "_link"]),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Similaridade": st.column_config.ProgressColumn("Similaridade", min_value=0, max_value=100, format="%.2f"),
        },
    )

    st.subheader("Obras similares")
    st.markdown(
        """
        <style>
        .oc-sim-meta-grid {display:grid; grid-template-columns:repeat(4,minmax(100px,1fr)); gap:12px; margin-top:14px; margin-bottom:14px;}
        .oc-sim-points-grid {display:grid; grid-template-columns:repeat(5,minmax(80px,1fr)); gap:10px; margin-top:10px; margin-bottom:14px;}
        .oc-sim-card {background:rgba(255,255,255,0.02); border:1px solid rgba(255,255,255,0.08); border-radius:12px; padding:12px 14px;}
        .oc-sim-label {font-size:0.86rem; opacity:.8; margin-bottom:4px; font-weight:600;}
        .oc-sim-value {font-size:0.95rem; font-weight:600; line-height:1.25; word-break:break-word;}
        .oc-sim-points .oc-sim-value {font-size:0.92rem;}
        </style>
        """,
        unsafe_allow_html=True,
    )
    for _, row in res.iterrows():
        a, b = st.columns([1.1, 1.9], gap="large")
        with a:
            img = imagem_para_exibir(row["_img"])
            if img:
                st.image(img, use_container_width=True)
        with b:
            st.markdown(f"## {limpar_texto(row['Nome da obra'])}")
            st.caption(f"{limpar_texto(row['Autor'])} · {limpar_texto(row['Site'])}")
            st.progress(min(max(float(row['Similaridade'])/100, 0.0), 1.0), text=f"{row['Nível']} · {row['Similaridade']:.2f}% de similaridade")

            meta_html = f"""
            <div class="oc-sim-meta-grid">
                <div class="oc-sim-card"><div class="oc-sim-label">Preço</div><div class="oc-sim-value">{row['Preço']}</div></div>
                <div class="oc-sim-card"><div class="oc-sim-label">Técnica</div><div class="oc-sim-value">{limpar_texto(row['Técnica'])}</div></div>
                <div class="oc-sim-card"><div class="oc-sim-label">Dimensões</div><div class="oc-sim-value">{limpar_texto(row['Dimensões'])}</div></div>
                <div class="oc-sim-card"><div class="oc-sim-label">Ano</div><div class="oc-sim-value">{limpar_texto(row['Ano'])}</div></div>
            </div>
            <div class="oc-sim-points-grid oc-sim-points">
                <div class="oc-sim-card"><div class="oc-sim-label">Pts Autor</div><div class="oc-sim-value">{row['Pts Autor']}</div></div>
                <div class="oc-sim-card"><div class="oc-sim-label">Pts Técnica</div><div class="oc-sim-value">{row['Pts Técnica']}</div></div>
                <div class="oc-sim-card"><div class="oc-sim-label">Pts Dimensões</div><div class="oc-sim-value">{row['Pts Dimensões']}</div></div>
                <div class="oc-sim-card"><div class="oc-sim-label">Pts Ano</div><div class="oc-sim-value">{row['Pts Ano']}</div></div>
                <div class="oc-sim-card"><div class="oc-sim-label">Pts Preço</div><div class="oc-sim-value">{row['Pts Preço']}</div></div>
            </div>
            """
            st.markdown(meta_html, unsafe_allow_html=True)
            if limpar_texto(row["_link"]) != SEM_INFO:
                st.link_button("Abrir obra similar", row["_link"], key=f"link_{hash(row['_link'])}")
        st.divider()

def aba_informacoes() -> None:
    st.header("Informações")
    st.markdown(
        """
        Esta aba resume a base didática usada pelo Oráculo Cultural para apoiar comparações de obras culturais e discussões de valor justo. O sistema não substitui laudo, perícia, avaliação profissional ou julgamento contábil; ele organiza evidências observáveis e aponta o grau de comparabilidade.
        """
    )
    st.subheader("Valor justo")
    st.markdown(
        """
        **Valor justo** é tratado aqui como medida orientada ao mercado: uma estimativa baseada em evidências de transações, ofertas e dados de obras comparáveis. Para obras de arte e bens culturais, raramente existe mercado perfeitamente ativo para itens idênticos; por isso, o Oráculo usa uma hierarquia prática de evidência.
        """
    )
    st.subheader("Heritage assets e bens culturais")
    st.markdown(
        """
        Obras culturais e heritage assets podem possuir valor financeiro, cultural, simbólico, social e de existência. Quando a mensuração confiável ainda não está disponível, o tratamento deve ser transparente, com documentação das evidências e limites de mensuração.
        """
    )
    st.subheader("Metodologia do programa")
    st.markdown(
        """
        1. **Coleta:** busca obras em fontes selecionadas.  
        2. **Normalização:** padroniza preço, dimensão, técnica e ano.  
        3. **Banco de evidências:** armazena dados em SQLite local.  
        4. **Comparação:** calcula similaridade objetiva por autoria, técnica, dimensões, ano e preço, conforme a ponderação metodológica definida.  
        5. **Classificação:** separa evidências em Nível 1, Nível 2 e Nível 3.
        """
    )
    st.subheader("Limites")
    st.markdown(
        """
        A pontuação de similaridade não é preço final. Ela é evidência organizada para triagem, documentação e comparação preliminar. Obras em Nível 3 exigem maior julgamento técnico, laudo, histórico de transações, proveniência, estado de conservação, raridade e autenticidade.
        """
    )


def aba_quem_somos() -> None:
    st.header("Quem somos")
    st.markdown(
        """
        O **Oráculo Cultural** é uma ferramenta acadêmica desenvolvida para apoiar a organização de evidências de mercado e a comparação de obras culturais em análises de valor justo.

        **Equipe acadêmica**

        - **Eduardo Guilherme de Matos Santos** — aluno de graduação em Ciências Contábeis da Universidade de Brasília (UnB).
        - **Professora Doutora Fátima de Souza Freire** — Departamento de Ciências Contábeis e Atuariais (UnB).
        - **Professor Doutor Jorge Madeira Nogueira** — Departamento de Economia da Universidade de Brasília (UnB).
        """
    )


def main() -> None:
    st.set_page_config(page_title=APP_NAME, page_icon="🔮", layout="wide")
    inicializar_banco()
    df = carregar_acervo()
    tabs = st.tabs(["Início", "Coleta", "Acervo", "Comparativo", "Informações", "Quem somos"])
    with tabs[0]:
        aba_inicio(df)
    with tabs[1]:
        aba_coleta(df)
    with tabs[2]:
        aba_acervo(carregar_acervo())
    with tabs[3]:
        aba_comparativo(carregar_acervo())
    with tabs[4]:
        aba_informacoes()
    with tabs[5]:
        aba_quem_somos()


if __name__ == "__main__":
    main()
