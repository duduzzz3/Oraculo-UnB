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
from pathlib import Path
from typing import Any

import pandas as pd
import requests
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
    if any(x in t for x in ["etching", "agua-forte", "aguaforte", "acid etching"]):
        return "Água-forte"
    if any(x in t for x in ["aquatint", "agua-tinta", "aguatinta"]):
        return "Água-tinta"
    if any(x in t for x in ["drypoint", "ponta-seca", "pontaseca"]):
        return "Ponta-seca"
    if any(x in t for x in ["monotype", "monoprint", "monotipia"]):
        return "Monotipia"
    if any(x in t for x in ["engraving", "intaglio", "gravura em metal", "metal engraving", "burin"]):
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


def preparar_dataframe_obras(df: pd.DataFrame) -> pd.DataFrame:
    df = adaptar_colunas_entrada(df)
    for col in ["Nome da obra", "Autor", "Ano da obra", "Técnica", "Técnica original", "Dimensões", "Preço", "Descrição", "Link da obra", "Link da imagem da obra"]:
        if col not in df.columns:
            df[col] = SEM_INFO
    out = pd.DataFrame()
    out["Nome da obra"] = df.apply(lambda r: corrigir_nome_obra(r.get("Nome da obra"), r.get("Link da obra")), axis=1)
    out["Autor"] = df.apply(lambda r: corrigir_autor(r.get("Autor"), r.get("Nome da obra")), axis=1)
    out["Preço"] = df["Preço"].apply(parse_preco)
    out["Dimensões"] = df["Dimensões"].apply(padronizar_dimensoes)
    out["Técnica original"] = df["Técnica original"].apply(limpar_texto)
    out["Técnica"] = df["Técnica original"].apply(normalizar_tecnica)
    out["Ano da obra"] = df["Ano da obra"].apply(normalizar_ano)
    out["Descrição"] = df["Descrição"].apply(limpar_texto)
    out["Link da obra"] = df["Link da obra"].apply(limpar_texto)
    out["Link da imagem da obra"] = df["Link da imagem da obra"].apply(limpar_texto)
    return out


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
                {"rate": float(valor), "updated_at": datetime.now().isoformat(timespec="seconds")},
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
    raw = cache.get("updated_at")
    try:
        if raw:
            return pd.to_datetime(raw).strftime("%d/%m/%Y %H:%M:%S")
    except Exception:
        pass
    return SEM_INFO

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
    df_norm = preparar_dataframe_obras(df).rename(columns=UI_TO_DB)
    agora = datetime.now().isoformat(timespec="seconds")
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


def executar_scraper(site: str, max_obras: int, headless: bool) -> tuple[bool, str, pd.DataFrame]:
    config = SITES[site]
    src = SCRAPERS_DIR / config["script"]
    workdir = COLETAS_DIR / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{chave_coluna(site)}"
    workdir.mkdir(parents=True, exist_ok=True)
    runner = config["runner"]
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["MAX_OBRAS"] = str(max_obras)
    if runner in {"saatchi_cli", "legacy_playwright"}:
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
    if texto == SEM_INFO:
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
    agora = datetime.now().isoformat(timespec="seconds")
    with conectar() as conn:
        rows = conn.execute("SELECT * FROM obras").fetchall()
        for row in rows:
            novo_nome = corrigir_nome_obra(row["nome_obra"], row["link_obra"])
            novo_autor = corrigir_autor(row["autor"], novo_nome)
            tecnica_original = limpar_texto(row["tecnica_original"] or row["tecnica"])
            nova_tecnica = normalizar_tecnica(tecnica_original)
            novo_ano = normalizar_ano(row["ano_obra"])
            nova_dim = padronizar_dimensoes(row["dimensoes"])
            if any([novo_nome != row["nome_obra"], novo_autor != row["autor"], nova_tecnica != row["tecnica"], novo_ano != row["ano_obra"], nova_dim != row["dimensoes"]]):
                conn.execute(
                    "UPDATE obras SET nome_obra=?, autor=?, tecnica=?, tecnica_original=?, ano_obra=?, dimensoes=?, atualizado_em=? WHERE id=?",
                    (novo_nome, novo_autor, nova_tecnica, tecnica_original, novo_ano, nova_dim, agora, row["id"]),
                )
                alterados += 1
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


def calcular_area(dim: str) -> float | None:
    nums = [float(x.replace(",", ".")) for x in re.findall(r"\d+(?:[.,]\d+)?", str(dim))]
    if len(nums) >= 2:
        return nums[0] * nums[1]
    return None


def similaridade(obra: pd.Series, comp: pd.Series) -> tuple[float, dict[str, float]]:
    pontos = {"Técnica": 0.0, "Dimensões": 0.0, "Preço": 0.0, "Ano": 0.0, "Autor": 0.0}
    if limpar_texto(obra.get("tecnica")) == limpar_texto(comp.get("tecnica")) and limpar_texto(obra.get("tecnica")) != SEM_INFO:
        pontos["Técnica"] = 28
    area1, area2 = calcular_area(obra.get("dimensoes", "")), calcular_area(comp.get("dimensoes", ""))
    if area1 and area2:
        ratio = min(area1, area2) / max(area1, area2)
        pontos["Dimensões"] = 25 * ratio
    p1, p2 = obra.get("preco"), comp.get("preco")
    try:
        if pd.notna(p1) and pd.notna(p2) and float(p1) > 0 and float(p2) > 0:
            dif = abs(math.log(float(p1)) - math.log(float(p2)))
            pontos["Preço"] = max(0, 22 * (1 - dif / 2.5))
    except Exception:
        pass
    a1, a2 = normalizar_ano(obra.get("ano_obra")), normalizar_ano(comp.get("ano_obra"))
    if a1 != SEM_INFO and a2 != SEM_INFO:
        dif_ano = abs(int(a1) - int(a2))
        pontos["Ano"] = max(0, 15 * (1 - dif_ano / 50))
    if limpar_texto(obra.get("autor")).lower() == limpar_texto(comp.get("autor")).lower() and limpar_texto(obra.get("autor")) != SEM_INFO:
        pontos["Autor"] = 10
    total = round(sum(pontos.values()), 2)
    return total, pontos


def classificar_nivel(score: float) -> str:
    if score >= 80:
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
        "Técnica igual": pontos.get("Técnica", 0) > 0,
        "Autor igual": pontos.get("Autor", 0) > 0,
        "Proximidade área": area_txt,
        "Diferença preço": preco_txt,
        "Diferença ano": dif_ano_txt,
    }



def aba_inicio(df: pd.DataFrame) -> None:
    st.title("🔮 Oráculo Cultural")
    st.markdown("Coleta obras culturais e organiza comparáveis para análise contábil a valor justo.")

    ultima_coleta_obras = SEM_INFO
    if not df.empty and "coletado_em" in df.columns and df["coletado_em"].notna().any():
        try:
            dt = pd.to_datetime(df["coletado_em"], errors="coerce").max()
            if pd.notna(dt):
                ultima_coleta_obras = dt.strftime("%d/%m/%Y %H:%M:%S")
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
    i1.info(f"**Última coleta de obras:** {ultima_coleta_obras}")
    i2.info(f"**Última atualização da cotação do dólar:** {ultima_atualizacao_dolar()}")

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
    st.caption("Use 0 para tentar coletar o máximo disponível. Os scrapers ativos são ArtSoul, Blombo, Gagosian e Saatchi Art.")
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
        preco_sem_info = st.checkbox("Preço sem informação", value=False)
        preco = None if preco_sem_info else st.number_input("Preço em R$", min_value=0.0, step=100.0)
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
            "Pts Técnica": round(pontos["Técnica"], 2),
            "Pts Dimensões": round(pontos["Dimensões"], 2),
            "Pts Preço": round(pontos["Preço"], 2),
            "Pts Ano": round(pontos["Ano"], 2),
            "Pts Autor": round(pontos["Autor"], 2),
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

    with st.expander("Metodologia de similaridade e níveis de valor justo", expanded=False):
        st.markdown(
            """
            O comparativo utiliza uma lógica de **abordagem de mercado**: a obra selecionada é confrontada com outras obras do acervo para verificar a força da evidência comparável. A pontuação total é de **0 a 100 pontos**, distribuída de forma objetiva entre os critérios abaixo.

            | Critério | Peso | Como é usado |
            |---|---:|---|
            | **Técnica** | **28 pts** | Compara o meio artístico e a materialidade da obra. Técnicas iguais ou muito próximas aumentam fortemente a similaridade. |
            | **Dimensões** | **25 pts** | Compara a escala física da obra pela área aproximada. Quanto mais próximas as dimensões, maior a pontuação. |
            | **Preço** | **22 pts** | Compara a ordem de grandeza dos preços observados, reduzindo a pontuação quando os valores estão muito distantes. |
            | **Ano** | **15 pts** | Compara a proximidade temporal da produção, considerando fase do artista, contexto e período de mercado. |
            | **Autor** | **10 pts** | Dá reforço quando a obra comparável é do mesmo autor, mas sem substituir a análise dos demais atributos. |

            **Leitura dos níveis:**
            - **Nível 1:** similaridade igual ou superior a 80. Evidência muito próxima, com pouco ajuste.
            - **Nível 2:** similaridade entre 50 e 79,99. Evidência observável comparável, mas com ajustes relevantes.
            - **Nível 3:** similaridade abaixo de 50. Baixa comparabilidade direta; exige maior julgamento técnico e documentação complementar.
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
            "Pts Técnica": round(pontos["Técnica"], 2),
            "Pts Dimensões": round(pontos["Dimensões"], 2),
            "Pts Preço": round(pontos["Preço"], 2),
            "Pts Ano": round(pontos["Ano"], 2),
            "Pts Autor": round(pontos["Autor"], 2),
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
                <div class="oc-sim-card"><div class="oc-sim-label">Pts Técnica</div><div class="oc-sim-value">{row['Pts Técnica']}</div></div>
                <div class="oc-sim-card"><div class="oc-sim-label">Pts Dimensões</div><div class="oc-sim-value">{row['Pts Dimensões']}</div></div>
                <div class="oc-sim-card"><div class="oc-sim-label">Pts Preço</div><div class="oc-sim-value">{row['Pts Preço']}</div></div>
                <div class="oc-sim-card"><div class="oc-sim-label">Pts Ano</div><div class="oc-sim-value">{row['Pts Ano']}</div></div>
                <div class="oc-sim-card"><div class="oc-sim-label">Pts Autor</div><div class="oc-sim-value">{row['Pts Autor']}</div></div>
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
        4. **Comparação:** calcula similaridade objetiva por técnica, dimensão, preço, ano e autoria.  
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
