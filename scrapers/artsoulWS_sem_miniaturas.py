from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Iterable
from urllib.parse import parse_qsl, urlencode, urljoin, urlsplit, urlunsplit

import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = "https://artsoul.com.br"
COLECAO_URL = "https://artsoul.com.br/obras/tecnica/pintura?tecnica%5B0%5D=pintura"
SEM_INFORMACAO = "Sem informação"

ARTWORK_PATH_RE = re.compile(r"^/obras/(?!tecnica(?:/|$)|tema(?:/|$)|categoria(?:/|$)|preco(?:/|$)|medida(?:/|$)|cidade(?:/|$)|selecoes?(?:/|$)|design(?:/|$)|$)[a-z0-9][a-z0-9\-]*$", re.I)
PRICE_RE = re.compile(r"R\$\s*\d{1,3}(?:\.\d{3})*,\d{2}")
YEAR_RE = re.compile(r"\b(?:19|20)\d{2}\b")
DIM_RE = re.compile(
    r"\b\d{1,4}(?:[,.]\d{1,2})?\s*(?:x|×|X|por)\s*\d{1,4}(?:[,.]\d{1,2})?(?:\s*(?:x|×|X|por)\s*\d{1,4}(?:[,.]\d{1,2})?)?\s*cm\b",
    re.I,
)
MATERIAL_RE = re.compile(
    r"\b(óleo|oleo|acrílic|acrilic|aquarela|guache|pastel|giz|caneta|hidrogr|nanquim|carvão|carvao|grafite|linho|tela|papel|madeira|colagem|spray|pigmento|resina|encáustica|encaustica|têmpera|tempera|vinil|mista)\b",
    re.I,
)
STOP_DESCRIPTION_RE = re.compile(
    r"^(Procurando a obra|Para mais|Inspiração|Veja também|Mais Obras|Qual a forma|Compartilhar|Falar com|Adicionar aos Favoritos)",
    re.I,
)
BAD_IMAGE_RE = re.compile(r"(logo|pagamento|seguranca|segurança|facebook|instagram|whatsapp|twitter|icon|sprite|avatar|placeholder|blank)", re.I)
IMAGE_EXT_RE = re.compile(r"\.(?:jpe?g|png|webp|avif)(?:[?#].*)?$", re.I)

# Rol fechado de técnicas aceitas pelo destino. A função normalize_technique_to_allowed()
# sempre retorna um destes valores.
ALLOWED_TECHNIQUES = (
    "Pintura acrílica",
    "Pintura a óleo",
    "Aquarela",
    "Guache",
    "Técnica mista",
    "Colagem",
    "Serigrafia",
    "Litografia",
    "Gravura",
    "Impressão fine art",
    "Giclée",
    "Fotografia",
    "Arte digital",
    "Escultura",
    "Cerâmica",
    "Metal",
    "Madeira",
    "Resina",
    "Têxtil",
    "Outros",
    "Sem informação",
)

# Mantém exatamente o mesmo padrão de colunas do Excel usado no scraper original.
# Isso evita quebrar sistemas/importadores que já leem a planilha antiga.
EXCEL_COLUMNS = [
    "Título",
    "Autor",
    "Ano",
    "Técnica",
    "Dimensões",
    "Preço",
    "Descrição",
    "Link da obra",
    "Link da imagem da obra",
]


@dataclass
class Obra:
    nome_da_obra: str = SEM_INFORMACAO
    autor: str = SEM_INFORMACAO
    preco_brl: str = SEM_INFORMACAO
    preco_numero: float | str = SEM_INFORMACAO
    dimensoes: str = SEM_INFORMACAO
    tecnica: str = SEM_INFORMACAO
    ano_da_obra: str = SEM_INFORMACAO
    descricao: str = SEM_INFORMACAO
    imagem_da_obra: str = SEM_INFORMACAO
    url_da_obra: str = SEM_INFORMACAO


def clean(value: Any) -> str:
    if value is None:
        return ""
    text = str(value)
    text = text.replace("\ufeff", "").replace("\xa0", " ")
    text = re.sub(r"[ \t\r\f\v]+", " ", text)
    text = re.sub(r"\n\s+", "\n", text)
    return text.strip()


def norm_space(value: Any) -> str:
    return re.sub(r"\s+", " ", clean(value)).strip()


def sem_info(value: Any) -> str:
    text = norm_space(value)
    return text if text else SEM_INFORMACAO


def absolute_url(url: str | None, base: str = BASE_URL) -> str:
    if not url:
        return ""
    url = clean(url)
    if url.startswith("//"):
        return "https:" + url
    return urljoin(base, url)


def linhas_texto(soup: BeautifulSoup | Tag) -> list[str]:
    return [norm_space(x) for x in soup.get_text("\n", strip=True).splitlines() if norm_space(x)]


def preco_para_float(preco: str) -> float | str:
    preco = clean(preco)
    m = PRICE_RE.search(preco)
    if not m:
        return SEM_INFORMACAO
    raw = m.group(0).replace("R$", "").strip()
    raw = raw.replace(".", "").replace(",", ".")
    try:
        return float(raw)
    except ValueError:
        return SEM_INFORMACAO


def make_session() -> requests.Session:
    retry = Retry(
        total=4,
        connect=4,
        read=4,
        status=4,
        backoff_factor=0.8,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=("GET", "HEAD"),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=12, pool_maxsize=12)
    session = requests.Session()
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0 Safari/537.36"
            ),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
            "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
            "Cache-Control": "no-cache",
        }
    )
    return session


def fetch(session: requests.Session, url: str, timeout: int = 35) -> str:
    resp = session.get(url, timeout=timeout)
    resp.raise_for_status()
    if not resp.encoding or resp.encoding.lower() in {"iso-8859-1", "ascii"}:
        resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


def parse_json_scripts(soup: BeautifulSoup) -> list[Any]:
    out: list[Any] = []
    for script in soup.select("script[type='application/ld+json'], script[type='application/json']"):
        txt = clean(script.string or script.get_text(" ", strip=True))
        if not txt:
            continue
        try:
            out.append(json.loads(txt))
        except Exception:
            # Alguns sites inserem mais de um objeto no mesmo script; tentamos recuperar objetos isolados.
            for m in re.finditer(r"\{.*?\}", txt, flags=re.S):
                try:
                    out.append(json.loads(m.group(0)))
                except Exception:
                    pass
    return out


def flatten_json(obj: Any) -> Iterable[Any]:
    if isinstance(obj, dict):
        yield obj
        for value in obj.values():
            yield from flatten_json(value)
    elif isinstance(obj, list):
        for item in obj:
            yield from flatten_json(item)


def first_json_value(jsons: list[Any], keys: tuple[str, ...]) -> str:
    for obj in flatten_json(jsons):
        if not isinstance(obj, dict):
            continue
        for key in keys:
            if key not in obj:
                continue
            value = obj.get(key)
            if isinstance(value, str) and norm_space(value):
                return norm_space(value)
            if isinstance(value, dict):
                for sub_key in ("name", "url", "contentUrl"):
                    sub = value.get(sub_key)
                    if isinstance(sub, str) and norm_space(sub):
                        return norm_space(sub)
            if isinstance(value, list):
                for item in value:
                    if isinstance(item, str) and norm_space(item):
                        return norm_space(item)
                    if isinstance(item, dict):
                        for sub_key in ("name", "url", "contentUrl"):
                            sub = item.get(sub_key)
                            if isinstance(sub, str) and norm_space(sub):
                                return norm_space(sub)
    return ""


def label_value(lines: list[str], labels: tuple[str, ...], max_lookahead: int = 5) -> str:
    normalized_labels = tuple(re.escape(label).replace("\\ ", r"\s+") for label in labels)
    label_re = re.compile(rf"^({'|'.join(normalized_labels)})(?:\s*\([^)]*\))?\s*:?\s*(.*)$", re.I)
    stop_re = re.compile(r"^(Ano|Técnicas?|Tecnicas?|Temas?|Cores?|Medidas|Descrição|Valor Total|Compartilhar|Comprar|Adicionar|Falar)\b", re.I)

    for i, line in enumerate(lines):
        m = label_re.match(line)
        if not m:
            continue
        inline = norm_space(m.group(2))
        if inline and not stop_re.match(inline):
            return inline
        values: list[str] = []
        for nxt in lines[i + 1 : i + 1 + max_lookahead]:
            if stop_re.match(nxt) and values:
                break
            if stop_re.match(nxt) and not values:
                continue
            if nxt in {":", "/", "-"}:
                continue
            values.append(nxt)
            # Para os rótulos da ficha técnica do Artsoul, normalmente o valor é só a próxima linha.
            if labels[0].lower().startswith(("ano", "técn", "tecn", "med")):
                break
        if values:
            return norm_space(" ".join(values))
    return ""


def clean_title(title: str, url: str = "") -> str:
    title = norm_space(title)
    title = re.sub(r"\s*-\s*Artsoul(?:\s*-\s*Marketplace.*)?$", "", title, flags=re.I)
    title = re.sub(r"\s*-\s*Marketplace de Arte Contemporânea$", "", title, flags=re.I)
    return title.strip(" -–—")


def extract_title(soup: BeautifulSoup, lines: list[str], jsons: list[Any], url: str) -> str:
    candidates = [
        first_json_value(jsons, ("name", "headline")),
        soup.select_one("h1").get_text(" ", strip=True) if soup.select_one("h1") else "",
        soup.select_one("meta[property='og:title']").get("content", "") if soup.select_one("meta[property='og:title']") else "",
        soup.title.string if soup.title and soup.title.string else "",
    ]
    for line in lines[:140]:
        if line.lower() in {"home", "obras", "valor total", "ficha técnica", "descricao", "descrição"}:
            continue
        if len(line) >= 2 and not PRICE_RE.search(line) and not re.match(r"^(SÃO PAULO|RIO|BRASIL|/|\d{4})$", line, re.I):
            # Linha após breadcrumb normalmente é o h1 quando h1 não foi capturado.
            if soup.select_one("h1") and norm_space(soup.select_one("h1").get_text(" ", strip=True)).lower() == line.lower():
                candidates.append(line)
    for c in candidates:
        c = clean_title(c, url)
        if c and c.lower() not in {"artsoul", "obras de arte", "obras"} and not PRICE_RE.search(c):
            return c
    slug = urlsplit(url).path.rstrip("/").split("/")[-1]
    slug = re.sub(r"-\d+$", "", slug)
    return sem_info(slug.replace("-", " ").title())


def extract_author(soup: BeautifulSoup, lines: list[str], jsons: list[Any], title: str) -> str:
    candidates: list[str] = []
    author_json = first_json_value(jsons, ("author", "creator", "artist", "brand"))
    if author_json:
        candidates.append(author_json)

    # Link de artista é o seletor mais estável na página de detalhe.
    for a in soup.select("a[href*='/artistas/'], a[href*='/artista/']"):
        text = norm_space(a.get_text(" ", strip=True))
        if text and text.lower() not in {"artistas", "artista"}:
            candidates.append(text)

    rotulo = label_value(lines, ("Artista", "Autor", "Criador"))
    if rotulo:
        candidates.append(rotulo)

    # Fallback: a linha logo após o título costuma ser o autor.
    title_norm = title.lower()
    for idx, line in enumerate(lines[:170]):
        if line.lower() == title_norm:
            for nxt in lines[idx + 1 : idx + 8]:
                if re.match(r"^(Valor Total|R\$|Comprar|Adicionar|Falar|Compartilhar|Ano:|Ficha Técnica|/|\d{4})", nxt, re.I):
                    continue
                if 2 <= len(nxt) <= 80 and not PRICE_RE.search(nxt):
                    candidates.append(nxt)
                    break

    for candidate in candidates:
        candidate = norm_space(candidate)
        candidate = re.sub(r"\s*-\s*Artsoul.*$", "", candidate, flags=re.I).strip()
        if candidate and candidate.lower() != title_norm and len(candidate) <= 100:
            return candidate
    return SEM_INFORMACAO


def extract_price(lines: list[str], soup: BeautifulSoup, jsons: list[Any]) -> tuple[str, float | str]:
    json_price = first_json_value(jsons, ("price", "lowPrice", "highPrice"))
    if json_price:
        if re.fullmatch(r"\d+(?:\.\d+)?", json_price):
            try:
                n = float(json_price)
                return f"R$ {n:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), n
            except ValueError:
                pass
        m = PRICE_RE.search(json_price)
        if m:
            return m.group(0), preco_para_float(m.group(0))

    for idx, line in enumerate(lines):
        if re.match(r"^Valor Total$", line, re.I):
            for nxt in lines[idx + 1 : idx + 5]:
                m = PRICE_RE.search(nxt)
                if m:
                    return m.group(0), preco_para_float(m.group(0))

    # Evita capturar parcelas primeiro; em geral o preço cheio aparece antes da frase "em até 12x".
    all_text = "\n".join(lines)
    matches = PRICE_RE.findall(all_text)
    if matches:
        return matches[0], preco_para_float(matches[0])

    for node in soup.select("[class*='price'], [class*='preco'], [class*='valor'], h2, strong, span"):
        m = PRICE_RE.search(node.get_text(" ", strip=True))
        if m:
            return m.group(0), preco_para_float(m.group(0))
    return SEM_INFORMACAO, SEM_INFORMACAO


def extract_year(lines: list[str], jsons: list[Any]) -> str:
    for key in ("dateCreated", "productionDate", "datePublished"):
        val = first_json_value(jsons, (key,))
        m = YEAR_RE.search(val)
        if m:
            return m.group(0)

    rotulo = label_value(lines, ("Ano",), max_lookahead=3)
    m = YEAR_RE.search(rotulo)
    if m:
        return m.group(0)

    # No detalhe do Artsoul aparece cidade / ano antes do h1.
    for line in lines[:130]:
        if re.fullmatch(r"(?:19|20)\d{2}", line):
            return line
    return SEM_INFORMACAO


def normalize_dimension_number(value: str) -> str:
    """Normaliza número de medida para o padrão brasileiro, sem unidade por item."""
    value = norm_space(value).lower().replace("cm", "").strip()
    value = value.replace(".", ",")
    if "," in value:
        value = value.rstrip("0").rstrip(",")
    return value


def dimension_numbers(value: str) -> list[str]:
    """Extrai até três números de um texto de medidas, preservando a ordem A/L/P."""
    value = norm_space(value)
    if not value:
        return []

    # Remove partes que atrapalham a leitura, mas mantém a ordem dos valores.
    value = re.sub(r"\bMedidas?\s*(?:\(A\s*/\s*L\s*/\s*P\))?\s*:?", " ", value, flags=re.I)
    value = re.sub(r"\b(?:Altura|Largura|Profundidade|Di[aâ]metro|Comprimento)\b\s*:?", " ", value, flags=re.I)
    value = value.replace("×", "x")
    value = re.sub(r"\bpor\b", "x", value, flags=re.I)

    # Preferência: números que aparecem junto de cm ou em uma sequência com x.
    cm_nums = re.findall(r"(?<!\d)(\d{1,4}(?:[,.]\d{1,2})?)\s*cm\b", value, flags=re.I)
    if len(cm_nums) >= 2:
        return [normalize_dimension_number(n) for n in cm_nums[:3]]

    seq = re.search(
        r"(?<!\d)(\d{1,4}(?:[,.]\d{1,2})?)\s*(?:cm)?\s*(?:x|/)\s*"
        r"(\d{1,4}(?:[,.]\d{1,2})?)\s*(?:cm)?"
        r"(?:\s*(?:x|/)\s*(\d{1,4}(?:[,.]\d{1,2})?)\s*(?:cm)?)?",
        value,
        flags=re.I,
    )
    if seq:
        nums = [g for g in seq.groups() if g]
        if len(nums) >= 2:
            return [normalize_dimension_number(n) for n in nums[:3]]

    # Fallback final: só aceita números soltos quando o texto deixa claro que são cm.
    if "cm" in value.lower():
        loose = re.findall(r"(?<!\d)(\d{1,4}(?:[,.]\d{1,2})?)(?!\d)", value)
        if len(loose) >= 2:
            return [normalize_dimension_number(n) for n in loose[:3]]
    return []


def format_dimensions_alp(nums: list[str]) -> str:
    """Formata no padrão Altura x Largura x Profundidade cm.

    O ArtSoul usa o rótulo "Medidas (A/L/P)". Em obras bidimensionais, o site
    frequentemente informa apenas A e L; nesse caso, a profundidade é padronizada
    como 0 para manter a estrutura A x L x P cm.
    """
    nums = [n for n in nums if n]
    if len(nums) >= 3:
        return f"{nums[0]} x {nums[1]} x {nums[2]} cm"
    if len(nums) == 2:
        return f"{nums[0]} x {nums[1]} x 0 cm"
    return ""


def normalize_dimensions_text(value: str) -> str:
    return format_dimensions_alp(dimension_numbers(value))


def dimension_text_from_lines(lines: list[str]) -> str:
    label_re = re.compile(r"^Medidas?\s*(?:\(A\s*/\s*L\s*/\s*P\))?\s*:?\s*(.*)$", re.I)
    stop_re = re.compile(r"^(Ano|T[eé]cnicas?|Tecnicas?|Temas?|Cores?|Descri[cç][aã]o|Valor Total|Compartilhar|Comprar|Adicionar|Falar)\b", re.I)

    for i, line in enumerate(lines):
        m = label_re.match(line)
        if not m:
            continue
        parts: list[str] = []
        inline = norm_space(m.group(1))
        if inline:
            parts.append(inline)
        for nxt in lines[i + 1 : i + 9]:
            if stop_re.match(nxt):
                break
            if re.search(r"\d", nxt) and ("cm" in nxt.lower() or len(parts) < 3):
                parts.append(nxt)
                if len(dimension_numbers(" ".join(parts))) >= 3:
                    break
                continue
            if parts:
                break
        candidate = norm_space(" ".join(parts))
        if candidate:
            return candidate
    return ""


def dimension_text_from_dom(soup: BeautifulSoup) -> str:
    # O HTML do ArtSoul costuma ficar em uma linha do tipo:
    # <div>Medidas (A/L/P):</div><p>72 cm ... 62 cm</p>
    # Por isso procuramos o rótulo e subimos poucos níveis até achar os valores.
    label_strings = soup.find_all(string=re.compile(r"Medidas?\s*(?:\(A\s*/\s*L\s*/\s*P\))?", re.I))
    for label in label_strings:
        parent = getattr(label, "parent", None)
        for _ in range(5):
            if not parent:
                break
            text = norm_space(parent.get_text(" ", strip=True))
            if "medida" in text.lower() and "cm" in text.lower():
                # Corta o bloco no próximo rótulo da ficha técnica, evitando capturar ano/preço.
                m = re.search(
                    r"Medidas?\s*(?:\(A\s*/\s*L\s*/\s*P\))?\s*:?\s*(.*?)"
                    r"(?:\b(?:Ano|T[eé]cnicas?|Tecnicas?|Temas?|Cores?|Descri[cç][aã]o|Valor Total)\b|$)",
                    text,
                    flags=re.I,
                )
                return m.group(1) if m else text
            parent = parent.parent
    return ""


def description_lines(lines: list[str]) -> list[str]:
    start = None
    for i, line in enumerate(lines):
        if re.match(r"^Descri[cç][aã]o$", line, re.I):
            start = i + 1
            break
    if start is None:
        return []
    out: list[str] = []
    for line in lines[start:]:
        if STOP_DESCRIPTION_RE.match(line):
            break
        if re.match(r"^#{1,6}\s*", line):
            continue
        out.append(line)
    return out


def extract_description(desc_lines: list[str]) -> str:
    description = norm_space(" ".join(desc_lines))
    return description if description else SEM_INFORMACAO


def extract_dimensions(soup: BeautifulSoup, lines: list[str], desc_lines: list[str]) -> str:
    # 1) Preferência absoluta: ficha técnica do site, que explicita A/L/P.
    for source in (dimension_text_from_dom(soup), dimension_text_from_lines(lines)):
        normalized = normalize_dimensions_text(source)
        if normalized:
            return normalized

    # 2) Fallback: descrição da obra.
    for line in desc_lines:
        normalized = normalize_dimensions_text(line)
        if normalized:
            return normalized

    # 3) Fallback global, usado apenas quando a ficha/descrição falharem.
    normalized = normalize_dimensions_text("\n".join(lines))
    if normalized:
        return normalized
    return SEM_INFORMACAO


def technique_signal_groups(text: str) -> set[str]:
    text = text.lower()
    groups: set[str] = set()
    patterns = {
        "acrilica": r"acr[ií]lic[ao]",
        "oleo": r"\b[óo]leo\b|oil\b",
        "aquarela": r"aquarela|watercolor|watercolour",
        "guache": r"guache|gouache",
        "colagem": r"colagem|collage",
        "serigrafia": r"serigrafia|silkscreen|screen\s*print",
        "litografia": r"litografia|lithograph",
        "gravura": r"gravura|xilogravura|linoleogravura|calcogravura|[áa]gua[- ]forte|ponta seca|etching|woodcut|linocut",
        "fine_art": r"fine\s*art",
        "giclee": r"gicl[ée]e|giclee",
        "fotografia": r"fotograf|photograph|photo\b",
        "digital": r"arte\s*digital|digital\s*art|\bdigital\b|nft\b",
        "escultura": r"escultura|sculpture",
        "ceramica": r"cer[aâ]mica|ceramic",
        "metal": r"\bmetal\b|a[çc]o|ferro|bronze|alum[ií]nio|lat[aã]o|cobre",
        "madeira": r"madeira|wood\b|mdf\b",
        "resina": r"resina|resin",
        "textil": r"t[eê]xtil|tecido|algod[aã]o|linho|bordado|tape[çc]aria|l[ãa]\b|textile|fabric|embroider",
        "pastel": r"pastel|giz",
        "caneta": r"caneta|hidrogr[aá]fica|marcador|marker",
        "nanquim": r"nanquim|nanquin|india ink",
        "grafite_carvao": r"grafite|carv[aã]o|charcoal|graphite",
        "spray": r"spray|aerossol|aerosol",
        "tempera": r"t[eê]mpera|tempera|enc[aá]ustica|encaustic",
    }
    for group, pattern in patterns.items():
        if re.search(pattern, text, flags=re.I):
            groups.add(group)
    return groups


def normalize_technique_to_allowed(raw_text: str) -> str:
    text = norm_space(raw_text)
    if not text:
        return SEM_INFORMACAO
    lower = text.lower()
    groups = technique_signal_groups(text)

    # Técnica mista explícita ou múltiplos meios de desenho/pintura na mesma obra.
    mixed_media_groups = {
        "acrilica", "oleo", "aquarela", "guache", "colagem", "pastel", "caneta",
        "nanquim", "grafite_carvao", "spray", "tempera",
    }
    if re.search(r"t[eé]cnica\s*mista|mixed\s*media|m[íi]dia\s*mista", lower, flags=re.I):
        return "Técnica mista"
    if len(groups.intersection(mixed_media_groups)) >= 2:
        return "Técnica mista"

    # Impressão/edições: termos mais específicos antes dos genéricos.
    if "giclee" in groups:
        return "Giclée"
    if "fine_art" in groups:
        return "Impressão fine art"
    if "serigrafia" in groups:
        return "Serigrafia"
    if "litografia" in groups:
        return "Litografia"
    if "gravura" in groups:
        return "Gravura"

    # Pintura/desenho.
    if "acrilica" in groups:
        return "Pintura acrílica"
    if "oleo" in groups:
        return "Pintura a óleo"
    if "aquarela" in groups:
        return "Aquarela"
    if "guache" in groups:
        return "Guache"
    if "colagem" in groups:
        return "Colagem"

    # Foto/digital.
    if "fotografia" in groups:
        return "Fotografia"
    if "digital" in groups:
        return "Arte digital"

    # Objeto/materialidade. Só cai aqui quando não houve técnica de pintura/edição acima.
    if "escultura" in groups:
        return "Escultura"
    if "ceramica" in groups:
        return "Cerâmica"
    if "metal" in groups:
        return "Metal"
    if "madeira" in groups:
        return "Madeira"
    if "resina" in groups:
        return "Resina"
    if "textil" in groups:
        return "Têxtil"

    # O site às vezes informa apenas "Pintura". Como o rol não aceita a categoria genérica,
    # mantemos um valor permitido sem inventar se é óleo/acrílica/etc.
    if re.search(r"\bpintura\b|painting", lower, flags=re.I):
        return "Outros"

    return "Outros"


def extract_technique(lines: list[str], desc_lines: list[str]) -> str:
    ficha = label_value(lines, ("Técnicas", "Técnica", "Tecnicas", "Tecnica"), max_lookahead=4)
    material_candidates: list[str] = []
    for line in desc_lines:
        if PRICE_RE.search(line) or YEAR_RE.fullmatch(line) or DIM_RE.search(line):
            continue
        if MATERIAL_RE.search(line) or technique_signal_groups(line):
            material_candidates.append(line)

    raw_text = " | ".join([x for x in [ficha, *material_candidates] if norm_space(x)])
    return normalize_technique_to_allowed(raw_text)


def srcset_best(srcset: str) -> str:
    best_url = ""
    best_width = -1
    for part in srcset.split(","):
        items = part.strip().split()
        if not items:
            continue
        url = items[0]
        width = 0
        if len(items) > 1:
            m = re.search(r"(\d+)w", items[1])
            if m:
                width = int(m.group(1))
        if width > best_width:
            best_url = url
            best_width = width
    return best_url


def add_img_candidate(candidates: list[tuple[int, str]], url: str, score: int) -> None:
    url = absolute_url(url)
    if not url:
        return
    if not (IMAGE_EXT_RE.search(url) or "digitaloceanspaces.com" in url or "/storage/" in url or "/uploads/" in url):
        return
    lower = url.lower()
    if BAD_IMAGE_RE.search(lower):
        score -= 60
    if "artsoul.nyc3.cdn.digitaloceanspaces.com" in lower:
        score += 70
    if IMAGE_EXT_RE.search(lower):
        score += 20
    candidates.append((score, url))


def extract_image(soup: BeautifulSoup, jsons: list[Any], title: str) -> str:
    candidates: list[tuple[int, str]] = []

    # 1) Dados estruturados/JSON-LD costumam ter a imagem original.
    for obj in flatten_json(jsons):
        if not isinstance(obj, dict):
            continue
        for key in ("image", "thumbnailUrl", "contentUrl", "url"):
            val = obj.get(key)
            if isinstance(val, str):
                add_img_candidate(candidates, val, 120)
            elif isinstance(val, list):
                for item in val:
                    if isinstance(item, str):
                        add_img_candidate(candidates, item, 120)
                    elif isinstance(item, dict):
                        for sub_key in ("url", "contentUrl"):
                            if isinstance(item.get(sub_key), str):
                                add_img_candidate(candidates, item[sub_key], 120)
            elif isinstance(val, dict):
                for sub_key in ("url", "contentUrl"):
                    if isinstance(val.get(sub_key), str):
                        add_img_candidate(candidates, val[sub_key], 120)

    # 2) OpenGraph/Twitter.
    for sel in (
        "meta[property='og:image']",
        "meta[property='og:image:secure_url']",
        "meta[name='twitter:image']",
        "link[rel='image_src']",
    ):
        node = soup.select_one(sel)
        if node:
            add_img_candidate(candidates, node.get("content") or node.get("href") or "", 110)

    title_words = {w for w in re.split(r"\W+", title.lower()) if len(w) >= 4}

    # 3) Tags visíveis: img/source com lazy-load/srcset e links diretos para CDN.
    for node in soup.select("img, source"):
        attrs = [
            "src",
            "data-src",
            "data-original",
            "data-lazy",
            "data-lazy-src",
            "data-full",
            "data-zoom-image",
            "content",
        ]
        score = 40
        alt_text = norm_space(node.get("alt") or node.get("title") or "")
        if alt_text and title_words.intersection(re.split(r"\W+", alt_text.lower())):
            score += 35
        for attr in attrs:
            val = node.get(attr)
            if val:
                add_img_candidate(candidates, val, score)
        for attr in ("srcset", "data-srcset"):
            val = node.get(attr)
            if val:
                add_img_candidate(candidates, srcset_best(val), score + 15)

    for a in soup.select("a[href]"):
        href = a.get("href") or ""
        if IMAGE_EXT_RE.search(href) or "digitaloceanspaces.com" in href:
            text = norm_space(a.get_text(" ", strip=True))
            score = 35
            if title_words and title_words.intersection(re.split(r"\W+", text.lower())):
                score += 35
            add_img_candidate(candidates, href, score)

    # Remove duplicados mantendo maior score.
    best_by_url: dict[str, int] = {}
    for score, url in candidates:
        best_by_url[url] = max(score, best_by_url.get(url, -999))
    if not best_by_url:
        return SEM_INFORMACAO
    return sorted(best_by_url.items(), key=lambda kv: kv[1], reverse=True)[0][0]


def is_artwork_url(href: str) -> bool:
    try:
        path = urlsplit(href).path.rstrip("/")
    except Exception:
        return False
    return bool(ARTWORK_PATH_RE.match(path))


def extract_artwork_links(soup: BeautifulSoup, current_url: str) -> list[str]:
    links: list[str] = []
    for a in soup.select("a[href]"):
        href = absolute_url(a.get("href"), current_url)
        if is_artwork_url(href):
            links.append(href.split("#", 1)[0])
    return list(dict.fromkeys(links))



def strip_page_params(url: str) -> str:
    """Remove parâmetros de página para sempre reconstruir URLs a partir da página-base."""
    parts = urlsplit(url)
    q = [
        (k, v)
        for k, v in parse_qsl(parts.query, keep_blank_values=True)
        if k.lower() not in {"page", "pagina", "p"}
    ]
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(q, doseq=True), parts.fragment))


def set_query_param(url: str, key: str, value: int | str) -> str:
    parts = urlsplit(url)
    q = [
        (k, v)
        for k, v in parse_qsl(parts.query, keep_blank_values=True)
        if k.lower() != key.lower()
    ]
    q.append((key, str(value)))
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(q, doseq=True), parts.fragment))


def with_page(url: str, page_number: int) -> str:
    """Compatibilidade com versões anteriores: usa o parâmetro page=N."""
    return set_query_param(strip_page_params(url), "page", page_number)


def page_url_candidates(start_url: str, page_number: int) -> list[str]:
    """Gera variações de URL para paginação.

    A forma usada hoje pelo ArtSoul é consulta com page=N, mas mantemos variações
    para resistir a pequenas mudanças do site sem perder a coleta.
    """
    base = strip_page_params(start_url)
    if page_number <= 1:
        return [base]

    parts = urlsplit(base)
    path = parts.path.rstrip("/")
    query = parts.query
    candidates = [
        set_query_param(base, "page", page_number),
        set_query_param(base, "pagina", page_number),
        set_query_param(base, "p", page_number),
        urlunsplit((parts.scheme, parts.netloc, f"{path}/page/{page_number}", query, parts.fragment)),
        urlunsplit((parts.scheme, parts.netloc, f"{path}/{page_number}", query, parts.fragment)),
    ]
    return list(dict.fromkeys(candidates))


def parse_total(lines: list[str]) -> int | None:
    text = "\n".join(lines)
    for pattern in (
        r"Encontramos um total de\s+(\d+)\s+obras",
        r"Encontramos um total de\s+(\d+)\s+obras de arte",
        r"Exibindo\s+\d+\s+at[eé]\s+\d+\s+de\s+(\d+)\s+registros",
    ):
        m = re.search(pattern, text, flags=re.I)
        if m:
            return int(m.group(1))
    return None


def parse_display_range(lines: list[str]) -> tuple[int, int, int] | None:
    """Lê o contador da listagem: Exibindo 1 até 32 de 788 registros."""
    text = "\n".join(lines)
    m = re.search(
        r"Exibindo\s+(\d+)\s+at[eé]\s+(\d+)\s+de\s+(\d+)\s+registros",
        text,
        flags=re.I,
    )
    if not m:
        return None
    return int(m.group(1)), int(m.group(2)), int(m.group(3))


def infer_total_pages(lines: list[str], links_on_page: int) -> int | None:
    """Calcula total de páginas a partir do contador do site ou da paginação visível."""
    display = parse_display_range(lines)
    if display:
        start, end, total = display
        per_page = max(1, end - start + 1)
        return max(1, (total + per_page - 1) // per_page)

    # Fallback: quando o contador some, tenta pegar o maior número na linha de paginação.
    text = " ".join(lines[-80:])
    nums = [int(n) for n in re.findall(r"(?<![#\w])(\d{1,4})(?![#\w])", text)]
    nums = [n for n in nums if n > 1]
    if nums:
        return max(nums)

    total = parse_total(lines)
    if total and links_on_page:
        return max(1, (total + links_on_page - 1) // links_on_page)
    return None


def next_page_url(soup: BeautifulSoup, current_url: str, current_page: int) -> str:
    """Tenta localizar um link real de próxima página no HTML."""
    selectors = [
        "a[rel='next']",
        "a[aria-label*='Próximo']",
        "a[aria-label*='Próxima']",
        "a[aria-label*='Next']",
        ".pagination a",
        "nav a",
    ]
    wanted = str(current_page + 1)
    for sel in selectors:
        for a in soup.select(sel):
            text = norm_space(a.get_text(" ", strip=True))
            href = a.get("href")
            if not href or href.startswith("#"):
                continue
            if sel not in {".pagination a", "nav a"} or text in {wanted, "Próximo", "Próxima", "Next", "›", "»"}:
                return absolute_url(href, current_url)

    for a in soup.select("a[href]"):
        text = norm_space(a.get_text(" ", strip=True))
        href = a.get("href")
        if href and text == wanted:
            return absolute_url(href, current_url)

    return ""


@dataclass
class ListingPageResult:
    page_number: int
    url: str
    soup: BeautifulSoup
    lines: list[str]
    links: list[str]
    display_range: tuple[int, int, int] | None


def fetch_listing_page(session: requests.Session, url: str, page_number: int) -> ListingPageResult:
    html = fetch(session, url)
    soup = BeautifulSoup(html, "lxml")
    lines = linhas_texto(soup)
    return ListingPageResult(
        page_number=page_number,
        url=url,
        soup=soup,
        lines=lines,
        links=extract_artwork_links(soup, url),
        display_range=parse_display_range(lines),
    )


def choose_listing_candidate(
    session: requests.Session,
    candidates: list[str],
    page_number: int,
    seen_links: set[str],
    expected_start: int | None,
    delay_between_candidates: float,
) -> ListingPageResult | None:
    """Escolhe a melhor variação de URL para uma página da listagem.

    Critérios: contador do site batendo com a página esperada, quantidade de links
    novos e quantidade total de links encontrados. Isso evita aceitar uma URL que
    redirecionou silenciosamente para a página 1.
    """
    best: tuple[int, ListingPageResult] | None = None
    for candidate_url in candidates:
        try:
            result = fetch_listing_page(session, candidate_url, page_number)
        except Exception as exc:
            print(f"  Falha testando paginação {candidate_url}: {exc}")
            continue

        new_count = len([link for link in result.links if link not in seen_links])
        score = new_count * 100 + len(result.links)
        if result.display_range:
            start, _end, _total = result.display_range
            if expected_start is not None and start == expected_start:
                score += 10000
            elif page_number > 1 and start <= 1:
                score -= 10000
            else:
                score += min(start, 999)
        elif page_number == 1:
            score += 1000

        if best is None or score > best[0]:
            best = (score, result)

        if delay_between_candidates:
            time.sleep(delay_between_candidates)

    if best is None:
        return None
    return best[1]


def discover_all_links(session: requests.Session, start_url: str, max_pages: int = 0, delay: float = 0.25) -> list[str]:
    """Descobre links de obras em todas as páginas da listagem.

    A página atual do ArtSoul informa um total e um intervalo por página
    (ex.: "Exibindo 1 até 32 de 788 registros"). Quando esse contador existe,
    o scraper calcula o total de páginas e varre page=1..N diretamente, em vez de
    depender apenas do botão "Próximo". Isso torna a coleta completa e previsível.
    """
    base_url = strip_page_params(start_url)
    seen_links: set[str] = set()
    all_links: list[str] = []
    total_expected: int | None = None
    total_pages: int | None = None
    per_page: int | None = None
    no_new_streak = 0
    last_result: ListingPageResult | None = None

    page_number = 1
    while True:
        if max_pages and page_number > max_pages:
            break
        if total_pages and page_number > total_pages:
            break
        if total_expected and len(all_links) >= total_expected:
            break

        expected_start = None
        if per_page:
            expected_start = (page_number - 1) * per_page + 1

        candidates = page_url_candidates(base_url, page_number)
        # Se o HTML trouxe um link real de próxima página, testa-o antes das variações.
        if last_result is not None:
            linked_next = next_page_url(last_result.soup, last_result.url, last_result.page_number)
            if linked_next:
                candidates = [linked_next, *candidates]
        candidates = list(dict.fromkeys(candidates))

        print(f"Listagem {page_number}: testando {len(candidates)} URL(s) de paginação")
        result = choose_listing_candidate(
            session=session,
            candidates=candidates,
            page_number=page_number,
            seen_links=seen_links,
            expected_start=expected_start,
            delay_between_candidates=0,
        )
        if result is None:
            print(f"  Não foi possível carregar a página {page_number} da listagem.")
            break

        last_result = result
        if total_expected is None:
            total_expected = parse_total(result.lines)
            if total_expected:
                print(f"Total informado pelo site: {total_expected} obras")

        if total_pages is None:
            total_pages = infer_total_pages(result.lines, len(result.links))
            if total_pages:
                print(f"Total de páginas estimado: {total_pages}")

        if per_page is None and result.display_range:
            start, end, _total = result.display_range
            per_page = max(1, end - start + 1)

        new_links = [link for link in result.links if link not in seen_links]
        print(
            f"  URL usada: {result.url}\n"
            f"  Links encontrados: {len(result.links)} | novos: {len(new_links)} | acumulado: {len(all_links) + len(new_links)}"
        )

        for link in new_links:
            seen_links.add(link)
            all_links.append(link)

        if not new_links:
            no_new_streak += 1
        else:
            no_new_streak = 0

        # Se não sabemos o total e duas páginas seguidas não trouxeram novidade, encerra.
        # Se sabemos o total, permite uma falha eventual, mas para após duas páginas inúteis.
        if no_new_streak >= 2:
            print("  Encerrando paginação: duas páginas seguidas sem links novos.")
            break

        page_number += 1
        if delay:
            time.sleep(delay)

    if total_expected and len(all_links) < total_expected:
        print(
            f"ATENÇÃO: o site informou {total_expected} obras, mas foram descobertas {len(all_links)} URLs únicas. "
            "Se isso acontecer repetidamente, aumente o delay ou rode com --max-pages 0 novamente."
        )
    return all_links

def parse_work_detail(html: str, url: str) -> Obra:
    soup = BeautifulSoup(html, "lxml")
    lines = linhas_texto(soup)
    jsons = parse_json_scripts(soup)

    title = extract_title(soup, lines, jsons, url)
    desc = description_lines(lines)
    price_text, price_number = extract_price(lines, soup, jsons)

    return Obra(
        nome_da_obra=title,
        autor=extract_author(soup, lines, jsons, title),
        preco_brl=price_text,
        preco_numero=price_number,
        dimensoes=extract_dimensions(soup, lines, desc),
        tecnica=extract_technique(lines, desc),
        ano_da_obra=extract_year(lines, jsons),
        descricao=extract_description(desc),
        imagem_da_obra=extract_image(soup, jsons, title),
        url_da_obra=url,
    )


def obra_to_excel_row(obra: Obra) -> dict[str, Any]:
    """Converte os dados internos para o layout histórico da planilha.

    Ordem e nomes das colunas:
    Título, Autor, Ano, Técnica, Dimensões, Preço, Descrição,
    Link da obra, Link da imagem da obra.
    """
    preco = obra.preco_numero if obra.preco_numero != SEM_INFORMACAO else SEM_INFORMACAO
    return {
        "Título": obra.nome_da_obra,
        "Autor": obra.autor,
        "Ano": obra.ano_da_obra,
        "Técnica": obra.tecnica,
        "Dimensões": obra.dimensoes,
        "Preço": preco,
        "Descrição": obra.descricao,
        "Link da obra": obra.url_da_obra,
        "Link da imagem da obra": obra.imagem_da_obra,
    }


def scrape_artsoul(
    start_url: str = COLECAO_URL,
    max_obras: int = 0,
    max_pages: int = 0,
    delay: float = 0.35,
) -> pd.DataFrame:
    session = make_session()
    links = discover_all_links(session, start_url, max_pages=max_pages, delay=delay)
    if max_obras and max_obras > 0:
        links = links[:max_obras]
    print(f"\nTotal de URLs de obras para detalhar: {len(links)}\n")

    rows: list[dict[str, Any]] = []
    for idx, link in enumerate(links, start=1):
        try:
            print(f"[{idx}/{len(links)}] Detalhando: {link}")
            html = fetch(session, link)
            obra = parse_work_detail(html, link)
            rows.append(obra_to_excel_row(obra))
            print(f"  OK: {obra.nome_da_obra} | {obra.autor} | {obra.preco_brl}")
        except Exception as exc:
            print(f"  ERRO em {link}: {exc}", file=sys.stderr)
            rows.append(obra_to_excel_row(Obra(url_da_obra=link)))
        if delay:
            time.sleep(delay)

    df = pd.DataFrame(rows, columns=EXCEL_COLUMNS)
    if not df.empty:
        df = df.drop_duplicates(subset=["Link da obra"], keep="first")
    return df


def exportar(df: pd.DataFrame, output: str) -> None:
    ext = os.path.splitext(output)[1].lower()
    if ext in {".xlsx", ".xlsm", ".xls"}:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="ArtSoul")
            ws = writer.sheets["ArtSoul"]
            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 12), 70)
            ws.freeze_panes = "A2"
        return
    df.to_csv(output, index=False, encoding="utf-8-sig")


def main() -> None:
    parser = argparse.ArgumentParser(description="Scraper ArtSoul - obras de pintura")
    parser.add_argument("--url", default=os.getenv("ARTSOUL_URL", COLECAO_URL), help="URL inicial da listagem")
    parser.add_argument("--output", default=os.getenv("OUTPUT", "artes_ArtSoul.xlsx"), help="Arquivo de saída .xlsx ou .csv")
    parser.add_argument("--max-obras", type=int, default=int(os.getenv("MAX_OBRAS", "0")), help="Limite de obras. 0 = sem limite")
    parser.add_argument("--max-pages", type=int, default=int(os.getenv("MAX_PAGES", "0")), help="Limite de páginas da listagem. 0 = sem limite")
    parser.add_argument("--delay", type=float, default=float(os.getenv("DELAY", "0.35")), help="Pausa entre requisições, em segundos")
    args = parser.parse_args()

    started = datetime.now()
    df = scrape_artsoul(args.url, max_obras=args.max_obras, max_pages=args.max_pages, delay=args.delay)
    exportar(df, args.output)
    elapsed = datetime.now() - started
    print(f"\nConcluído: {len(df)} obras exportadas para {args.output} em {elapsed}.")


if __name__ == "__main__":
    main()
