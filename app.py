
from datetime import datetime
from zoneinfo import ZoneInfo

FUSO_BRASILIA = ZoneInfo("America/Sao_Paulo")

def agora_brasilia() -> datetime:
    return datetime.now(FUSO_BRASILIA)

def agora_brasilia_iso() -> str:
    return agora_brasilia().isoformat(timespec="seconds")

def formatar_data_brasilia(valor) -> str:
    if valor is None or str(valor).strip() in {"", "nan", "None", "Sem informação"}:
        return "Sem informação"
    try:
        dt = pd.to_datetime(valor, errors="coerce")
        if pd.isna(dt):
            return "Sem informação"
        if dt.tzinfo is None:
            dt = dt.tz_localize("UTC")
        dt = dt.tz_convert("America/Sao_Paulo")
        return dt.strftime("%d/%m/%Y %H:%M:%S")
    except Exception:
        return "Sem informação"

def ultima_atualizacao_dolar() -> str:
    cache = _ler_cache_dolar()
    return formatar_data_brasilia(cache.get("updated_at"))
