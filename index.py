"""
FollowOp BI — Backend FastAPI com integração Google Sheets API + Cruzamento de Dados.

Uso:
  1. Coloque o JSON da Service Account em credentials.json (ou defina GOOGLE_CREDENTIALS_PATH)
  2. Defina SPREADSHEET_ID no .env ou como variável de ambiente
  3. Execute: uvicorn server:app --reload --port 8000

Endpoints — Sheets (leitura direta):
  GET /api/sheets/painel            → KPIs consolidados + séries diárias
  GET /api/sheets/vendedor/{nome}   → Dados individuais (Vitor ou Davi)
  GET /api/sheets/plano             → Metas e parâmetros do plano de rampa
  GET /api/sheets/completo          → Todos os dados Sheets em uma chamada

Endpoints — Cruzamento (7 pilares):
  GET /api/cruzamento/funil         → Funil completo: prospecção → lead → venda
  GET /api/cruzamento/cac           → CAC por canal com custo de prospecção
  GET /api/cruzamento/roi           → ROI por canal incluindo custo WhatsApp
  GET /api/cruzamento/vendedor/{n}  → Performance individual cruzada
  GET /api/cruzamento/completo      → Todos os cruzamentos unificados
  GET /health                       → Status da conexão
"""
import os, json, re, logging
from datetime import datetime, timedelta
from typing import Optional
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

# Google API
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("followop-bi")

# ══════════════════════════════════════
# CONFIG
# ══════════════════════════════════════
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "19l6cAYq3TP0j7rcNOkUaN0iZf7pUEhA-L5Car8Tkqvk")
CREDENTIALS_PATH = os.getenv("GOOGLE_CREDENTIALS_PATH", "credentials.json")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# Custo de prospecção: USD 0.005 por disparo (meio centavo de dólar)
CUSTO_DISPARO_USD = float(os.getenv("CUSTO_DISPARO_USD", "0.005"))
COTACAO_USD_BRL = float(os.getenv("COTACAO_USD_BRL", "5.70"))  # Atualizar conforme necessidade

# Cache simples em memória (TTL = 30s para não sobrecarregar a API)
_cache: dict = {}
CACHE_TTL = 30  # segundos

app = FastAPI(title="FollowOp BI — Google Sheets Bridge", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Em produção, restrinja ao domínio do frontend
    allow_methods=["GET"],
    allow_headers=["*"],
)

# ══════════════════════════════════════
# GOOGLE SHEETS CLIENT
# ══════════════════════════════════════
_sheets_service = None

def get_sheets_service():
    """Inicializa o client Google Sheets com Service Account."""
    global _sheets_service
    if _sheets_service:
        return _sheets_service

    creds_path = Path(CREDENTIALS_PATH)
    if not creds_path.exists():
        # Tenta também via variável de ambiente com JSON inline
        creds_json = os.getenv("GOOGLE_CREDENTIALS_JSON")
        if creds_json:
            creds_data = json.loads(creds_json)
            creds = Credentials.from_service_account_info(creds_data, scopes=SCOPES)
        else:
            raise FileNotFoundError(
                f"Arquivo de credenciais não encontrado: {CREDENTIALS_PATH}. "
                "Defina GOOGLE_CREDENTIALS_PATH ou GOOGLE_CREDENTIALS_JSON."
            )
    else:
        creds = Credentials.from_service_account_file(str(creds_path), scopes=SCOPES)

    _sheets_service = build("sheets", "v4", credentials=creds)
    logger.info("Google Sheets API conectada com sucesso.")
    return _sheets_service


def read_sheet(range_name: str) -> list[list[str]]:
    """Lê um range da planilha. Retorna lista de linhas (lista de strings)."""
    cache_key = f"sheet:{range_name}"
    now = datetime.now()

    if cache_key in _cache:
        data, ts = _cache[cache_key]
        if (now - ts).total_seconds() < CACHE_TTL:
            return data

    try:
        service = get_sheets_service()
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name,
        ).execute()
        rows = result.get("values", [])
        _cache[cache_key] = (rows, now)
        return rows
    except HttpError as e:
        if e.resp.status == 403:
            raise HTTPException(403, "Sem permissão. Verifique se a planilha foi compartilhada com a Service Account.")
        elif e.resp.status == 404:
            raise HTTPException(404, f"Aba não encontrada: {range_name}")
        raise HTTPException(502, f"Erro Google Sheets API: {e}")
    except FileNotFoundError as e:
        raise HTTPException(500, str(e))
    except Exception as e:
        logger.error(f"Erro ao ler planilha: {e}")
        raise HTTPException(502, f"Erro de conexão com Google Sheets: {str(e)}")


# ══════════════════════════════════════
# PARSERS
# ══════════════════════════════════════
def safe_int(val: str) -> int:
    """Converte string para int, removendo separadores."""
    if not val or val.strip() in ("", "—", "-"):
        return 0
    return int(re.sub(r"[^\d-]", "", val.strip()))

def safe_float(val: str) -> float:
    """Converte string para float (ex: '24%' → 24.0)."""
    if not val or val.strip() in ("", "—", "-"):
        return 0.0
    cleaned = val.strip().replace("%", "").replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return 0.0

def safe_str(val: str) -> str:
    return val.strip() if val else ""

def cell(row: list, idx: int) -> str:
    """Acessa célula com fallback para string vazia."""
    return row[idx].strip() if idx < len(row) and row[idx] else ""


def parse_painel(rows: list[list[str]]) -> dict:
    """
    Parseia a aba Painel.
    Estrutura fixa:
      Linha 0: título
      Linha 1: Hoje | data | Dia do plano | D##
      Linha 2: Meta final | descrição
      Linha 3: header bloco "Hoje" (Hoje | Vitor | Davi | Equipe)
      Linhas 4-8: métricas do dia
      Linha 9: header bloco "Acumulado"
      Linhas 10-22: métricas acumuladas
      Linha 23: nota
      Linha 24: header série diária
      Linhas 25+: D1, D2, ... dados diários
    """
    result = {
        "titulo": cell(rows[0], 0) if len(rows) > 0 else "",
        "data_hoje": cell(rows[1], 1) if len(rows) > 1 else "",
        "dia_plano": cell(rows[1], 3) if len(rows) > 1 else "",
        "meta_final": cell(rows[2], 1) if len(rows) > 2 else "",
    }

    # Bloco "Hoje" — procura as linhas por label
    hoje = {}
    acumulado = {}
    serie_diaria = []
    current_block = None

    for row in rows[3:]:
        label = cell(row, 0).lower()

        # Detecta headers de bloco
        # "Hoje | Vitor | Davi | Equipe" → header do bloco hoje (label exata "hoje")
        # "Acumulado | Vitor | Davi | Equipe" → header acumulado
        # "Dia" sozinho → header serie diária
        non_empty = sum(1 for c in row if c.strip())
        if label == "hoje":
            current_block = "hoje"
            continue
        if label.startswith("acumulado") and cell(row, 1).lower() in ("vitor", "davi", ""):
            current_block = "acumulado"
            continue
        if label == "dia" or (label.startswith("dia") and non_empty <= 1):
            current_block = "diaria"
            continue
        if label.startswith("na coluna equipe"):
            continue

        if current_block == "hoje":
            lbl = cell(row, 0)
            if lbl:
                hoje[lbl] = {
                    "vitor": cell(row, 1) if len(row) > 1 else "",
                    "davi": cell(row, 2) if len(row) > 2 else "",
                    "equipe": cell(row, 3) if len(row) > 3 else "",
                }
        elif current_block == "acumulado":
            lbl = cell(row, 0)
            if lbl:
                acumulado[lbl] = {
                    "vitor": cell(row, 1) if len(row) > 1 else "",
                    "davi": cell(row, 2) if len(row) > 2 else "",
                    "equipe": cell(row, 3) if len(row) > 3 else "",
                }
        elif current_block == "diaria":
            dia = cell(row, 0)
            if dia.startswith("D"):
                serie_diaria.append({
                    "dia": dia,
                    "data": cell(row, 1),
                    "meta_equipe": safe_int(cell(row, 2)),
                    "feito_equipe": safe_int(cell(row, 3)),
                })

    # Estruturar KPIs do dia
    kpis_hoje = {
        "ativar_chip_meta": hoje.get("Ativar chip hoje (meta)", {}).get("equipe", "—"),
        "chips_aquecer_meta": hoje.get("Chips a aquecer hoje (meta)", {}).get("equipe", "—"),
        "envios_meta": safe_int(hoje.get("Envios de hoje (meta)", {}).get("equipe", "0")),
        "envios_feitos": safe_int(hoje.get("Envios de hoje (feitos)", {}).get("equipe", "0")),
        "pct_meta_hoje": hoje.get("% da meta de hoje", {}).get("equipe", "0%"),
        "vitor": {
            "envios_meta": safe_int(hoje.get("Envios de hoje (meta)", {}).get("vitor", "0")),
            "envios_feitos": safe_int(hoje.get("Envios de hoje (feitos)", {}).get("vitor", "0")),
            "pct_meta": hoje.get("% da meta de hoje", {}).get("vitor", "0%"),
        },
        "davi": {
            "envios_meta": safe_int(hoje.get("Envios de hoje (meta)", {}).get("davi", "0")),
            "envios_feitos": safe_int(hoje.get("Envios de hoje (feitos)", {}).get("davi", "0")),
            "pct_meta": hoje.get("% da meta de hoje", {}).get("davi", "0%"),
        },
    }

    # Estruturar acumulado
    kpis_acumulado = {
        "meta_envios": safe_int(acumulado.get("Meta de envios até hoje", {}).get("equipe", "0")),
        "envios_feitos": safe_int(acumulado.get("Envios feitos", {}).get("equipe", "0")),
        "pct_meta": acumulado.get("% da meta até hoje", {}).get("equipe", "0%"),
        "dias_meta_batida": safe_int(acumulado.get("Dias com meta batida", {}).get("equipe", "0")),
        "dias_passados": safe_int(acumulado.get("Dias já passados", {}).get("equipe", "0")),
        "chips_ativados": safe_int(acumulado.get("Chips ativados (✔)", {}).get("equipe", "0")),
        "chips_previstos": safe_int(acumulado.get("Chips previstos até hoje", {}).get("equipe", "0")),
        "dias_aquecimento_feitos": safe_int(acumulado.get("Dias de aquecimento feitos (✔)", {}).get("equipe", "0")),
        "dias_aquecimento_previstos": safe_int(acumulado.get("Dias de aquecimento previstos até hoje", {}).get("equipe", "0")),
        "respostas": safe_int(acumulado.get("Respostas", {}).get("equipe", "0")),
        "taxa_resposta": acumulado.get("Taxa de resposta", {}).get("equipe", "0%"),
        "bloqueios": safe_int(acumulado.get("Bloqueios", {}).get("equipe", "0")),
        "registros_dia": safe_int(acumulado.get("Registros do dia feitos (✔)", {}).get("equipe", "0")),
        "vitor": {
            "envios_feitos": safe_int(acumulado.get("Envios feitos", {}).get("vitor", "0")),
            "pct_meta": acumulado.get("% da meta até hoje", {}).get("vitor", "0%"),
            "dias_meta_batida": safe_int(acumulado.get("Dias com meta batida", {}).get("vitor", "0")),
            "chips_ativados": safe_int(acumulado.get("Chips ativados (✔)", {}).get("vitor", "0")),
            "respostas": safe_int(acumulado.get("Respostas", {}).get("vitor", "0")),
            "taxa_resposta": acumulado.get("Taxa de resposta", {}).get("vitor", "0%"),
            "bloqueios": safe_int(acumulado.get("Bloqueios", {}).get("vitor", "0")),
        },
        "davi": {
            "envios_feitos": safe_int(acumulado.get("Envios feitos", {}).get("davi", "0")),
            "pct_meta": acumulado.get("% da meta até hoje", {}).get("davi", "0%"),
            "dias_meta_batida": safe_int(acumulado.get("Dias com meta batida", {}).get("davi", "0")),
            "chips_ativados": safe_int(acumulado.get("Chips ativados (✔)", {}).get("davi", "0")),
            "respostas": safe_int(acumulado.get("Respostas", {}).get("davi", "0")),
            "taxa_resposta": acumulado.get("Taxa de resposta", {}).get("davi", "0%"),
            "bloqueios": safe_int(acumulado.get("Bloqueios", {}).get("davi", "0")),
        },
    }

    result["hoje"] = kpis_hoje
    result["acumulado"] = kpis_acumulado
    result["serie_diaria"] = serie_diaria
    result["updated_at"] = datetime.now().isoformat()

    return result


def parse_vendedor(rows: list[list[str]], nome: str) -> dict:
    """Parseia aba individual de vendedor (Vitor ou Davi)."""
    dias = []
    totais = {}

    for row in rows[3:]:  # Pula header
        dia = cell(row, 0)
        if not dia.startswith("D"):
            # Linha de totais (última linha com dados)
            if len(row) >= 8 and cell(row, 4):
                totais = {
                    "meta_total": safe_int(cell(row, 4)),
                    "envios_total": safe_int(cell(row, 7)),
                    "respostas_total": safe_int(cell(row, 8)) if len(row) > 8 else 0,
                    "bloqueios_total": safe_int(cell(row, 9)) if len(row) > 9 else 0,
                    "pct_geral": cell(row, 11) if len(row) > 11 else "",
                    "status_geral": cell(row, 12) if len(row) > 12 else "",
                }
            continue

        dias.append({
            "dia": dia,
            "data": cell(row, 1),
            "chips_ativar": cell(row, 2),
            "chips_aquecer": cell(row, 3),
            "meta_envios": safe_int(cell(row, 4)),
            "chip_ativado": cell(row, 5) == "✔",
            "aquecimento_feito": cell(row, 6) == "✔",
            "envios_feitos": safe_int(cell(row, 7)),
            "respostas": safe_int(cell(row, 8)) if len(row) > 8 else 0,
            "bloqueios": safe_int(cell(row, 9)) if len(row) > 9 else 0,
            "registro_feito": cell(row, 10) == "✔" if len(row) > 10 else False,
            "pct_meta": cell(row, 11) if len(row) > 11 else "",
            "status": cell(row, 12) if len(row) > 12 else "",
        })

    return {
        "nome": nome,
        "dias": dias,
        "totais": totais,
        "updated_at": datetime.now().isoformat(),
    }


def parse_plano(rows: list[list[str]]) -> dict:
    """Parseia aba Plano."""
    dias = []
    for row in rows[3:]:  # Pula headers
        dia = cell(row, 0)
        if not dia.startswith("D"):
            continue
        dias.append({
            "dia": dia,
            "data": cell(row, 1),
            "chips_ativar_equipe": cell(row, 2),
            "chips_aquecer_equipe": cell(row, 3),
            "chips_maduros_equipe": cell(row, 4),
            "meta_envios_equipe": safe_int(cell(row, 5)),
            "chips_ativar_vendedor": cell(row, 6),
            "chips_aquecer_vendedor": cell(row, 7),
            "chips_maduros_vendedor": cell(row, 8),
            "meta_envios_vendedor": safe_int(cell(row, 9)),
        })

    return {
        "data_inicio": cell(rows[1], 1) if len(rows) > 1 else "",
        "data_meta_final": cell(rows[2], 1) if len(rows) > 2 else "",
        "dias": dias,
        "updated_at": datetime.now().isoformat(),
    }


# ══════════════════════════════════════
# ENDPOINTS
# ══════════════════════════════════════
@app.get("/health")
async def health():
    """Verifica se a conexão com Google Sheets está funcionando."""
    try:
        service = get_sheets_service()
        meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        sheets = [s["properties"]["title"] for s in meta.get("sheets", [])]
        return {"status": "ok", "spreadsheet": meta.get("properties", {}).get("title", ""), "sheets": sheets}
    except FileNotFoundError as e:
        return JSONResponse(status_code=500, content={"status": "error", "detail": str(e)})
    except Exception as e:
        return JSONResponse(status_code=502, content={"status": "error", "detail": str(e)})


@app.get("/api/sheets/painel")
async def get_painel():
    """Retorna dados consolidados da aba Painel."""
    rows = read_sheet("Painel!A1:D50")
    return parse_painel(rows)


@app.get("/api/sheets/vendedor/{nome}")
async def get_vendedor(nome: str):
    """Retorna dados individuais de um vendedor (Vitor ou Davi)."""
    nome_cap = nome.capitalize()
    if nome_cap not in ("Vitor", "Davi"):
        raise HTTPException(400, "Vendedor deve ser 'Vitor' ou 'Davi'.")
    rows = read_sheet(f"{nome_cap}!A1:M30")
    return parse_vendedor(rows, nome_cap)


@app.get("/api/sheets/plano")
async def get_plano():
    """Retorna o plano de rampa com metas por dia."""
    rows = read_sheet("Plano!A1:J30")
    return parse_plano(rows)


@app.get("/api/sheets/completo")
async def get_completo():
    """Retorna todos os dados em uma única chamada (usado pelo frontend para refresh)."""
    try:
        painel_rows = read_sheet("Painel!A1:D50")
        vitor_rows = read_sheet("Vitor!A1:M30")
        davi_rows = read_sheet("Davi!A1:M30")
        plano_rows = read_sheet("Plano!A1:J30")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(502, f"Erro ao buscar dados: {str(e)}")

    return {
        "painel": parse_painel(painel_rows),
        "vitor": parse_vendedor(vitor_rows, "Vitor"),
        "davi": parse_vendedor(davi_rows, "Davi"),
        "plano": parse_plano(plano_rows),
        "source": "google_sheets",
        "updated_at": datetime.now().isoformat(),
    }


# ══════════════════════════════════════
# CALCULADORAS DE CRUZAMENTO
# ══════════════════════════════════════

def custo_disparo_brl() -> float:
    """Custo por disparo de prospecção em BRL."""
    return CUSTO_DISPARO_USD * COTACAO_USD_BRL


def calcular_custo_prospeccao(envios: int) -> dict:
    """Calcula custo total de prospecção baseado em envios."""
    custo_unitario = custo_disparo_brl()
    total = envios * custo_unitario
    return {
        "envios": envios,
        "custo_unitario_usd": CUSTO_DISPARO_USD,
        "custo_unitario_brl": round(custo_unitario, 4),
        "cotacao_usd_brl": COTACAO_USD_BRL,
        "custo_total_brl": round(total, 2),
    }


def calcular_metricas_vendedor(vendedor_data: dict, plano_data: dict) -> dict:
    """Calcula métricas cruzadas para um vendedor individual."""
    dias = vendedor_data.get("dias", [])
    totais = vendedor_data.get("totais", {})
    nome = vendedor_data.get("nome", "")

    total_envios = totais.get("envios_total", 0)
    total_respostas = totais.get("respostas_total", 0)
    total_bloqueios = totais.get("bloqueios_total", 0)
    meta_total = totais.get("meta_total", 0)

    # Custo de prospecção deste vendedor
    custo = calcular_custo_prospeccao(total_envios)

    # Taxa de resposta
    taxa_resposta = round((total_respostas / total_envios * 100) if total_envios > 0 else 0, 2)

    # Taxa de bloqueio
    taxa_bloqueio = round((total_bloqueios / total_envios * 100) if total_envios > 0 else 0, 2)

    # Aderência à meta (% cumprido do plano)
    aderencia_meta = round((total_envios / meta_total * 100) if meta_total > 0 else 0, 1)

    # Dias produtivos (com envio > 0) vs dias totais
    dias_com_envio = sum(1 for d in dias if d.get("envios_feitos", 0) > 0)
    dias_totais = len(dias)
    dias_meta_batida = sum(1 for d in dias if d.get("status", "").lower().startswith("meta batida"))

    # Custo por resposta (CPR) — similar ao CPL
    cpr = round(custo["custo_total_brl"] / total_respostas, 2) if total_respostas > 0 else 0

    # Eficiência diária
    media_envios_dia = round(total_envios / dias_com_envio, 1) if dias_com_envio > 0 else 0

    # Série diária com custo
    serie_diaria = []
    for d in dias:
        envios = d.get("envios_feitos", 0)
        serie_diaria.append({
            "dia": d.get("dia"),
            "data": d.get("data"),
            "meta": d.get("meta_envios", 0),
            "envios": envios,
            "respostas": d.get("respostas", 0),
            "bloqueios": d.get("bloqueios", 0),
            "custo_brl": round(envios * custo_disparo_brl(), 2),
            "status": d.get("status", ""),
            "pct_meta": d.get("pct_meta", ""),
        })

    return {
        "nome": nome,
        "totais": {
            "envios": total_envios,
            "respostas": total_respostas,
            "bloqueios": total_bloqueios,
            "meta_total": meta_total,
        },
        "custos": custo,
        "taxas": {
            "resposta": taxa_resposta,
            "bloqueio": taxa_bloqueio,
            "aderencia_meta": aderencia_meta,
        },
        "performance": {
            "dias_com_envio": dias_com_envio,
            "dias_totais": dias_totais,
            "dias_meta_batida": dias_meta_batida,
            "media_envios_dia": media_envios_dia,
            "custo_por_resposta_brl": cpr,
        },
        "serie_diaria": serie_diaria,
    }


def calcular_funil_prospeccao(painel: dict, vitor: dict, davi: dict, plano: dict) -> dict:
    """
    Constrói o funil completo: Prospecção → Resposta → Lead → Conversão.
    Integra dados de envio (Sheets) com taxas de conversão do BI.
    """
    acum = painel.get("acumulado", {})

    total_envios = acum.get("envios_feitos", 0)
    total_respostas = acum.get("respostas", 0)
    total_bloqueios = acum.get("bloqueios", 0)

    # Custo total de prospecção
    custo_total = calcular_custo_prospeccao(total_envios)

    # Taxas de conversão do funil de prospecção
    taxa_envio_resposta = round((total_respostas / total_envios * 100) if total_envios > 0 else 0, 2)

    # CPL da prospecção = custo total / respostas (respostas = leads qualificados)
    cpl_prospeccao = round(custo_total["custo_total_brl"] / total_respostas, 2) if total_respostas > 0 else 0

    # Métricas do plano
    meta_envios_acum = acum.get("meta_envios", 0)
    pct_meta = acum.get("pct_meta", "0%")
    dias_meta_batida = acum.get("dias_meta_batida", 0)
    dias_passados = acum.get("dias_passados", 0)

    # Chips
    chips_ativados = acum.get("chips_ativados", 0)
    chips_previstos = acum.get("chips_previstos", 0)

    # Série diária da equipe com custos
    serie = painel.get("serie_diaria", [])
    serie_com_custo = []
    for s in serie:
        feito = s.get("feito_equipe", 0)
        serie_com_custo.append({
            **s,
            "custo_brl": round(feito * custo_disparo_brl(), 2),
        })

    # Funil por etapa
    funil_etapas = [
        {"etapa": "Envios", "qtd": total_envios, "pct": 100.0},
        {"etapa": "Respostas", "qtd": total_respostas,
         "pct": round(taxa_envio_resposta, 1)},
        {"etapa": "Bloqueios", "qtd": total_bloqueios,
         "pct": round((total_bloqueios / total_envios * 100) if total_envios > 0 else 0, 1)},
    ]

    return {
        "resumo": {
            "total_envios": total_envios,
            "total_respostas": total_respostas,
            "total_bloqueios": total_bloqueios,
            "taxa_envio_resposta": taxa_envio_resposta,
            "meta_envios_acum": meta_envios_acum,
            "pct_meta": pct_meta,
            "dias_meta_batida": dias_meta_batida,
            "dias_passados": dias_passados,
            "chips_ativados": chips_ativados,
            "chips_previstos": chips_previstos,
        },
        "custos": {
            **custo_total,
            "cpl_prospeccao": cpl_prospeccao,
        },
        "funil_etapas": funil_etapas,
        "serie_diaria": serie_com_custo,
        "updated_at": datetime.now().isoformat(),
    }


def calcular_cac_com_prospeccao(painel: dict, vitor: dict, davi: dict) -> dict:
    """
    CAC incluindo custo de prospecção por vendedor e canal.
    CAC = (custo prospecção) ÷ respostas qualificadas
    """
    acum = painel.get("acumulado", {})

    # Equipe
    envios_equipe = acum.get("envios_feitos", 0)
    respostas_equipe = acum.get("respostas", 0)
    custo_equipe = calcular_custo_prospeccao(envios_equipe)
    cac_equipe = round(custo_equipe["custo_total_brl"] / respostas_equipe, 2) if respostas_equipe > 0 else 0

    # Por vendedor
    vendedores_cac = []
    for vdata, nome in [(vitor, "Vitor"), (davi, "Davi")]:
        totais = vdata.get("totais", {})
        envios = totais.get("envios_total", 0)
        respostas = totais.get("respostas_total", 0)
        bloqueios = totais.get("bloqueios_total", 0)
        custo = calcular_custo_prospeccao(envios)
        cac = round(custo["custo_total_brl"] / respostas, 2) if respostas > 0 else 0
        cpl = round(custo["custo_total_brl"] / respostas, 2) if respostas > 0 else 0

        vendedores_cac.append({
            "nome": nome,
            "envios": envios,
            "respostas": respostas,
            "bloqueios": bloqueios,
            "custo_total_brl": custo["custo_total_brl"],
            "cac_brl": cac,
            "cpl_brl": cpl,
            "taxa_resposta": round((respostas / envios * 100) if envios > 0 else 0, 2),
            "taxa_bloqueio": round((bloqueios / envios * 100) if envios > 0 else 0, 2),
        })

    return {
        "equipe": {
            "envios": envios_equipe,
            "respostas": respostas_equipe,
            "custo_total_brl": custo_equipe["custo_total_brl"],
            "cac_brl": cac_equipe,
            "custo_unitario_brl": custo_equipe["custo_unitario_brl"],
            "cotacao_usd_brl": COTACAO_USD_BRL,
        },
        "vendedores": vendedores_cac,
        "updated_at": datetime.now().isoformat(),
    }


def calcular_roi_prospeccao(painel: dict, vitor: dict, davi: dict) -> dict:
    """
    ROI da operação de prospecção.
    Como as respostas ainda não têm valor de receita associado,
    calcula métricas de eficiência operacional:
    - Custo por resposta (CPR)
    - Custo por envio efetivo
    - Projeção de ROI se respostas converterem em vendas
    """
    acum = painel.get("acumulado", {})
    envios = acum.get("envios_feitos", 0)
    respostas = acum.get("respostas", 0)
    bloqueios = acum.get("bloqueios", 0)

    custo = calcular_custo_prospeccao(envios)
    custo_total = custo["custo_total_brl"]

    # Custo por resposta
    cpr = round(custo_total / respostas, 2) if respostas > 0 else 0

    # Envios efetivos (descontando bloqueios proporcionalmente)
    envios_efetivos = envios - bloqueios
    custo_por_envio_efetivo = round(custo_total / envios_efetivos, 4) if envios_efetivos > 0 else 0

    # Projeção: se cada resposta gerar em média 1 venda com ticket médio configurável
    ticket_medio_estimado = float(os.getenv("TICKET_MEDIO_ESTIMADO", "800"))
    taxa_conversao_estimada = float(os.getenv("TAXA_CONVERSAO_ESTIMADA", "20"))  # %
    clientes_estimados = int(respostas * taxa_conversao_estimada / 100)
    receita_projetada = clientes_estimados * ticket_medio_estimado
    lucro_projetado = receita_projetada - custo_total
    roi_projetado = round((lucro_projetado / custo_total * 100) if custo_total > 0 else 0, 1)

    return {
        "investimento": {
            **custo,
        },
        "eficiencia": {
            "custo_por_resposta_brl": cpr,
            "custo_por_envio_efetivo_brl": custo_por_envio_efetivo,
            "envios_efetivos": envios_efetivos,
            "taxa_bloqueio": round((bloqueios / envios * 100) if envios > 0 else 0, 2),
        },
        "projecao": {
            "ticket_medio_estimado": ticket_medio_estimado,
            "taxa_conversao_estimada": taxa_conversao_estimada,
            "respostas": respostas,
            "clientes_estimados": clientes_estimados,
            "receita_projetada": receita_projetada,
            "lucro_projetado": round(lucro_projetado, 2),
            "roi_projetado": roi_projetado,
        },
        "updated_at": datetime.now().isoformat(),
    }


# ══════════════════════════════════════
# ENDPOINTS — CRUZAMENTO DE DADOS
# ══════════════════════════════════════

def _fetch_all_sheets():
    """Helper: busca todos os dados do Sheets de uma vez."""
    painel_rows = read_sheet("Painel!A1:D50")
    vitor_rows = read_sheet("Vitor!A1:M30")
    davi_rows = read_sheet("Davi!A1:M30")
    plano_rows = read_sheet("Plano!A1:J30")
    return (
        parse_painel(painel_rows),
        parse_vendedor(vitor_rows, "Vitor"),
        parse_vendedor(davi_rows, "Davi"),
        parse_plano(plano_rows),
    )


@app.get("/api/cruzamento/funil")
async def get_funil_cruzado():
    """Funil completo: prospecção → resposta → lead (7 pilares)."""
    try:
        painel, vitor, davi, plano = _fetch_all_sheets()
        return calcular_funil_prospeccao(painel, vitor, davi, plano)
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Erro no cruzamento do funil: {e}")
        raise HTTPException(502, f"Erro ao calcular funil: {str(e)}")


@app.get("/api/cruzamento/cac")
async def get_cac_cruzado():
    """CAC por vendedor incluindo custo de prospecção WhatsApp."""
    try:
        painel, vitor, davi, plano = _fetch_all_sheets()
        return calcular_cac_com_prospeccao(painel, vitor, davi)
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Erro no cruzamento do CAC: {e}")
        raise HTTPException(502, f"Erro ao calcular CAC: {str(e)}")


@app.get("/api/cruzamento/roi")
async def get_roi_cruzado():
    """ROI da prospecção com projeção de conversão."""
    try:
        painel, vitor, davi, plano = _fetch_all_sheets()
        return calcular_roi_prospeccao(painel, vitor, davi)
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Erro no cruzamento do ROI: {e}")
        raise HTTPException(502, f"Erro ao calcular ROI: {str(e)}")


@app.get("/api/cruzamento/vendedor/{nome}")
async def get_vendedor_cruzado(nome: str):
    """Performance cruzada de um vendedor individual."""
    nome_cap = nome.capitalize()
    if nome_cap not in ("Vitor", "Davi"):
        raise HTTPException(400, "Vendedor deve ser 'Vitor' ou 'Davi'.")
    try:
        painel, vitor, davi, plano = _fetch_all_sheets()
        vendedor_data = vitor if nome_cap == "Vitor" else davi
        return calcular_metricas_vendedor(vendedor_data, plano)
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Erro no cruzamento do vendedor: {e}")
        raise HTTPException(502, f"Erro ao calcular métricas do vendedor: {str(e)}")


@app.get("/api/cruzamento/completo")
async def get_cruzamento_completo():
    """Todos os cruzamentos unificados em um único JSON."""
    try:
        painel, vitor, davi, plano = _fetch_all_sheets()
        return {
            "funil": calcular_funil_prospeccao(painel, vitor, davi, plano),
            "cac": calcular_cac_com_prospeccao(painel, vitor, davi),
            "roi": calcular_roi_prospeccao(painel, vitor, davi),
            "vendedores": {
                "vitor": calcular_metricas_vendedor(vitor, plano),
                "davi": calcular_metricas_vendedor(davi, plano),
            },
            "config": {
                "custo_disparo_usd": CUSTO_DISPARO_USD,
                "cotacao_usd_brl": COTACAO_USD_BRL,
                "custo_disparo_brl": round(custo_disparo_brl(), 4),
                "ticket_medio_estimado": float(os.getenv("TICKET_MEDIO_ESTIMADO", "800")),
                "taxa_conversao_estimada": float(os.getenv("TAXA_CONVERSAO_ESTIMADA", "20")),
            },
            "source": "google_sheets + calculadoras",
            "updated_at": datetime.now().isoformat(),
        }
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Erro no cruzamento completo: {e}")
        raise HTTPException(502, f"Erro ao calcular cruzamento: {str(e)}")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
