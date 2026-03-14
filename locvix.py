"""
Dashboard de Análise de Dados — Locvix
ERP: GestãoClick via API REST
Gera relatório Excel + Dashboard HTML interativo
"""

import os
import sys
import json
import hashlib
import threading
import concurrent.futures
import html as html_mod
from datetime import datetime, timedelta
import math
import time
import base64
import tempfile

# Fix Windows console encoding to support Unicode / emoji in print()
if sys.platform == "win32":
    import io
    if hasattr(sys.stdout, "buffer"):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "buffer"):
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import requests
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

# ══════════════════════════════════════════════════════════════════
#  CONFIGURAÇÕES — preencha as credenciais ou use variáveis de ambiente
# ══════════════════════════════════════════════════════════════════
# As credenciais ficam dentro do ERP GestãoClick:
#   Configurações → Integrações → API → Access Token / Secret Token
GCK_ACCESS_TOKEN  = os.getenv("GCK_ACCESS_TOKEN",  "SEU_ACCESS_TOKEN_AQUI")
GCK_SECRET_TOKEN  = os.getenv("GCK_SECRET_TOKEN",  "SEU_SECRET_TOKEN_AQUI")

# URL base da API GestãoClick
GCK_BASE_URL = "https://api.beteltecnologia.net/api/integracao_api"

# Período padrão (ano corrente)
_ano_atual  = datetime.now().year
DATA_INI    = os.getenv("GCK_DATA_INI", f"01/01/{_ano_atual}")
DATA_FIM    = os.getenv("GCK_DATA_FIM", datetime.now().strftime("%d/%m/%Y"))

# Pastas de saída
_agora       = datetime.now().strftime("%Y%m%d_%H%M%S")
_BASE_SAIDA  = r"Z:\codigos\Locvix GPT" if os.path.isdir(r"Z:\codigos\Locvix GPT") else os.path.dirname(os.path.abspath(__file__))
SAIDA_EXCEL  = os.path.join(_BASE_SAIDA, f"Relatorio_Locvix_{_agora}.xlsx")
SAIDA_HTML   = os.path.join(_BASE_SAIDA, f"Dashboard_Locvix_{_agora}.html")

# Cache em disco (pasta _cache_gck/ junto ao script)
_BASE_DIR   = _BASE_SAIDA
_CACHE_DIR  = os.path.join(
    tempfile.gettempdir() if os.name != "nt" else _BASE_DIR,
    "_cache_gck"
)
os.makedirs(_CACHE_DIR, exist_ok=True)

# TTL de cache (em segundos)
_TTL_VENDAS    = 1800   # 30 min
_TTL_OUTROS    = 86400  # 24 h

# Callback de progresso (usado pelo Streamlit ou similar)
_progresso = None

def _prog(pct: float, msg: str = ""):
    if callable(_progresso):
        try:
            _progresso(min(float(pct), 1.0), msg)
        except Exception:
            pass


# ══════════════════════════════════════════════════════════════════
#  CLIENTE HTTP — GestãoClick API
# ══════════════════════════════════════════════════════════════════
class GCKClient:
    """
    Cliente para a API REST do GestãoClick.
    Autenticação via headers access-token + secret-access-token.
    Suporte a paginação automática.
    """

    def __init__(self, access_token: str, secret_token: str, base_url: str = GCK_BASE_URL):
        self.base_url = base_url.rstrip("/")
        self.session  = requests.Session()
        self.session.headers.update({
            "access-token":        access_token,
            "secret-access-token": secret_token,
            "Content-Type":        "application/json",
            "Accept":              "application/json",
        })
        self._rate_lock = threading.Lock()
        self._last_call = 0.0

    def _throttle(self, intervalo: float = 0.35):
        with self._rate_lock:
            espera = intervalo - (time.time() - self._last_call)
            if espera > 0:
                time.sleep(espera)
            self._last_call = time.time()

    def get(self, endpoint: str, params: dict | None = None, tentativas: int = 5) -> dict | None:
        """GET com retry e rate-limit."""
        url = f"{self.base_url}/{endpoint.lstrip('/')}"
        for t in range(tentativas):
            self._throttle()
            try:
                r = self.session.get(url, params=params or {}, timeout=60)
                if r.status_code == 429:
                    espera = 30 * (t + 1)
                    print(f"  [AVISO] Rate limit (429), aguardando {espera}s...")
                    time.sleep(espera)
                    continue
                r.raise_for_status()
                return r.json()
            except requests.exceptions.HTTPError as e:
                print(f"  [ERRO] HTTP {e.response.status_code} em {url}")
                if e.response.status_code in (401, 403):
                    print("  ⚠ Credenciais inválidas ou sem permissão de API.")
                    return None
                time.sleep(2 * (t + 1))
            except Exception as e:
                print(f"  [AVISO] {endpoint} tentativa {t+1}: {e}")
                time.sleep(2 * (t + 1))
        return None

    def paginar(self, endpoint: str, params: dict | None = None,
                campo_dados: str = "data", limite: int = 100) -> list[dict]:
        """
        Busca todas as páginas de um endpoint paginado.
        Parâmetros de paginação: pagina / limite.
        """
        todos: list[dict] = []
        params = dict(params or {})
        params["limite"] = limite
        pag = 1
        tot_pags = 1

        while pag <= tot_pags:
            params["pagina"] = pag
            resp = self.get(endpoint, params)
            if resp is None:
                break
            items = resp.get(campo_dados, resp.get("data", []))
            if isinstance(items, list):
                todos.extend(items)
            meta = resp.get("meta", {})
            tot_pags = meta.get("last_page", meta.get("ultima_pagina", 1))
            total    = meta.get("total", "?")
            if pag == 1:
                print(f"  {endpoint}: {total} registros ({tot_pags} pág.)")
            pag += 1

        return todos


# Instância global (inicializada no main)
_client: GCKClient | None = None

def _gck() -> GCKClient:
    global _client
    if _client is None:
        _client = GCKClient(GCK_ACCESS_TOKEN, GCK_SECRET_TOKEN)
    return _client


# ══════════════════════════════════════════════════════════════════
#  CACHE EM DISCO
# ══════════════════════════════════════════════════════════════════
def _cache_path(chave: str) -> str:
    h = hashlib.md5(chave.encode()).hexdigest()
    return os.path.join(_CACHE_DIR, f"{h}.json")

def _cache_load(chave: str, ttl: int) -> list | dict | None:
    p = _cache_path(chave)
    if not os.path.exists(p):
        return None
    if time.time() - os.path.getmtime(p) > ttl:
        return None
    try:
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None

def _cache_save(chave: str, data) -> None:
    try:
        with open(_cache_path(chave), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, default=str)
    except Exception as e:
        print(f"  [AVISO] Não foi possível salvar cache: {e}")


# ══════════════════════════════════════════════════════════════════
#  BUSCA DE DADOS — VENDAS
# ══════════════════════════════════════════════════════════════════
def buscar_vendas(data_ini: str, data_fim: str) -> list[dict]:
    """
    Retorna lista de itens de venda do GestãoClick no período.
    Tenta os endpoints: vendas, pedidos (fallback).
    Campos normalizados: id, data, cliente, status, produto, qtd, v_bruto, v_desc, v_liq, vendedor, categoria
    """
    _prog(0.05, "Buscando vendas...")
    chave = f"vendas|{data_ini}|{data_fim}"
    cached = _cache_load(chave, _TTL_VENDAS)
    if cached is not None:
        print(f"  ✔ Vendas (cache): {len(cached)} registros")
        _prog(0.30, "Vendas carregadas do cache")
        return cached

    # Converte datas para o formato esperado pela API (YYYY-MM-DD)
    def br_to_iso(d: str) -> str:
        try:
            return datetime.strptime(d, "%d/%m/%Y").strftime("%Y-%m-%d")
        except Exception:
            return d

    params = {
        "data_inicio": br_to_iso(data_ini),
        "data_fim":    br_to_iso(data_fim),
    }

    # Tenta endpoint /vendas; se falhar, tenta /pedidos
    for endpoint in ["vendas", "pedidos"]:
        raw = _gck().paginar(endpoint, params)
        if raw:
            break

    if not raw:
        print("  [AVISO] Nenhuma venda encontrada no período ou credenciais inválidas.")
        return []

    registros: list[dict] = []
    for v in raw:
        # Normaliza campos (GestãoClick pode variar o nome dos campos entre planos)
        data_raw  = v.get("data_emissao") or v.get("data_venda") or v.get("data") or ""
        try:
            data_dt = datetime.strptime(data_raw[:10], "%Y-%m-%d")
        except Exception:
            data_dt = None

        cliente   = (v.get("cliente_nome") or v.get("nome_cliente") or v.get("cliente") or "")
        status    = (v.get("status") or v.get("situacao") or "").upper()
        vendedor  = (v.get("vendedor_nome") or v.get("vendedor") or "Sem Vendedor")
        cat       = (v.get("categoria") or v.get("grupo") or "SEM CATEGORIA").upper()
        nf        = str(v.get("numero_nf") or v.get("numero") or v.get("id") or "")
        id_venda  = str(v.get("id") or "")

        # Itens da venda
        itens = v.get("itens") or v.get("produtos") or []
        if itens:
            for it in itens:
                cod   = str(it.get("codigo") or it.get("produto_codigo") or "")
                desc  = str(it.get("descricao") or it.get("produto_nome") or it.get("nome") or "")
                unid  = str(it.get("unidade") or "UN")
                try:    qtd    = float(it.get("quantidade") or it.get("qtd") or 0)
                except: qtd    = 0.0
                try:    v_unit = float(it.get("valor_unitario") or it.get("preco") or 0)
                except: v_unit = 0.0
                try:    v_bruto= float(it.get("valor_total") or it.get("total") or qtd * v_unit)
                except: v_bruto= qtd * v_unit
                try:    v_desc = float(it.get("desconto") or it.get("valor_desconto") or 0)
                except: v_desc = 0.0
                registros.append({
                    "ID":           id_venda,
                    "NF":           nf,
                    "Data":         data_dt,
                    "Cliente":      cliente,
                    "Status":       status,
                    "Vendedor":     vendedor,
                    "Categoria":    cat,
                    "Cod. Produto": cod,
                    "Produto":      desc,
                    "Unidade":      unid,
                    "Qtd":          qtd,
                    "Vlr Unitário": v_unit,
                    "Vlr Bruto":    v_bruto,
                    "Desconto":     v_desc,
                    "Vlr Líquido":  v_bruto - v_desc,
                })
        else:
            # Venda sem detalhamento de itens — registra a venda como um todo
            try:    v_bruto = float(v.get("valor_total") or v.get("total") or 0)
            except: v_bruto = 0.0
            try:    v_desc  = float(v.get("desconto") or v.get("valor_desconto") or 0)
            except: v_desc  = 0.0
            registros.append({
                "ID":           id_venda,
                "NF":           nf,
                "Data":         data_dt,
                "Cliente":      cliente,
                "Status":       status,
                "Vendedor":     vendedor,
                "Categoria":    cat,
                "Cod. Produto": "",
                "Produto":      "",
                "Unidade":      "",
                "Qtd":          1.0,
                "Vlr Unitário": v_bruto - v_desc,
                "Vlr Bruto":    v_bruto,
                "Desconto":     v_desc,
                "Vlr Líquido":  v_bruto - v_desc,
            })

    _cache_save(chave, [
        {**r, "Data": r["Data"].isoformat() if isinstance(r["Data"], datetime) else r["Data"]}
        for r in registros
    ])

    # Converte datas de volta
    for r in registros:
        if isinstance(r["Data"], str) and r["Data"]:
            try:
                r["Data"] = datetime.fromisoformat(r["Data"][:19])
            except Exception:
                pass

    print(f"  ✔ {len(registros)} itens de venda")
    _prog(0.30, f"Vendas: {len(registros)} itens carregados")
    return registros


# ══════════════════════════════════════════════════════════════════
#  BUSCA DE DADOS — FINANCEIRO
# ══════════════════════════════════════════════════════════════════
def buscar_financeiro(data_ini: str, data_fim: str) -> dict:
    """
    Retorna {'receber': [...], 'pagar': [...]}
    Campos: id, descricao, cliente/fornecedor, valor, data_vencimento, data_pagamento, status, categoria
    """
    _prog(0.32, "Buscando financeiro...")
    resultado = {}

    def br_to_iso(d: str) -> str:
        try:
            return datetime.strptime(d, "%d/%m/%Y").strftime("%Y-%m-%d")
        except Exception:
            return d

    params = {"data_inicio": br_to_iso(data_ini), "data_fim": br_to_iso(data_fim)}

    for tipo, endpoint in [("receber", "contas_receber"), ("pagar", "contas_pagar")]:
        chave  = f"{tipo}|{data_ini}|{data_fim}"
        cached = _cache_load(chave, _TTL_VENDAS)
        if cached is not None:
            resultado[tipo] = cached
            print(f"  ✔ {tipo} (cache): {len(cached)} registros")
            continue

        raw = _gck().paginar(endpoint, params)
        normalizado = []
        for c in raw:
            def _float(k):
                try:    return float(c.get(k) or 0)
                except: return 0.0
            def _data(k):
                v = c.get(k, "")
                try:    return datetime.strptime(str(v)[:10], "%Y-%m-%d")
                except: return None

            normalizado.append({
                "ID":           str(c.get("id") or ""),
                "Descrição":    c.get("descricao") or c.get("historico") or "",
                "Pessoa":       (c.get("cliente_nome") or c.get("fornecedor_nome")
                                 or c.get("nome") or ""),
                "Valor":        _float("valor"),
                "Valor Pago":   _float("valor_pago") or _float("valor_recebido"),
                "Saldo":        _float("saldo") or _float("valor") - _float("valor_pago"),
                "Vencimento":   _data("data_vencimento"),
                "Pagamento":    _data("data_pagamento") or _data("data_recebimento"),
                "Status":       (c.get("status") or c.get("situacao") or "").upper(),
                "Categoria":    (c.get("categoria") or c.get("plano_conta") or "OUTROS").upper(),
                "Centro Custo": (c.get("centro_custo") or "").upper(),
            })

        _cache_save(chave, [
            {**r,
             "Vencimento": r["Vencimento"].isoformat() if isinstance(r["Vencimento"], datetime) else r["Vencimento"],
             "Pagamento":  r["Pagamento"].isoformat()  if isinstance(r["Pagamento"],  datetime) else r["Pagamento"],
            } for r in normalizado
        ])
        resultado[tipo] = normalizado

    _prog(0.45, "Financeiro carregado")
    return resultado


# ══════════════════════════════════════════════════════════════════
#  BUSCA DE DADOS — CLIENTES
# ══════════════════════════════════════════════════════════════════
def buscar_clientes() -> list[dict]:
    """Lista de clientes cadastrados."""
    _prog(0.47, "Buscando clientes...")
    chave  = "clientes"
    cached = _cache_load(chave, _TTL_OUTROS)
    if cached:
        print(f"  ✔ Clientes (cache): {len(cached)}")
        return cached
    raw = _gck().paginar("clientes")
    normalizado = []
    for c in raw:
        normalizado.append({
            "ID":      str(c.get("id") or ""),
            "Nome":    c.get("nome") or c.get("razao_social") or "",
            "Cidade":  c.get("cidade") or "",
            "UF":      c.get("uf") or c.get("estado") or "",
            "CNPJ":    c.get("cpf_cnpj") or "",
            "Tipo":    (c.get("tipo") or "PF").upper(),
            "Grupo":   (c.get("grupo") or c.get("categoria") or "").upper(),
            "Ativo":   c.get("ativo", True),
        })
    _cache_save(chave, normalizado)
    print(f"  ✔ {len(normalizado)} clientes")
    _prog(0.52, "Clientes carregados")
    return normalizado


# ══════════════════════════════════════════════════════════════════
#  BUSCA DE DADOS — PRODUTOS / ESTOQUE
# ══════════════════════════════════════════════════════════════════
def buscar_produtos() -> list[dict]:
    """Catálogo de produtos com estoque."""
    _prog(0.53, "Buscando produtos...")
    chave  = "produtos"
    cached = _cache_load(chave, _TTL_OUTROS)
    if cached:
        print(f"  ✔ Produtos (cache): {len(cached)}")
        return cached
    raw = _gck().paginar("produtos")
    normalizado = []
    for p in raw:
        def _float(k):
            try:    return float(p.get(k) or 0)
            except: return 0.0
        normalizado.append({
            "Código":     str(p.get("codigo") or p.get("id") or ""),
            "Descrição":  p.get("descricao") or p.get("nome") or "",
            "Categoria":  (p.get("categoria") or p.get("grupo") or "SEM CATEGORIA").upper(),
            "Marca":      (p.get("marca") or "").upper(),
            "Unidade":    p.get("unidade") or "UN",
            "Estoque":    _float("estoque") or _float("quantidade_estoque"),
            "Preco Venda":_float("preco_venda") or _float("valor_venda"),
            "Preco Custo":_float("preco_custo") or _float("valor_custo"),
            "Ativo":      p.get("ativo", True),
        })
    _cache_save(chave, normalizado)
    print(f"  ✔ {len(normalizado)} produtos")
    _prog(0.60, "Produtos carregados")
    return normalizado


# ══════════════════════════════════════════════════════════════════
#  BUSCA DE DADOS — CONTRATOS
# ══════════════════════════════════════════════════════════════════
def buscar_contratos() -> list[dict]:
    """Contratos ativos (recorrência)."""
    _prog(0.61, "Buscando contratos...")
    chave  = "contratos"
    cached = _cache_load(chave, _TTL_OUTROS)
    if cached:
        print(f"  ✔ Contratos (cache): {len(cached)}")
        return cached
    raw = _gck().paginar("contratos")
    normalizado = []
    for c in raw:
        def _float(k):
            try:    return float(c.get(k) or 0)
            except: return 0.0
        def _data(k):
            v = c.get(k, "")
            try:    return datetime.strptime(str(v)[:10], "%Y-%m-%d")
            except: return None
        normalizado.append({
            "ID":          str(c.get("id") or ""),
            "Número":      str(c.get("numero") or c.get("id") or ""),
            "Cliente":     c.get("cliente_nome") or c.get("nome") or "",
            "Descricao":   c.get("descricao") or c.get("objeto") or "",
            "Valor":       _float("valor") or _float("valor_total"),
            "Periodicidade":(c.get("periodicidade") or c.get("recorrencia") or "MENSAL").upper(),
            "Inicio":      _data("data_inicio"),
            "Fim":         _data("data_fim") or _data("data_vencimento"),
            "Status":      (c.get("status") or c.get("situacao") or "").upper(),
        })
    _cache_save(chave, [
        {**r,
         "Inicio": r["Inicio"].isoformat() if isinstance(r["Inicio"], datetime) else r["Inicio"],
         "Fim":    r["Fim"].isoformat()    if isinstance(r["Fim"],    datetime) else r["Fim"],
        } for r in normalizado
    ])
    print(f"  ✔ {len(normalizado)} contratos")
    _prog(0.68, "Contratos carregados")
    return normalizado


# ══════════════════════════════════════════════════════════════════
#  BUSCA DE DADOS — ORDENS DE SERVIÇO
# ══════════════════════════════════════════════════════════════════
def buscar_ordens_servico(data_ini: str, data_fim: str) -> list[dict]:
    """Ordens de serviço no período."""
    _prog(0.70, "Buscando ordens de serviço...")
    chave  = f"os|{data_ini}|{data_fim}"
    cached = _cache_load(chave, _TTL_VENDAS)
    if cached:
        print(f"  ✔ OS (cache): {len(cached)}")
        _prog(0.78, "OS carregadas do cache")
        return cached

    def br_to_iso(d: str) -> str:
        try:    return datetime.strptime(d, "%d/%m/%Y").strftime("%Y-%m-%d")
        except: return d

    params = {"data_inicio": br_to_iso(data_ini), "data_fim": br_to_iso(data_fim)}
    raw = _gck().paginar("ordens_servicos", params)

    normalizado = []
    for o in raw:
        def _float(k):
            try:    return float(o.get(k) or 0)
            except: return 0.0
        def _data(k):
            v = o.get(k, "")
            try:    return datetime.strptime(str(v)[:10], "%Y-%m-%d")
            except: return None

        normalizado.append({
            "ID":        str(o.get("id") or ""),
            "Número":    str(o.get("numero") or o.get("id") or ""),
            "Data":      _data("data_abertura") or _data("data"),
            "Cliente":   o.get("cliente_nome") or o.get("nome") or "",
            "Tecnico":   o.get("tecnico_nome") or o.get("tecnico") or o.get("responsavel") or "",
            "Descricao": o.get("descricao") or o.get("servico") or "",
            "Status":    (o.get("status") or o.get("situacao") or "").upper(),
            "Prioridade":(o.get("prioridade") or "NORMAL").upper(),
            "Valor":     _float("valor") or _float("valor_total"),
            "Fechamento":_data("data_fechamento") or _data("data_conclusao"),
        })

    _cache_save(chave, [
        {**r,
         "Data":      r["Data"].isoformat()      if isinstance(r["Data"],      datetime) else r["Data"],
         "Fechamento":r["Fechamento"].isoformat() if isinstance(r["Fechamento"], datetime) else r["Fechamento"],
        } for r in normalizado
    ])
    print(f"  ✔ {len(normalizado)} ordens de serviço")
    _prog(0.78, "OS carregadas")
    return normalizado


# ══════════════════════════════════════════════════════════════════
#  GERADOR DE EXCEL
# ══════════════════════════════════════════════════════════════════
COR_HD   = "1F4E79"
COR_HD_F = "FFFFFF"
COR_TOTL = "D6E4F0"
COR_ZEBR = "EBF4FB"
COR_SUBT = "BDD7EE"

def _hdr(ws, row, c1, c2, texto):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    cell = ws.cell(row=row, column=c1, value=texto)
    cell.font  = Font(bold=True, color=COR_HD_F, size=12)
    cell.fill  = PatternFill("solid", fgColor=COR_HD)
    cell.alignment = Alignment(horizontal="center", vertical="center")

def _borda(ws, r1, r2, c1, c2):
    brd = Border(left=Side(style="thin"), right=Side(style="thin"),
                 top=Side(style="thin"), bottom=Side(style="thin"))
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).border = brd

def _larguras(ws, d: dict):
    for col, w in d.items():
        ws.column_dimensions[col].width = w

def _sheet_vendas(wb, df: pd.DataFrame):
    ws = wb.create_sheet("Vendas por Produto")
    grp = (df.groupby(["Cod. Produto","Produto","Categoria"])
            .agg(Qtd=("Qtd","sum"), V_Bruto=("Vlr Bruto","sum"),
                 Desc=("Desconto","sum"), V_Liq=("Vlr Líquido","sum"), NFs=("ID","nunique"))
            .reset_index().sort_values("V_Liq", ascending=False))
    grp["Part %"] = grp["V_Liq"] / grp["V_Liq"].sum() * 100

    cols  = ["Cod. Produto","Produto","Categoria","Qtd","V_Bruto","Desc","V_Liq","Part %","NFs"]
    hdrs  = ["Código","Produto","Categoria","Qtd Total","Vlr Bruto","Desconto","Vlr Líquido","Part. %","Qtd NFs"]
    _hdr(ws, 1, 1, len(cols), "RANKING DE VENDAS POR PRODUTO — LOCVIX")
    ws.row_dimensions[1].height = 25
    hf = PatternFill("solid", fgColor=COR_HD)
    for ci, h in enumerate(hdrs, 1):
        c = ws.cell(row=2, column=ci, value=h)
        c.fill = hf; c.font = Font(bold=True, color="FFFFFF", size=10)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 28
    zf = PatternFill("solid", fgColor=COR_ZEBR)
    for ri, row in enumerate(grp[cols].itertuples(index=False), start=3):
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=ri, column=ci, value=round(val, 4) if isinstance(val, float) else val)
            cell.alignment = Alignment(vertical="center")
            if ri % 2 == 0: cell.fill = zf
            if ci == 4:   cell.number_format = "#,##0.00"
            elif ci in (5,6,7): cell.number_format = "R$ #,##0.00"
            elif ci == 8:
                cell.number_format = "0.00%"
                cell.value = (val/100) if isinstance(val, float) else val
    total_row = len(grp) + 3
    ws.cell(row=total_row, column=3, value="TOTAL").font = Font(bold=True)
    for ci, cl in [(4,"D"),(5,"E"),(6,"F"),(7,"G")]:
        c = ws.cell(row=total_row, column=ci,
                    value=f"=SUM({cl}3:{cl}{total_row-1})")
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor=COR_TOTL)
        c.number_format = "#,##0.00" if ci == 4 else "R$ #,##0.00"
    _borda(ws, 2, total_row, 1, len(cols))
    _larguras(ws, {"A":14,"B":44,"C":22,"D":10,"E":14,"F":14,"G":14,"H":10,"I":8})
    ws.freeze_panes = "A3"
    return grp

def _sheet_clientes(wb, df_vendas: pd.DataFrame):
    ws = wb.create_sheet("Vendas por Cliente")
    grp = (df_vendas.groupby(["Cliente"])
            .agg(NFs=("ID","nunique"), V_Bruto=("Vlr Bruto","sum"),
                 Desc=("Desconto","sum"), V_Liq=("Vlr Líquido","sum"))
            .reset_index().sort_values("V_Liq", ascending=False))
    grp["Part %"] = grp["V_Liq"] / grp["V_Liq"].sum() * 100
    cols  = ["Cliente","NFs","V_Bruto","Desc","V_Liq","Part %"]
    hdrs  = ["Cliente","Nº NFs / Pedidos","Vlr Bruto","Desconto","Vlr Líquido","Part. %"]
    _hdr(ws, 1, 1, len(cols), "RANKING DE VENDAS POR CLIENTE — LOCVIX")
    ws.row_dimensions[1].height = 25
    hf = PatternFill("solid", fgColor=COR_HD)
    for ci, h in enumerate(hdrs, 1):
        c = ws.cell(row=2, column=ci, value=h)
        c.fill = hf; c.font = Font(bold=True, color="FFFFFF", size=10)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 28
    zf = PatternFill("solid", fgColor=COR_ZEBR)
    for ri, row in enumerate(grp[cols].itertuples(index=False), start=3):
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=ri, column=ci, value=round(val, 4) if isinstance(val, float) else val)
            cell.alignment = Alignment(vertical="center")
            if ri % 2 == 0: cell.fill = zf
            if ci in (3,4,5): cell.number_format = "R$ #,##0.00"
            elif ci == 6:
                cell.number_format = "0.00%"
                cell.value = (val/100) if isinstance(val, float) else val
    total_row = len(grp) + 3
    for ci, cl in [(3,"C"),(4,"D"),(5,"E")]:
        c = ws.cell(row=total_row, column=ci,
                    value=f"=SUM({cl}3:{cl}{total_row-1})")
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor=COR_TOTL)
        c.number_format = "R$ #,##0.00"
    _borda(ws, 2, total_row, 1, len(cols))
    _larguras(ws, {"A":42,"B":14,"C":14,"D":14,"E":14,"F":10})
    ws.freeze_panes = "A3"

def _sheet_financeiro(wb, receber: list, pagar: list):
    for dados, nome, titulo in [
        (receber, "Contas a Receber", "CONTAS A RECEBER — LOCVIX"),
        (pagar,   "Contas a Pagar",   "CONTAS A PAGAR — LOCVIX"),
    ]:
        ws = wb.create_sheet(nome)
        if not dados:
            ws.cell(row=1, column=1, value=f"Sem dados de {nome}")
            continue
        cols  = ["Descrição","Pessoa","Valor","Valor Pago","Saldo","Vencimento","Status","Categoria"]
        _hdr(ws, 1, 1, len(cols), titulo)
        ws.row_dimensions[1].height = 25
        hf = PatternFill("solid", fgColor=COR_HD)
        for ci, h in enumerate(cols, 1):
            c = ws.cell(row=2, column=ci, value=h)
            c.fill = hf; c.font = Font(bold=True, color="FFFFFF", size=10)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[2].height = 28
        zf = PatternFill("solid", fgColor=COR_ZEBR)
        for ri, r in enumerate(dados, start=3):
            vals = [
                r.get("Descrição",""), r.get("Pessoa",""),
                r.get("Valor",0), r.get("Valor Pago",0), r.get("Saldo",0),
                r.get("Vencimento"), r.get("Status",""), r.get("Categoria",""),
            ]
            for ci, val in enumerate(vals, 1):
                cell = ws.cell(row=ri, column=ci, value=val)
                cell.alignment = Alignment(vertical="center")
                if ri % 2 == 0: cell.fill = zf
                if ci in (3,4,5): cell.number_format = "R$ #,##0.00"
                elif ci == 6 and val: cell.number_format = "DD/MM/YYYY"
        _borda(ws, 2, len(dados)+2, 1, len(cols))
        _larguras(ws, {"A":30,"B":34,"C":14,"D":14,"E":14,"F":12,"G":16,"H":22})
        ws.freeze_panes = "A3"

def _sheet_os(wb, os_list: list):
    ws = wb.create_sheet("Ordens de Serviço")
    if not os_list:
        ws.cell(row=1, column=1, value="Sem dados de OS no período")
        return
    cols  = ["Número","Data","Cliente","Tecnico","Descricao","Status","Prioridade","Valor","Fechamento"]
    _hdr(ws, 1, 1, len(cols), "ORDENS DE SERVIÇO — LOCVIX")
    ws.row_dimensions[1].height = 25
    hf = PatternFill("solid", fgColor=COR_HD)
    for ci, h in enumerate(cols, 1):
        c = ws.cell(row=2, column=ci, value=h)
        c.fill = hf; c.font = Font(bold=True, color="FFFFFF", size=10)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 28
    zf = PatternFill("solid", fgColor=COR_ZEBR)
    for ri, r in enumerate(os_list, start=3):
        vals = [r.get(k) for k in cols]
        for ci, val in enumerate(vals, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.alignment = Alignment(vertical="center")
            if ri % 2 == 0: cell.fill = zf
            if ci == 8: cell.number_format = "R$ #,##0.00"
            elif ci in (2,9) and val: cell.number_format = "DD/MM/YYYY"
    _borda(ws, 2, len(os_list)+2, 1, len(cols))
    _larguras(ws, {"A":10,"B":12,"C":34,"D":22,"E":36,"F":16,"G":12,"H":12,"I":12})
    ws.freeze_panes = "A3"

def _sheet_contratos(wb, contratos: list):
    ws = wb.create_sheet("Contratos")
    if not contratos:
        ws.cell(row=1, column=1, value="Sem contratos cadastrados")
        return
    cols = ["Número","Cliente","Descricao","Valor","Periodicidade","Inicio","Fim","Status"]
    _hdr(ws, 1, 1, len(cols), "CONTRATOS / RECORRÊNCIA — LOCVIX")
    ws.row_dimensions[1].height = 25
    hf = PatternFill("solid", fgColor=COR_HD)
    for ci, h in enumerate(cols, 1):
        c = ws.cell(row=2, column=ci, value=h)
        c.fill = hf; c.font = Font(bold=True, color="FFFFFF", size=10)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 28
    zf = PatternFill("solid", fgColor=COR_ZEBR)
    for ri, r in enumerate(contratos, start=3):
        vals = [r.get(k) for k in cols]
        for ci, val in enumerate(vals, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.alignment = Alignment(vertical="center")
            if ri % 2 == 0: cell.fill = zf
            if ci == 4: cell.number_format = "R$ #,##0.00"
            elif ci in (6,7) and val: cell.number_format = "DD/MM/YYYY"
    _borda(ws, 2, len(contratos)+2, 1, len(cols))
    _larguras(ws, {"A":10,"B":34,"C":36,"D":14,"E":16,"F":12,"G":12,"H":16})
    ws.freeze_panes = "A3"

def gerar_excel(df_vendas: pd.DataFrame, receber: list, pagar: list,
                os_list: list, contratos: list, caminho: str) -> str:
    """Gera arquivo Excel completo com todas as abas."""
    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)  # remove a aba padrão

    grp_prod = _sheet_vendas(wb, df_vendas)
    _sheet_clientes(wb, df_vendas)
    _sheet_financeiro(wb, receber, pagar)
    _sheet_os(wb, os_list)
    _sheet_contratos(wb, contratos)

    wb.save(caminho)
    print(f"  ✔ Excel salvo: {caminho}")
    return caminho


# ══════════════════════════════════════════════════════════════════
#  DASHBOARD HTML
# ══════════════════════════════════════════════════════════════════
def gerar_dashboard_html(
    df_vendas: pd.DataFrame,
    receber:   list,
    pagar:     list,
    os_list:   list,
    contratos: list,
    caminho:   str,
    data_ini:  str,
    data_fim:  str,
) -> str:
    """Gera dashboard HTML interativo completo para Locvix."""
    import json as _json

    # ── Datas ───────────────────────────────────────────────────────
    try:
        dt_min = df_vendas["Data"].dropna().min()
        dt_max = df_vendas["Data"].dropna().max()
    except Exception:
        dt_min = dt_max = datetime.now()
    dt_min_iso = dt_min.strftime("%Y-%m-%d") if isinstance(dt_min, datetime) else data_ini
    dt_max_iso = dt_max.strftime("%Y-%m-%d") if isinstance(dt_max, datetime) else data_fim
    periodo    = f"{data_ini} a {data_fim}"
    agora_str  = datetime.now().strftime("%d/%m/%Y %H:%M")

    # ── Logo em base64 ──────────────────────────────────────────────
    _logo_b64   = ""
    _logo_mime  = "image/png"
    _logo_candidates = [
        (os.path.join(_BASE_DIR, "logo_locvix.png"),  "image/png"),
        (os.path.join(_BASE_DIR, "logo_locvix.jpg"),  "image/jpeg"),
        (os.path.join(_BASE_DIR, "logo_locvix.jfif"), "image/jpeg"),
        (os.path.join(_BASE_DIR, "logo locvix.png"),  "image/png"),
        (os.path.join(_BASE_DIR, "logo locvix.jpg"),  "image/jpeg"),
        (os.path.join(_BASE_DIR, "logo locvix.jfif"), "image/jpeg"),
    ]
    for _lp, _lm in _logo_candidates:
        if os.path.exists(_lp):
            with open(_lp, "rb") as f:
                _logo_b64  = base64.b64encode(f.read()).decode()
            _logo_mime = _lm
            break

    _logo_alfa_b64  = ""
    _logo_alfa_path = os.path.join(_BASE_DIR, "logo_alfa.jpg")
    if not os.path.exists(_logo_alfa_path):
        _logo_alfa_path = r"Z:\codigos\Fabio\Logo Alfa.jpg"
    if os.path.exists(_logo_alfa_path):
        with open(_logo_alfa_path, "rb") as f:
            _logo_alfa_b64 = base64.b64encode(f.read()).decode()

    logo_tag = (f'<img src="data:{_logo_mime};base64,{_logo_b64}" '
                f'style="height:52px;width:auto;border-radius:6px;'
                f'box-shadow:0 2px 8px rgba(0,0,0,.3);object-fit:contain;"/>'
               ) if _logo_b64 else '<span style="font-size:24px;font-weight:800">LOCVIX</span>'
    logo_alfa_tag = (f'<img src="data:image/jpeg;base64,{_logo_alfa_b64}" '
                     f'style="height:30px;width:auto;object-fit:contain;border-radius:4px;"/>'
                    ) if _logo_alfa_b64 else "<strong>Alfa Soluções</strong>"

    # ── Dados para o JS ─────────────────────────────────────────────
    def prep_vendas():
        rows = []
        for _, r in df_vendas.iterrows():
            d = r.get("Data")
            rows.append({
                "id":        str(r.get("ID","") or ""),
                "data":      d.strftime("%Y-%m-%d") if isinstance(d, datetime) else "",
                "cliente":   str(r.get("Cliente",""))[:50],
                "cod":       str(r.get("Cod. Produto","") or ""),
                "produto":   str(r.get("Produto","") or "")[:50],
                "categoria": str(r.get("Categoria","") or "SEM CATEGORIA"),
                "vendedor":  str(r.get("Vendedor","") or "Sem Vendedor"),
                "status":    str(r.get("Status","") or ""),
                "qtd":       round(float(r.get("Qtd",0) or 0), 4),
                "bruto":     round(float(r.get("Vlr Bruto",0) or 0), 2),
                "desc":      round(float(r.get("Desconto",0) or 0), 2),
                "liq":       round(float(r.get("Vlr Líquido",0) or 0), 2),
            })
        return rows

    def prep_financeiro(lista: list, tipo: str):
        rows = []
        for r in lista:
            def _dt(k):
                v = r.get(k)
                if isinstance(v, datetime): return v.strftime("%Y-%m-%d")
                if isinstance(v, str) and v: return v[:10]
                return ""
            rows.append({
                "id":       r.get("ID",""),
                "desc":     r.get("Descrição","")[:60],
                "pessoa":   r.get("Pessoa","")[:50],
                "valor":    round(float(r.get("Valor",0) or 0), 2),
                "pago":     round(float(r.get("Valor Pago",0) or 0), 2),
                "saldo":    round(float(r.get("Saldo",0) or 0), 2),
                "venc":     _dt("Vencimento"),
                "pgto":     _dt("Pagamento"),
                "status":   r.get("Status",""),
                "cat":      r.get("Categoria",""),
                "tipo":     tipo,
            })
        return rows

    def prep_os():
        rows = []
        for r in os_list:
            def _dt(k):
                v = r.get(k)
                if isinstance(v, datetime): return v.strftime("%Y-%m-%d")
                return ""
            rows.append({
                "id":   r.get("Número",""),
                "data": _dt("Data"),
                "cli":  r.get("Cliente","")[:40],
                "tec":  r.get("Tecnico","")[:30],
                "desc": r.get("Descricao","")[:60],
                "st":   r.get("Status",""),
                "prio": r.get("Prioridade",""),
                "val":  round(float(r.get("Valor",0) or 0), 2),
                "fech": _dt("Fechamento"),
            })
        return rows

    raw_vendas = prep_vendas()
    raw_rec    = prep_financeiro(receber, "receber")
    raw_pag    = prep_financeiro(pagar, "pagar")
    raw_os     = prep_os()
    raw_contr  = [{
        "id":     c.get("Número",""),
        "cli":    c.get("Cliente","")[:40],
        "val":    round(float(c.get("Valor",0) or 0), 2),
        "period": c.get("Periodicidade",""),
        "st":     c.get("Status",""),
    } for c in contratos]

    jv = lambda v: _json.dumps(v, ensure_ascii=False)

    # ── Categorias únicas ───────────────────────────────────────────
    categorias   = sorted(df_vendas["Categoria"].dropna().unique().tolist())
    vendedores   = sorted(df_vendas["Vendedor"].dropna().unique().tolist())
    opt_cat  = "\n".join(f'<option value="{c}">{c}</option>' for c in categorias)
    opt_vend = "\n".join(f'<option value="{v}">{v}</option>' for v in vendedores)

    # ── HTML completo ───────────────────────────────────────────────
    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Dashboard — Locvix</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{font-family:'Segoe UI',Arial,sans-serif;background:#f0f4f8;color:#1e293b;font-size:14px;}}

/* ── TOPBAR ── */
.topbar{{background:linear-gradient(135deg,#0f2027 0%,#1a3a4a 60%,#203a43 100%);
  color:#fff;padding:18px 32px;display:flex;align-items:center;justify-content:space-between;
  box-shadow:0 2px 12px rgba(0,0,0,.35);}}
.topbar-title{{font-size:21px;font-weight:800;letter-spacing:.8px;}}
.topbar .sub{{font-size:12px;color:#90cdf4;margin-top:3px;}}
.topbar .periodo{{font-size:12px;color:#7dd3fc;text-align:right;}}

/* ── FILTER BAR ── */
.filter-bar{{background:#fff;border-bottom:2px solid #e2e8f0;padding:12px 32px;
  display:flex;align-items:flex-end;gap:14px;flex-wrap:wrap;
  box-shadow:0 1px 4px rgba(0,0,0,.07);position:sticky;top:0;z-index:100;}}
.filter-group{{display:flex;flex-direction:column;gap:3px;}}
.filter-group label{{font-size:11px;font-weight:700;color:#718096;text-transform:uppercase;letter-spacing:.4px;}}
.filter-group input[type=date],
.filter-group select{{border:1px solid #e2e8f0;border-radius:7px;padding:7px 10px;
  font-size:13px;color:#1e293b;background:#f8fafc;outline:none;min-width:140px;
  cursor:pointer;transition:border-color .15s;}}
.filter-group input:focus,.filter-group select:focus{{border-color:#1a3a4a;background:#fff;}}
.filter-sep{{width:1px;height:36px;background:#e2e8f0;margin:0 4px;align-self:center;}}
.btn{{padding:8px 18px;border:none;border-radius:7px;font-size:13px;font-weight:700;cursor:pointer;transition:all .15s;}}
.btn-apply{{background:#1a3a4a;color:#fff;}}
.btn-apply:hover{{background:#0f2027;}}
.btn-clear{{background:#e2e8f0;color:#4a5568;}}
.btn-clear:hover{{background:#cbd5e0;}}
.filter-info{{font-size:12px;color:#059669;font-weight:600;margin-left:auto;align-self:center;white-space:nowrap;}}

/* ── LAYOUT ── */
.container{{max-width:1520px;margin:0 auto;padding:22px 20px;}}
.section-title{{font-size:14px;font-weight:800;color:#1a3a4a;margin:26px 0 11px;
  padding-left:10px;border-left:4px solid #1a3a4a;letter-spacing:.5px;text-transform:uppercase;}}

/* ── KPI CARDS ── */
.kpi-grid{{display:grid;gap:14px;margin-bottom:24px;}}
.kpi-grid.col3{{grid-template-columns:repeat(3,1fr);}}
.kpi-grid.col4{{grid-template-columns:repeat(4,1fr);}}
.kpi-grid.col5{{grid-template-columns:repeat(5,1fr);}}
.kpi-grid.col6{{grid-template-columns:repeat(6,1fr);}}
@media(max-width:1100px){{.kpi-grid.col6,.kpi-grid.col5{{grid-template-columns:repeat(3,1fr);}}}}
@media(max-width:700px){{.kpi-grid{{grid-template-columns:repeat(2,1fr)!important;}}}}
.kpi-card{{background:#fff;border-radius:10px;padding:16px 14px;
  box-shadow:0 1px 4px rgba(0,0,0,.1);border-top:4px solid #1a3a4a;text-align:center;
  transition:transform .15s;}}
.kpi-card:hover{{transform:translateY(-3px);box-shadow:0 4px 12px rgba(0,0,0,.15);}}
.kpi-card.green{{border-top-color:#059669;}}
.kpi-card.teal{{border-top-color:#0891b2;}}
.kpi-card.orange{{border-top-color:#d97706;}}
.kpi-card.red{{border-top-color:#dc2626;}}
.kpi-card.purple{{border-top-color:#7c3aed;}}
.kpi-card.blue{{border-top-color:#2563eb;}}
.kpi-label{{font-size:11px;color:#718096;font-weight:700;text-transform:uppercase;letter-spacing:.4px;margin-bottom:6px;}}
.kpi-value{{font-size:20px;font-weight:800;color:#1e293b;}}
.kpi-value.small{{font-size:15px;}}

/* ── CHARTS ── */
.chart-row{{display:grid;gap:16px;margin-bottom:16px;}}
.chart-row.col2{{grid-template-columns:1fr 1fr;}}
.chart-row.col3{{grid-template-columns:1fr 1fr 1fr;}}
@media(max-width:900px){{.chart-row.col2,.chart-row.col3{{grid-template-columns:1fr;}}}}
.chart-card{{background:#fff;border-radius:10px;padding:18px 20px;box-shadow:0 1px 4px rgba(0,0,0,.1);}}
.chart-card h3{{font-size:13px;font-weight:700;color:#4a5568;margin-bottom:14px;
  text-transform:uppercase;letter-spacing:.4px;}}

/* ── TABLE CARD ── */
.table-card{{background:#fff;border-radius:10px;padding:18px 20px;
  box-shadow:0 1px 4px rgba(0,0,0,.1);margin-bottom:16px;overflow-x:auto;}}
.table-card h3{{font-size:13px;font-weight:700;color:#4a5568;margin-bottom:14px;
  text-transform:uppercase;letter-spacing:.4px;}}
table.data-tbl{{width:100%;border-collapse:collapse;font-size:13px;}}
table.data-tbl thead th{{background:#1a3a4a;color:#fff;padding:9px 14px;text-align:left;
  font-size:11px;font-weight:700;letter-spacing:.3px;white-space:nowrap;}}
table.data-tbl thead th.num{{text-align:right;}}
table.data-tbl tbody tr:nth-child(even){{background:#f8fafc;}}
table.data-tbl tbody tr:hover{{background:#edf2f7;}}
table.data-tbl tbody td{{padding:8px 14px;border-bottom:1px solid #e2e8f0;color:#1e293b;}}
table.data-tbl tbody td.num{{text-align:right;font-variant-numeric:tabular-nums;font-weight:600;}}
table.data-tbl tfoot td{{background:#e2e8f0;font-weight:800;padding:9px 14px;}}
table.data-tbl tfoot td.num{{text-align:right;}}

/* ── STATUS BADGES ── */
.badge{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;}}
.badge.verde{{background:#d1fae5;color:#065f46;}}
.badge.vermelho{{background:#fef2f2;color:#991b1b;}}
.badge.amarelo{{background:#fef3c7;color:#92400e;}}
.badge.cinza{{background:#f1f5f9;color:#64748b;}}

/* ── FOOTER ── */
.footer{{margin-top:32px;padding:20px 32px;border-top:1px solid #e2e8f0;background:#f8fafc;
  display:flex;align-items:center;justify-content:space-between;gap:16px;flex-wrap:wrap;}}
.footer-dev{{display:flex;align-items:center;gap:10px;font-size:12px;color:#94a3b8;}}
.footer-dev strong{{color:#64748b;}}
.footer-sep{{width:1px;height:32px;background:#e2e8f0;}}
.footer-gen{{font-size:11px;color:#b0bec5;text-align:right;}}

/* ── DARK MODE BUTTON ── */
#btn-theme{{position:fixed;bottom:24px;right:24px;z-index:10000;width:48px;height:48px;
  border-radius:50%;border:2px solid rgba(255,255,255,.18);cursor:pointer;font-size:20px;
  box-shadow:0 4px 18px rgba(0,0,0,.35);background:#1a3a4a;color:#f1f5f9;
  transition:all .2s;display:flex;align-items:center;justify-content:center;line-height:1;}}
#btn-theme:hover{{transform:scale(1.12);}}

/* ── DARK MODE ── */
body[data-theme="dark"]{{background:#0f172a;color:#e2e8f0;}}
body[data-theme="dark"] .filter-bar{{background:#1e293b;border-color:#334155;}}
body[data-theme="dark"] .filter-group label{{color:#94a3b8;}}
body[data-theme="dark"] .filter-group input,
body[data-theme="dark"] .filter-group select{{background:#0f172a;color:#e2e8f0;border-color:#475569;}}
body[data-theme="dark"] .filter-sep{{background:#334155;}}
body[data-theme="dark"] .btn-clear{{background:#334155;color:#cbd5e0;}}
body[data-theme="dark"] .btn-apply{{background:#3b82f6;}}
body[data-theme="dark"] .section-title{{color:#93c5fd;border-color:#3b82f6;}}
body[data-theme="dark"] .kpi-card{{background:#1e293b;box-shadow:0 2px 8px rgba(0,0,0,.5);}}
body[data-theme="dark"] .kpi-label{{color:#94a3b8;}}
body[data-theme="dark"] .kpi-value{{color:#f1f5f9;}}
body[data-theme="dark"] .chart-card{{background:#1e293b;}}
body[data-theme="dark"] .chart-card h3{{color:#94a3b8;}}
body[data-theme="dark"] .table-card{{background:#1e293b;}}
body[data-theme="dark"] .table-card h3{{color:#94a3b8;}}
body[data-theme="dark"] table.data-tbl thead th{{background:#0f2027;}}
body[data-theme="dark"] table.data-tbl tbody td{{color:#e2e8f0;border-color:#334155;}}
body[data-theme="dark"] table.data-tbl tbody tr:nth-child(even){{background:#1e293b;}}
body[data-theme="dark"] table.data-tbl tbody tr:hover{{background:#334155;}}
body[data-theme="dark"] table.data-tbl tfoot td{{background:#334155;}}
body[data-theme="dark"] .footer{{background:#1e293b;border-color:#334155;}}
body[data-theme="dark"] .footer-dev{{color:#64748b;}}
body[data-theme="dark"] .footer-sep{{background:#334155;}}
body[data-theme="dark"] #btn-theme{{background:#e2e8f0;color:#1e293b;}}

/* ── SENHA (Password Gate) ── */
#pg-overlay{{position:fixed;inset:0;z-index:99999;display:flex;align-items:center;
  justify-content:center;background:linear-gradient(135deg,#0f2027 0%,#203a43 100%);}}
#pg-box{{background:#fff;border-radius:18px;padding:42px 48px;
  box-shadow:0 12px 60px rgba(0,0,0,.6);text-align:center;width:90%;max-width:380px;}}
#pg-box .pg-logo{{height:54px;margin-bottom:18px;border-radius:6px;object-fit:contain;}}
#pg-box h2{{font-size:20px;font-weight:800;color:#1e293b;margin-bottom:6px;}}
#pg-box .pg-sub{{font-size:13px;color:#64748b;margin-bottom:22px;line-height:1.55;}}
#pg-inp{{width:100%;padding:13px 16px;border:2px solid #e2e8f0;border-radius:10px;
  font-size:17px;text-align:center;letter-spacing:4px;outline:none;font-weight:600;
  transition:border-color .2s;}}
#pg-inp:focus{{border-color:#1a3a4a;}}
#pg-inp.err-shake{{border-color:#dc2626;animation:shake .35s ease;}}
@keyframes shake{{0%,100%{{transform:translateX(0)}}25%{{transform:translateX(-6px)}}75%{{transform:translateX(6px)}}}}
#pg-btn{{width:100%;margin-top:14px;padding:13px;border:none;border-radius:10px;
  background:#1a3a4a;color:#fff;font-size:15px;font-weight:700;cursor:pointer;
  transition:background .2s;letter-spacing:.5px;}}
#pg-btn:hover{{background:#0f2027;}}
#pg-err{{color:#dc2626;font-size:13px;margin-top:10px;min-height:18px;font-weight:600;}}
body.pg-locked{{overflow:hidden;}}
</style>
</head>
<body class="pg-locked">

<button id="btn-theme" onclick="toggleTheme()" title="Alternar modo claro/escuro">🌙</button>

<!-- PASSWORD GATE -->
<div id="pg-overlay">
  <div id="pg-box">
    {(f'<img class="pg-logo" src="data:image/png;base64,{_logo_b64}"/>') if _logo_b64 else '<div style="font-size:36px;margin-bottom:16px">🔒</div>'}
    <h2>&#128274; Dashboard Protegido</h2>
    <p class="pg-sub">Insira a senha para acessar<br/>o painel da <strong>Locvix</strong></p>
    <input id="pg-inp" type="password" placeholder="&#8226;&#8226;&#8226;&#8226;&#8226;&#8226;&#8226;&#8226;"
      autocomplete="off" onkeydown="if(event.key==='Enter')pgEntrar()"/>
    <button id="pg-btn" onclick="pgEntrar()">&#128275; Entrar</button>
    <div id="pg-err"></div>
  </div>
</div>

<!-- TOPBAR -->
<div class="topbar">
  <div style="display:flex;align-items:center;gap:18px;">
    {logo_tag}
    <div>
      <div class="topbar-title">📊 DASHBOARD — LOCVIX</div>
      <div class="sub">Análise de Vendas · Financeiro · Clientes · OS · Contratos · GestãoClick ERP</div>
    </div>
  </div>
  <div class="periodo">
    Período: <strong>{periodo}</strong><br/>
    <span style="color:#94a3b8;font-size:11px;">Gerado em {agora_str}</span>
  </div>
</div>

<!-- FILTROS -->
<div class="filter-bar">
  <div class="filter-group">
    <label>📅 Data Início</label>
    <input type="date" id="fDateIni" value="{dt_min_iso}"/>
  </div>
  <div class="filter-group">
    <label>📅 Data Fim</label>
    <input type="date" id="fDateFim" value="{dt_max_iso}"/>
  </div>
  <div class="filter-sep"></div>
  <div class="filter-group">
    <label>🏷 Categoria</label>
    <select id="fCat">
      <option value="">— Todas —</option>
      {opt_cat}
    </select>
  </div>
  <div class="filter-group">
    <label>👤 Vendedor</label>
    <select id="fVend">
      <option value="">— Todos —</option>
      {opt_vend}
    </select>
  </div>
  <div class="filter-sep"></div>
  <button class="btn btn-apply" onclick="aplicarFiltros()">▶ Aplicar</button>
  <button class="btn btn-clear" onclick="limparFiltros()">✕ Limpar</button>
  <div class="filter-info" id="filtroInfo"></div>
</div>

<div class="container">

  <!-- ── KPIs DE VENDAS ── -->
  <div class="section-title">💰 Vendas — Resumo do Período</div>
  <div class="kpi-grid col6">
    <div class="kpi-card green">
      <div class="kpi-label">Fat. Líquido</div>
      <div class="kpi-value small" id="kLiq">—</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Fat. Bruto</div>
      <div class="kpi-value small" id="kBruto">—</div>
    </div>
    <div class="kpi-card orange">
      <div class="kpi-label">Desconto Total</div>
      <div class="kpi-value small" id="kDesc">—</div>
    </div>
    <div class="kpi-card blue">
      <div class="kpi-label">Qtd Vendida</div>
      <div class="kpi-value" id="kQtd">—</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Clientes Ativos</div>
      <div class="kpi-value" id="kCli">—</div>
    </div>
    <div class="kpi-card teal">
      <div class="kpi-label">Pedidos / NFs</div>
      <div class="kpi-value" id="kNFs">—</div>
    </div>
  </div>

  <!-- ── KPIs FINANCEIRO ── -->
  <div class="section-title">🏦 Financeiro — Resumo</div>
  <div class="kpi-grid col4">
    <div class="kpi-card green">
      <div class="kpi-label">A Receber (Total)</div>
      <div class="kpi-value small" id="kRecTotal">—</div>
    </div>
    <div class="kpi-card teal">
      <div class="kpi-label">A Receber (Recebido)</div>
      <div class="kpi-value small" id="kRecPago">—</div>
    </div>
    <div class="kpi-card red">
      <div class="kpi-label">A Pagar (Total)</div>
      <div class="kpi-value small" id="kPagTotal">—</div>
    </div>
    <div class="kpi-card orange">
      <div class="kpi-label">A Pagar (Pago)</div>
      <div class="kpi-value small" id="kPagPago">—</div>
    </div>
  </div>

  <!-- ── KPIs OS + CONTRATOS ── -->
  <div class="section-title">🔧 Operações — OS & Contratos</div>
  <div class="kpi-grid col4">
    <div class="kpi-card blue">
      <div class="kpi-label">OS Abertas</div>
      <div class="kpi-value" id="kOsAbert">—</div>
    </div>
    <div class="kpi-card green">
      <div class="kpi-label">OS Concluídas</div>
      <div class="kpi-value" id="kOsConc">—</div>
    </div>
    <div class="kpi-card purple">
      <div class="kpi-label">Contratos Ativos</div>
      <div class="kpi-value" id="kContrAt">—</div>
    </div>
    <div class="kpi-card teal">
      <div class="kpi-label">MRR (Recorrência)</div>
      <div class="kpi-value small" id="kMRR">—</div>
    </div>
  </div>

  <!-- ── GRÁFICOS DE VENDAS ── -->
  <div class="section-title">📈 Análise de Vendas</div>
  <div class="chart-row col2">
    <div class="chart-card">
      <h3>📅 Faturamento Líquido por Mês</h3>
      <canvas id="chartMensal" height="140"></canvas>
    </div>
    <div class="chart-card">
      <h3>🍩 Participação por Categoria</h3>
      <canvas id="chartCategoria" height="140"></canvas>
    </div>
  </div>
  <div class="chart-row col2">
    <div class="chart-card">
      <h3>🏆 Top 10 Produtos — Fat. Líquido</h3>
      <canvas id="chartProdutos" height="150"></canvas>
    </div>
    <div class="chart-card">
      <h3>👥 Top 10 Clientes — Fat. Líquido</h3>
      <canvas id="chartClientes" height="150"></canvas>
    </div>
  </div>

  <!-- ── GRÁFICOS FINANCEIROS ── -->
  <div class="section-title">💳 Financeiro</div>
  <div class="chart-row col2">
    <div class="chart-card">
      <h3>📥 Contas a Receber por Categoria</h3>
      <canvas id="chartReceber" height="130"></canvas>
    </div>
    <div class="chart-card">
      <h3>📤 Contas a Pagar por Categoria</h3>
      <canvas id="chartPagar" height="130"></canvas>
    </div>
  </div>

  <!-- ── TABELA TOP PRODUTOS ── -->
  <div class="section-title">📋 Top 20 Produtos</div>
  <div class="table-card">
    <h3>Ranking de produtos por faturamento líquido</h3>
    <table class="data-tbl">
      <thead>
        <tr>
          <th>#</th>
          <th>Código</th>
          <th>Produto</th>
          <th>Categoria</th>
          <th class="num">Qtd</th>
          <th class="num">Fat. Bruto</th>
          <th class="num">Desconto</th>
          <th class="num">Fat. Líquido</th>
          <th class="num">Part. %</th>
        </tr>
      </thead>
      <tbody id="tblProdCorpo"></tbody>
      <tfoot><tr id="tblProdTotal"></tr></tfoot>
    </table>
  </div>

  <!-- ── TABELA TOP CLIENTES ── -->
  <div class="section-title">👥 Top 20 Clientes</div>
  <div class="table-card">
    <h3>Ranking de clientes por faturamento líquido</h3>
    <table class="data-tbl">
      <thead>
        <tr>
          <th>#</th>
          <th>Cliente</th>
          <th>Vendedor</th>
          <th class="num">Pedidos</th>
          <th class="num">Fat. Bruto</th>
          <th class="num">Desconto</th>
          <th class="num">Fat. Líquido</th>
          <th class="num">Part. %</th>
        </tr>
      </thead>
      <tbody id="tblCliCorpo"></tbody>
      <tfoot><tr id="tblCliTotal"></tr></tfoot>
    </table>
  </div>

  <!-- ── TABELA CONTAS A RECEBER ── -->
  <div class="section-title">📥 Contas a Receber</div>
  <div class="table-card">
    <h3>Últimas 30 contas mais relevantes</h3>
    <table class="data-tbl">
      <thead>
        <tr>
          <th>Descrição</th>
          <th>Pessoa</th>
          <th class="num">Valor</th>
          <th class="num">Recebido</th>
          <th class="num">Saldo</th>
          <th>Vencimento</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody id="tblRecCorpo"></tbody>
    </table>
  </div>

  <!-- ── TABELA OS ── -->
  <div class="section-title">🔧 Ordens de Serviço</div>
  <div class="table-card">
    <h3>Últimas 30 OS no período</h3>
    <table class="data-tbl">
      <thead>
        <tr>
          <th>Nº</th>
          <th>Data</th>
          <th>Cliente</th>
          <th>Técnico</th>
          <th>Descrição</th>
          <th>Status</th>
          <th class="num">Valor</th>
        </tr>
      </thead>
      <tbody id="tblOsCorpo"></tbody>
    </table>
  </div>

</div>

<div class="footer">
  <div class="footer-dev">
    {logo_alfa_tag}
    <div>
      <div>Desenvolvido por <strong>Fabrício Zamprogno</strong></div>
      <div>em parceria com <strong>Alfa Soluções Consultoria</strong></div>
    </div>
  </div>
  <div class="footer-sep"></div>
  <div class="footer-gen">
    Dashboard Locvix — GestãoClick ERP<br/>
    Gerado em {agora_str}
  </div>
</div>

<script>
// ═══════════════════════════════════════════════
//  DADOS BRUTOS
// ═══════════════════════════════════════════════
const VENDAS    = {jv(raw_vendas)};
const RECEBER   = {jv(raw_rec)};
const PAGAR     = {jv(raw_pag)};
const OS_LIST   = {jv(raw_os)};
const CONTRATOS = {jv(raw_contr)};

const BRL = v => 'R$\u00a0' + v.toLocaleString('pt-BR',{{minimumFractionDigits:2,maximumFractionDigits:2}});
const NUM = v => v.toLocaleString('pt-BR');
const PCT = (a,b) => b > 0 ? (a/b*100).toFixed(1)+'%' : '—';

// ── Cores ──────────────────────────────────────
const CORES = ['#1a3a4a','#0891b2','#059669','#d97706','#7c3aed',
               '#2563eb','#dc2626','#0d9488','#65a30d','#ea580c',
               '#db2777','#6366f1','#78716c','#0369a1','#15803d'];

// ═══════════════════════════════════════════════
//  SENHA (SHA-256)
//  Padrão: "locvix2026" — altere o hash para mudar a senha
// ═══════════════════════════════════════════════
const _PG = 'd7e7e89803ee3dae7236276af4188e5a3a0c3b95107f17bd4bcd47a272c33d0b';
// ↑ hash de "locvix2026" — substitua pelo hash SHA-256 da sua senha

async function pgEntrar() {{
  const inp = document.getElementById('pg-inp');
  const err = document.getElementById('pg-err');
  err.textContent = '';
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(inp.value));
  const hex = Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,'0')).join('');
  if (hex === _PG) {{
    const ov = document.getElementById('pg-overlay');
    ov.style.transition = 'opacity .4s';
    ov.style.opacity = '0';
    setTimeout(()=>{{ov.remove();document.body.classList.remove('pg-locked');}}, 420);
  }} else {{
    err.textContent = '\u274c Senha incorreta.';
    inp.value = '';
    inp.classList.add('err-shake');
    setTimeout(()=>inp.classList.remove('err-shake'), 400);
    inp.focus();
  }}
}}
document.addEventListener('DOMContentLoaded', ()=>{{
  const i = document.getElementById('pg-inp'); if(i) i.focus();
}});

// ═══════════════════════════════════════════════
//  ESTADO DE FILTROS
// ═══════════════════════════════════════════════
let dadosFilt = VENDAS;

function filtrar() {{
  const ini  = document.getElementById('fDateIni').value;
  const fim  = document.getElementById('fDateFim').value;
  const cat  = document.getElementById('fCat').value;
  const vend = document.getElementById('fVend').value;
  dadosFilt = VENDAS.filter(r => {{
    if (ini  && r.data < ini)  return false;
    if (fim  && r.data > fim)  return false;
    if (cat  && r.categoria !== cat)  return false;
    if (vend && r.vendedor  !== vend) return false;
    return true;
  }});
  const info = document.getElementById('filtroInfo');
  info.textContent = dadosFilt.length === VENDAS.length ? '' :
    `\u2714 ${{NUM(dadosFilt.length)}} de ${{NUM(VENDAS.length)}} itens filtrados`;
}}

// ═══════════════════════════════════════════════
//  CHARTS
// ═══════════════════════════════════════════════
const charts = {{}};
function destroyChart(id) {{ if(charts[id]) {{ charts[id].destroy(); delete charts[id]; }} }}

function mkMensal(rows) {{
  const m = {{}};
  rows.forEach(r => {{
    const ym = (r.data||'').substring(0,7);
    if (!ym) return;
    m[ym] = (m[ym]||0) + r.liq;
  }});
  const entries = Object.entries(m).sort((a,b)=>a[0]>b[0]?1:-1);
  const meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  const labels = entries.map(([ym]) => {{
    const [y,mo] = ym.split('-');
    return meses[parseInt(mo)-1]+'/'+y.substring(2);
  }});
  destroyChart('chartMensal');
  charts['chartMensal'] = new Chart(document.getElementById('chartMensal'), {{
    type: 'bar',
    data: {{
      labels,
      datasets: [{{
        label: 'Fat. Líquido',
        data: entries.map(e=>Math.round(e[1])),
        backgroundColor: '#0891b2', borderRadius: 5, borderSkipped: false,
      }}]
    }},
    options: {{
      plugins: {{legend:{{display:false}},
        tooltip:{{callbacks:{{label:c=>BRL(c.raw)}}}}}},
      scales: {{
        y: {{ticks:{{callback:v=>'R$'+(v>=1000?(v/1000).toFixed(0)+'k':v)}},grid:{{color:'#e2e8f0'}}}},
        x: {{grid:{{display:false}}}}
      }}
    }}
  }});
}}

function mkCategoria(rows) {{
  const m = {{}};
  rows.forEach(r => {{ m[r.categoria||'SEM CATEGORIA'] = (m[r.categoria||'SEM CATEGORIA']||0)+r.liq; }});
  const entries = Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,10);
  destroyChart('chartCategoria');
  charts['chartCategoria'] = new Chart(document.getElementById('chartCategoria'), {{
    type: 'doughnut',
    data: {{
      labels: entries.map(e=>e[0]),
      datasets: [{{ data: entries.map(e=>Math.round(e[1])),
        backgroundColor: CORES, borderWidth: 2, borderColor: '#fff', hoverOffset: 10 }}]
    }},
    options: {{
      cutout: '55%',
      plugins: {{
        legend: {{position:'right',labels:{{font:{{size:11}},boxWidth:14}}}},
        tooltip: {{callbacks:{{label:c=>c.label+': '+BRL(c.raw)}}}}
      }}
    }}
  }});
}}

function mkHorizBar(id, entries) {{
  destroyChart(id);
  charts[id] = new Chart(document.getElementById(id), {{
    type: 'bar',
    data: {{
      labels: entries.map(e=>e[0].length>28?e[0].substring(0,25)+'...':e[0]),
      datasets: [{{ data: entries.map(e=>Math.round(e[1])),
        backgroundColor: CORES, borderRadius: 4, borderSkipped: false }}]
    }},
    options: {{
      indexAxis: 'y',
      plugins: {{legend:{{display:false}},
        tooltip:{{callbacks:{{label:c=>BRL(c.raw)}}}}}},
      scales: {{
        x: {{ticks:{{callback:v=>'R$'+(v>=1000?(v/1000).toFixed(0)+'k':v)}},grid:{{color:'#e2e8f0'}}}},
        y: {{grid:{{display:false}}}}
      }}
    }}
  }});
}}

function mkDonut(id, entries) {{
  destroyChart(id);
  charts[id] = new Chart(document.getElementById(id), {{
    type: 'doughnut',
    data: {{
      labels: entries.map(e=>e[0]),
      datasets: [{{ data: entries.map(e=>Math.round(e[1])),
        backgroundColor: CORES, borderWidth: 2, borderColor: '#fff' }}]
    }},
    options: {{
      cutout: '55%',
      plugins: {{
        legend: {{position:'right',labels:{{font:{{size:11}},boxWidth:14}}}},
        tooltip: {{callbacks:{{label:c=>c.label+': '+BRL(c.raw)}}}}
      }}
    }}
  }});
}}

// ═══════════════════════════════════════════════
//  ATUALIZAÇÃO DOS KPIs
// ═══════════════════════════════════════════════
function atualizarKPIVendas(rows) {{
  const liq   = rows.reduce((s,r)=>s+r.liq,0);
  const bruto = rows.reduce((s,r)=>s+r.bruto,0);
  const desc  = rows.reduce((s,r)=>s+r.desc,0);
  const qtd   = rows.reduce((s,r)=>s+r.qtd,0);
  const clis  = new Set(rows.map(r=>r.cliente)).size;
  const nfs   = new Set(rows.map(r=>r.id)).size;
  document.getElementById('kLiq').textContent   = BRL(liq);
  document.getElementById('kBruto').textContent = BRL(bruto);
  document.getElementById('kDesc').textContent  = BRL(desc);
  document.getElementById('kQtd').textContent   = NUM(Math.round(qtd));
  document.getElementById('kCli').textContent   = NUM(clis);
  document.getElementById('kNFs').textContent   = NUM(nfs);
}}

function atualizarKPIFinanceiro() {{
  const recTot  = RECEBER.reduce((s,r)=>s+r.valor,0);
  const recPago = RECEBER.reduce((s,r)=>s+r.pago,0);
  const pagTot  = PAGAR.reduce((s,r)=>s+r.valor,0);
  const pagPago = PAGAR.reduce((s,r)=>s+r.pago,0);
  document.getElementById('kRecTotal').textContent = BRL(recTot);
  document.getElementById('kRecPago').textContent  = BRL(recPago);
  document.getElementById('kPagTotal').textContent = BRL(pagTot);
  document.getElementById('kPagPago').textContent  = BRL(pagPago);
}}

function atualizarKPIOS() {{
  const total   = OS_LIST.length;
  const abertas = OS_LIST.filter(o=>!['CONCLUIDO','FECHADA','CANCELADA'].includes(o.st)).length;
  const conc    = OS_LIST.filter(o=>['CONCLUIDO','CONCLUÍDA','FECHADA'].includes(o.st)).length;
  const contativos = CONTRATOS.filter(c=>c.st && !['CANCELADO','INATIVO'].includes(c.st.toUpperCase())).length;
  const mrr = CONTRATOS
    .filter(c=>c.st && !['CANCELADO','INATIVO'].includes(c.st.toUpperCase()))
    .filter(c=>c.period && (c.period.includes('MENSAL') || c.period.includes('MÊS')))
    .reduce((s,c)=>s+c.val,0);
  document.getElementById('kOsAbert').textContent  = NUM(abertas);
  document.getElementById('kOsConc').textContent   = NUM(conc);
  document.getElementById('kContrAt').textContent  = NUM(contativos);
  document.getElementById('kMRR').textContent      = BRL(mrr);
}}

// ═══════════════════════════════════════════════
//  TABELAS
// ═══════════════════════════════════════════════
function renderTblProdutos(rows) {{
  const m = {{}};
  rows.forEach(r => {{
    const k = r.cod || r.produto;
    if (!m[k]) m[k] = {{cod:r.cod,prod:r.produto,cat:r.categoria,qtd:0,bruto:0,desc:0,liq:0}};
    m[k].qtd+=r.qtd; m[k].bruto+=r.bruto; m[k].desc+=r.desc; m[k].liq+=r.liq;
  }});
  const arr = Object.values(m).sort((a,b)=>b.liq-a.liq).slice(0,20);
  const totLiq = arr.reduce((s,r)=>s+r.liq,0);
  let html = '';
  arr.forEach((r,i) => {{
    html += `<tr>
      <td>${{i+1}}</td><td><code style="font-size:11px;color:#64748b">${{r.cod||'—'}}</code></td>
      <td>${{r.prod||'—'}}</td><td>${{r.cat||'—'}}</td>
      <td class="num">${{r.qtd.toLocaleString('pt-BR',{{maximumFractionDigits:0}})}}</td>
      <td class="num">${{BRL(r.bruto)}}</td><td class="num">${{BRL(r.desc)}}</td>
      <td class="num">${{BRL(r.liq)}}</td>
      <td class="num">${{PCT(r.liq,totLiq)}}</td>
    </tr>`;
  }});
  document.getElementById('tblProdCorpo').innerHTML = html;
  const tTot = arr.reduce((s,r)=>s+r.bruto,0);
  const tDesc = arr.reduce((s,r)=>s+r.desc,0);
  const tQtd = arr.reduce((s,r)=>s+r.qtd,0);
  document.getElementById('tblProdTotal').innerHTML =
    `<td colspan="4">TOTAL (TOP 20)</td>
     <td class="num">${{tQtd.toLocaleString('pt-BR',{{maximumFractionDigits:0}})}}</td>
     <td class="num">${{BRL(tTot)}}</td><td class="num">${{BRL(tDesc)}}</td>
     <td class="num">${{BRL(totLiq)}}</td><td class="num">100%</td>`;
}}

function renderTblClientes(rows) {{
  const m = {{}};
  rows.forEach(r => {{
    const k = r.cliente;
    if (!m[k]) m[k] = {{cli:r.cliente,vend:r.vendedor,ids:new Set(),bruto:0,desc:0,liq:0}};
    m[k].ids.add(r.id); m[k].bruto+=r.bruto; m[k].desc+=r.desc; m[k].liq+=r.liq;
  }});
  const arr = Object.values(m).sort((a,b)=>b.liq-a.liq).slice(0,20);
  const totLiq = arr.reduce((s,r)=>s+r.liq,0);
  let html = '';
  arr.forEach((r,i) => {{
    html += `<tr>
      <td>${{i+1}}</td><td>${{r.cli||'—'}}</td><td>${{r.vend||'—'}}</td>
      <td class="num">${{NUM(r.ids.size)}}</td>
      <td class="num">${{BRL(r.bruto)}}</td><td class="num">${{BRL(r.desc)}}</td>
      <td class="num">${{BRL(r.liq)}}</td>
      <td class="num">${{PCT(r.liq,totLiq)}}</td>
    </tr>`;
  }});
  document.getElementById('tblCliCorpo').innerHTML = html;
  const tTot = arr.reduce((s,r)=>s+r.bruto,0);
  const tDesc = arr.reduce((s,r)=>s+r.desc,0);
  document.getElementById('tblCliTotal').innerHTML =
    `<td colspan="3">TOTAL (TOP 20)</td><td class="num">—</td>
     <td class="num">${{BRL(tTot)}}</td><td class="num">${{BRL(tDesc)}}</td>
     <td class="num">${{BRL(totLiq)}}</td><td class="num">100%</td>`;
}}

function statusBadge(st) {{
  const s = (st||'').toUpperCase();
  if (['PAGO','RECEBIDO','QUITADO','CONCLUIDO','CONCLUÍDA','ATIVO'].some(x=>s.includes(x)))
    return `<span class="badge verde">${{st}}</span>`;
  if (['VENCIDO','INADIMPLENTE','ATRASADO','CANCELADO'].some(x=>s.includes(x)))
    return `<span class="badge vermelho">${{st}}</span>`;
  if (['ABERTO','PENDENTE','AGUARDANDO'].some(x=>s.includes(x)))
    return `<span class="badge amarelo">${{st}}</span>`;
  return `<span class="badge cinza">${{st||'—'}}</span>`;
}}

function renderTblReceber() {{
  const sorted = [...RECEBER].sort((a,b)=>b.valor-a.valor).slice(0,30);
  let html = '';
  sorted.forEach(r => {{
    html += `<tr>
      <td>${{r.desc||'—'}}</td><td>${{r.pessoa||'—'}}</td>
      <td class="num">${{BRL(r.valor)}}</td>
      <td class="num">${{BRL(r.pago)}}</td>
      <td class="num" style="color:${{r.saldo>0?'#dc2626':'#059669'}}">${{BRL(r.saldo)}}</td>
      <td>${{r.venc?r.venc.split('-').reverse().join('/'):'—'}}</td>
      <td>${{statusBadge(r.status)}}</td>
    </tr>`;
  }});
  document.getElementById('tblRecCorpo').innerHTML = html;
}}

function renderTblOS() {{
  const sorted = [...OS_LIST].sort((a,b)=>(b.data||'')>(a.data||'')?1:-1).slice(0,30);
  let html = '';
  sorted.forEach(r => {{
    html += `<tr>
      <td>${{r.id||'—'}}</td>
      <td>${{r.data?r.data.split('-').reverse().join('/'):'—'}}</td>
      <td>${{r.cli||'—'}}</td><td>${{r.tec||'—'}}</td>
      <td title="${{r.desc||''}}">${{(r.desc||'').substring(0,45)}}...</td>
      <td>${{statusBadge(r.st)}}</td>
      <td class="num">${{BRL(r.val)}}</td>
    </tr>`;
  }});
  document.getElementById('tblOsCorpo').innerHTML = html;
}}

// ═══════════════════════════════════════════════
//  GRÁFICOS FINANCEIROS
// ═══════════════════════════════════════════════
function mkFinanceiro() {{
  // Contas a receber por categoria
  const mRec = {{}};
  RECEBER.forEach(r => {{ mRec[r.cat||'OUTROS'] = (mRec[r.cat||'OUTROS']||0)+r.valor; }});
  const recEntries = Object.entries(mRec).sort((a,b)=>b[1]-a[1]).slice(0,8);
  mkDonut('chartReceber', recEntries);

  // Contas a pagar por categoria
  const mPag = {{}};
  PAGAR.forEach(r => {{ mPag[r.cat||'OUTROS'] = (mPag[r.cat||'OUTROS']||0)+r.valor; }});
  const pagEntries = Object.entries(mPag).sort((a,b)=>b[1]-a[1]).slice(0,8);
  mkDonut('chartPagar', pagEntries);
}}

// ═══════════════════════════════════════════════
//  ATUALIZAÇÃO GERAL
// ═══════════════════════════════════════════════
function atualizar() {{
  const rows = dadosFilt;
  atualizarKPIVendas(rows);
  atualizarKPIFinanceiro();
  atualizarKPIOS();

  // Gráficos de vendas
  mkMensal(rows);
  mkCategoria(rows);

  // Top 10 produtos e clientes
  const mProd = {{}};
  rows.forEach(r=>{{ const k=r.cod||r.produto;
    if(!mProd[k]) mProd[k]=[r.produto,0]; mProd[k][1]+=r.liq; }});
  mkHorizBar('chartProdutos', Object.values(mProd).sort((a,b)=>b[1]-a[1]).slice(0,10));

  const mCli = {{}};
  rows.forEach(r=>{{ mCli[r.cliente]=(mCli[r.cliente]||0)+r.liq; }});
  mkHorizBar('chartClientes', Object.entries(mCli).sort((a,b)=>b[1]-a[1]).slice(0,10));

  // Tabelas
  renderTblProdutos(rows);
  renderTblClientes(rows);
  renderTblReceber();
  renderTblOS();
  mkFinanceiro();
}}

function aplicarFiltros() {{ filtrar(); atualizar(); }}
function limparFiltros() {{
  document.getElementById('fDateIni').value = '{dt_min_iso}';
  document.getElementById('fDateFim').value = '{dt_max_iso}';
  document.getElementById('fCat').value  = '';
  document.getElementById('fVend').value = '';
  document.getElementById('filtroInfo').textContent = '';
  dadosFilt = VENDAS;
  atualizar();
}}

// ═══════════════════════════════════════════════
//  DARK MODE
// ═══════════════════════════════════════════════
function toggleTheme() {{
  const body = document.body;
  const isDark = body.getAttribute('data-theme') === 'dark';
  const next = isDark ? 'light' : 'dark';
  body.setAttribute('data-theme', next);
  document.getElementById('btn-theme').textContent = isDark ? '\ud83c\udf19' : '\u2600\ufe0f';
  try {{ localStorage.setItem('locvix-theme', next); }} catch(e){{}}
  const gc = isDark ? '#e2e8f0' : '#e2e8f0';
  const lc = isDark ? '#94a3b8' : '#4a5568';
  Chart.defaults.color = isDark ? '#94a3b8' : '#4a5568';
  Object.values(Chart.instances).forEach(ch => {{
    if (ch.options.scales) {{
      Object.values(ch.options.scales).forEach(sc => {{
        if (sc.grid) sc.grid.color = isDark ? '#334155' : '#e2e8f0';
        if (sc.ticks) sc.ticks.color = isDark ? '#94a3b8' : '#4a5568';
      }});
    }}
    if (ch.options.plugins?.legend) ch.options.plugins.legend.labels.color = isDark ? '#94a3b8' : '#4a5568';
    ch.update();
  }});
}}

document.addEventListener('DOMContentLoaded', () => {{
  const i = document.getElementById('pg-inp'); if(i) i.focus();
  try {{
    const saved = localStorage.getItem('locvix-theme');
    if (saved === 'dark') {{
      document.body.setAttribute('data-theme','dark');
      const btn = document.getElementById('btn-theme');
      if (btn) btn.textContent = '\u2600\ufe0f';
    }}
  }} catch(e) {{}}
}});

// ─── Init ───────────────────────────────────────
Chart.defaults.font.family = "'Segoe UI', Arial, sans-serif";
Chart.defaults.font.size   = 12;
dadosFilt = VENDAS;
atualizar();
</script>
</body>
</html>"""

    # sanitize surrogates that Windows may inject
    html_clean = html.encode("utf-8", errors="xmlcharrefreplace").decode("utf-8")
    with open(caminho, "w", encoding="utf-8") as f:
        f.write(html_clean)
    print(f"  OK Dashboard HTML salvo: {caminho}")
    return html


# ══════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════
def main(
    saida_html:  str | None = None,
    saida_excel: str | None = None,
    data_ini:    str | None = None,
    data_fim:    str | None = None,
) -> str | None:
    """
    Executa a coleta completa de dados e gera Excel + Dashboard HTML.

    Parâmetros
    ----------
    saida_html  : caminho do HTML gerado (usa SAIDA_HTML global se None)
    saida_excel : caminho do Excel (usa SAIDA_EXCEL global se None; "" = pula Excel)
    data_ini    : início do período "DD/MM/AAAA" (usa DATA_INI global se None)
    data_fim    : fim do período   "DD/MM/AAAA" (usa DATA_FIM global se None)

    Retorna o conteúdo HTML como string (útil para Streamlit).
    """
    global _client

    d_ini  = data_ini  or DATA_INI
    d_fim  = data_fim  or DATA_FIM
    h_path = saida_html  or SAIDA_HTML
    x_path = saida_excel if saida_excel is not None else SAIDA_EXCEL

    print(f"\n{'='*60}")
    print(f"  LOCVIX — ANÁLISE DE DADOS VIA GESTÃOCLICK API")
    print(f"  Período: {d_ini} → {d_fim}")
    print(f"{'='*60}")

    # Verifica credenciais
    if GCK_ACCESS_TOKEN == "SEU_ACCESS_TOKEN_AQUI":
        print("\n  ⚠ ATENÇÃO: Credenciais não configuradas!")
        print("  Configure as variáveis de ambiente:")
        print("    GCK_ACCESS_TOKEN  = seu Access Token do GestãoClick")
        print("    GCK_SECRET_TOKEN  = seu Secret Token do GestãoClick")
        print("  Ou edite as constantes GCK_ACCESS_TOKEN / GCK_SECRET_TOKEN")
        print("  no início deste arquivo.\n")
        print("  Gerando dashboard com dados de DEMONSTRAÇÃO...\n")
        return _main_demo(h_path, d_ini, d_fim)

    # Reinicia o cliente com as credenciais atuais
    _client = GCKClient(GCK_ACCESS_TOKEN, GCK_SECRET_TOKEN)

    # ── Coleta de dados ────────────────────────────────────────────
    vendas    = buscar_vendas(d_ini, d_fim)
    financ    = buscar_financeiro(d_ini, d_fim)
    clientes  = buscar_clientes()
    produtos  = buscar_produtos()
    contratos = buscar_contratos()
    os_list   = buscar_ordens_servico(d_ini, d_fim)

    receber = financ.get("receber", [])
    pagar   = financ.get("pagar", [])

    _prog(0.82, "Processando dados...")

    # ── DataFrame de vendas ─────────────────────────────────────────
    if vendas:
        df = pd.DataFrame(vendas)
        # Garante coluna Data como datetime
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    else:
        df = pd.DataFrame(columns=[
            "ID","NF","Data","Cliente","Status","Vendedor","Categoria",
            "Cod. Produto","Produto","Unidade","Qtd",
            "Vlr Unitário","Vlr Bruto","Desconto","Vlr Líquido"
        ])
        print("  ⚠ Sem dados de vendas no período.")

    # ── Gera arquivos ───────────────────────────────────────────────
    _prog(0.88, "Gerando Excel...")
    if x_path:
        gerar_excel(df, receber, pagar, os_list, contratos, x_path)

    _prog(0.93, "Gerando Dashboard HTML...")
    html_content = gerar_dashboard_html(
        df_vendas=df, receber=receber, pagar=pagar,
        os_list=os_list, contratos=contratos,
        caminho=h_path, data_ini=d_ini, data_fim=d_fim,
    )

    _prog(1.0, "✔ Concluído!")
    print(f"\n  ✅ Tudo pronto!")
    print(f"  📊 Dashboard: {h_path}")
    if x_path:
        print(f"  📋 Excel:     {x_path}")
    print(f"{'='*60}\n")
    return html_content


# ══════════════════════════════════════════════════════════════════
#  MODO DEMONSTRAÇÃO (sem credenciais)
# ══════════════════════════════════════════════════════════════════
def _main_demo(h_path: str, d_ini: str, d_fim: str) -> str:
    """
    Gera um dashboard de demonstração com dados fictícios
    para validar o layout sem precisar das credenciais.
    """
    import random
    rng = random.Random(42)

    clientes_demo = ["Empresa Alpha Ltda", "Beta Soluções S/A", "Gamma Tech",
                     "Delta Serviços", "Epsilon Comercial", "Zeta Distribuidora",
                     "Eta Logística", "Theta Consultoria", "Iota Indústria"]
    cats_demo     = ["SERVIÇOS", "PRODUTOS", "SUPORTE", "INSTALAÇÃO", "MANUTENÇÃO"]
    vendedores_demo = ["João Silva", "Maria Souza", "Carlos Oliveira", "Ana Santos"]
    produtos_demo = [
        {"cod":"PROD001","desc":"Sistema ERP Módulo"},
        {"cod":"PROD002","desc":"Licença Software Anual"},
        {"cod":"SERV001","desc":"Hora de Suporte Técnico"},
        {"cod":"SERV002","desc":"Implantação"},
        {"cod":"SERV003","desc":"Treinamento Usuários"},
        {"cod":"PROD003","desc":"Firewall Hardware"},
        {"cod":"PROD004","desc":"Switch 24 portas"},
        {"cod":"SERV004","desc":"Monitoramento Remoto"},
    ]

    # Gera vendas fictícias
    vendas_demo = []
    try:
        d1 = datetime.strptime(d_ini, "%d/%m/%Y")
        d2 = datetime.strptime(d_fim, "%d/%m/%Y")
    except Exception:
        d1 = datetime(datetime.now().year, 1, 1)
        d2 = datetime.now()

    delta = (d2 - d1).days
    for i in range(min(delta * 3, 500)):
        data = d1 + timedelta(days=rng.randint(0, max(delta, 1)))
        prod = rng.choice(produtos_demo)
        qtd  = rng.choice([1, 1, 1, 2, 5, 10, 20])
        prec = rng.uniform(50, 2000)
        desc = rng.uniform(0, 0.1) * prec * qtd
        vbruto = round(prec * qtd, 2)
        vliq   = round(vbruto - desc, 2)
        vendas_demo.append({
            "ID":           f"PED{i+1:04d}",
            "NF":           f"{1000+i}",
            "Data":         data,
            "Cliente":      rng.choice(clientes_demo),
            "Status":       rng.choice(["EMITIDO","EMITIDO","PAGO","CANCELADO"]),
            "Vendedor":     rng.choice(vendedores_demo),
            "Categoria":    rng.choice(cats_demo),
            "Cod. Produto": prod["cod"],
            "Produto":      prod["desc"],
            "Unidade":      "UN",
            "Qtd":          float(qtd),
            "Vlr Unitário": round(prec, 2),
            "Vlr Bruto":    vbruto,
            "Desconto":     round(desc, 2),
            "Vlr Líquido":  vliq,
        })

    df = pd.DataFrame(vendas_demo)

    # Financeiro fictício
    receber_demo = [
        {"ID": f"R{i}", "Descrição": f"NF {1000+i}", "Pessoa": rng.choice(clientes_demo),
         "Valor": round(rng.uniform(500,5000),2), "Valor Pago": 0.0,
         "Saldo": round(rng.uniform(500,5000),2),
         "Vencimento": (datetime.now() + timedelta(days=rng.randint(-30,60))),
         "Pagamento": None, "Status": rng.choice(["ABERTO","VENCIDO","RECEBIDO"]),
         "Categoria": rng.choice(cats_demo), "Centro Custo": ""}
        for i in range(30)
    ]
    pagar_demo = [
        {"ID": f"P{i}", "Descrição": f"Fornecedor {i}", "Pessoa": f"Fornecedor {i+1}",
         "Valor": round(rng.uniform(200,3000),2), "Valor Pago": round(rng.uniform(0,200),2),
         "Saldo": round(rng.uniform(0,3000),2),
         "Vencimento": (datetime.now() + timedelta(days=rng.randint(-15,45))),
         "Pagamento": None, "Status": rng.choice(["ABERTO","PAGO","VENCIDO"]),
         "Categoria": rng.choice(["PESSOAL","INFRA","FORNECEDOR","IMPOSTOS"]), "Centro Custo": ""}
        for i in range(20)
    ]

    # OS fictícias
    os_demo = [
        {"ID": f"OS{i:03d}", "Número": f"OS{i:03d}",
         "Data": d1 + timedelta(days=rng.randint(0, max(delta,1))),
         "Cliente": rng.choice(clientes_demo),
         "Tecnico": rng.choice(["Carlos", "Ana", "Marcos"]),
         "Descricao": rng.choice(["Manutenção preventiva","Instalação","Suporte remoto","Atualização"]),
         "Status": rng.choice(["ABERTA","CONCLUÍDO","AGUARDANDO","CANCELADA"]),
         "Prioridade": rng.choice(["ALTA","NORMAL","BAIXA"]),
         "Valor": round(rng.uniform(100,800),2),
         "Fechamento": None}
        for i in range(50)
    ]

    # Contratos fictícios
    contratos_demo = [
        {"ID": f"C{i:03d}", "Número": f"C{i:03d}",
         "Cliente": rng.choice(clientes_demo),
         "Descricao": "Contrato de manutenção e suporte",
         "Valor": rng.choice([500,800,1200,1500,2000,3000]),
         "Periodicidade": "MENSAL",
         "Inicio": d1 - timedelta(days=rng.randint(30,365)),
         "Fim": d1 + timedelta(days=rng.randint(180,730)),
         "Status": rng.choice(["ATIVO","ATIVO","ATIVO","VENCIDO","CANCELADO"])}
        for i in range(15)
    ]

    html_content = gerar_dashboard_html(
        df_vendas=df, receber=receber_demo, pagar=pagar_demo,
        os_list=os_demo, contratos=contratos_demo,
        caminho=h_path, data_ini=d_ini, data_fim=d_fim,
    )
    print(f"\n  📊 Dashboard DEMO gerado: {h_path}")
    return html_content


# ══════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    import webbrowser
    html = main()
    # Abre automaticamente no navegador padrão
    if os.path.exists(SAIDA_HTML):
        webbrowser.open(f"file:///{SAIDA_HTML.replace(chr(92),'/')}")
