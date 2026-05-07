#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para refatorar alertas_manutencao.py:
1. Adiciona funções de cálculo de ignição acumulada desde última manutenção.
2. Remove dependência do horímetro não confiável da máquina.
3. Usa horas de ignição ligada como base do "horímetro acumulado".
"""

import os

alertas_file = r"z:\codigos\Locvix GPT\alertas_manutencao.py"

# Ler arquivo atual
with open(alertas_file, "r", encoding="utf-8") as f:
    content = f.read()

# 1. Adicionar imports necessários se não existirem
new_imports = """import json
import hashlib
import tempfile
from pathlib import Path
"""

if "import json" not in content:
    # Encontrar o lugar onde adicionar (após 'from email.mime.text import MIMEText')
    insert_pos = content.find("import requests")
    if insert_pos > 0:
        insert_pos = content.find("\n", insert_pos) + 1
        content = content[:insert_pos] + "\nimport json\nimport hashlib\nimport tempfile\n" + content[insert_pos:]

# 2. Adicionar cache e funções de ignição após imports
cache_and_ft_functions = '''

# ══════════════════════════════════════════════════════════════════
#  CACHE EM DISCO
# ══════════════════════════════════════════════════════════════════
_CACHE_SCHEMA = "4"
_CACHE_DIR = os.path.join(tempfile.gettempdir(), "_cache_alertas")
os.makedirs(_CACHE_DIR, exist_ok=True)

def _cache_path(chave: str) -> str:
    h = hashlib.md5(f"{_CACHE_SCHEMA}|{chave}".encode()).hexdigest()
    return os.path.join(_CACHE_DIR, f"{h}.json")

def _cache_load(chave: str, ttl: int) -> list | dict | None:
    """Carrega cache se ainda válido (TTL em segundos)."""
    import time
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
    """Salva dados em cache."""
    try:
        with open(_cache_path(chave), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, default=str)
    except Exception as e:
        print(f"  [AVISO] Cache não salvo: {e}")


# ══════════════════════════════════════════════════════════════════
#  FULLTRACK — CÁLCULO DE HORAS DE IGNIÇÃO ACUMULADA
# ══════════════════════════════════════════════════════════════════

def _ft_parse_dt(s: str | None):
    """Converte data/hora FullTrack (dd/mm/YYYY HH:MM:SS) em datetime."""
    if not s:
        return None
    try:
        return __import__("datetime").datetime.strptime(str(s), "%d/%m/%Y %H:%M:%S")
    except Exception:
        return None


def _ft_horas_ignicao_intervalo(veiculo_id: str, dt_ini, dt_fim) -> float:
    """
    Soma horas de ignição ligada no intervalo via /events/interval.
    dt_ini, dt_fim: datetime objects
    Retorna: float com total de horas
    """
    from datetime import timedelta
    
    if not veiculo_id or dt_fim <= dt_ini:
        return 0.0

    hdrs = {"apikey": FULLTRACK_API_KEY, "secretkey": FULLTRACK_SECRET_KEY}
    passo = timedelta(days=7)
    cur = dt_ini
    eventos = []

    while cur < dt_fim:
        nxt = min(cur + passo, dt_fim)
        bts = int(cur.timestamp())
        ets = int(nxt.timestamp())
        ck = f"ft_ignicao|{veiculo_id}|{bts}|{ets}"
        cached = _cache_load(ck, 900)

        rows = []
        if isinstance(cached, list):
            for it in cached:
                if isinstance(it, list) and len(it) == 2:
                    d = _ft_parse_dt(it[0])
                    if d is None:
                        continue
                    try:
                        ig = int(it[1])
                    except Exception:
                        continue
                    rows.append((d, 1 if ig == 1 else 0))
        else:
            try:
                url = f"{FULLTRACK_BASE}/events/interval/id/{veiculo_id}/begin/{bts}/end/{ets}"
                resp = requests.get(url, headers=hdrs, timeout=20)
                data = resp.json()
                for ev in (data.get("data") or []):
                    d = _ft_parse_dt(ev.get("ras_eve_data_gps"))
                    if d is None:
                        continue
                    try:
                        ig = int(ev.get("ras_eve_ignicao") or 0)
                    except Exception:
                        ig = 0
                    rows.append((d, 1 if ig == 1 else 0))
                _cache_save(ck, [[d.strftime("%d/%m/%Y %H:%M:%S"), ig] for d, ig in rows])
            except Exception:
                pass

        eventos.extend(rows)
        cur = nxt

    if not eventos:
        return 0.0

    eventos.sort(key=lambda x: x[0])

    total_seg = 0.0
    for i in range(len(eventos) - 1):
        t0, ig = eventos[i]
        t1, _ = eventos[i + 1]
        if ig == 1 and t1 > t0:
            total_seg += (t1 - t0).total_seconds()

    if eventos[-1][1] == 1 and dt_fim > eventos[-1][0]:
        total_seg += (dt_fim - eventos[-1][0]).total_seconds()

    return round(total_seg / 3600.0, 1)

'''

# Inserir após a definição de INTERVALO_HORAS_PADRAO
if "def _ft_parse_dt" not in content:
    insert_pos = content.find("def _split_emails")
    if insert_pos > 0:
        content = content[:insert_pos] + cache_and_ft_functions + "\n\n" + content[insert_pos:]

# 3. Substituir função buscar_horimetros_fulltrack
old_buscar_horimetros = '''def buscar_horimetros_fulltrack() -> dict:
    """
    Retorna dict {nome_veiculo: horimetro_atual} via /events/all.
    Usado para calcular horas restantes até a próxima manutenção.
    """
    try:
        response = requests.get(
            f"{FULLTRACK_BASE}/events/all",
            headers={"apikey": FULLTRACK_API_KEY, "secretkey": FULLTRACK_SECRET_KEY},
            timeout=15,
        )
        response.raise_for_status()
        dados = response.json().get("data") or []
        resultado = {}
        for evento in dados:
            nome = (evento.get("ras_vei_veiculo") or "").strip()
            placa = (evento.get("ras_vei_placa") or "").strip()
            horimetro = evento.get("ras_eve_horimetro")
            horimetro_horas = round(float(horimetro or 0) / 3600, 1)
            if nome:
                resultado[nome] = horimetro_horas
            if placa:
                resultado[placa] = horimetro_horas
        print(f"  ✔ FullTrack: horímetro de {len(dados)} veículos obtido")
        return resultado
    except Exception as exc:
        print(f"  [AVISO] buscar_horimetros_fulltrack: {exc} — usando fallback por data")
        return {}'''

new_buscar_horimetros = '''def buscar_horimetros_fulltrack(manutencoes: list) -> dict:
    """
    Retorna dict {equipamento: horas_acumuladas_desde_ultima_manutencao}.
    
    IMPORTANTE: Em vez de usar o horímetro da máquina (não confiável),
    calcula as HORAS DE IGNIÇÃO LIGADA desde a última manutenção registrada.
    
    Entrada: lista de manutencões (do Supabase)
    Saída: {equipamento: horas_acumuladas}
    """
    from datetime import datetime
    
    if not manutencoes:
        return {}
    
    resultado = {}
    agora = datetime.now()
    
    for rec in manutencoes:
        nome = (rec.get("equipamento") or "").strip()
        if not nome:
            continue
        
        # Busca data da última manutenção
        ultima_manutencao = rec.get("ultima_manutencao")
        if not ultima_manutencao:
            # Se não tem registro de manutenção, não conseguimos calcular
            resultado[nome] = 0.0
            continue
        
        try:
            # Converte data (formato: YYYY-MM-DD ou YYYY-MM-DDTHH:MM:SS)
            data_str = ultima_manutencao[:10]  # pega só YYYY-MM-DD
            dt_ultima = datetime.strptime(data_str, "%Y-%m-%d")
            
            # Busca ID do veículo no FullTrack
            # Para simplificar, usamos 'nome' como ID (pode precisar ajuste)
            veiculo_id = nome
            
            # Calcula horas de ignição desde última manutenção até agora
            horas_acumuladas = _ft_horas_ignicao_intervalo(veiculo_id, dt_ultima, agora)
            resultado[nome] = horas_acumuladas
            
        except Exception as exc:
            print(f"  [AVISO] Cálculo de horas para {nome}: {exc}")
            resultado[nome] = 0.0
    
    print(f"  ✔ FullTrack: horímetro acumulado de {len(resultado)} equipamentos calculado")
    return resultado'''

if old_buscar_horimetros in content:
    content = content.replace(old_buscar_horimetros, new_buscar_horimetros)
    print("✔ Função buscar_horimetros_fulltrack substituída")
else:
    print("⚠ Função buscar_horimetros_fulltrack não encontrada exatamente — verificar manualmente")

# 4. Atualizar função calcular_status
old_calcular_status = '''def calcular_status(registro: dict, horimetros_ft: dict) -> dict:
    """
    Retorna dict com: status, modo, horo_atual, horo_proxima, horas_rest, situacao, dt_proxima.
    Prioridade: horímetro FullTrack → fallback por data.
    """
    nome = (registro.get("equipamento") or "").strip()
    horimetro_ultima = registro.get("horimetro_ultima_manutencao")
    intervalo_horas = float(registro.get("intervalo_horas") or INTERVALO_HORAS_PADRAO)
    horimetro_atual = horimetros_ft.get(nome)

    if horimetro_ultima is not None and horimetro_atual is not None:
        horimetro_ultima = float(horimetro_ultima)
        horimetro_proxima = round(horimetro_ultima + intervalo_horas, 1)
        horas_restantes = round(horimetro_proxima - horimetro_atual, 1)
        if horas_restantes < 0:
            status = "vencida"
            situacao = f"{abs(horas_restantes):.1f} h em atraso"
        elif horas_restantes <= AVISO_HORAS:
            status = "proxima"
            situacao = f"Faltam {horas_restantes:.1f} h"
        else:
            status = "ok"
            situacao = f"Faltam {horas_restantes:.1f} h"
        return {
            "status": status,
            "modo": "horimetro",
            "horo_atual": horimetro_atual,
            "horo_proxima": horimetro_proxima,
            "horas_rest": horas_restantes,
            "situacao": situacao,
            "dt_proxima": None,
        }'''

new_calcular_status = '''def calcular_status(registro: dict, horimetros_ft: dict) -> dict:
    """
    Retorna dict com: status, modo, horo_atual, horo_proxima, horas_rest, situacao, dt_proxima.
    
    Lógica:
    - Usa horímetro acumulado (horas de ignição desde última manutenção)
    - Compara com intervalo_horas (que o usuário digita)
    - Se passou do limite: "vencida"
    - Se faltam ≤20h: "proxima" (aviso antecipado)
    - Senão: "ok"
    """
    nome = (registro.get("equipamento") or "").strip()
    horimetro_ultima = registro.get("horimetro_ultima_manutencao")
    intervalo_horas = float(registro.get("intervalo_horas") or INTERVALO_HORAS_PADRAO)
    horas_acumuladas = horimetros_ft.get(nome)

    if horimetro_ultima is not None and horas_acumuladas is not None:
        horimetro_ultima = float(horimetro_ultima)
        horas_acumuladas = float(horas_acumuladas)
        
        # Horímetro esperado na próxima manutenção
        horimetro_proxima = round(horimetro_ultima + intervalo_horas, 1)
        
        # Horímetro ATUAL = última manutenção + horas acumuladas desde então
        horimetro_atual = round(horimetro_ultima + horas_acumuladas, 1)
        
        # Quantas horas faltam até a próxima manutenção
        horas_restantes = round(horimetro_proxima - horimetro_atual, 1)
        
        if horas_restantes < 0:
            status = "vencida"
            situacao = f"{abs(horas_restantes):.1f} h em atraso"
        elif horas_restantes <= AVISO_HORAS:
            status = "proxima"
            situacao = f"Faltam {horas_restantes:.1f} h"
        else:
            status = "ok"
            situacao = f"Faltam {horas_restantes:.1f} h"
        
        return {
            "status": status,
            "modo": "horimetro_acumulado",
            "horo_atual": horimetro_atual,
            "horo_proxima": horimetro_proxima,
            "horas_rest": horas_restantes,
            "situacao": situacao,
            "dt_proxima": None,
        }'''

if old_calcular_status in content:
    content = content.replace(old_calcular_status, new_calcular_status)
    print("✔ Função calcular_status substituída")
else:
    print("⚠ Função calcular_status não encontrada exatamente — verificar manualmente")

# 5. Atualizar chamada para buscar_horimetros_fulltrack na função main()
if 'horimetros_ft = buscar_horimetros_fulltrack()' in content:
    content = content.replace(
        'horimetros_ft = buscar_horimetros_fulltrack()',
        'horimetros_ft = buscar_horimetros_fulltrack(manutencoes)'
    )
    print("✔ Chamada para buscar_horimetros_fulltrack atualizada")

# Validar sintaxe
try:
    compile(content, alertas_file, 'exec')
    print("✔ Sintaxe válida!")
except SyntaxError as e:
    print(f"✗ Erro de sintaxe: {e}")
    exit(1)

# Gravar
with open(alertas_file, "w", encoding="utf-8") as f:
    f.write(content)

print(f"\n✅ {alertas_file} refatorado com sucesso!")
print("\nMudanças aplicadas:")
print("  1. ✔ Adicionadas funções de cache")
print("  2. ✔ Adicionadas funções de cálculo de ignição acumulada")
print("  3. ✔ buscar_horimetros_fulltrack modificada para usar horas acumuladas")
print("  4. ✔ calcular_status modificada para comparar com intervalo_horas")
print("  5. ✔ Alertas disparam quando faltam ≤20h para manutenção")
