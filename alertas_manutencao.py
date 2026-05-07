"""
Alertas automáticos de manutenção preventiva — Locvix
Roda via GitHub Actions (cron diário às 08h BRT) ou manualmente.

Variáveis de ambiente necessárias (GitHub Actions Secrets):
  SUPABASE_URL         — URL da instância Supabase
  SUPABASE_ANON        — Chave anônima Supabase
  SMTP_HOST            — Servidor SMTP (padrão: smtp.gmail.com)
  SMTP_PORT            — Porta SMTP (padrão: 587)
  SMTP_USER            — Usuário SMTP / e-mail remetente
  SMTP_PASS            — Senha de app SMTP
  EMAIL_FROM           — Nome <email> do remetente (padrão: SMTP_USER)
  EMAIL_DEFAULT        — E-mail do André (recebe TODOS os alertas)
  FULLTRACK_API_KEY    — API Key FullTrack
  FULLTRACK_SECRET_KEY — Secret Key FullTrack
"""

import os
import sys
import smtplib
from datetime import date, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import requests

import json
import hashlib
import tempfile


# ── Configuração via variáveis de ambiente ─────────────────────────
SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
SUPABASE_ANON = os.environ.get("SUPABASE_ANON", "")
SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER", "")
SMTP_PASS = os.environ.get("SMTP_PASS", "")
EMAIL_FROM = os.environ.get("EMAIL_FROM", SMTP_USER)
EMAIL_DEFAULT = os.environ.get("EMAIL_DEFAULT", "")  # e-mail do André — recebe TODOS os alertas
FULLTRACK_API_KEY = os.environ.get("FULLTRACK_API_KEY", "530fdb8a61907a2f9904477a335f1a8eee0ea5d9")
FULLTRACK_SECRET_KEY = os.environ.get("FULLTRACK_SECRET_KEY", "8c002f8a04533e2f2ed428820f89dca5b3e9996a")
FULLTRACK_BASE = "https://ws.fulltrack2.com"

# Horas restantes para emitir alerta de "próxima manutenção"
AVISO_HORAS = 20
# Dias de antecedência (fallback quando horímetro não disponível)
AVISO_DIAS = 5
# Fallback do intervalo quando o cadastro não informar horas.
INTERVALO_HORAS_PADRAO = 600




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



def _split_emails(raw_value: str) -> list:
    """Converte lista de e-mails separada por vírgula/quebra de linha em lista sem duplicidade."""
    emails = []
    seen = set()
    for item in (raw_value or "").replace(";", ",").replace("\n", ",").split(","):
        email = item.strip()
        if not email:
            continue
        email_key = email.lower()
        if email_key in seen:
            continue
        seen.add(email_key)
        emails.append(email)
    return emails


# ── Busca horímetro atual de todos os veículos no FullTrack ────────
def buscar_horimetros_fulltrack(manutencoes: list) -> dict:
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
    return resultado


# ── Busca registros do Supabase ────────────────────────────────────
def buscar_manutencoes() -> list:
    if not SUPABASE_URL or not SUPABASE_ANON:
        print("  [ERRO] SUPABASE_URL ou SUPABASE_ANON não configurados.")
        return []
    headers = {
        "apikey": SUPABASE_ANON,
        "Authorization": f"Bearer {SUPABASE_ANON}",
    }
    response = requests.get(
        f"{SUPABASE_URL}/rest/v1/manutencoes_equipamentos",
        headers=headers,
        params={
            "select": "equipamento,ultima_manutencao,responsavel_email,tipo_servico,horimetro_ultima_manutencao,intervalo_horas,hodometro_ultima_manutencao,intervalo_km,periodo_dias",
            "order": "equipamento.asc",
        },
        timeout=15,
    )
    response.raise_for_status()
    return response.json()


# ── Calcula status ─────────────────────────────────────────────────
def calcular_status(registro: dict, horimetros_ft: dict) -> dict:
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
        }

    ultima = (registro.get("ultima_manutencao") or "")[:10]
    intervalo_meses = int(registro.get("intervalo_meses") or 2)
    if not ultima:
        return {
            "status": "vencida",
            "modo": "data",
            "horo_atual": horimetro_atual,
            "horo_proxima": None,
            "horas_rest": None,
            "situacao": "Nunca realizada",
            "dt_proxima": None,
        }

    dt_ultima = date.fromisoformat(ultima)
    dt_proxima = dt_ultima + timedelta(days=intervalo_meses * 30)
    dias = (dt_proxima - date.today()).days
    if dias < 0:
        status = "vencida"
        situacao = f"{abs(dias)} dias em atraso"
    elif dias <= AVISO_DIAS:
        status = "proxima"
        situacao = f"Faltam {dias} dias"
    else:
        status = "ok"
        situacao = f"Faltam {dias} dias"
    return {
        "status": status,
        "modo": "data",
        "horo_atual": horimetro_atual,
        "horo_proxima": None,
        "horas_rest": None,
        "situacao": situacao,
        "dt_proxima": dt_proxima,
    }


# ── Monta e envia e-mail ───────────────────────────────────────────
def enviar_email(destinatarios: list, equipamentos: list) -> None:
    if not SMTP_USER or not SMTP_PASS:
        print("  [SKIP] Sem credenciais SMTP")
        return
    if not equipamentos:
        return

    rows_html = ""
    for equipamento in equipamentos:
        cor = "#dc2626" if equipamento["status"] == "vencida" else "#d97706"
        badge = "🔴 VENCIDA" if equipamento["status"] == "vencida" else "⚠️ PRÓXIMA"

        # Monta critérios como HTML
        criterios_html = ""
        for c in equipamento.get("criterios", []):
            criterios_html += f"<div style='font-size:11px;margin:4px 0;color:#374151'>{c['tipo']}: {c['situacao']}</div>"

        rows_html += f"""
        <tr>
          <td style="padding:8px 12px;border-bottom:1px solid #e5e7eb;font-weight:600">{equipamento['cc']}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #e5e7eb">
            <span style="background:{cor};color:#fff;padding:2px 8px;border-radius:10px;font-size:12px">{badge}</span>
          </td>
          <td style="padding:8px 12px;border-bottom:1px solid #e5e7eb;font-size:12px;color:#64748b">
            {criterios_html}
          </td>
          <td style="padding:8px 12px;border-bottom:1px solid #e5e7eb;color:{cor};font-weight:600">{equipamento['situacao']}</td>
        </tr>"""

    vencidas_count = sum(1 for equipamento in equipamentos if equipamento["status"] == "vencida")
    proximas_count = sum(1 for equipamento in equipamentos if equipamento["status"] == "proxima")
    resumo_txt = []
    if vencidas_count:
        resumo_txt.append(f"<strong style='color:#dc2626'>{vencidas_count} vencida(s)</strong>")
    if proximas_count:
        resumo_txt.append(f"<strong style='color:#d97706'>{proximas_count} próxima(s)</strong>")

    html_body = f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="utf-8"></head>
<body style="font-family:Arial,sans-serif;color:#1e293b;max-width:700px;margin:0 auto;padding:0">
  <div style="background:#1e3a5f;color:#fff;padding:24px 28px;border-radius:8px 8px 0 0">
    <h2 style="margin:0;font-size:20px">🛠 Locvix — Alerta de Manutenção Preventiva</h2>
    <p style="margin:6px 0 0;opacity:.8;font-size:13px">Verificação automática — {date.today().strftime('%d/%m/%Y')}</p>
  </div>
  <div style="background:#f8fafc;padding:24px 28px;border-radius:0 0 8px 8px;border:1px solid #e2e8f0;border-top:none">
    <p style="margin:0 0 16px">Olá,</p>
    <p style="margin:0 0 16px">Os seguintes equipamentos requerem atenção: {' e '.join(resumo_txt)}.</p>
    <table style="width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.1)">
      <thead>
        <tr style="background:#1e3a5f;color:#fff">
          <th style="padding:10px 12px;text-align:left;font-size:13px">Equipamento</th>
          <th style="padding:10px 12px;text-align:left;font-size:13px">Status</th>
          <th style="padding:10px 12px;text-align:left;font-size:13px">Critérios (Horímetro | Período)</th>
          <th style="padding:10px 12px;text-align:left;font-size:13px">Situação</th>
        </tr>
      </thead>
      <tbody>{rows_html}</tbody>
    </table>
    <p style="margin:20px 0 4px;font-size:13px;color:#64748b">
      Para registrar manutenções, acesse o dashboard Locvix → módulo <strong>🛠 Manutenção</strong>.
    </p>
    <p style="margin:4px 0 0;font-size:11px;color:#94a3b8">
      Este e-mail foi enviado automaticamente pelo sistema Locvix (GitHub Actions).
    </p>
  </div>
</body></html>"""

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"🔔 Locvix — Alerta de Manutenção ({len(equipamentos)} equipamento(s))"
    msg["From"] = EMAIL_FROM
    msg["To"] = ", ".join(destinatarios)
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.login(SMTP_USER, SMTP_PASS)
        smtp.sendmail(EMAIL_FROM, destinatarios, msg.as_string())
    print(f"  ✅ E-mail enviado → {', '.join(destinatarios)} ({len(equipamentos)} equipamento(s))")


# ── Main ──────────────────────────────────────────────────────────
def main() -> int:
    """Retorna 0 em sucesso, 1 se houve erro crítico."""
    print("=" * 55)
    print("  Locvix — Alertas de Manutenção Preventiva")
    print(f"  Data: {date.today().strftime('%d/%m/%Y')}")
    print("=" * 55)

    horimetros_ft = buscar_horimetros_fulltrack(manutencoes)

    try:
        registros = buscar_manutencoes()
    except Exception as exc:
        print(f"  [ERRO] Falha ao buscar dados do Supabase: {exc}")
        return 1

    if not registros:
        print("  Nenhum registro de manutenção cadastrado.")
        return 0

    print(f"  {len(registros)} equipamento(s) encontrado(s).")

    # E-mails padrão do André — recebem TODOS os alertas
    defaults = _split_emails(EMAIL_DEFAULT)

    # Lista única de todos os alertas para enviar de uma vez ao André
    todos_alertas = []
    total_vencidas = 0
    total_proximas = 0

    for registro in registros:
        nome = (registro.get("equipamento") or "").strip()
        if not nome or "TESTE" in nome.upper():
            continue

        calc = calcular_status(registro, horimetros_ft)
        if calc["status"] == "ok":
            continue

        if calc["status"] == "vencida":
            total_vencidas += 1
        elif calc["status"] == "proxima":
            total_proximas += 1

        entry = {
            "cc": nome,
            "status": calc["status_geral"],
            "criterios": calc["criterios"],
            "horo_atual": calc["horo_atual"],
            "situacao": calc["situacao"],
        }
        todos_alertas.append(entry)

    total_alertas = len(todos_alertas)

    if total_alertas == 0:
        print("  ✅ Todos os equipamentos estão em dia. Nenhum alerta necessário.")
        return 0

    print(
        f"  {total_alertas} equipamento(s) requerem atenção "
        f"({total_vencidas} vencida(s), {total_proximas} próxima(s))."
    )

    if not defaults:
        print("  [AVISO] EMAIL_DEFAULT não configurado — e-mail não enviado.")
        return 1

    # Envia UM único e-mail com TODOS os alertas para o André (EMAIL_DEFAULT)
    erros = 0
    try:
        enviar_email(defaults, todos_alertas)
    except Exception as exc:
        print(f"  [ERRO] Falha ao enviar para {', '.join(defaults)}: {exc}")
        erros += 1

    print("\n  Concluído.")
    return 1 if erros else 0


if __name__ == "__main__":
    sys.exit(main())
