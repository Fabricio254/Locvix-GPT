"""
Alertas automáticos de manutenção preventiva — Locvix
Roda via GitHub Actions (cron diário às 08h BRT) ou manualmente.

Variáveis de ambiente necessárias (GitHub Actions Secrets):
  SUPABASE_URL   — URL da instância Supabase
  SUPABASE_ANON  — Chave anônima Supabase
  SMTP_HOST      — Servidor SMTP (padrão: smtp.gmail.com)
  SMTP_PORT      — Porta SMTP (padrão: 587)
  SMTP_USER      — Usuário SMTP / e-mail remetente
  SMTP_PASS      — Senha de app SMTP (Gmail: senha de app, não a senha da conta)
  EMAIL_FROM     — Nome <email> do remetente (padrão: SMTP_USER)
  EMAIL_DEFAULT  — E-mail padrão para equipamentos sem responsável definido (opcional)
"""

import os
import sys
import smtplib
import requests
from datetime import date, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ── Configuração via variáveis de ambiente ─────────────────────────
SUPABASE_URL    = os.environ.get("SUPABASE_URL", "")
SUPABASE_ANON   = os.environ.get("SUPABASE_ANON", "")
SMTP_HOST       = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT       = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER       = os.environ.get("SMTP_USER", "")
SMTP_PASS       = os.environ.get("SMTP_PASS", "")
EMAIL_FROM      = os.environ.get("EMAIL_FROM", SMTP_USER)
EMAIL_DEFAULT   = os.environ.get("EMAIL_DEFAULT", "")  # fallback para equipamentos sem e-mail

# Dias de antecedência para emitir alerta de "próxima"
AVISO_DIAS = 5


# ── Busca registros ────────────────────────────────────────────────
def buscar_manutencoes() -> list[dict]:
    if not SUPABASE_URL or not SUPABASE_ANON:
        print("  [ERRO] SUPABASE_URL ou SUPABASE_ANON não configurados.")
        return []
    hdrs = {
        "apikey":        SUPABASE_ANON,
        "Authorization": f"Bearer {SUPABASE_ANON}",
    }
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/manutencoes_equipamentos",
        headers=hdrs,
        params={
            "select": "equipamento,ultima_manutencao,responsavel_email,intervalo_meses",
            "order":  "equipamento.asc",
        },
        timeout=15,
    )
    r.raise_for_status()
    return r.json()


# ── Calcula status ─────────────────────────────────────────────────
def calcular_status(rec: dict):
    """Retorna (status, dt_proxima, dias_restantes)."""
    ultima = (rec.get("ultima_manutencao") or "")[:10]
    intervalo = int(rec.get("intervalo_meses") or 2)
    if not ultima:
        return "vencida", None, -9999
    dt_ultima  = date.fromisoformat(ultima)
    dt_proxima = dt_ultima + timedelta(days=intervalo * 30)
    dias       = (dt_proxima - date.today()).days
    status     = "vencida" if dias < 0 else ("proxima" if dias <= AVISO_DIAS else "ok")
    return status, dt_proxima, dias


# ── Monta e envia e-mail ───────────────────────────────────────────
def enviar_email(destinatario: str, equipamentos: list[dict]) -> None:
    if not SMTP_USER or not SMTP_PASS:
        print(f"  [SKIP] Sem credenciais SMTP — pulando {destinatario}")
        return
    if not equipamentos:
        return

    rows_html = ""
    for e in equipamentos:
        if e["status"] == "vencida":
            cor   = "#dc2626"
            badge = "🔴 VENCIDA"
            dias_txt = (
                f"{abs(e['dias'])} dias em atraso"
                if e["dias"] != -9999
                else "Nunca realizada"
            )
        else:
            cor   = "#d97706"
            badge = "⚠️ PRÓXIMA"
            dias_txt = f"{e['dias']} dias restantes"

        ultima_fmt  = e["ultima"].replace("-", "/")[:10][::-1].replace("/", "/")  # dd/mm/yyyy
        proxima_fmt = (
            e["proxima"].strftime("%d/%m/%Y") if e.get("proxima") else "—"
        )
        # Corrigir formato da data
        if ultima_fmt and len(ultima_fmt) == 8 and "-" not in (e["ultima"] or ""):
            # já veio como YYYY-MM-DD, reformatar para DD/MM/YYYY
            partes = (e["ultima"] or "")[:10].split("-")
            ultima_fmt = f"{partes[2]}/{partes[1]}/{partes[0]}" if len(partes) == 3 else "—"
        else:
            ultima_fmt = "Não registrada"

        rows_html += f"""
        <tr>
          <td style="padding:8px 12px;border-bottom:1px solid #e5e7eb;font-weight:600">{e['cc']}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #e5e7eb">
            <span style="background:{cor};color:#fff;padding:2px 8px;border-radius:10px;font-size:12px">{badge}</span>
          </td>
          <td style="padding:8px 12px;border-bottom:1px solid #e5e7eb">{ultima_fmt}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #e5e7eb">{proxima_fmt}</td>
          <td style="padding:8px 12px;border-bottom:1px solid #e5e7eb;color:{cor};font-weight:600">{dias_txt}</td>
        </tr>"""

    vencidas_count = sum(1 for e in equipamentos if e["status"] == "vencida")
    proximas_count = sum(1 for e in equipamentos if e["status"] == "proxima")
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
          <th style="padding:10px 12px;text-align:left;font-size:13px">Última Manutenção</th>
          <th style="padding:10px 12px;text-align:left;font-size:13px">Próxima Manutenção</th>
          <th style="padding:10px 12px;text-align:left;font-size:13px">Situação</th>
        </tr>
      </thead>
      <tbody>{rows_html}</tbody>
    </table>
    <p style="margin:20px 0 4px;font-size:13px;color:#64748b">
      Para registrar manutenções, acesse:
      <a href="https://locvix.streamlit.app" style="color:#1e3a5f;font-weight:600">locvix.streamlit.app</a>
      → módulo <strong>🛠 Manutenção</strong> → expandir <em>"Registrar / Atualizar Manutenção"</em>.
    </p>
    <p style="margin:4px 0 0;font-size:11px;color:#94a3b8">
      Este e-mail foi enviado automaticamente pelo sistema Locvix (GitHub Actions).
    </p>
  </div>
</body></html>"""

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"🔔 Locvix — Alerta de Manutenção ({len(equipamentos)} equipamento(s))"
    msg["From"]    = EMAIL_FROM
    msg["To"]      = destinatario if isinstance(destinatario, str) else ", ".join(destinatario)
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    rcpt_list = [destinatario] if isinstance(destinatario, str) else list(destinatario)
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.login(SMTP_USER, SMTP_PASS)
        smtp.sendmail(EMAIL_FROM, rcpt_list, msg.as_string())
    print(f"  ✅ E-mail enviado → {', '.join(rcpt_list)} ({len(equipamentos)} equipamento(s))")


# ── Main ──────────────────────────────────────────────────────────
def main() -> int:
    """Retorna 0 em sucesso, 1 se houve erro crítico."""
    print("=" * 55)
    print("  Locvix — Alertas de Manutenção Preventiva")
    print(f"  Data: {date.today().strftime('%d/%m/%Y')}")
    print("=" * 55)

    try:
        registros = buscar_manutencoes()
    except Exception as e:
        print(f"  [ERRO] Falha ao buscar dados do Supabase: {e}")
        return 1

    if not registros:
        print("  Nenhum registro de manutenção cadastrado.")
        return 0

    print(f"  {len(registros)} equipamento(s) encontrado(s).")

    # Agrupa alertas por e-mail responsável
    por_email: dict[str, list] = {}
    sem_email: list = []

    for rec in registros:
        status, dt_proxima, dias = calcular_status(rec)
        if status == "ok":
            continue   # em dia, sem alerta

        entry = {
            "cc":     rec["equipamento"],
            "status": status,
            "ultima": (rec.get("ultima_manutencao") or "")[:10],
            "proxima": dt_proxima,
            "dias":   dias,
        }
        email = (rec.get("responsavel_email") or "").strip()
        # Sempre envia para os e-mails padrão
        defaults = [e.strip() for e in EMAIL_DEFAULT.replace("\n", ",").split(",") if e.strip()]
        destinos = set(defaults)
        if email:
            destinos.add(email)
        for dest in destinos:
            por_email.setdefault(dest, []).append(entry)
        if not destinos:
            sem_email.append(entry)

    total_alertas = sum(len(v) for v in por_email.values()) + len(sem_email)
    if total_alertas == 0:
        print("  ✅ Todos os equipamentos estão em dia. Nenhum alerta necessário.")
        return 0

    # Consolida: todos os equipamentos vão num único e-mail para todos os destinos padrão
    defaults = [e.strip() for e in EMAIL_DEFAULT.replace("\n", ",").split(",") if e.strip()]
    todos_equipamentos = []
    seen = set()
    for eqs in por_email.values():
        for eq in eqs:
            if eq["cc"] not in seen:
                seen.add(eq["cc"])
                todos_equipamentos.append(eq)

    print(f"  {len(todos_equipamentos)} equipamento(s) requerem atenção.")
    print(f"  Destinatários: {', '.join(defaults) if defaults else '(nenhum)'}")

    erros = 0
    if defaults and todos_equipamentos:
        try:
            enviar_email(defaults, todos_equipamentos)
        except Exception as e:
            print(f"  [ERRO] Falha ao enviar: {e}")
            erros += 1

    if sem_email:
        print(f"  [AVISO] {len(sem_email)} equipamento(s) sem e-mail responsável definido.")

    print("\n  Concluído.")
    return 1 if erros else 0


if __name__ == "__main__":
    sys.exit(main())
