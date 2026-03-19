"""
App principal — Dashboard LOCVIX
ERP: GestãoClick via API REST
Executa via: streamlit run app.py
"""
import os
import sys
import io
import queue
import tempfile
import threading
import contextlib
import traceback as _tb
import importlib
from datetime import datetime, date, timezone, timedelta

_BRT = timezone(timedelta(hours=-3))

import base64
import streamlit as st
import streamlit.components.v1 as components

# ─── Configuração da página ────────────────────────────────────────
st.set_page_config(
    page_title="Dashboard Locvix",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS global ───────────────────────────────────────────────────
st.markdown("""
<style>
.block-container{padding-top:.6rem!important;padding-bottom:0!important;}
iframe{border:none!important;}
div[data-testid="stSidebar"]{background:#0d1b2a;}
div[data-testid="stSidebar"] *{color:#e0e8f0!important;}
div[data-testid="stSidebar"] hr{border-color:#1e3550!important;}
div[data-testid="stSidebar"] .stButton button{
    background:#1e40af;color:#fff;border:none;border-radius:6px;
}
div[data-testid="stSidebar"] .stButton button:hover{background:#2563eb;}
/* Oculta toolbar superior (Share, Edit, Deploy, Github) */
[data-testid="stToolbar"]{display:none!important;}
[data-testid="stDecoration"]{display:none!important;}
[data-testid="stHeader"]{display:none!important;}
#MainMenu{display:none!important;}
footer{display:none!important;}
</style>
""", unsafe_allow_html=True)

# ─── Utilitário de logo ───────────────────────────────────────────
_BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
_LOGO_EXTS = ["logo locvix.jfif", "logo_locvix.png", "logo_locvix.jpg", "logo_locvix.jfif"]

def _logo_path():
    for nome in _LOGO_EXTS:
        p = os.path.join(_BASE_DIR, nome)
        if os.path.exists(p):
            return p
    return None

# ─── Tela de login ────────────────────────────────────────────────
_SENHA_CORRETA = "zampa254"

if not st.session_state.get("_autenticado"):
    _lp = _logo_path()
    col_l, col_c, col_r = st.columns([1, 1, 1])
    with col_c:
        st.markdown("<div style='margin-top:60px'></div>", unsafe_allow_html=True)
        if _lp:
            with open(_lp, "rb") as _f:
                _b64 = base64.b64encode(_f.read()).decode()
            _ext = _lp.rsplit(".", 1)[-1].lower().replace("jfif", "jpeg")
            st.markdown(
                f"<div style='text-align:center'><img src='data:image/{_ext};base64,{_b64}' width='180' style='border-radius:12px'/></div>",
                unsafe_allow_html=True,
            )
        st.markdown(
            "<h2 style='text-align:center;margin-top:16px'>🔒 Área Restrita — LOCVIX</h2>",
            unsafe_allow_html=True,
        )
        with st.form("_login_form"):
            senha = st.text_input("Senha de acesso", type="password")
            submitted = st.form_submit_button("Entrar", use_container_width=True, type="primary")
        if submitted:
            if senha == _SENHA_CORRETA:
                st.session_state["_autenticado"] = True
                st.rerun()
            else:
                st.error("Senha incorreta.")
    st.stop()

# ─── Injeta credenciais GestãoClick via st.secrets ────────────────
# Quando rodando no Streamlit Cloud, usa .streamlit/secrets.toml.
# Localmente usa as variáveis de ambiente (ou placeholder no locvix.py).
try:
    if "GCK_ACCESS_TOKEN" in st.secrets:
        os.environ["GCK_ACCESS_TOKEN"] = st.secrets["GCK_ACCESS_TOKEN"]
    if "GCK_SECRET_TOKEN" in st.secrets:
        os.environ["GCK_SECRET_TOKEN"] = st.secrets["GCK_SECRET_TOKEN"]
except Exception:
    pass

# ─── Sidebar ──────────────────────────────────────────────────────
with st.sidebar:
    _lp = _logo_path()
    if _lp:
        _cl, _cr = st.columns([1, 1])
        with _cl:
            st.image(_lp, width=150)
    st.markdown("## 📊 Dashboard LOCVIX")
    st.markdown("---")
    st.markdown("##### 🌐 ERP GestãoClick")
    st.markdown(
        "<small>Dados obtidos em tempo real via API GestãoClick.</small>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    # ── Período ───────────────────────────────────────────────────
    st.markdown("### 📅 Período")
    _hoje    = date.today()
    _ini_def = date(_hoje.year, 1, 1)

    col_d1, col_d2 = st.columns(2)
    with col_d1:
        data_ini_sel = st.date_input(
            "Início",
            value=_ini_def,
            min_value=date(2020, 1, 1),
            max_value=_hoje,
            format="DD/MM/YYYY",
            key="periodo_ini",
        )
    with col_d2:
        data_fim_sel = st.date_input(
            "Fim",
            value=_hoje,
            min_value=date(2020, 1, 1),
            max_value=_hoje,
            format="DD/MM/YYYY",
            key="periodo_fim",
        )

    st.markdown("---")

    # ── Módulos ───────────────────────────────────────────────────
    st.markdown("##### 📋 Módulos")
    modulos = {
        "geral":       "📊 Visão Geral",
        "vendas":      "💰 Vendas",
        "financeiro":  "💳 Financeiro",
        "operacoes":   "🔧 Operações",
        "ponto":       "🕐 Ponto Colaborador",
    }
    modulo_sel = st.radio(
        "Selecionar módulo",
        options=list(modulos.keys()),
        format_func=lambda k: modulos[k],
        label_visibility="collapsed",
        key="modulo_ativo",
    )

    st.markdown("---")

    # ── Fonte das Vendas ──────────────────────────────────────────
    st.markdown("##### 📂 Fonte das Vendas")
    fonte_vendas_sel = st.radio(
        "Fonte dos dados de vendas",
        options=["api", "excel"],
        format_func=lambda k: "API GestaoClick" if k == "api" else "Excel Manual",
        index=1,
        label_visibility="collapsed",
        key="fonte_vendas",
    )
    if fonte_vendas_sel == "excel":
        st.info("Planilhas: CONTROLE CONTAS A RECEBER 2025 e 2026")

    st.markdown("---")
    st.caption(f"Última execução: {datetime.now(_BRT).strftime('%d/%m/%Y %H:%M')}")

# ═══════════════════════════════════════════════════════════════════
#  PÁGINA PRINCIPAL — Dashboard Locvix
# ═══════════════════════════════════════════════════════════════════
st.title("📊 Dashboard — LOCVIX")
st.caption("Análise de Vendas · Financeiro · Clientes · OS · Contratos via GestãoClick ERP")

HTML_KEY   = "locvix_html"
STATUS_KEY = "locvix_status"
TIME_KEY   = "locvix_time"
PERIOD_KEY = "locvix_period"

_data_ini_str = data_ini_sel.strftime("%d/%m/%Y")
_data_fim_str = data_fim_sel.strftime("%d/%m/%Y")
_period_id    = f"{_data_ini_str}_{_data_fim_str}"

# Invalida cache se período mudou
if st.session_state.get(PERIOD_KEY) != _period_id:
    st.session_state.pop(HTML_KEY, None)
    st.session_state.pop(STATUS_KEY, None)

# ── Barra de ação ─────────────────────────────────────────────────
col1, col2 = st.columns([3, 1])
with col2:
    btn_atualizar = st.button(
        "🔄 Atualizar Dados",
        use_container_width=True,
        type="primary",
    )

# ── Executa coleta ────────────────────────────────────────────────
if btn_atualizar or HTML_KEY not in st.session_state:
    _prog_bar    = st.progress(0, text="⏳ Iniciando...")
    _prog_status = st.empty()
    log_buf      = io.StringIO()

    # Garante que o diretório do módulo está no path
    if _BASE_DIR not in sys.path:
        sys.path.insert(0, _BASE_DIR)

    # Carrega / recarrega o módulo locvix
    if "locvix" in sys.modules:
        lv = sys.modules["locvix"]
        # Ao pressionar Atualizar, força fetch fresco (limpa cache em memória)
        if btn_atualizar:
            if hasattr(lv, "_client"):
                lv._client = None
        importlib.reload(lv)
    else:
        import locvix as lv

    # Injeta credenciais nas variáveis globais do módulo
    lv.GCK_ACCESS_TOKEN = os.getenv("GCK_ACCESS_TOKEN", lv.GCK_ACCESS_TOKEN)
    lv.GCK_SECRET_TOKEN = os.getenv("GCK_SECRET_TOKEN", lv.GCK_SECRET_TOKEN)

    # Se clicou Atualizar, ignora todo cache em disco (flag no módulo)
    lv._SKIP_CACHE       = btn_atualizar
    lv._SKIP_PONTO_CACHE = btn_atualizar

    # Fila de progresso: (float, str) ou None (sentinela)
    _q: queue.Queue = queue.Queue()

    def _cb(pct: float, msg: str):
        _q.put((float(pct), str(msg)))

    lv._progresso = _cb

    # Arquivo temporário para o HTML (por compatibilidade)
    with tempfile.NamedTemporaryFile(
        suffix=".html", delete=False, mode="w", encoding="utf-8"
    ) as tmp:
        tmp_path = tmp.name

    _result: list = [None, None]   # [html_str | None, exc | None]

    _fonte_vendas = fonte_vendas_sel

    def _worker():
        try:
            with contextlib.redirect_stdout(log_buf):
                _result[0] = lv.main(
                    saida_html=tmp_path,
                    saida_excel=None,          # sem Excel no Streamlit
                    data_ini=_data_ini_str,
                    data_fim=_data_fim_str,
                    fonte_vendas=_fonte_vendas,
                )
        except Exception as exc:
            _result[1] = exc
        finally:
            lv._progresso = None
            _q.put(None)   # sentinela: fim da thread

    t = threading.Thread(target=_worker, daemon=True)
    t.start()

    import time as _time
    while True:
        try:
            item = _q.get(timeout=0.15)
        except queue.Empty:
            continue
        if item is None:
            break
        pct, msg = item
        _prog_bar.progress(min(pct, 1.0), text=f"⏳ {msg}")
        _prog_status.caption(msg)

    t.join()

    # Limpeza do temporário
    try:
        os.unlink(tmp_path)
    except Exception:
        pass

    if _result[1] is not None:
        st.session_state[STATUS_KEY] = "erro"
        st.error(f"❌ Erro ao gerar dashboard: {_result[1]}")
        with st.expander("🔍 Traceback completo", expanded=True):
            st.code(
                "".join(_tb.format_exception(type(_result[1]), _result[1],
                                             _result[1].__traceback__)),
                language=None,
            )
    elif _result[0]:
        st.session_state[HTML_KEY]   = _result[0]
        st.session_state[STATUS_KEY] = "ok"
        st.session_state[PERIOD_KEY] = _period_id
        st.session_state[TIME_KEY]   = datetime.now(_BRT).strftime("%d/%m/%Y às %H:%M:%S")
        _log_txt = log_buf.getvalue()
        if _log_txt.strip():
            st.session_state["locvix_log"] = _log_txt
        st.rerun()
    else:
        st.session_state[STATUS_KEY] = "erro"
        st.error("❌ Nenhum dado retornado. Verifique as credenciais GestãoClick e o período.")

    log_txt = log_buf.getvalue()
    if log_txt.strip():
        with st.expander("📋 Log de execução", expanded=False):
            st.code(log_txt, language=None)

# ── Exibe o dashboard HTML ────────────────────────────────────────
if HTML_KEY in st.session_state and st.session_state.get(STATUS_KEY) == "ok":
    with col1:
        st.caption(
            f"🕐 Gerado em: {st.session_state.get(TIME_KEY, '')}  |  "
            f"📅 Período: {st.session_state.get(PERIOD_KEY,'').replace('_',' a ')}"
        )
    with col2:
        _periodo_label = st.session_state.get(PERIOD_KEY, "dashboard").replace("_", "_a_").replace("/", "-")
        st.download_button(
            label="💾 Salvar HTML",
            data=st.session_state[HTML_KEY].encode("utf-8", errors="xmlcharrefreplace"),
            file_name=f"dashboard_locvix_{_periodo_label}.html",
            mime="text/html",
            use_container_width=True,
        )
    st.success("✅ Dashboard gerado com sucesso!")

    _html_safe = st.session_state[HTML_KEY].encode("utf-8", errors="replace").decode("utf-8")

    # Inject module-activation script based on sidebar selection
    _modulo_sel = st.session_state.get("modulo_ativo", "geral")
    _allowed    = {"geral", "vendas", "financeiro", "operacoes", "ponto"}
    if _modulo_sel not in _allowed:
        _modulo_sel = "geral"
    _inject = (
        f"<script>(function(){{function _am(){{if(typeof setModulo==='function')"
        f"setModulo('{_modulo_sel}');else setTimeout(_am,80);}}"
        f"if(document.readyState==='complete')_am();"
        f"else window.addEventListener('load',_am);}})()</script>"
    )
    _html_safe = _html_safe.replace("</body>", f"{_inject}</body>")

    components.html(
        _html_safe,
        height=980,
        scrolling=True,
    )

    _log_saved = st.session_state.get("locvix_log", "")
    if _log_saved.strip():
        with st.expander("📋 Log de execução", expanded=False):
            st.code(_log_saved, language=None)

elif HTML_KEY not in st.session_state:
    st.info(
        "Clique em **🔄 Atualizar Dados** para buscar os dados da API GestãoClick "
        "e gerar o dashboard completo."
    )
