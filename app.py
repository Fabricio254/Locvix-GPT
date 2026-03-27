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

_ALL_MODULES = ["geral", "vendas", "financeiro", "operacoes", "manutencao", "ponto", "orcamento"]


def _secrets_get(*keys, default=None):
    current = st.secrets
    try:
        for key in keys:
            current = current[key]
        return current
    except Exception:
        return default


def _load_users():
    users = {}
    raw_users = _secrets_get("users", default={}) or {}
    for username in raw_users:
        item = raw_users[username]
        modules = [m for m in list(item.get("modules", [])) if m in _ALL_MODULES]
        users[username.lower()] = {
            "name": item.get("name", username.title()),
            "password": str(item.get("password", "")),
            "modules": modules or list(_ALL_MODULES),
        }

    legacy_password = _secrets_get("auth", "password")
    legacy_username = _secrets_get("auth", "username")
    if legacy_password and legacy_username:
        users.setdefault(str(legacy_username).lower(), {
            "name": str(legacy_username),
            "password": str(legacy_password),
            "modules": list(_ALL_MODULES),
        })
    return users

# ─── Configuração da página ────────────────────────────────────────
st.set_page_config(
    page_title="Dashboard Locvix",
    page_icon="logo_locvix.png",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS global ───────────────────────────────────────────────────
st.markdown("""
<style>
.block-container{padding-top:.6rem!important;padding-bottom:0!important;}
iframe{border:none!important;}
/* Sidebar sempre visível — impede colapso */
div[data-testid="stSidebar"]{
    background:#0d1b2a;
    min-width:280px!important;
    max-width:320px!important;
    transform:none!important;
    visibility:visible!important;
    position:relative!important;
    display:flex!important;
}
div[data-testid="stSidebar"] *{color:#e0e8f0!important;}
div[data-testid="stSidebar"] hr{border-color:#1e3550!important;}
/* Esconde TODOS os botões de colapso sidebar */
button[data-testid="stSidebarCollapseButton"]{display:none!important;}
button[data-testid="baseButton-headerNoPadding"]{display:none!important;}
[data-testid="collapsedControl"]{display:none!important;}
/* Oculta toolbar superior (Share, Edit, Deploy, Github) */
[data-testid="stToolbar"]{display:none!important;}
[data-testid="stDecoration"]{display:none!important;}
[data-testid="stHeader"]{background:transparent!important;}
#MainMenu{display:none!important;}
footer{display:none!important;}
/* Botão verde para Novo Orçamento na sidebar */
.orc-btn-sidebar button{
    background:#16a34a!important;color:#fff!important;border:none!important;
    border-radius:8px!important;padding:.6rem 1rem!important;
    font-weight:700!important;font-size:1rem!important;
}
/* Botão Criar Orçamento na área principal */
.orc-btn-main button{
    background:#16a34a!important;color:#fff!important;border:none!important;
    border-radius:10px!important;padding:.75rem 1.5rem!important;
    font-weight:700!important;font-size:1.1rem!important;
    margin:.5rem 0!important;
}
/* Alinha os dois botões de ação do módulo Orçamento */
div[data-testid="stHorizontalBlock"] div[data-testid="stColumn"] button{
    height:3.2rem!important;
    font-size:1rem!important;
    font-weight:600!important;
    border-radius:8px!important;
}
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
_USERS = _load_users()

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
            usuario = st.text_input("Usuário")
            senha = st.text_input("Senha de acesso", type="password")
            submitted = st.form_submit_button("Entrar", use_container_width=True, type="primary")
        if submitted:
            user_cfg = _USERS.get(usuario.strip().lower())
            if user_cfg and senha == user_cfg["password"]:
                st.session_state["_autenticado"] = True
                st.session_state["_usuario_login"] = usuario.strip().lower()
                st.session_state["_usuario_nome"] = user_cfg["name"]
                st.session_state["_usuario_modulos"] = list(user_cfg["modules"])
                st.rerun()
            else:
                st.error("Usuário ou senha incorretos.")
    st.stop()

# ─── Injeta credenciais GestãoClick via st.secrets ────────────────
# Quando rodando no Streamlit Cloud, usa .streamlit/secrets.toml.
# Localmente usa as variáveis de ambiente (ou placeholder no locvix.py).
try:
    _gck_access = _secrets_get("GCK_ACCESS_TOKEN") or _secrets_get("credentials", "GCK_ACCESS_TOKEN")
    _gck_secret = _secrets_get("GCK_SECRET_TOKEN") or _secrets_get("credentials", "GCK_SECRET_TOKEN")
    _supabase_url = _secrets_get("SUPABASE_URL") or _secrets_get("credentials", "SUPABASE_URL")
    _supabase_anon = _secrets_get("SUPABASE_ANON") or _secrets_get("credentials", "SUPABASE_ANON")
    if _gck_access:
        os.environ["GCK_ACCESS_TOKEN"] = _gck_access
    if _gck_secret:
        os.environ["GCK_SECRET_TOKEN"] = _gck_secret
    if _supabase_url:
        os.environ["SUPABASE_URL"] = _supabase_url
    if _supabase_anon:
        os.environ["SUPABASE_ANON"] = _supabase_anon
except Exception:
    pass

_usuario_nome = st.session_state.get("_usuario_nome", "Usuário")
_usuario_modulos = [m for m in st.session_state.get("_usuario_modulos", list(_ALL_MODULES)) if m in _ALL_MODULES]
if not _usuario_modulos:
    _usuario_modulos = ["geral"]

# ─── Sidebar ──────────────────────────────────────────────────────
with st.sidebar:
    _lp = _logo_path()
    if _lp:
        _cl, _cr = st.columns([1, 1])
        with _cl:
            st.image(_lp, width=150)
    st.markdown("## 📊 Dashboard LOCVIX")
    st.caption(f"👤 {_usuario_nome}")
    if st.button("🚪 Sair", use_container_width=True):
        for _k in ["_autenticado", "_usuario_login", "_usuario_nome", "_usuario_modulos", "modulo_ativo"]:
            st.session_state.pop(_k, None)
        st.rerun()
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
    _ini_def = date(2025, 1, 1)
    _fim_def = date(2026, 12, 31)

    col_d1, col_d2 = st.columns(2)
    with col_d1:
        data_ini_sel = st.date_input(
            "Início",
            value=_ini_def,
            min_value=date(2020, 1, 1),
            max_value=date(2099, 12, 31),
            format="DD/MM/YYYY",
            key="periodo_ini",
        )
    with col_d2:
        data_fim_sel = st.date_input(
            "Fim",
            value=_fim_def,
            min_value=date(2020, 1, 1),
            max_value=date(2099, 12, 31),
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
        "manutencao":  "🛠 Manutenção",
        "ponto":       "🕐 Ponto Colaborador",
        "orcamento":   "📋 Orçamento",
    }
    _modulos_permitidos = [m for m in modulos if m in _usuario_modulos]
    if not _modulos_permitidos:
        _modulos_permitidos = ["geral"]
    if st.session_state.get("modulo_ativo") not in _modulos_permitidos:
        st.session_state["modulo_ativo"] = _modulos_permitidos[0]
    modulo_sel = st.radio(
        "Selecionar módulo",
        options=_modulos_permitidos,
        format_func=lambda k: modulos[k],
        label_visibility="collapsed",
        key="modulo_ativo",
    )

    if modulo_sel == "orcamento":
        st.markdown("")
        with st.container():
            st.markdown('<div class="orc-btn-sidebar">', unsafe_allow_html=True)
            if st.button("➕ Criar Novo Orçamento", use_container_width=True,
                         key="_btn_orc_sidebar"):
                st.session_state["_show_orc_form"] = True
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
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

# ── Força sidebar sempre aberta via JS no iframe (mesmo domínio) ──
components.html("""<script>
(function(){
    function forceOpen(){
        var p = window.parent.document;
        // Se o botão de expandir está visível, clica nele
        var btn = p.querySelector('[data-testid="collapsedControl"] button') ||
                  p.querySelector('[data-testid="collapsedControl"]');
        if(btn){ btn.click(); return; }
        // Sobrepõe transform inline da sidebar
        var sb = p.querySelector('[data-testid="stSidebar"]');
        if(sb){
            sb.style.setProperty('transform','none','important');
            sb.style.setProperty('min-width','280px','important');
            sb.style.setProperty('visibility','visible','important');
            // Esconde botão de colapsar
            var cb = p.querySelector('[data-testid="stSidebarCollapseButton"]');
            if(cb) cb.style.setProperty('display','none','important');
        }
    }
    setTimeout(forceOpen, 100);
    setTimeout(forceOpen, 600);
    setTimeout(forceOpen, 1500);
})();
</script>""", height=0, scrolling=False)

# ═══════════════════════════════════════════════════════════════════
#  PÁGINA PRINCIPAL — Dashboard Locvix
# ═══════════════════════════════════════════════════════════════════
st.title("📊 Dashboard — LOCVIX")
st.caption("Análise de Vendas · Financeiro · Clientes · OS · Contratos via GestãoClick ERP")

HTML_KEY   = "locvix_html_v9"   # bump this when JS/HTML changes break cached output
STATUS_KEY = "locvix_status"
TIME_KEY   = "locvix_time"
PERIOD_KEY = "locvix_period"

_data_ini_str = data_ini_sel.strftime("%d/%m/%Y")
_data_fim_str = data_fim_sel.strftime("%d/%m/%Y")
_period_id    = f"{_data_ini_str}_{_data_fim_str}"

# Compute a hash of locvix.py — if the source changed, invalidate the cache
import hashlib as _hashlib, pathlib as _pathlib
_src_hash = _hashlib.md5(
    _pathlib.Path(__file__).with_name("locvix.py").read_bytes()
).hexdigest()[:8]
_CODE_VER_KEY = "locvix_code_ver"
if st.session_state.get(_CODE_VER_KEY) != _src_hash:
    st.session_state.pop(HTML_KEY, None)
    st.session_state.pop(STATUS_KEY, None)
    st.session_state[_CODE_VER_KEY] = _src_hash

# Invalida cache se período mudou
if st.session_state.get(PERIOD_KEY) != _period_id:
    st.session_state.pop(HTML_KEY, None)
    st.session_state.pop(STATUS_KEY, None)

# Remove stale HTML from old key versions
for _old_key in ["locvix_html", "locvix_html_v2", "locvix_html_v3", "locvix_html_v4", "locvix_html_v5", "locvix_html_v6", "locvix_html_v7", "locvix_html_v8"]:
    st.session_state.pop(_old_key, None)

# ── Barra de ação ─────────────────────────────────────────────────
col1, col2, col3 = st.columns([2.5, 1, 1])
with col2:
    btn_atualizar = st.button(
        "🔄 Atualizar Dados",
        use_container_width=True,
        type="primary",
    )
with col3:
    btn_atualizar_orc = st.button(
        "📋 Atualizar Orçamentos",
        use_container_width=True,
        help="Atualiza somente os dados de orçamentos (mais rápido)",
        disabled=(HTML_KEY not in st.session_state),
    )

# Sempre carrega as duas lojas; seleção de loja fica dentro do HTML
_loja_filtro = "ambas"

# ── Atualização parcial — somente Orçamentos ─────────────────────
if btn_atualizar_orc and HTML_KEY in st.session_state:
    import re as _re, json as _json
    _prog_orc = st.progress(0, text="⏳ Buscando orçamentos...")
    try:
        if _BASE_DIR not in sys.path:
            sys.path.insert(0, _BASE_DIR)
        if "locvix" in sys.modules:
            lv = sys.modules["locvix"]
            importlib.reload(lv)
        else:
            import locvix as lv
        lv.GCK_ACCESS_TOKEN  = os.getenv("GCK_ACCESS_TOKEN", lv.GCK_ACCESS_TOKEN)
        lv.GCK_SECRET_TOKEN  = os.getenv("GCK_SECRET_TOKEN", lv.GCK_SECRET_TOKEN)
        lv.SUPABASE_URL      = os.getenv("SUPABASE_URL",     lv.SUPABASE_URL)
        lv.SUPABASE_ANON     = os.getenv("SUPABASE_ANON",    lv.SUPABASE_ANON)
        lv._SKIP_CACHE       = True   # ignora cache para forçar fetch fresco
        lv._empresa_cache    = {}     # limpa cache da empresa também

        _prog_orc.progress(0.2, text="⏳ Conectando ao ERP...")
        novos_orc = lv.buscar_orcamentos()

        _prog_orc.progress(0.85, text="⏳ Atualizando dashboard...")
        # Serializa igual ao jv() do locvix.py
        _orc_json = _json.dumps(novos_orc, ensure_ascii=True, default=str).replace("</" , r"<\/")
        # Substitui o bloco JSON de orçamentos no HTML armazenado
        _html_atual = st.session_state[HTML_KEY]
        _html_novo = _re.sub(
            r'(<script type="application/json" id="_dORCAMENTOS">).*?(</script>)',
            r'\g<1>' + _orc_json + r'\2',
            _html_atual,
            flags=_re.DOTALL,
        )
        st.session_state[HTML_KEY] = _html_novo
        st.session_state[TIME_KEY] = datetime.now(_BRT).strftime("%d/%m/%Y %H:%M:%S")
        _prog_orc.progress(1.0, text=f"✅ {len(novos_orc)} orçamentos atualizados!")
        st.success(f"✅ Orçamentos atualizados: {len(novos_orc)} registros.")
        st.rerun()
    except Exception as _e_orc:
        _prog_orc.empty()
        st.error(f"❌ Erro ao atualizar orçamentos: {_e_orc}")

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
    lv.SUPABASE_URL     = os.getenv("SUPABASE_URL",     lv.SUPABASE_URL)
    lv.SUPABASE_ANON    = os.getenv("SUPABASE_ANON",    lv.SUPABASE_ANON)

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
                    loja_filtro=_loja_filtro,
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

    # ── Ações do módulo Orçamento ──
    if st.session_state.get("modulo_ativo") == "orcamento" and not st.session_state.get("_show_orc_form") and not st.session_state.get("_show_del_orc"):
        _col_criar, _col_excluir = st.columns(2)
        with _col_criar:
            st.button("➕ Criar Novo Orçamento",
                      use_container_width=True, key="_btn_orc_main",
                      on_click=lambda: st.session_state.update({"_show_orc_form": True}))
        with _col_excluir:
            st.button("🗑️ Excluir Orçamento Existente",
                      use_container_width=True, key="_btn_del_orc",
                      on_click=lambda: st.session_state.update({"_show_del_orc": True}))

    # ── Painel Excluir Orçamento ──
    if st.session_state.get("modulo_ativo") == "orcamento" and st.session_state.get("_show_del_orc"):
        with st.container(border=True):
            _ct, _cf = st.columns([6, 1])
            with _ct:
                st.markdown("### 🗑️ Excluir Orçamento")
            with _cf:
                if st.button("❌ Fechar", key="_btn_fechar_del"):
                    st.session_state["_show_del_orc"] = False
                    st.session_state.pop("_del_confirmado", None)
                    st.session_state.pop("_del_orc_dados", None)
                    st.rerun()
            st.warning("⚠️ **Atenção:** a exclusão é permanente e não pode ser desfeita.")
            _loja_del = st.selectbox("🏢 Loja do orçamento",
                                      ["W & A Locações", "G & J — Locvix"],
                                      key="_del_loja")
            _c_num, _c_buscar = st.columns([4, 1])
            with _c_num:
                _orc_num = st.text_input("Nº do Orçamento (ID ou código)",
                                          key="_del_orc_num", placeholder="Ex: 7")
            with _c_buscar:
                st.markdown("<div style='margin-top:1.75rem'>", unsafe_allow_html=True)
                _btn_buscar = st.button("🔍 Buscar", use_container_width=True, key="_btn_del_buscar")
                st.markdown("</div>", unsafe_allow_html=True)

            # ── Busca dados do orçamento ──
            if _btn_buscar:
                if not _orc_num.strip():
                    st.error("❌ Informe o número do orçamento.")
                    st.session_state.pop("_del_orc_dados", None)
                else:
                    import locvix as _lv3, importlib as _il3, os as _os3, sys as _sys3
                    _base3 = _os3.path.dirname(_os3.path.abspath(__file__))
                    if _base3 not in _sys3.path: _sys3.path.insert(0, _base3)
                    if "locvix" in _sys3.modules:
                        _lv3 = _sys3.modules["locvix"]; _il3.reload(_lv3)
                    else:
                        import locvix as _lv3
                    _lv3.GCK_ACCESS_TOKEN = _os3.getenv("GCK_ACCESS_TOKEN", _lv3.GCK_ACCESS_TOKEN)
                    _lv3.GCK_SECRET_TOKEN = _os3.getenv("GCK_SECRET_TOKEN", _lv3.GCK_SECRET_TOKEN)
                    _loja3 = _lv3.LOJA_WA_ID if "W & A" in _loja_del else _lv3.LOJA_GJ_ID
                    with st.spinner("🔍 Buscando orçamento..."):
                        _dados = _lv3.buscar_orcamento_por_id(_orc_num.strip(), loja_id=_loja3)
                    st.session_state["_del_orc_dados"] = _dados
                    st.session_state.pop("_del_confirmado", None)

            # ── Exibe dados encontrados ──
            _dados_orc = st.session_state.get("_del_orc_dados")
            if _dados_orc:
                if not _dados_orc["ok"]:
                    st.error(f"❌ {_dados_orc['msg']}")
                else:
                    _valor_fmt = f"R$ {_dados_orc['valor']:,.2f}".replace(",","X").replace(".",",").replace("X",".")
                    st.info(
                        f"**Orçamento nº {_dados_orc['codigo']}**  \n"
                        f"👤 Cliente: **{_dados_orc['cliente']}**  \n"
                        f"📅 Data: {_dados_orc['data']}  |  "
                        f"📌 Situação: {_dados_orc['situacao']}  |  "
                        f"💰 Valor: {_valor_fmt}"
                    )
                    if not st.session_state.get("_del_confirmado"):
                        if st.button("🗑️ Excluir este orçamento", type="primary",
                                      use_container_width=True, key="_btn_del_confirmar"):
                            st.session_state["_del_confirmado"] = True
                            st.rerun()
                    else:
                        st.error(f"⚠️ Confirma a exclusão do orçamento **nº {_dados_orc['codigo']}** "
                                  f"do cliente **{_dados_orc['cliente']}**? Esta ação é irreversível!")
                        _c1, _c2 = st.columns(2)
                        with _c1:
                            if st.button("✅ Sim, excluir definitivamente", type="primary",
                                          use_container_width=True, key="_btn_del_sim"):
                                import locvix as _lv3, importlib as _il3, os as _os3, sys as _sys3
                                _base3 = _os3.path.dirname(_os3.path.abspath(__file__))
                                if _base3 not in _sys3.path: _sys3.path.insert(0, _base3)
                                if "locvix" in _sys3.modules:
                                    _lv3 = _sys3.modules["locvix"]; _il3.reload(_lv3)
                                else:
                                    import locvix as _lv3
                                _lv3.GCK_ACCESS_TOKEN = _os3.getenv("GCK_ACCESS_TOKEN", _lv3.GCK_ACCESS_TOKEN)
                                _lv3.GCK_SECRET_TOKEN = _os3.getenv("GCK_SECRET_TOKEN", _lv3.GCK_SECRET_TOKEN)
                                _loja3 = _lv3.LOJA_WA_ID if "W & A" in _loja_del else _lv3.LOJA_GJ_ID
                                with st.spinner("⏳ Excluindo..."):
                                    _res_del = _lv3.deletar_orcamento_api(_dados_orc["id"], loja_id=_loja3)
                                if _res_del["ok"]:
                                    st.success(f"✅ {_res_del['msg']}")
                                    st.session_state["_show_del_orc"] = False
                                    st.session_state.pop("_del_confirmado", None)
                                    st.session_state.pop("_del_orc_dados", None)
                                    st.rerun()
                                else:
                                    st.error(f"❌ {_res_del['msg']}")
                                    st.session_state.pop("_del_confirmado", None)
                        with _c2:
                            if st.button("↩️ Cancelar", use_container_width=True, key="_btn_del_nao"):
                                st.session_state.pop("_del_confirmado", None)
                                st.rerun()

    # ── Formulário Novo Orçamento (visível na área principal quando acionado) ──
    if st.session_state.get("modulo_ativo") == "orcamento" and st.session_state.get("_show_orc_form"):
        with st.container(border=True):
            _col_tit, _col_fechar = st.columns([6, 1])
            with _col_tit:
                st.markdown("### ➕ Criar Novo Orçamento")
            with _col_fechar:
                if st.button("❌ Fechar", key="_btn_fechar_orc"):
                    st.session_state["_show_orc_form"] = False
                    st.rerun()

            @st.cache_data(ttl=300, show_spinner="Carregando dados do ERP...")
            def _load_aux_orc():
                import locvix as _lv, importlib as _il, os as _os, sys as _sys
                _base = _os.path.dirname(_os.path.abspath(__file__))
                if _base not in _sys.path:
                    _sys.path.insert(0, _base)
                if "locvix" in _sys.modules:
                    _lv = _sys.modules["locvix"]; _il.reload(_lv)
                else:
                    import locvix as _lv
                _lv.GCK_ACCESS_TOKEN = _os.getenv("GCK_ACCESS_TOKEN", _lv.GCK_ACCESS_TOKEN)
                _lv.GCK_SECRET_TOKEN = _os.getenv("GCK_SECRET_TOKEN", _lv.GCK_SECRET_TOKEN)
                return {
                    "clientes":   _lv.buscar_clientes(),
                    "centros_cc": _lv.buscar_centros_custo(),
                    "situacoes":  _lv.buscar_situacoes_orcamento(),
                    "formas_pag": _lv.buscar_formas_pagamento(),
                    "servicos":   _lv.buscar_servicos(),
                    "vendedores": _lv.buscar_vendedores(),
                }

            _aux = _load_aux_orc()
            _clientes   = _aux["clientes"]
            _centros    = _aux["centros_cc"]
            _situacoes  = _aux["situacoes"]
            _formas     = _aux["formas_pag"]
            _servicos   = _aux["servicos"]
            _vendedores = _aux["vendedores"]

            def _opts(lista, campo="nome"):
                # suporta chave "nome" (minúsculo) ou "Nome" (maiúsculo)
                return [""] + [x.get(campo) or x.get(campo.title()) or x.get(campo.upper()) or "" for x in lista]
            def _id_por_nome(lista, nome):
                return next((x.get("id") or x.get("ID") or "" for x in lista if (x.get("nome") or x.get("Nome") or "") == nome), "")
            def _preco_serv(nome): return next((x["preco"] for x in _servicos if x.get("nome") == nome), 0.0)

            st.markdown("#### 📋 Dados do Orçamento")
            _loja_opts = ["G & J — Locvix (padrão)", "W & A Locações"]
            col_a, col_b, col_c = st.columns([2, 1, 1])
            with col_a:
                loja_sel = st.selectbox("🏢 Empresa (Loja)", _loja_opts, key="_orc_loja")
            with col_b:
                data_orc = st.date_input("📅 Data", value=date.today(), format="DD/MM/YYYY", key="_orc_data")
            with col_c:
                validade_orc = st.text_input("⏳ Validade", value="30 dias", key="_orc_validade")

            col_d, col_e = st.columns([3, 2])
            with col_d:
                cli_nome = st.selectbox("👤 Cliente *", _opts(_clientes), key="_orc_cliente",
                                        help="Comece a digitar para filtrar")
            with col_e:
                vend_nome = st.selectbox("🧑\u200d💼 Vendedor", _opts(_vendedores), key="_orc_vendedor")

            col_f, col_g = st.columns([2, 2])
            with col_f:
                cc_nome = st.selectbox("🏷️ Centro de Custo / Equipamento", _opts(_centros), key="_orc_cc")
            with col_g:
                sit_nome = st.selectbox("📌 Situação", _opts(_situacoes),
                                        index=1 if _situacoes else 0, key="_orc_sit")

            cuidados_orc = st.text_input("📬 Aos cuidados de", key="_orc_cuidados",
                                         placeholder="Nome do responsável no cliente (opcional)")
            intro_orc = st.text_area("📝 Introdução", key="_orc_intro", height=80,
                                     placeholder="Texto introdutório da proposta (opcional)")
            obs_orc = st.text_area("📄 Observações", key="_orc_obs", height=60,
                                   placeholder="Observações adicionais (opcional)")

            st.markdown("#### 🔧 Serviços (Horas de Máquina)")
            n_serv = st.number_input("Quantidade de itens", min_value=1, max_value=20,
                                     value=1, step=1, key="_orc_n_serv")
            servicos_form = []
            for _i in range(int(n_serv)):
                st.markdown(f"**Item {_i + 1}**")
                _c1, _c2, _c3, _c4 = st.columns([3, 1, 1.2, 1.2])
                with _c1:
                    _svc_nome = st.selectbox("Serviço", _opts(_servicos), key=f"_orc_svc_nome_{_i}")
                with _c2:
                    _svc_qtd = st.number_input("Qtd (h)", min_value=0.0, value=1.0,
                                               step=0.5, format="%.2f", key=f"_orc_svc_qtd_{_i}")
                with _c3:
                    _svc_preco = st.number_input("Vlr Unit (R$)", min_value=0.0,
                                                 value=float(_preco_serv(_svc_nome) if _svc_nome else 0.0),
                                                 step=0.01, format="%.2f", key=f"_orc_svc_preco_{_i}")
                with _c4:
                    _svc_desc = st.number_input("Desconto (R$)", min_value=0.0, value=0.0,
                                                step=0.01, format="%.2f", key=f"_orc_svc_desc_{_i}")
                _svc_det = st.text_input(f"Detalhes item {_i+1}", key=f"_orc_svc_det_{_i}",
                                         placeholder="Ex: Turno diurno 10h/dia")
                _svc_id = _id_por_nome(_servicos, _svc_nome) if _svc_nome else ""
                if _svc_nome:
                    servicos_form.append({"servico": {
                        "id": _svc_id, "servico_id": _svc_id, "nome_servico": _svc_nome,
                        "detalhes": _svc_det, "sigla_unidade": "H",
                        "quantidade": str(_svc_qtd), "valor_venda": str(_svc_preco),
                        "tipo_desconto": "R$", "desconto_valor": str(_svc_desc),
                        "desconto_porcentagem": "0",
                    }})

            _total = sum(
                float(s["servico"]["quantidade"]) * float(s["servico"]["valor_venda"])
                - float(s["servico"]["desconto_valor"]) for s in servicos_form
            ) if servicos_form else 0.0
            st.markdown(f"**💰 Total estimado: R$ {_total:,.2f}**".replace(",","X").replace(".",",").replace("X","."))

            st.markdown("#### 💳 Pagamento")
            _pc1, _pc2 = st.columns(2)
            with _pc1:
                cond_pag = st.radio("Condição", ["À vista", "Parcelado"], horizontal=True, key="_orc_cond")
            with _pc2:
                forma_nome = st.selectbox("Forma de Pagamento", _opts(_formas), key="_orc_forma")
            if cond_pag == "Parcelado":
                _pp1, _pp2 = st.columns(2)
                with _pp1:
                    n_parcelas = st.number_input("Nº de parcelas", min_value=1, max_value=60,
                                                 value=1, key="_orc_parcelas")
                with _pp2:
                    d1_parcela = st.date_input("Data 1ª parcela", value=date.today(),
                                               format="DD/MM/YYYY", key="_orc_d1parc")
                intervalo_dias = st.number_input("Intervalo (dias)", min_value=0, value=30, key="_orc_intervalo")
            else:
                n_parcelas = 1; d1_parcela = date.today(); intervalo_dias = 0

            st.markdown("")
            if st.button("📤 Criar Orçamento", type="primary",
                         use_container_width=True, key="_btn_criar_orc"):
                if not cli_nome:
                    st.error("❌ Selecione um cliente.")
                elif not servicos_form:
                    st.error("❌ Adicione pelo menos um serviço.")
                elif not sit_nome:
                    st.error("❌ Selecione a situação.")
                else:
                    import locvix as _lv2, importlib as _il2, os as _os2, sys as _sys2
                    _base2 = _os2.path.dirname(_os2.path.abspath(__file__))
                    if _base2 not in _sys2.path:
                        _sys2.path.insert(0, _base2)
                    if "locvix" in _sys2.modules:
                        _lv2 = _sys2.modules["locvix"]; _il2.reload(_lv2)
                    else:
                        import locvix as _lv2
                    _lv2.GCK_ACCESS_TOKEN = _os2.getenv("GCK_ACCESS_TOKEN", _lv2.GCK_ACCESS_TOKEN)
                    _lv2.GCK_SECRET_TOKEN = _os2.getenv("GCK_SECRET_TOKEN", _lv2.GCK_SECRET_TOKEN)

                    _cli_id   = _id_por_nome(_clientes, cli_nome)
                    _vend_id  = _id_por_nome(_vendedores, vend_nome) if vend_nome else ""
                    _cc_id    = _id_por_nome(_centros, cc_nome) if cc_nome else ""
                    _sit_id   = _id_por_nome(_situacoes, sit_nome) if sit_nome else ""
                    _forma_id = _id_por_nome(_formas, forma_nome) if forma_nome else ""
                    _loja_id  = _lv2.LOJA_WA_ID if "W & A" in loja_sel else _lv2.LOJA_GJ_ID

                    _payload = {
                        "cliente_id": int(_cli_id) if _cli_id else None,
                        "data": data_orc.strftime("%Y-%m-%d"),
                        "validade": validade_orc or "30 dias",
                        "situacao_id": _sit_id,
                        "condicao_pagamento": "a_vista" if cond_pag == "À vista" else "parcelado",
                        "servicos": servicos_form,
                    }
                    if _vend_id:     _payload["vendedor_id"]      = _vend_id
                    if _cc_id:       _payload["centro_custo_id"]  = int(_cc_id)
                    if cuidados_orc: _payload["aos_cuidados_de"]  = cuidados_orc
                    if intro_orc:    _payload["introducao"]        = intro_orc
                    if obs_orc:      _payload["observacoes"]       = obs_orc
                    if _forma_id:
                        _payload["forma_pagamento_id"]    = _forma_id
                        _payload["numero_parcelas"]       = int(n_parcelas)
                        _payload["data_primeira_parcela"] = d1_parcela.strftime("%Y-%m-%d")
                        if intervalo_dias:
                            _payload["intervalo_dias"] = int(intervalo_dias)

                    with st.spinner("⏳ Enviando orçamento ao GestãoClick..."):
                        _res = _lv2.criar_orcamento_api(_payload, loja_id=_loja_id)

                    if _res["ok"]:
                        st.success(f"✅ {_res['msg']}")
                        if _res["pdf_bytes"]:
                            st.download_button(
                                label=f"📄 Baixar PDF — Orçamento Nº {_res['codigo']}",
                                data=_res["pdf_bytes"],
                                file_name=f"Orcamento_{_res['codigo']}.pdf",
                                mime="application/pdf",
                                use_container_width=True,
                            )
                    else:
                        st.error(f"❌ Erro: {_res['msg']}")

    _ctx_dash = st.container()

    with _ctx_dash:
        _html_safe = st.session_state[HTML_KEY].encode("ascii", errors="xmlcharrefreplace").decode("ascii")

        # Inject module-activation script based on sidebar selection
        _modulo_sel = st.session_state.get("modulo_ativo", "geral")
        _allowed    = set(_usuario_modulos)
        if _modulo_sel not in _allowed:
            _modulo_sel = next(iter(_allowed), "geral")
        _inject = (
            f"<script>(function(){{function _am(){{if(typeof setModulo==='function')"
            f"setModulo('{_modulo_sel}');else setTimeout(_am,80);}}"
            f"if(document.readyState==='complete')_am();"
            f"else window.addEventListener('load',_am);}})()</script>"
        )
        _html_safe = _html_safe.replace("</body>", f"{_inject}</body>")

        components.html(
            _html_safe,
            height=1600,
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
