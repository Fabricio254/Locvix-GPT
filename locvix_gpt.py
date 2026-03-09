"""
LOCVIX GPT - Sistema de Perguntas e Respostas
LOCVIX Guindastes e Serviços
Desenvolvido por Fabrício Zamprogno
"""

import streamlit as st
import streamlit.components.v1 as components
import json
import unicodedata
import os
from datetime import datetime
import base64

# Cores da LOCVIX
COLORS = {
    "primary": "#F47920",       # Laranja LOCVIX
    "primary_dark": "#D4661A",
    "secondary": "#1B2A6B",     # Azul escuro LOCVIX
    "light_bg": "#F5F5F5",
    "white": "#FFFFFF",
    "text": "#333333",
    "success": "#00A86B",
    "error": "#FF6B6B"
}

def get_logo_base64():
    """Retorna a logo LOCVIX em base64 para embed em HTML"""
    logo_path = "logo locvix.jfif"
    if os.path.exists(logo_path):
        with open(logo_path, "rb") as f:
            data = base64.b64encode(f.read()).decode()
        return f"data:image/jpeg;base64,{data}"
    return None

def normalize_text(text):
    """Normaliza texto removendo acentos, til, cedilha, etc."""
    if not text:
        return text
    text = text.lower()
    text = unicodedata.normalize('NFD', text)
    text = ''.join(char for char in text if unicodedata.category(char) != 'Mn')
    return text

@st.cache_resource
def load_qa_data():
    """Carrega dados QA do JSON"""
    try:
        if os.path.exists("data/qa_data.json"):
            with open("data/qa_data.json", "r", encoding="utf-8") as f:
                return json.load(f)
        else:
            st.error("❌ Arquivo data/qa_data.json não encontrado")
            return None
    except Exception as e:
        st.error(f"❌ Erro ao carregar dados: {str(e)}")
        return None

def ask_question(question, qa_data):
    """Busca resposta para uma pergunta"""
    if not qa_data:
        return None, []

    question_lower = normalize_text(question)

    # Busca direta por substring
    for question_obj in qa_data.get("questions", []):
        if question_lower in normalize_text(question_obj['question']):
            return question_obj['answer'], [question_obj]

    best_match = None
    best_score = 0

    generic_words = {'o', 'a', 'os', 'as', 'um', 'uma', 'de', 'da', 'do', 'das', 'dos',
                     'em', 'na', 'no', 'nas', 'nos', 'para', 'por', 'com', 'como', 'quando',
                     'onde', 'que', 'qual', 'quais', 'quem', 'e', 'sao', 'foi', 'sera',
                     'tem', 'ter', 'fazer', 'feito', 'pode', 'posso', 'devo', 'deve'}

    for question_obj in qa_data.get("questions", []):
        pergunta_excel = normalize_text(question_obj['question'])
        words_user = set(question_lower.split()) - generic_words
        words_excel = set(pergunta_excel.split()) - generic_words

        if len(words_user) == 0 or len(words_excel) == 0:
            continue

        common_words = words_user.intersection(words_excel)

        if len(common_words) > 0:
            score = len(common_words) / len(words_user)
            if len(words_user) > 3 and len(common_words) < 2:
                continue
            if score > best_score:
                best_score = score
                best_match = question_obj

    if best_match and best_score > 0.3:
        return best_match['answer'], [best_match]

    return "Não encontrei uma resposta exata para sua pergunta. Tente usar as sugestões no painel lateral ou reformule sua pergunta.", []

def setup_page():
    """Configura a página Streamlit"""
    st.set_page_config(
        page_title="LOCVIX GPT",
        page_icon="🏗️",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    st.markdown(f"""
    <style>
    :root {{
        --primary: {COLORS['primary']};
        --secondary: {COLORS['secondary']};
        --light-bg: {COLORS['light_bg']};
    }}

    [data-testid="stAppViewContainer"] {{
        background-color: white;
    }}

    [data-testid="stHeader"] {{
        background-color: {COLORS['secondary']};
    }}

    [data-testid="stSidebar"] {{
        background-color: {COLORS['light_bg']};
        border-right: 3px solid {COLORS['primary']};
    }}

    /* Forçar sidebar sempre visível no desktop */
    @media (min-width: 769px) {{
        section[data-testid="stSidebar"] {{
            transform: translateX(0) !important;
            display: flex !important;
            min-width: 21rem !important;
        }}
        [data-testid="stSidebarCollapseButton"] {{
            display: none !important;
        }}
    }}

    /* Botões */
    .stButton > button {{
        background-color: {COLORS['secondary']};
        color: white;
        border: 2px solid {COLORS['secondary']};
        border-radius: 8px;
        font-weight: bold;
        transition: all 0.3s ease;
    }}

    .stButton > button:hover {{
        background-color: {COLORS['primary']};
        border-color: {COLORS['primary']};
        box-shadow: 0 4px 8px rgba(244, 121, 32, 0.3);
    }}

    /* Input */
    .stTextInput > div > div > input {{
        border: 2px solid {COLORS['secondary']} !important;
        border-radius: 8px;
    }}

    .stTextInput > div > div > input:focus {{
        border-color: {COLORS['primary']} !important;
        box-shadow: 0 0 0 3px rgba(244, 121, 32, 0.15);
    }}

    h1, h2, h3 {{
        color: {COLORS['secondary']};
    }}

    /* ===== RESPONSIVO MOBILE ===== */
    @media (max-width: 768px) {{

        [data-testid="stAppViewContainer"] {{
            overflow-x: hidden;
        }}

        [data-testid="stMain"] > div:first-child {{
            padding-left: 0.5rem !important;
            padding-right: 0.5rem !important;
        }}

        .stButton > button {{
            font-size: 1em !important;
            padding: 10px 4px !important;
            min-height: 48px;
        }}

        .stTextInput > div > div > input {{
            font-size: 1em !important;
            padding: 12px !important;
            min-height: 48px;
        }}

        .header-subtitle {{
            display: none !important;
        }}

        h3 {{ font-size: 1em !important; }}
        h2 {{ font-size: 1.1em !important; }}
    }}
    </style>
    """, unsafe_allow_html=True)

def create_sidebar(qa_data):
    """Cria sidebar com sugestões"""
    with st.sidebar:
        # Logo LOCVIX no topo da sidebar
        logo_path = "logo locvix.jfif"
        if os.path.exists(logo_path):
            col_logo, col_txt = st.columns([1, 2])
            with col_logo:
                st.image(logo_path, width=110)
            with col_txt:
                st.markdown(f"""
                <div style="padding-top:14px; color:{COLORS['secondary']}; font-weight:bold; font-size:1.1em; line-height:1.3;">
                LOCVIX<br><span style="font-size:0.72em; font-weight:normal; color:#888;">Guindastes e Serviços</span>
                </div>""", unsafe_allow_html=True)
            st.divider()
        st.markdown(f"## 🎯 Sugestões de Perguntas")

        search_text = st.text_input(
            "🔍 Buscar tópico ou pergunta",
            placeholder="Digite aqui..."
        ).lower()

        if not search_text:
            topics = qa_data.get("topics", {})
        else:
            topics = {}
            for topic, data in qa_data.get("topics", {}).items():
                if search_text in normalize_text(topic):
                    topics[topic] = data
                else:
                    filtered_questions = [
                        q for q in data.get("questions", [])
                        if search_text in normalize_text(q['question'])
                    ]
                    if filtered_questions:
                        topics[topic] = {"questions": filtered_questions}

        for topic, data in topics.items():
            with st.expander(f"**{topic}**", expanded=False):
                for question in data.get("questions", [])[:5]:
                    if st.button(
                        f"• {question['question'][:60]}...",
                        key=f"btn_{question['id']}",
                        use_container_width=True
                    ):
                        st.session_state.user_question = question['question']
                        st.rerun()

        st.divider()

        st.markdown(f"""
        <div style="padding: 12px; background-color: #f0f0f0; border-radius: 8px;">
        <small>
        📊 <strong>Total:</strong> {qa_data.get('total_questions', 0)} perguntas<br>
        📂 <strong>Tópicos:</strong> {len(qa_data.get('topics', {}))} tópicos<br>
        🕐 <strong>Atualizado:</strong> {qa_data.get('last_updated', 'N/A')}
        </small>
        </div>
        """, unsafe_allow_html=True)

        st.divider()

        st.markdown(f"""
        <div style="background-color: {COLORS['primary']}; color: white;
                    padding: 14px; border-radius: 8px; margin-top: 4px;">
        <strong>💡 Dica</strong><br>
        <small>Use as sugestões acima para explorar <strong>{qa_data.get('total_questions', 0)}</strong>
        perguntas pré-organizadas. É mais rápido! 🚀</small>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(f"""
        <div style="margin-top: 24px; padding: 10px 12px; border-top: 1px solid #ddd;
                    text-align: center; color: #aaa; font-size: 0.72em; line-height: 1.6;">
            Desenvolvido por<br>
            <strong style="color: {COLORS['secondary']};">Fabrício Zamprogno</strong><br>
            📱 27-996076278
        </div>
        """, unsafe_allow_html=True)

def check_password():
    """Tela de login com usuário e senha"""
    if st.session_state.get("authenticated"):
        return True

    logo_b64 = get_logo_base64()
    logo_html = f'<img src="{logo_b64}" style="max-width:200px; margin-bottom:12px;" />' if logo_b64 else ""

    st.markdown(f"""
    <style>
    [data-testid="stSidebar"] {{ display: none !important; }}
    .login-box {{
        max-width: 400px;
        margin: 60px auto;
        padding: 36px 32px;
        border-radius: 14px;
        border: 2px solid {COLORS['primary']};
        background: white;
        text-align: center;
        box-shadow: 0 4px 24px rgba(0,0,0,0.10);
    }}
    .login-title {{ color: {COLORS['secondary']}; font-size: 1.4em; font-weight: 800; margin-top: 8px; margin-bottom: 4px; }}
    .login-sub {{ color: #888; font-size: 0.95em; margin-bottom: 20px; }}
    </style>
    <div class="login-box">
        {logo_html}
        <div class="login-title">LOCVIX GPT</div>
        <div class="login-sub">Digite suas credenciais para acessar</div>
    </div>
    """, unsafe_allow_html=True)

    with st.form("login_form"):
        usuario = st.text_input("Usuário", placeholder="Digite o usuário")
        senha = st.text_input("Senha", type="password", placeholder="Digite a senha")
        entrar = st.form_submit_button("🔐 Entrar", use_container_width=True)

    if entrar:
        correct_user = st.secrets.get("auth", {}).get("username", "locvix")
        correct_pass = st.secrets.get("auth", {}).get("password", "zampa254")
        if usuario == correct_user and senha == correct_pass:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("❌ Usuário ou senha incorretos!")

    return False


def main():
    """Função principal"""
    setup_page()

    if not check_password():
        return

    qa_data = load_qa_data()

    if qa_data is None:
        st.error("❌ Falha ao carregar dados. Verifique se o arquivo data/qa_data.json existe.")
        return

    # Sidebar PRIMEIRO
    create_sidebar(qa_data)

    # Header
    logo_b64 = get_logo_base64()
    logo_img_html = f'<img src="{logo_b64}" style="height:38px; object-fit:contain; vertical-align:middle; margin-right:10px;" />' if logo_b64 else ""
    st.markdown(f"""
    <div style="background: linear-gradient(90deg, {COLORS['secondary']} 0%, {COLORS['primary']} 100%);
                padding: 8px 16px; border-radius: 6px; margin-bottom: 12px;
                display: flex; align-items: center; gap: 10px; flex-wrap: wrap;">
        {logo_img_html}
        <span style="color: white; font-size: 1.1em; font-weight: bold; white-space: nowrap;">LOCVIX GPT</span>
        <span class="header-subtitle" style="color: rgba(255,255,255,0.80); font-size: 0.8em;">
            | Sistema de Perguntas e Respostas - LOCVIX Guindastes e Serviços
        </span>
    </div>
    """, unsafe_allow_html=True)

    # Inicializar session state
    if 'messages' not in st.session_state:
        st.session_state.messages = []
    if 'user_question' not in st.session_state:
        st.session_state.user_question = ""
    if 'input_key' not in st.session_state:
        st.session_state.input_key = 0
    if 'enter_pressed' not in st.session_state:
        st.session_state.enter_pressed = False
    if 'input_value' not in st.session_state:
        st.session_state.input_value = ""

    def on_input_submit():
        val = st.session_state.get(f"user_input_{st.session_state.input_key}", "")
        if val.strip():
            st.session_state.input_value = val
            st.session_state.enter_pressed = True

    # Expander de sugestões (visível no mobile)
    with st.expander("🎯 Ver Sugestões de Perguntas", expanded=False):
        search_mobile = st.text_input("🔍 Buscar", placeholder="Digite um tópico...", key="search_mobile").lower()
        topics = qa_data.get("topics", {})
        if search_mobile:
            topics = {t: d for t, d in topics.items()
                      if search_mobile in normalize_text(t) or
                      any(search_mobile in normalize_text(q['question']) for q in d.get('questions', []))}
        for topic, data in topics.items():
            st.markdown(f"**{topic}**")
            for q in data.get("questions", [])[:4]:
                if st.button(f"• {q['question'][:55]}...", key=f"mob_{q['id']}", use_container_width=True):
                    st.session_state.user_question = q['question']
                    st.rerun()

    # Histórico de mensagens
    if st.session_state.messages:
        st.subheader("💬 Histórico de Conversa")
        for msg in st.session_state.messages:
            if msg['role'] == 'user':
                st.markdown(f"""
                <div style="background-color: {COLORS['secondary']}; color: white;
                            padding: 12px; border-radius: 8px; margin: 8px 0;">
                <strong>Você:</strong> {msg['content']}
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div style="background-color: {COLORS['light_bg']};
                            padding: 12px; border-left: 4px solid {COLORS['primary']};
                            border-radius: 8px; margin: 8px 0;">
                <strong>Assistente:</strong><br>{msg['content']}
                </div>
                """, unsafe_allow_html=True)
                if 'source' in msg and msg['source']:
                    st.markdown(f"""
                    <div style="background-color: {COLORS['light_bg']};
                                border-left: 4px solid {COLORS['secondary']};
                                padding: 8px; border-radius: 8px; font-size: 0.9em; margin: 4px 0;">
                    📋 <strong>Tópico:</strong> {msg['source']['topic']}
                    </div>
                    """, unsafe_allow_html=True)

        # Scroll automático ao final do chat
        st.markdown('<div id="chat-bottom"></div>', unsafe_allow_html=True)
        components.html("""
            <script>
                window.parent.document.getElementById('chat-bottom').scrollIntoView({behavior: 'smooth', block: 'end'});
            </script>
        """, height=0)

    # Input
    st.subheader("❓ Faça sua Pergunta")
    user_input = st.text_input(
        "Clique aqui e digite sua pergunta",
        placeholder="Ex: Quais equipamentos a LOCVIX possui?",
        key=f"user_input_{st.session_state.input_key}",
        on_change=on_input_submit
    )
    col_btn, col_clear = st.columns(2)
    with col_btn:
        send_button = st.button("📤 Enviar", use_container_width=True)
    with col_clear:
        if st.button("🗑️ Limpar Chat", use_container_width=True):
            st.session_state.messages = []
            st.rerun()

    # Enter pressionado
    if st.session_state.enter_pressed:
        user_input = st.session_state.input_value
        send_button = True
        st.session_state.enter_pressed = False

    # Se tem pergunta do sidebar
    if st.session_state.user_question:
        user_input = st.session_state.user_question
        send_button = True

    # Processar pergunta
    if send_button and user_input.strip():
        st.session_state.messages.append({"role": "user", "content": user_input})
        answer, sources = ask_question(user_input, qa_data)
        source_info = sources[0] if sources else None
        st.session_state.messages.append({
            "role": "assistant",
            "content": answer,
            "source": source_info
        })
        st.session_state.user_question = ""
        st.session_state.input_key += 1
        st.rerun()

    # Footer
    st.divider()
    st.markdown(f"""
    <div style="text-align: center; padding: 20px; color: #666;">
    <small>
    🏗️ <strong>LOCVIX Guindastes e Serviços</strong> | LOCVIX GPT v1.0<br>
    Sistema de Atendimento Inteligente | Desenvolvido por Fabrício Zamprogno
    </small>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
