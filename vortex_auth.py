import hmac, hashlib, os, base64
from datetime import datetime
import streamlit as st

AUTH_SECRET = os.environ.get("VORTEX_AUTH_SECRET", "change-this-secret-in-production")
PARAM_NAME  = "t"   # query param: ?t=usuario|hmac
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo.png")

@st.cache_data(show_spinner=False)
def _logo_b64() -> str:
    try:
        with open(LOGO_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except FileNotFoundError:
        return ""

def logo_html(max_height_px: int = 70, fallback: str = '<div style="font-size:20px; font-weight:800; color:white; letter-spacing:-0.5px;">⚡ Vórtex</div>') -> str:
    b64 = _logo_b64()
    if not b64:
        return fallback
    return f'<img src="data:image/png;base64,{b64}" style="max-height:{max_height_px}px; width:auto; display:block; margin:0 auto; border-radius:8px;" alt="Vórtex" />'

USERS = {
    "admin":      {"senha": "admin123",    "perfil": "admin",        "nome": "Admin User"},
    
    "planner":   {"senha": "planner123", "perfil": "visualizacao", "nome": "Planning Team"},
    "buyer":     {"senha": "buyer123",   "perfil": "compras",      "nome": "Purchasing Team"},
}

def assinar(usuario: str) -> str:
    sig = hmac.new(AUTH_SECRET.encode(), usuario.encode(), hashlib.sha256).hexdigest()[:16]
    return f"{usuario}|{sig}"

def validar(token: str):
    if not token or "|" not in token: return None
    user, sig = token.rsplit("|", 1)
    esperado = hmac.new(AUTH_SECRET.encode(), user.encode(), hashlib.sha256).hexdigest()[:16]
    if hmac.compare_digest(sig, esperado) and user in USERS:
        return user
    return None

def gravar_token(usuario: str):
    st.query_params[PARAM_NAME] = assinar(usuario)

def apagar_token():
    try:
        del st.query_params[PARAM_NAME]
    except Exception:
        pass

def restaurar_sessao():
    # Se a sessão já está logada, garanta que o token esteja sempre na URL
    # (F5 ou navegação entre páginas pode descartar query params).
    if st.session_state.get("logado"):
        user = st.session_state.get("usuario")
        if user and st.query_params.get(PARAM_NAME) != assinar(user):
            try: st.query_params[PARAM_NAME] = assinar(user)
            except Exception: pass
        return
    token = st.query_params.get(PARAM_NAME)
    user = validar(token) if token else None
    if user:
        st.session_state["logado"]  = True
        st.session_state["usuario"] = user
        st.session_state["perfil"]  = USERS[user]["perfil"]
        st.session_state["nome"]    = USERS[user]["nome"]

GLOBAL_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

    html, body, [class*="css"], .stApp, .stMarkdown, .stDataFrame, button, input, select, textarea {
        font-family: 'Inter', system-ui, -apple-system, sans-serif !important;
    }
    code, pre, .stCode {
        font-family: 'JetBrains Mono', 'SF Mono', Consolas, monospace !important;
        font-variant-numeric: tabular-nums;
    }

    .block-container { padding: 4.5rem 1.75rem 2.5rem; max-width: 1600px; }

    /* ── Sidebar ─────────────────────────────────────────── */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0b1220 0%, #111a2e 100%);
        border-right: 1px solid rgba(255,255,255,0.06);
    }
    [data-testid="stSidebar"] * { color: #e2e8f0 !important; }
    [data-testid="stSidebar"] .stButton > button {
        background: rgba(255,255,255,0.08) !important;
        color: #e2e8f0 !important;
        border: 1px solid rgba(255,255,255,0.15) !important;
    }
    [data-testid="stSidebar"] .stButton > button:hover {
        background: rgba(239,68,68,0.2) !important;
        border-color: rgba(239,68,68,0.5) !important;
        color: #fca5a5 !important;
    }
    [data-testid="stSidebarNav"] a {
        border-radius: 8px; margin: 2px 10px;
        padding: 9px 14px !important; transition: background .15s ease;
        font-size: 13.5px; font-weight: 500;
    }
    [data-testid="stSidebarNav"] a:hover { background: rgba(37,99,235,0.16) !important; }
    [data-testid="stSidebarNav"] a[aria-selected="true"] {
        background: rgba(37,99,235,0.22) !important;
        box-shadow: inset 2px 0 0 #2563eb;
    }

    /* ── Page header ─────────────────────────────────────── */
    .page-header {
        display: flex; align-items: center; gap: 14px;
        margin-bottom: 22px; padding-bottom: 14px;
        border-bottom: 1px solid rgba(128,128,128,0.2);
    }
    .page-header h1 {
        font-size: 22px !important; font-weight: 700 !important;
        letter-spacing: -0.01em !important;
        margin: 0 !important; padding: 0 !important;
    }
    .page-header .subtitle {
        font-size: 13px; opacity: 0.72; margin-top: 2px;
    }

    /* ── Cards ───────────────────────────────────────────── */
    .pf-card {
        background: rgba(128,128,128,0.05); border-radius: 12px;
        padding: 18px 22px; border: 1px solid rgba(128,128,128,0.15);
        margin-bottom: 14px;
    }
    .pf-card-title {
        font-size: 11px; font-weight: 600; opacity: 0.65;
        text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 6px;
    }
    .pf-card-value {
        font-size: 26px; font-weight: 700;
        line-height: 1.1; letter-spacing: -0.01em;
    }
    .pf-card-sub { font-size: 12px; opacity: 0.6; margin-top: 4px; }

    /* ── Badges ──────────────────────────────────────────── */
    .badge-red    { background:rgba(220,38,38,0.15); color:#ef4444; padding:3px 9px; border-radius:999px; font-size:11px; font-weight:600; border:1px solid rgba(220,38,38,0.3); }
    .badge-yellow { background:rgba(202,138,4,0.15); color:#eab308; padding:3px 9px; border-radius:999px; font-size:11px; font-weight:600; border:1px solid rgba(202,138,4,0.3); }
    .badge-green  { background:rgba(22,163,74,0.15); color:#22c55e; padding:3px 9px; border-radius:999px; font-size:11px; font-weight:600; border:1px solid rgba(22,163,74,0.3); }
    .badge-blue   { background:rgba(37,99,235,0.15); color:#60a5fa; padding:3px 9px; border-radius:999px; font-size:11px; font-weight:600; border:1px solid rgba(37,99,235,0.3); }
    .badge-gray   { background:rgba(100,116,139,0.15); color:#94a3b8; padding:3px 9px; border-radius:999px; font-size:11px; font-weight:600; border:1px solid rgba(100,116,139,0.3); }

    /* ── Metrics (st.metric) — só bordas/espaçamento, sem forçar cor ─ */
    [data-testid="stMetric"] {
        border-radius: 12px;
        padding: 14px 16px;
        border: 1px solid rgba(128,128,128,0.18);
    }
    [data-testid="stMetricLabel"] {
        font-size: 11.5px !important; font-weight: 600 !important;
        text-transform: uppercase; letter-spacing: 0.05em;
        opacity: 0.75;
    }
    [data-testid="stMetricValue"] {
        font-size: 26px !important; font-weight: 700 !important;
        letter-spacing: -0.01em;
    }
    [data-testid="stMetricDelta"] { font-size: 12px !important; }

    /* ── Inputs — só borda e focus, sem forçar background ─── */
    .stTextInput input, .stNumberInput input, .stDateInput input,
    .stTextArea textarea {
        border-radius: 8px !important;
        font-size: 13.5px !important;
    }
    .stTextInput input:focus, .stNumberInput input:focus,
    .stDateInput input:focus, .stTextArea textarea:focus {
        border-color: #2563eb !important;
        box-shadow: 0 0 0 3px rgba(37,99,235,0.15) !important;
        outline: none !important;
    }

    /* ── Buttons — só radius/peso, mantém cor do tema ─────── */
    .stButton > button, .stDownloadButton > button {
        border-radius: 8px; font-weight: 600;
        font-size: 13.5px;
        transition: all .15s ease;
    }
    .stButton > button[kind="primary"]:hover,
    .stDownloadButton > button[kind="primary"]:hover {
        box-shadow: 0 4px 12px rgba(37,99,235,0.22);
        transform: translateY(-1px);
    }

    /* ── Checkboxes ──────────────────────────────────────── */
    .stCheckbox > label { font-size: 13.5px !important; }

    /* ── Tabs ────────────────────────────────────────────── */
    [data-baseweb="tab-list"] {
        border-bottom: 1px solid rgba(128,128,128,0.2);
        gap: 4px;
    }
    [data-baseweb="tab"] {
        font-weight: 600 !important; font-size: 13.5px !important;
        padding: 10px 16px !important;
        border-radius: 8px 8px 0 0 !important;
    }
    [data-baseweb="tab"][aria-selected="true"] {
        color: #2563eb !important;
        background: rgba(37,99,235,0.08) !important;
    }

    /* ── Expanders ───────────────────────────────────────── */
    [data-testid="stExpander"] {
        border: 1px solid rgba(128,128,128,0.2) !important;
        border-radius: 10px !important;
    }
    [data-testid="stExpander"] summary {
        font-size: 13px !important; font-weight: 500 !important;
        padding: 8px 14px !important;
    }

    /* ── DataFrame / Tables ──────────────────────────────── */
    [data-testid="stDataFrame"] {
        border-radius: 10px; overflow: hidden;
        border: 1px solid rgba(128,128,128,0.2);
    }

    /* ── Alerts ──────────────────────────────────────────── */
    [data-testid="stAlert"] { border-radius: 10px; }

    /* ── Dividers ────────────────────────────────────────── */
    hr { margin: 18px 0 !important; }

    /* ── Plotly backgrounds ──────────────────────────────── */
    .js-plotly-plot .plotly .bg { fill: transparent !important; }

    /* ── Tabelas da Home — sem quebra de texto ───────────── */
    .pf-home-table {
        background: rgba(128,128,128,0.04);
        border-radius: 10px;
        border: 1px solid rgba(128,128,128,0.18);
        overflow-x: auto; overflow-y: hidden;
        max-width: 100%;
    }
    .pf-home-table table { width: 100%; border-collapse: collapse; font-size: 12.5px; margin: 0; }
    .pf-home-table table th {
        background: rgba(128,128,128,0.08); font-weight: 600;
        text-transform: uppercase; font-size: 11px; letter-spacing: 0.05em;
        padding: 10px 14px; text-align: left;
        border-bottom: 1px solid rgba(128,128,128,0.2); white-space: nowrap;
        opacity: 0.85;
    }
    .pf-home-table table td {
        padding: 9px 14px; border-bottom: 1px solid rgba(128,128,128,0.1);
        white-space: nowrap;
    }
    .pf-home-table table tr:last-child td { border-bottom: none; }
    .pf-home-table table tr:hover td { background: rgba(128,128,128,0.06); }
    .pf-home-table table td:nth-child(n+4) { text-align: right; font-variant-numeric: tabular-nums; }
</style>
"""

def inject_style():
    st.markdown(GLOBAL_CSS, unsafe_allow_html=True)

def render_sidebar():
    """Sidebar padronizada usada em todas as páginas."""
    if not st.session_state.get("logado"):
        return
    with st.sidebar:
        st.markdown(f"""
        <div style="padding: 12px 8px 18px; border-bottom: 1px solid #2a3f5f; margin-bottom: 16px; text-align:center;">
            {logo_html(max_height_px=72)}
            <div style="font-size:11px; color:#64748b; margin-top:8px;">Sistema de Planning</div>
        </div>
        """, unsafe_allow_html=True)

        perfil_badge = "🔑 Admin" if st.session_state.get("perfil") == "admin" else "👁️ Visualização"
        st.markdown(f"""
        <div style="padding:10px 8px; margin-bottom:16px; background:rgba(255,255,255,0.05); border-radius:8px;">
            <div style="font-size:13px; font-weight:600; color:white;">{st.session_state.get('nome','')}</div>
            <div style="font-size:11px; color:#64748b; margin-top:2px;">{perfil_badge}</div>
        </div>
        """, unsafe_allow_html=True)

        dados_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dados", "dados_sap.xlsx")
        if os.path.exists(dados_path):
            mt = datetime.fromtimestamp(os.path.getmtime(dados_path))
            st.markdown(f"""
            <div style="padding:8px; margin-bottom:16px; background:rgba(22,163,74,0.15); border-radius:8px; border:1px solid rgba(22,163,74,0.3);">
                <div style="font-size:11px; color:#4ade80; font-weight:600;">✓ Dados atualizados</div>
                <div style="font-size:10px; color:#64748b;">{mt.strftime('%d/%m/%Y %H:%M')}</div>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div style="padding:8px; margin-bottom:16px; background:rgba(239,68,68,0.15); border-radius:8px; border:1px solid rgba(239,68,68,0.3);">
                <div style="font-size:11px; color:#f87171; font-weight:600;">⚠ Dados não encontrados</div>
                <div style="font-size:10px; color:#64748b;">Execute atualizar_dados.py</div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<div style='flex:1'></div>", unsafe_allow_html=True)
        if st.button("Sair", use_container_width=True, key=f"btn_sair_{st.session_state.get('usuario','x')}"):
            apagar_token()
            for k in ["logado","usuario","perfil","nome"]:
                st.session_state.pop(k, None)
            st.rerun()

# ── FILTROS COMPACTOS ─────────────────────────────────────────

def filtro_checkbox_dropdown(label, opcoes, key, default=None):
    """
    Dropdown flutuante com checkboxes estilo Excel/SAP.
    Usa st.popover — não empurra o conteúdo abaixo.
    """
    if default is None:
        default = list(opcoes)

    sel_key   = f"_cksel_{key}"
    nonce_key = f"_cknonce_{key}"

    if sel_key not in st.session_state:
        st.session_state[sel_key] = list(default)
    if nonce_key not in st.session_state:
        st.session_state[nonce_key] = 0

    selecionados = st.session_state[sel_key]
    n_total = len(opcoes)
    n_sel   = len([s for s in selecionados if s in opcoes])
    nonce   = st.session_state[nonce_key]

    # Label do botão mostra o estado
    if n_sel == 0:
        btn_label = f"⚠️ {label}: Nenhum"
    elif n_sel == n_total:
        btn_label = f"{label}: Todas ({n_total}) ▾"
    else:
        nomes = ", ".join(str(s) for s in opcoes if s in selecionados)[:28]
        extra = f" +{n_sel - len(nomes.split(','))}" if n_sel > 2 else ""
        btn_label = f"{label}: {nomes}{'...' if len(nomes)==28 else ''} ({n_sel}) ▾"

    with st.popover(btn_label, use_container_width=True):
        # Pesquisa
        busca = st.text_input(
            "Pesquisar", key=f"{key}_busca",
            placeholder="🔍 Pesquisar...",
            label_visibility="collapsed"
        )
        opcoes_filtradas = [o for o in opcoes if busca.lower() in str(o).lower()] if busca else list(opcoes)

        todos_marcados = all(o in selecionados for o in opcoes_filtradas) and len(opcoes_filtradas) > 0

        def _sel_tudo(sel_key=sel_key, nonce_key=nonce_key, opcoes_filtradas=opcoes_filtradas):
            atuais = st.session_state.get(sel_key, [])
            st.session_state[sel_key] = list(set(atuais) | set(opcoes_filtradas))
            st.session_state[nonce_key] = st.session_state.get(nonce_key, 0) + 1

        def _limpar(sel_key=sel_key, nonce_key=nonce_key, opcoes_filtradas=opcoes_filtradas):
            atuais = st.session_state.get(sel_key, [])
            st.session_state[sel_key] = [s for s in atuais if s not in opcoes_filtradas]
            st.session_state[nonce_key] = st.session_state.get(nonce_key, 0) + 1

        b1, b2 = st.columns(2)
        with b1:
            st.button(
                "☑ Selecionar Tudo",
                key=f"{key}_btn_tudo_{nonce}",
                use_container_width=True,
                type="secondary",
                on_click=_sel_tudo,
            )
        with b2:
            st.button(
                "☐ Limpar",
                key=f"{key}_btn_limpar_{nonce}",
                use_container_width=True,
                type="secondary",
                on_click=_limpar,
            )

        st.markdown("<hr style='margin:6px 0; opacity:0.2'>", unsafe_allow_html=True)

        novos   = list(selecionados)
        changed = False
        for op in opcoes_filtradas:
            checked = op in selecionados
            novo = st.checkbox(
                str(op), value=checked,
                key=f"{key}_op_{str(op).replace(' ','_')}_n{nonce}"
            )
            if novo and op not in novos:
                novos.append(op); changed = True
            elif not novo and op in novos:
                novos.remove(op); changed = True

        if changed:
            st.session_state[sel_key] = novos
            st.rerun()

    return [s for s in st.session_state[sel_key] if s in opcoes] or list(opcoes)


def filtro_familia(df, key="fam", label="Família"):
    opcoes = sorted(df["Família"].dropna().unique().tolist()) if "Família" in df.columns else []
    return filtro_checkbox_dropdown(label, opcoes, key=key)


def filtro_categoria(key="cat", default=None):
    """Popover com checkboxes para Categoria — mesmo padrão do filtro_familia."""
    if default is None:
        default = ["PA", "Revenda"]
    return filtro_checkbox_dropdown("Categoria", ["PA", "Revenda"], key=key, default=default)


def filtro_categoria_mp(key="cat_mp", default=None):
    """Popover com checkboxes para Categoria de MP (grupos OITM 119/120/121/122)."""
    opcoes = ["Mat. Prima", "Componente", "Embalagem", "Consumível"]
    if default is None:
        default = list(opcoes)
    return filtro_checkbox_dropdown("Categoria", opcoes, key=key, default=default)


def filtro_abc(key="abc", default=None):
    """Popover com checkboxes para ABC — mesmo padrão do filtro_familia."""
    if default is None:
        default = ["A", "B", "C"]
    return filtro_checkbox_dropdown("Classe ABC", ["A", "B", "C"], key=key, default=default)


def filtro_meses(opcoes, key="meses", default_n=4, label="Meses a exibir"):
    """Popover com checkboxes para meses — mesmo padrão de filtro_familia/filtro_categoria."""
    default = list(opcoes[:default_n]) if default_n else list(opcoes)
    return filtro_checkbox_dropdown(label, list(opcoes), key=key, default=default)
