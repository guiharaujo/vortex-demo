import sys


import streamlit as st
import pandas as pd
import os
from datetime import datetime
from auth import USERS, gravar_token, apagar_token, restaurar_sessao, render_sidebar, inject_style, logo_html, LOGO_PATH

# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================
st.set_page_config(
    page_title="Vórtex",
    page_icon=LOGO_PATH if os.path.exists(LOGO_PATH) else "⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

inject_style()

# ============================================================
# AUTENTICAÇÃO
# ============================================================
restaurar_sessao()

def login():
    # CSS local da tela de login — adapta ao tema (light/dark)
    st.markdown("""
    <style>
        [data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"] { display: none !important; }
        header[data-testid="stHeader"] { background: transparent !important; }
        .block-container { padding-top: 6vh !important; max-width: 900px !important; }

        /* Form wrapper — translúcido, funciona nos 2 temas */
        [data-testid="stForm"] {
            background: rgba(128,128,128,0.05) !important;
            border: 1px solid rgba(128,128,128,0.2) !important;
            border-radius: 14px !important;
            padding: 24px 28px !important;
            box-shadow: 0 8px 24px rgba(0,0,0,0.12) !important;
        }

        /* Botão Entrar — azul forte, funciona nos 2 temas */
        [data-testid="stForm"] button[kind="primaryFormSubmit"],
        [data-testid="stForm"] button[kind="primary"] {
            background: #2563eb !important;
            color: #ffffff !important;
            border: none !important;
            font-weight: 600 !important;
        }
        [data-testid="stForm"] button[kind="primaryFormSubmit"]:hover,
        [data-testid="stForm"] button[kind="primary"]:hover {
            background: #1d4ed8 !important;
        }
        [data-testid="stForm"] button[kind="primaryFormSubmit"] *,
        [data-testid="stForm"] button[kind="primary"] * { color: #ffffff !important; }

        .pf-login-title   { font-size: 28px; font-weight: 800; letter-spacing: -0.02em; }
        .pf-login-sub     { font-size: 13px; opacity: 0.7; margin-top: 4px; }
        .pf-login-foot    { text-align:center; margin-top:20px; font-size:11px; opacity: 0.5; }
    </style>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown(f"""
        <div style="text-align:center; padding: 16px 0 18px;">
            {logo_html(max_height_px=140)}
            <div class="pf-login-sub" style="margin-top:12px;">Sistema de Planning Inteligente</div>
        </div>
        """, unsafe_allow_html=True)

        with st.form("login_form"):
            usuario = st.text_input("Usuário", placeholder="seu.usuario")
            senha   = st.text_input("Senha", type="password", placeholder="••••••••")
            submit  = st.form_submit_button("Entrar", use_container_width=True, type="primary")

            if submit:
                if usuario in USERS and USERS[usuario]["senha"] == senha:
                    st.session_state["logado"]  = True
                    st.session_state["usuario"] = usuario
                    st.session_state["perfil"]  = USERS[usuario]["perfil"]
                    st.session_state["nome"]    = USERS[usuario]["nome"]
                    gravar_token(usuario)
                    st.rerun()
                else:
                    st.error("Usuário ou senha incorretos.")

        st.markdown("""
        <div class="pf-login-foot">Vórtex v1.0 · 2026</div>
        """, unsafe_allow_html=True)

# ============================================================
# SIDEBAR
# ============================================================
def sidebar():
    with st.sidebar:
        st.markdown(f"""
        <div style="padding: 16px 8px 24px; border-bottom: 1px solid #2a3f5f; margin-bottom: 16px;">
            <div style="font-size:20px; font-weight:800; color:white; letter-spacing:-0.5px;">⚡ Vórtex</div>
            <div style="font-size:11px; color:#64748b; margin-top:2px;">Sistema de Planning</div>
        </div>
        """, unsafe_allow_html=True)

        # Info do usuário
        perfil_badge = "🔑 Admin" if st.session_state["perfil"] == "admin" else "👁️ Visualização"
        st.markdown(f"""
        <div style="padding:10px 8px; margin-bottom:16px; background:rgba(255,255,255,0.05); border-radius:8px;">
            <div style="font-size:13px; font-weight:600; color:white;">{st.session_state['nome']}</div>
            <div style="font-size:11px; color:#64748b; margin-top:2px;">{perfil_badge}</div>
        </div>
        """, unsafe_allow_html=True)

        # Status dos dados
        dados_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dados", "dados_sap.xlsx")
        if os.path.exists(dados_path):
            mod_time = datetime.fromtimestamp(os.path.getmtime(dados_path))
            st.markdown(f"""
            <div style="padding:8px; margin-bottom:16px; background:rgba(22,163,74,0.15); border-radius:8px; border:1px solid rgba(22,163,74,0.3);">
                <div style="font-size:11px; color:#4ade80; font-weight:600;">✓ Dados atualizados</div>
                <div style="font-size:10px; color:#64748b;">{mod_time.strftime('%d/%m/%Y %H:%M')}</div>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div style="padding:8px; margin-bottom:16px; background:rgba(239,68,68,0.15); border-radius:8px; border:1px solid rgba(239,68,68,0.3);">
                <div style="font-size:11px; color:#f87171; font-weight:600;">⚠ Dados não encontrados</div>
                <div style="font-size:10px; color:#64748b;">Execute atualizar_dados.py</div>
            </div>
            """, unsafe_allow_html=True)

        # Logout
        st.markdown("<div style='flex:1'></div>", unsafe_allow_html=True)
        if st.button("Sair", use_container_width=True):
            apagar_token()
            for key in ["logado","usuario","perfil","nome"]:
                st.session_state.pop(key, None)
            st.rerun()

# ============================================================
# PÁGINA HOME
# ============================================================
def pagina_home():
    st.markdown("""
    <div class="page-header">
        <span style="font-size:28px;">🏠</span>
        <div>
            <h1>Dashboard Principal</h1>
            <div class="subtitle">Visão geral do planejamento — atualizado diariamente às 18h</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    dados_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dados", "dados_sap.xlsx")

    # ── Barra de atualização de dados ─────────────────────────
    ba1, ba2, ba3 = st.columns([2, 2, 1])
    with ba1:
        if os.path.exists(dados_path):
            mt = datetime.fromtimestamp(os.path.getmtime(dados_path))
            st.markdown(f"**📅 Última atualização:** `{mt.strftime('%d/%m/%Y %H:%M')}`")
        else:
            st.markdown("**📅 Última atualização:** _nunca_")
    if st.session_state.get("perfil") == "admin":
        with ba3:
            if st.button("🔄 Atualizar Dados", use_container_width=True, key="btn_atualizar_home"):
                import subprocess, sys as _sys
                with st.spinner("Atualizando dados do SAP... (pode levar 1-2 min)"):
                    try:
                        r = subprocess.run(
                            [_sys.executable, os.path.join(os.path.dirname(os.path.abspath(__file__)), "atualizar_dados.py")],
                            capture_output=True, text=True, timeout=600,
                        )
                        if r.returncode == 0:
                            st.cache_data.clear()
                            st.success("✅ Dados atualizados!")
                            st.rerun()
                        else:
                            st.error(f"Falha na atualização:\n{r.stderr[-800:]}")
                    except Exception as e:
                        st.error(f"Erro: {e}")

    if not os.path.exists(dados_path):
        st.warning("⚠️ Arquivo de dados não encontrado. Clique em **Atualizar Dados** acima.")
        return

    try:
        df_cap  = pd.read_excel(dados_path, sheet_name="capacidade")
        df_mrp  = pd.read_excel(dados_path, sheet_name="mrp")
        df_hist = pd.read_excel(dados_path, sheet_name="historico")
    except Exception as e:
        st.error(f"Erro ao ler dados: {e}")
        return

    # KPIs
    total_pas     = len(df_cap)
    pas_criticos  = len(df_cap[df_cap.get("Peças Possíveis", pd.Series(dtype=int)) == 0]) if "Peças Possíveis" in df_cap.columns else 0
    ins_urgentes  = len(df_mrp[df_mrp.get("Urgência", pd.Series(dtype=str)).str.contains("URGENTE", na=False)]) if "Urgência" in df_mrp.columns else 0
    pas_com_cart  = len(df_cap[df_cap.get("Carteira", pd.Series(dtype=int)) > 0]) if "Carteira" in df_cap.columns else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("PAs Ativos", f"{total_pas:,}", help="Total de produtos acabados ativos")
    with c2:
        st.metric("PAs Críticos", f"{pas_criticos:,}", delta=f"-{pas_criticos} sem produção", delta_color="inverse")
    with c3:
        st.metric("Insumos Urgentes", f"{ins_urgentes:,}", delta="Comprar hoje", delta_color="inverse")
    with c4:
        st.metric("PAs com Carteira", f"{pas_com_cart:,}", help="PAs com pedidos em aberto")

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 🚨 PAs Críticos — Ação Imediata")
        if "Peças Possíveis" in df_cap.columns and "Carteira" in df_cap.columns:
            criticos = df_cap[
                (df_cap["Peças Possíveis"] == 0) & (df_cap["Carteira"] > 0)
            ].head(10)
            if len(criticos) > 0:
                cols_show = [c for c in ["Código PA","Descrição PA","Classe ABC","Carteira","Insumo Gargalo"] if c in criticos.columns]
                st.dataframe(criticos[cols_show], use_container_width=True, hide_index=True, height=380, key="grid_home_criticos")
            else:
                st.success("✅ Nenhum PA crítico com carteira em aberto!")
        else:
            st.info("Dados de capacidade não disponíveis.")

    with col2:
        st.markdown("### ⚙️ Insumos — Comprar Urgente")
        if "Urgência" in df_mrp.columns:
            urgentes = df_mrp[df_mrp["Urgência"].str.contains("URGENTE", na=False)].head(10)
            if len(urgentes) > 0:
                cols_show = [c for c in ["Insumo","Descrição","Estoque Atual","Nec. Líquida","Data Necessidade"] if c in urgentes.columns]
                st.dataframe(urgentes[cols_show], use_container_width=True, hide_index=True, height=380, key="grid_home_urgentes")
            else:
                st.success("✅ Nenhum insumo em situação urgente!")
        else:
            st.info("Dados de MRP não disponíveis.")

# ============================================================
# MAIN
# ============================================================
def main():
    if "logado" not in st.session_state:
        login()
        return

    render_sidebar()
    pagina_home()

if __name__ == "__main__":
    main()
