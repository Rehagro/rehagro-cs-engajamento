import streamlit as st
import base64


@st.cache_resource
def _get_store():
    return {}


def get_store():
    return _get_store()


def _logo_b64(path):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return ""


def require_login():
    """Mostra tela de login se não há usuário na sessão e retorna o nome."""
    if st.session_state.get('rh_usuario'):
        return st.session_state['rh_usuario']

    logo_str  = _logo_b64("Logo-Rehagro-branca-transp.png")
    logo_html = (
        f'<img src="data:image/png;base64,{logo_str}" '
        f'style="height:60px;opacity:0.92;display:block;margin:0 auto 14px"/>'
        if logo_str else ''
    )

    st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@700;800;900&family=Inter:wght@400;500;600&display=swap');
html, body, [class*="css"] { font-family:'Inter',sans-serif !important; background:#F2EDE4 !important; }
.stApp { background:#F2EDE4 !important; }
#MainMenu, footer, header { visibility:hidden; }
[data-testid="stAppViewContainer"] { padding-top:0 !important; }
[data-testid="block-container"]    { padding-top:0 !important; }
.rh-lc {
    background:#fff; border-radius:16px; overflow:hidden;
    box-shadow:0 8px 40px rgba(0,0,0,.10); margin:0 auto;
}
.rh-lh {
    background:#1B3D2A; padding:32px 36px 26px; text-align:center;
    position:relative; overflow:hidden;
}
.rh-lh-bg {
    position:absolute; inset:0; opacity:.05;
    background-image:repeating-linear-gradient(45deg,#fff 0,#fff 1px,transparent 0,transparent 50%);
    background-size:12px 12px;
}
.rh-le {
    color:#C8A532; font-family:'Montserrat',sans-serif; font-size:10px;
    font-weight:700; letter-spacing:2px; text-transform:uppercase;
    margin:0 0 6px; position:relative;
}
.rh-lt {
    font-family:'Montserrat',sans-serif; font-weight:900; font-size:20px;
    color:#fff; margin:0; position:relative; text-transform:uppercase; letter-spacing:1px;
}
.rh-lb { padding:24px 32px 8px; }
.rh-ld { font-size:13px; color:#777; line-height:1.6; margin-bottom:4px; }
[data-testid="stTextInput"] input {
    border-radius:8px !important; border:1.5px solid #ddd !important;
    font-size:13px !important; padding:10px 14px !important;
}
.stButton > button[kind="primary"] {
    font-family:'Montserrat',sans-serif !important; font-weight:700 !important;
    border-radius:8px !important; background:#1B3D2A !important;
    color:#fff !important; border:none !important;
    font-size:13px !important; box-shadow:0 4px 14px rgba(27,61,42,.3) !important;
}
</style>
""", unsafe_allow_html=True)

    _, col, _ = st.columns([1, 2, 1])
    with col:
        st.markdown(f"""
<div style="margin-top:8vh">
  <div class="rh-lc">
    <div class="rh-lh">
      <div class="rh-lh-bg"></div>
      {logo_html}
      <p class="rh-le">Customer Success · Rehagro</p>
      <p class="rh-lt">Agente de Engajamento</p>
    </div>
    <div class="rh-lb">
      <p class="rh-ld">Informe seu nome ou e-mail para acessar. Sua última análise ficará salva e pode ser retomada mesmo após fechar o navegador.</p>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)
        nome = st.text_input(
            "", placeholder="Seu nome ou e-mail…",
            key="rh_login_input", label_visibility="collapsed"
        )
        if st.button("Acessar →", type="primary", use_container_width=True, key="rh_login_btn"):
            if nome.strip():
                st.session_state['rh_usuario'] = nome.strip()
                st.rerun()
            else:
                st.error("Informe seu nome ou e-mail para continuar.")
    st.stop()
