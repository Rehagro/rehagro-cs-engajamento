import streamlit as st
import pandas as pd
import io

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Clash+Display:wght@500;600;700&family=Cabinet+Grotesk:wght@400;500;700;800&display=swap');
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Outfit:wght@300;400;500;600;700&display=swap');

:root {
  --verde:   #0F3D20;
  --verde2:  #1B5E35;
  --verde3:  #2E7D32;
  --ouro:    #C8A951;
  --ouro2:   #E8C96A;
  --creme:   #F4F0E6;
  --cinza:   #EDEAE0;
  --texto:   #111111;
  --sub:     #5A5A4A;
  --branco:  #FFFFFF;
}

*, *::before, *::after { box-sizing: border-box; }
html, body, [class*="css"] {
    font-family: 'Outfit', sans-serif !important;
    background: var(--creme) !important;
    color: var(--texto) !important;
}
.stApp { background: var(--creme) !important; }
#MainMenu, footer, header { visibility: hidden; }
[data-testid="stAppViewContainer"] { padding-top: 0 !important; }
[data-testid="block-container"] { padding-top: 0 !important; }
button[aria-label="Close sidebar"] { display: none !important; }
button[aria-label="Fechar barra lateral"] { display: none !important; }
[data-testid="stAppViewContainer"] { padding-top: 0 !important; }
[data-testid="block-container"] { padding-top: 0 !important; }

.rh-hero {
    background: var(--verde);
    padding: 18px 40px 20px;
    margin: -1rem -1rem 0 -1rem;
    position: relative;
    overflow: hidden;
    border-bottom: 3px solid var(--ouro);
}
.rh-hero-bg {
    position: absolute; inset: 0;
    background:
        radial-gradient(ellipse 80% 60% at 110% 50%, rgba(200,169,81,0.18) 0%, transparent 60%),
        radial-gradient(ellipse 40% 80% at -10% 50%, rgba(46,125,50,0.3) 0%, transparent 60%);
}
.rh-hero-grid {
    position: absolute; inset: 0; opacity: 0.04;
    background-image: linear-gradient(var(--ouro) 1px, transparent 1px),
                      linear-gradient(90deg, var(--ouro) 1px, transparent 1px);
    background-size: 40px 40px;
}
.rh-hero-inner { position: relative; z-index: 1; }
.rh-eyebrow {
    font-family: 'Outfit', sans-serif;
    font-size: 0.7rem; font-weight: 600;
    letter-spacing: 4px; text-transform: uppercase;
    color: var(--ouro); margin: 0 0 6px 0;
    display: flex; align-items: center; gap: 10px;
}
.rh-eyebrow::before {
    content: ''; display: inline-block;
    width: 32px; height: 2px; background: var(--ouro);
}
.rh-hero-title {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 2.2rem; line-height: 1;
    color: var(--branco); margin: 0 0 4px 0;
    letter-spacing: 2px; white-space: nowrap;
}
.rh-hero-title span { color: var(--ouro); }
.rh-hero-sub {
    font-size: 0.85rem; color: rgba(255,255,255,0.6);
    font-weight: 300; margin: 4px 0 0 0;
    line-height: 1.4; white-space: nowrap;
}
.rh-hero-pills { display: flex; gap: 6px; margin-top: 10px; flex-wrap: nowrap; }
.rh-pill {
    background: rgba(200,169,81,0.15);
    border: 1px solid rgba(200,169,81,0.35);
    color: var(--ouro2); font-size: 0.72rem;
    font-weight: 600; letter-spacing: 1px;
    text-transform: uppercase; padding: 5px 14px;
    border-radius: 100px;
}
.rh-body { padding: 20px 0 0 0; }
.rh-section {
    font-family: 'Outfit', sans-serif;
    font-size: 0.65rem; font-weight: 700;
    letter-spacing: 3.5px; text-transform: uppercase;
    color: var(--ouro); margin: 0 0 10px 0;
    display: flex; align-items: center; gap: 10px;
}
.rh-section::after {
    content: ''; flex: 1; height: 1px;
    background: linear-gradient(90deg, rgba(200,169,81,0.4), transparent);
}
.rh-divider {
    height: 1px;
    background: linear-gradient(90deg, var(--ouro), transparent);
    margin: 12px 0; opacity: 0.35;
}
[data-testid="stFileUploader"] label,
[data-testid="stFileUploader"] > div > label {
    font-size: 1rem !important; font-weight: 600 !important;
    color: var(--verde) !important; margin-bottom: 6px !important;
}
[data-testid="stFileUploader"] {
    background: var(--branco) !important;
    border-radius: 10px !important;
    border: 1.5px dashed rgba(15,61,32,0.2) !important;
    padding: 6px 10px !important; margin-bottom: 12px !important;
}
[data-testid="stTextInput"] label {
    font-size: 1rem !important; font-weight: 600 !important;
    color: var(--verde) !important;
}
[data-testid="stTextInput"] input {
    border-radius: 8px !important;
    border: 1.5px solid rgba(15,61,32,0.2) !important;
    font-size: 0.95rem !important;
}
.stDownloadButton > button {
    background: transparent !important;
    color: var(--verde) !important;
    border: 2px solid var(--verde) !important;
    border-radius: 8px !important;
    font-family: 'Outfit', sans-serif !important;
    font-size: 0.95rem !important; font-weight: 600 !important;
    padding: 12px 24px !important;
}
.stDownloadButton > button:hover {
    background: var(--verde) !important;
    color: var(--branco) !important;
}
[data-testid="stDataFrame"] {
    border-radius: 10px !important;
    overflow: hidden !important;
    border: 1px solid var(--cinza) !important;
}
[data-testid="stAlert"] { border-radius: 8px !important; font-size: 0.9rem !important; }

/* progresso de vistos */
.rh-progresso-card {
    background: var(--branco);
    border-radius: 12px;
    padding: 16px 20px;
    border: 1px solid var(--cinza);
    border-left: 4px solid var(--ouro);
}
.rh-progresso-titulo {
    font-size: 0.65rem; font-weight: 700;
    letter-spacing: 2px; text-transform: uppercase;
    color: var(--sub); margin-bottom: 8px;
}
.rh-progresso-nums {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 1.8rem; color: var(--verde);
    line-height: 1;
}
.rh-progresso-sub { font-size: 0.8rem; color: var(--sub); margin-top: 2px; }

/* card de aluno selecionado */
.rh-aluno-card {
    background: var(--verde);
    border-radius: 12px;
    padding: 12px 20px;
    margin-bottom: 12px;
    border-bottom: 3px solid var(--ouro);
}
.rh-aluno-nome {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 1.8rem; color: var(--branco);
    letter-spacing: 2px; margin: 0;
}
.rh-aluno-badge {
    display: inline-block;
    background: rgba(200,169,81,0.2);
    border: 1px solid rgba(200,169,81,0.5);
    color: var(--ouro2);
    font-size: 0.72rem; font-weight: 600;
    letter-spacing: 1.5px; text-transform: uppercase;
    padding: 3px 10px; border-radius: 100px; margin-top: 6px;
}

/* dashboard block */
.rh-dash-bloco {
    background: var(--branco);
    border-radius: 12px;
    padding: 12px 16px;
    margin-bottom: 10px;
    border: 1px solid var(--cinza);
    border-top: 3px solid var(--ouro);
}
.rh-dash-bloco-titulo {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 1.1rem; letter-spacing: 2px;
    color: var(--verde); margin-bottom: 8px;
    display: flex; align-items: center; gap: 10px;
}

/* info cards canvas */
.rh-info-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: 8px; margin-bottom: 4px;
}
.rh-info-item {
    background: var(--creme);
    border-radius: 8px;
    padding: 8px 12px;
    border: 1px solid var(--cinza);
}
.rh-info-label {
    font-size: 0.62rem; font-weight: 700;
    letter-spacing: 2px; text-transform: uppercase;
    color: var(--sub); margin-bottom: 4px;
}
.rh-info-valor {
    font-size: 1rem; font-weight: 600;
    color: var(--verde);
}
.rh-info-valor.alerta { color: #C62828; }
.rh-info-valor.ok     { color: var(--verde3); }

.rh-footer {
    display: flex; align-items: center; justify-content: center;
    gap: 16px; padding: 32px 0 16px;
    border-top: 1px solid var(--cinza); margin-top: 48px;
    color: var(--sub); font-size: 0.8rem;
}
.rh-footer-dot {
    width: 4px; height: 4px; border-radius: 50%;
    background: var(--ouro); display: inline-block;
}

/* orientação de navegação no BI */
.rh-nav-hint {
    display: flex; align-items: flex-start; gap: 8px;
    background: rgba(15,61,32,0.05);
    border: 1px solid rgba(15,61,32,0.12);
    border-left: 3px solid var(--verde);
    border-radius: 6px;
    padding: 7px 10px;
    margin: 0 0 6px 0;
}
.rh-nav-icon { font-size: 0.85rem; flex-shrink: 0; margin-top: 1px; }
.rh-nav-body { display: flex; flex-direction: column; gap: 2px; }
.rh-nav-bi {
    font-size: 0.62rem; font-weight: 700;
    letter-spacing: 1.5px; text-transform: uppercase;
    color: var(--verde);
}
.rh-nav-page {
    font-size: 0.75rem; font-weight: 500;
    color: var(--texto); line-height: 1.3;
}
.rh-nav-dash {
    font-size: 0.68rem; color: var(--sub);
    font-style: italic; line-height: 1.3;
}

/* filtros de exportação */
.rh-filtros-bloco {
    background: rgba(15,61,32,0.04);
    border: 1px solid rgba(15,61,32,0.12);
    border-left: 3px solid var(--ouro);
    border-radius: 8px;
    padding: 10px 14px;
    margin: 6px 0 12px 0;
}
.rh-filtros-titulo {
    font-size: 0.6rem; font-weight: 700;
    letter-spacing: 2.5px; text-transform: uppercase;
    color: var(--sub); margin-bottom: 8px;
    display: flex; align-items: center; gap: 6px;
}
.rh-filtros-titulo::before { content: '🔍'; font-size: 0.75rem; }
.rh-filtros-chips { display: flex; flex-wrap: wrap; gap: 6px; }
.rh-filtro-chip {
    display: inline-flex; align-items: center; gap: 0;
    border-radius: 6px; overflow: hidden;
    border: 1px solid rgba(15,61,32,0.15);
    font-size: 0.72rem; font-weight: 500;
}
.rh-filtro-label {
    background: var(--verde); color: rgba(255,255,255,0.85);
    padding: 3px 8px; font-weight: 600; font-size: 0.68rem;
    letter-spacing: 0.5px;
}
.rh-filtro-val {
    background: var(--branco); color: var(--verde);
    padding: 3px 10px; font-weight: 600;
}
.rh-filtro-chip.variavel .rh-filtro-val {
    background: rgba(200,169,81,0.12);
    color: var(--sub); font-style: italic;
}
</style>
""", unsafe_allow_html=True)

# ── Constantes ────────────────────────────────────────────
_MESES = {
    'janeiro': 1, 'fevereiro': 2, 'março': 3, 'abril': 4,
    'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8,
    'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12,
}

# ── Session state ─────────────────────────────────────────
if 'ca_vistos' not in st.session_state:
    st.session_state.ca_vistos = set()

# ── Helpers ───────────────────────────────────────────────
def _norm_disc(series):
    return series.astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)

def _montar_data_nps(row):
    try:
        ano = int(row['Data da aula - Ano'])
        mes = _MESES.get(str(row['Data da aula - Mês']).strip().lower())
        dia = int(row['Data da aula - Dia'])
        if mes:
            return pd.Timestamp(year=ano, month=mes, day=dia)
    except Exception:
        pass
    return pd.NaT

def _fmt_data(ts):
    try:
        return pd.Timestamp(ts).strftime('%d/%m/%Y')
    except Exception:
        return ''

# ── Funções de carga ──────────────────────────────────────
@st.cache_data(show_spinner=False)
def carregar_canvas(arquivo_bytes):
    df = pd.read_excel(io.BytesIO(arquivo_bytes), skiprows=2)
    df.columns = df.columns.str.strip()
    return df

@st.cache_data(show_spinner=False)
def carregar_status_modulo(arquivo_bytes):
    df = pd.read_excel(io.BytesIO(arquivo_bytes), skiprows=2)
    df.columns = df.columns.str.strip()
    df['DISCIPLINA'] = _norm_disc(df['DISCIPLINA'])
    df = df[df['STATUS'].isin(['1 - Em Andamento', '3 - Finalizado'])].copy()
    excluir_pattern = 'Objetivos de Aprendizagem|Gravações de aula online ao vivo'
    df = df[~df['MÓDULO'].str.contains(excluir_pattern, case=False, na=False)].copy()
    return df

@st.cache_data(show_spinner=False)
def carregar_envio_tarefas(arquivo_bytes):
    df = pd.read_excel(io.BytesIO(arquivo_bytes), skiprows=2)
    df.columns = df.columns.str.strip()
    df['DISCIPLINA'] = _norm_disc(df['DISCIPLINA'])
    return df

@st.cache_data(show_spinner=False)
def carregar_nps(arquivo_bytes):
    df = pd.read_excel(io.BytesIO(arquivo_bytes), skiprows=2)
    df.columns = df.columns.str.strip()
    col_item = next((c for c in df.columns if c.upper() == 'ITEM'), None)
    if col_item:
        df = df[df[col_item].astype(str).str.contains('0 a 10', case=False, na=False)].copy()
    df['_aluno_norm'] = df['Aluno'].astype(str).str.strip().str.lower()
    df['_prof_norm'] = df['Professor'].astype(str).str.strip().str.lower()
    df['_topico_norm'] = df['Tópico'].astype(str).str.strip().str.lower()
    return df

@st.cache_data(show_spinner=False)
def carregar_comentarios(arquivo_bytes):
    df = pd.read_excel(io.BytesIO(arquivo_bytes), skiprows=2)
    df.columns = df.columns.str.strip()
    df['_aluno_norm'] = df['Aluno'].astype(str).str.strip().str.lower()
    df['_prof_norm'] = df['Professor'].astype(str).str.strip().str.lower()
    df['_topico_norm'] = df['Tópico'].astype(str).str.strip().str.lower()
    return df

# ── Hero ──────────────────────────────────────────────────
import base64, os

def _logo_b64(path):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return ""

_logo_hero = _logo_b64("Logo-Rehagro-chapada-branca.png")
_logo_img  = f'<img src="data:image/png;base64,{_logo_hero}" style="height:76px;opacity:0.92;" />' if _logo_hero else ""

st.markdown(f"""
<div class="rh-hero">
  <div class="rh-hero-bg"></div>
  <div class="rh-hero-grid"></div>
  <div class="rh-hero-inner" style="display:flex;justify-content:space-between;align-items:center;">
    <div>
      <p class="rh-eyebrow">Rehagro · Customer Success</p>
      <h1 class="rh-hero-title">COMPORTAMENTO <span>DO ALUNO</span></h1>
      <p class="rh-hero-sub">Panorama completo do aluno para preparar o contato proativo do CS.</p>
      <div class="rh-hero-pills">
        <span class="rh-pill">Acesso Canvas</span>
        <span class="rh-pill">Módulos</span>
        <span class="rh-pill">NPS</span>
        <span class="rh-pill">Comentários</span>
      </div>
    </div>
    <div style="flex-shrink:0;padding-left:24px;">{_logo_img}</div>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="rh-body">', unsafe_allow_html=True)

if st.button("← Monitoramento de Alunos", key="nav_mon"):
    st.switch_page("pages/monitoramento.py")

# ── Upload ────────────────────────────────────────────────
st.markdown('<p class="rh-section">Carregar arquivos</p>', unsafe_allow_html=True)

col_u1, col_u2, col_u3 = st.columns(3)
with col_u1:
    st.markdown("""
    <div class="rh-nav-hint">
        <div class="rh-nav-icon">📍</div>
        <div class="rh-nav-body">
            <span class="rh-nav-bi">BI Rehagro Canvas</span>
            <span class="rh-nav-page">Resumo de dados dos alunos (Almas e turmas)</span>
            <span class="rh-nav-dash">Avaliar: Acesso ao Canvas</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    f_canvas  = st.file_uploader("Acesso ao Canvas",   type=["xlsx"], key="ca_canvas")
    st.markdown("""
    <div class="rh-filtros-bloco">
        <div class="rh-filtros-titulo">Filtros antes de exportar</div>
        <div class="rh-filtros-chips">
            <span class="rh-filtro-chip">
                <span class="rh-filtro-label">Status do aluno</span>
                <span class="rh-filtro-val">Ativo</span>
            </span>
            <span class="rh-filtro-chip">
                <span class="rh-filtro-label">Função na disciplina</span>
                <span class="rh-filtro-val">Aluno</span>
            </span>
            <span class="rh-filtro-chip variavel">
                <span class="rh-filtro-label">Curso</span>
                <span class="rh-filtro-val">seu curso</span>
            </span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("""
    <div class="rh-nav-hint">
        <div class="rh-nav-icon">📍</div>
        <div class="rh-nav-body">
            <span class="rh-nav-bi">BI Rehagro Canvas</span>
            <span class="rh-nav-page">Resumo de dados dos alunos (Almas e turmas)</span>
            <span class="rh-nav-dash">Avaliar: Qual o status do módulo para cada aluno x disciplina</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    f_status  = st.file_uploader("Status dos Módulos", type=["xlsx"], key="ca_status")
    st.markdown("""
    <div class="rh-filtros-bloco">
        <div class="rh-filtros-titulo">Filtros antes de exportar</div>
        <div class="rh-filtros-chips">
            <span class="rh-filtro-chip">
                <span class="rh-filtro-label">Status do aluno</span>
                <span class="rh-filtro-val">Ativo</span>
            </span>
            <span class="rh-filtro-chip">
                <span class="rh-filtro-label">Função na disciplina</span>
                <span class="rh-filtro-val">Aluno</span>
            </span>
            <span class="rh-filtro-chip variavel">
                <span class="rh-filtro-label">Curso</span>
                <span class="rh-filtro-val">seu curso</span>
            </span>
        </div>
    </div>
    """, unsafe_allow_html=True)
with col_u2:
    st.markdown("""
    <div class="rh-nav-hint">
        <div class="rh-nav-icon">📍</div>
        <div class="rh-nav-body">
            <span class="rh-nav-bi">BI Rehagro Canvas</span>
            <span class="rh-nav-page">Resumo de dados dos alunos (Almas e turmas)</span>
            <span class="rh-nav-dash">Avaliar: Entrega de atividades</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    f_tarefas = st.file_uploader("Envio de Tarefas",   type=["xlsx"], key="ca_tarefas")
    st.markdown("""
    <div class="rh-filtros-bloco">
        <div class="rh-filtros-titulo">Filtros antes de exportar</div>
        <div class="rh-filtros-chips">
            <span class="rh-filtro-chip">
                <span class="rh-filtro-label">Status do aluno</span>
                <span class="rh-filtro-val">Ativo</span>
            </span>
            <span class="rh-filtro-chip">
                <span class="rh-filtro-label">Função na disciplina</span>
                <span class="rh-filtro-val">Aluno</span>
            </span>
            <span class="rh-filtro-chip variavel">
                <span class="rh-filtro-label">Curso</span>
                <span class="rh-filtro-val">seu curso</span>
            </span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("""
    <div class="rh-nav-hint">
        <div class="rh-nav-icon">📍</div>
        <div class="rh-nav-body">
            <span class="rh-nav-bi">BI Rehagro Educação — Avaliação de Aula</span>
            <span class="rh-nav-page">Tabela respostas objetivas</span>
            <span class="rh-nav-dash">Avaliação de aula/etapa | Detalhamento</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    f_nps     = st.file_uploader("Avaliação de Aula/Etapa", type=["xlsx"], key="ca_nps")
    st.markdown("""
    <div class="rh-filtros-bloco">
        <div class="rh-filtros-titulo">Filtros antes de exportar</div>
        <div class="rh-filtros-chips">
            <span class="rh-filtro-chip">
                <span class="rh-filtro-label">Item</span>
                <span class="rh-filtro-val">De 0 a 10 · Nota da aula</span>
            </span>
            <span class="rh-filtro-chip">
                <span class="rh-filtro-label">Status de matrícula</span>
                <span class="rh-filtro-val">Matriculado</span>
            </span>
            <span class="rh-filtro-chip">
                <span class="rh-filtro-label">Tipo de aula</span>
                <span class="rh-filtro-val">Gravada</span>
            </span>
            <span class="rh-filtro-chip variavel">
                <span class="rh-filtro-label">Curso</span>
                <span class="rh-filtro-val">seu curso</span>
            </span>
        </div>
    </div>
    """, unsafe_allow_html=True)
with col_u3:
    st.markdown("""
    <div class="rh-nav-hint">
        <div class="rh-nav-icon">📍</div>
        <div class="rh-nav-body">
            <span class="rh-nav-bi">BI Rehagro Educação — Avaliação de Aula</span>
            <span class="rh-nav-page">Tabela Comentários</span>
            <span class="rh-nav-dash">Relatório do Professor | Comentários</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    f_coment  = st.file_uploader("Comentários", type=["xlsx"], key="ca_coment")
    st.markdown("""
    <div class="rh-filtros-bloco">
        <div class="rh-filtros-titulo">Filtros antes de exportar</div>
        <div class="rh-filtros-chips">
            <span class="rh-filtro-chip">
                <span class="rh-filtro-label">Status de matrícula</span>
                <span class="rh-filtro-val">Matriculado</span>
            </span>
            <span class="rh-filtro-chip">
                <span class="rh-filtro-label">Tipo de aula</span>
                <span class="rh-filtro-val">Gravada</span>
            </span>
            <span class="rh-filtro-chip variavel">
                <span class="rh-filtro-label">Curso</span>
                <span class="rh-filtro-val">seu curso</span>
            </span>
        </div>
    </div>
    """, unsafe_allow_html=True)

if not f_canvas:
    st.info("Carregue pelo menos o arquivo **Acesso ao Canvas** para liberar o filtro de aluno.")
    st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# ── Carregar dados ────────────────────────────────────────
with st.spinner("Lendo arquivos..."):
    df_canvas  = carregar_canvas(f_canvas.read())
    df_status  = carregar_status_modulo(f_status.read())   if f_status  else None
    df_tarefas = carregar_envio_tarefas(f_tarefas.read())  if f_tarefas else None
    df_nps_raw = carregar_nps(f_nps.read())                if f_nps     else None
    df_coment  = carregar_comentarios(f_coment.read())     if f_coment  else None

col_nome = next((c for c in df_canvas.columns if 'NOME' in c.upper()), None)
if not col_nome:
    st.error("Coluna de nome não encontrada no arquivo Canvas.")
    st.stop()

nomes_todos = sorted(
    df_canvas[col_nome].dropna().astype(str).str.strip().unique().tolist()
)

# ── Seleção de aluno com rastreamento de vistos ───────────
st.markdown('<div class="rh-divider"></div>', unsafe_allow_html=True)
st.markdown('<p class="rh-section">Selecionar aluno</p>', unsafe_allow_html=True)

col_sel, col_prog = st.columns([3, 1], gap="large")

with col_sel:
    col_busca_row, col_filtro_row = st.columns([2, 1])
    with col_busca_row:
        busca = st.text_input("Buscar por nome:", placeholder="Digite parte do nome...", key="ca_busca")
    with col_filtro_row:
        so_nao_vistos = st.checkbox(
            "Mostrar apenas não vistos",
            value=False,
            key="ca_filtro_vistos",
            help="Oculta alunos que já foram visualizados nesta sessão."
        )

    nomes_filtrados = nomes_todos
    if busca.strip():
        nomes_filtrados = [n for n in nomes_filtrados if busca.strip().lower() in n.lower()]
    if so_nao_vistos:
        nomes_filtrados = [n for n in nomes_filtrados if n not in st.session_state.ca_vistos]

    if not nomes_filtrados:
        if so_nao_vistos:
            st.success("✅ Todos os alunos visíveis já foram visualizados nesta sessão.")
        else:
            st.warning("Nenhum aluno encontrado para esta busca.")
        st.markdown('</div>', unsafe_allow_html=True)
        st.stop()

    def _fmt_nome(nome):
        if nome in st.session_state.ca_vistos:
            return f"✅  {nome}"
        return f"○  {nome}"

    aluno_sel = st.selectbox(
        "Selecionar aluno:",
        nomes_filtrados,
        format_func=_fmt_nome,
        key="ca_aluno_sel",
    )

with col_prog:
    n_vistos = len(st.session_state.ca_vistos)
    n_total  = len(nomes_todos)
    pct      = n_vistos / n_total if n_total > 0 else 0

    st.markdown(f"""
    <div class="rh-progresso-card">
        <div class="rh-progresso-titulo">Progresso da sessão</div>
        <div class="rh-progresso-nums">{n_vistos}<span style="font-size:1rem;color:var(--sub);font-family:'Outfit',sans-serif;font-weight:400;"> / {n_total}</span></div>
        <div class="rh-progresso-sub">alunos visualizados</div>
    </div>
    """, unsafe_allow_html=True)
    st.progress(pct)

    if st.button("↺  Reiniciar progresso", use_container_width=True, key="ca_limpar_vistos"):
        st.session_state.ca_vistos = set()
        st.rerun()

# Marcar como visto automaticamente ao visualizar
if aluno_sel and aluno_sel not in st.session_state.ca_vistos:
    st.session_state.ca_vistos.add(aluno_sel)
    st.rerun()

# ── Card de identificação do aluno ────────────────────────
st.markdown('<div class="rh-divider"></div>', unsafe_allow_html=True)
st.markdown(f"""
<div class="rh-aluno-card">
    <div style="font-size:0.65rem;font-weight:700;letter-spacing:3px;text-transform:uppercase;
                color:rgba(200,169,81,0.7);margin-bottom:4px;">Aluno selecionado</div>
    <div class="rh-aluno-nome">{aluno_sel}</div>
    <span class="rh-aluno-badge">Comportamento completo</span>
</div>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# DASHBOARD 1 — ACESSO AO CANVAS
# ════════════════════════════════════════════════════════════
st.markdown('<p class="rh-section">Dashboard 1 · Acesso ao Canvas</p>', unsafe_allow_html=True)

linha_canvas = df_canvas[df_canvas[col_nome].astype(str).str.strip() == aluno_sel]

if linha_canvas.empty:
    st.warning("Aluno não encontrado no arquivo de Acesso ao Canvas.")
else:
    row = linha_canvas.iloc[0]

    col_ultima = next((c for c in df_canvas.columns if 'LTIMA' in c.upper()), None)
    col_dias   = next((c for c in df_canvas.columns if 'DIAS' in c.upper()), None)
    col_email  = next((c for c in df_canvas.columns if 'MAIL' in c.upper()), None)
    col_curso  = next((c for c in df_canvas.columns if c.upper() == 'CURSO'), None)
    col_turma  = next((c for c in df_canvas.columns if c.upper() == 'TURMA'), None)

    ultima_raw = row.get(col_ultima) if col_ultima else None
    if pd.isna(ultima_raw) if ultima_raw is not None else True:
        ultimo_acesso = "Nunca acessou"
        cls_acesso    = "alerta"
    else:
        ultimo_acesso = _fmt_data(ultima_raw)
        cls_acesso    = "ok"

    dias_raw = row.get(col_dias) if col_dias else None
    dias_val = int(dias_raw) if pd.notna(dias_raw) else None
    dias_str = f"{dias_val} dias" if dias_val is not None else "—"
    cls_dias = "alerta" if (dias_val or 0) >= 30 else ("ok" if (dias_val or 0) <= 7 else "")

    email_val = str(row.get(col_email, '')).strip() if col_email else '—'
    curso_val = str(row.get(col_curso, '')).strip() if col_curso else '—'
    turma_val = str(row.get(col_turma, '')).strip() if col_turma else '—'

    st.markdown(f"""
    <div class="rh-dash-bloco">
        <div class="rh-dash-bloco-titulo">📅 Acesso ao Canvas</div>
        <div class="rh-info-grid">
            <div class="rh-info-item">
                <div class="rh-info-label">Curso</div>
                <div class="rh-info-valor">{curso_val}</div>
            </div>
            <div class="rh-info-item">
                <div class="rh-info-label">Turma</div>
                <div class="rh-info-valor">{turma_val}</div>
            </div>
            <div class="rh-info-item">
                <div class="rh-info-label">E-mail</div>
                <div class="rh-info-valor" style="font-size:0.85rem">{email_val}</div>
            </div>
            <div class="rh-info-item">
                <div class="rh-info-label">Último acesso</div>
                <div class="rh-info-valor {cls_acesso}">{ultimo_acesso}</div>
            </div>
            <div class="rh-info-item">
                <div class="rh-info-label">Dias sem acesso</div>
                <div class="rh-info-valor {cls_dias}">{dias_str}</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# DASHBOARD 2 — MÓDULOS E ATIVIDADES
# ════════════════════════════════════════════════════════════
st.markdown('<p class="rh-section">Dashboard 2 · Módulos e Atividades</p>', unsafe_allow_html=True)

if df_status is None and df_tarefas is None:
    st.info("Arquivo não carregado")
else:
    # Filtrar status por aluno
    if df_status is not None:
        col_aluno_mod = next((c for c in df_status.columns if c.upper() == 'ALUNO'), None)
        df_mod_aluno  = df_status[
            df_status[col_aluno_mod].astype(str).str.strip() == aluno_sel
        ].copy() if col_aluno_mod else pd.DataFrame()
    else:
        df_mod_aluno = pd.DataFrame()

    # Filtrar tarefas por aluno
    if df_tarefas is not None:
        col_aluno_tar = next((c for c in df_tarefas.columns if c.upper() == 'ALUNO'), None)
        df_tar_aluno  = df_tarefas[
            df_tarefas[col_aluno_tar].astype(str).str.strip() == aluno_sel
        ].copy() if col_aluno_tar else pd.DataFrame()
    else:
        df_tar_aluno = pd.DataFrame()

    if df_mod_aluno.empty and df_tar_aluno.empty:
        st.info("Nenhum módulo ou atividade registrado para este aluno.")
    else:
        # Extrair notas e pontualidade por disciplina
        if not df_tar_aluno.empty:
            col_grupo = next((c for c in df_tar_aluno.columns if 'GRUPO' in c.upper()), None)
            col_nota  = next((c for c in df_tar_aluno.columns if 'NOTA'  in c.upper()), None)
            col_pont  = next((c for c in df_tar_aluno.columns if 'PONTUALIDADE' in c.upper()), None)
            col_disc_tar = 'DISCIPLINA'

            if col_grupo and col_nota:
                prat_sel = [col_disc_tar, col_nota] + ([col_pont] if col_pont else [])
                prat_rename = {col_nota: 'Nota Atividade Prática'}
                if col_pont:
                    prat_rename[col_pont] = 'Pontualidade'
                prat = (
                    df_tar_aluno[df_tar_aluno[col_grupo] == 'Atividade Pratica']
                    [prat_sel]
                    .rename(columns=prat_rename)
                    .drop_duplicates(subset=[col_disc_tar])
                )
                teste = (
                    df_tar_aluno[df_tar_aluno[col_grupo] == 'Teste seu conhecimento']
                    [[col_disc_tar, col_nota]]
                    .rename(columns={col_nota: 'Nota Teste seu Conhecimento'})
                    .drop_duplicates(subset=[col_disc_tar])
                )
            else:
                prat  = pd.DataFrame(columns=['DISCIPLINA', 'Nota Atividade Prática'])
                teste = pd.DataFrame(columns=['DISCIPLINA', 'Nota Teste seu Conhecimento'])
        else:
            prat  = pd.DataFrame(columns=['DISCIPLINA', 'Nota Atividade Prática'])
            teste = pd.DataFrame(columns=['DISCIPLINA', 'Nota Teste seu Conhecimento'])

        if not df_mod_aluno.empty:
            tabela = df_mod_aluno[['ALUNO', 'DISCIPLINA', 'STATUS']].copy()
        else:
            # Só tem tarefas — montar disciplinas únicas
            discips = pd.concat([prat[['DISCIPLINA']], teste[['DISCIPLINA']]]).drop_duplicates()
            tabela  = discips.copy()
            tabela.insert(0, 'ALUNO', aluno_sel)
            tabela['STATUS'] = 'Sem módulo registrado'

        tabela = tabela.merge(prat,  on='DISCIPLINA', how='left')
        tabela = tabela.merge(teste, on='DISCIPLINA', how='left')

        # Formatar notas: NaN → ''
        for col_n in ['Nota Atividade Prática', 'Nota Teste seu Conhecimento']:
            if col_n in tabela.columns:
                tabela[col_n] = tabela[col_n].apply(
                    lambda x: '' if pd.isna(x) else (int(x) if float(x) == int(float(x)) else x)
                )
            else:
                tabela[col_n] = ''
        if 'Pontualidade' not in tabela.columns:
            tabela['Pontualidade'] = ''
        else:
            tabela['Pontualidade'] = tabela['Pontualidade'].fillna('').astype(str).str.strip()

        display_cols = ['ALUNO', 'DISCIPLINA', 'STATUS', 'Nota Atividade Prática', 'Pontualidade', 'Nota Teste seu Conhecimento']
        display_cols = [c for c in display_cols if c in tabela.columns]

        st.markdown('<div class="rh-dash-bloco"><div class="rh-dash-bloco-titulo">📚 Módulos e Atividades</div>', unsafe_allow_html=True)
        st.dataframe(
            tabela[display_cols],
            use_container_width=True,
            hide_index=True,
            height=200,
        )
        st.markdown('</div>', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# DASHBOARD 3 — AVALIAÇÃO E COMENTÁRIOS
# ════════════════════════════════════════════════════════════
st.markdown('<p class="rh-section">Dashboard 3 · Avaliação de Aula/Etapa e Comentários</p>', unsafe_allow_html=True)

if df_nps_raw is None and df_coment is None:
    st.info("Arquivo não carregado")
else:
    aluno_norm = aluno_sel.strip().lower()

    aval_aluno   = df_nps_raw[df_nps_raw['_aluno_norm'] == aluno_norm].copy() if df_nps_raw is not None else pd.DataFrame()
    coment_aluno = df_coment[df_coment['_aluno_norm']   == aluno_norm].copy() if df_coment  is not None else pd.DataFrame()

    if aval_aluno.empty and coment_aluno.empty:
        st.info("Sem avaliações registradas")
    else:
        # Preparar base avaliação
        if not aval_aluno.empty:
            base = aval_aluno[['_aluno_norm', '_prof_norm', '_topico_norm',
                                'Turma', 'Disciplina', 'Tópico', 'Professor', 'Resposta']].copy()
            base = base.rename(columns={'Resposta': 'Nota'})
        else:
            base = pd.DataFrame(columns=['_aluno_norm', '_prof_norm', '_topico_norm',
                                         'Turma', 'Disciplina', 'Tópico', 'Professor', 'Nota'])

        # Preparar comentários (traz Turma/Tópico/Professor para preencher linhas sem avaliação)
        if not coment_aluno.empty:
            comt = coment_aluno[['_aluno_norm', '_prof_norm', '_topico_norm',
                                  'Turma', 'Tópico', 'Professor', 'Resposta']].copy()
            comt = comt.rename(columns={
                'Resposta':  'Comentário',
                'Turma':     '_turma_c',
                'Tópico':    '_topico_c',
                'Professor': '_prof_c',
            })
        else:
            comt = pd.DataFrame(columns=['_aluno_norm', '_prof_norm', '_topico_norm',
                                         '_turma_c', '_topico_c', '_prof_c', 'Comentário'])

        # Join triplo: Aluno + Tópico + Professor
        merged = base.merge(comt, on=['_aluno_norm', '_prof_norm', '_topico_norm'], how='outer')

        merged['Turma']      = merged['Turma'].fillna(merged.get('_turma_c', '')).fillna('').astype(str).str.strip()
        merged['Tópico']     = merged['Tópico'].fillna(merged.get('_topico_c', '')).fillna('').astype(str).str.strip()
        merged['Professor']  = merged['Professor'].fillna(merged.get('_prof_c', '')).fillna('').astype(str).str.strip()
        merged['Disciplina'] = merged.get('Disciplina', pd.Series(dtype=str)).fillna('').astype(str).str.strip()
        merged['Nota']       = merged.get('Nota', pd.Series(dtype=str)).fillna('').astype(str).str.strip()
        merged['Comentário'] = merged.get('Comentário', pd.Series(dtype=str)).fillna('').astype(str).str.strip()
        merged['Aluno']      = aluno_sel

        tabela_nps = merged[['Turma', 'Aluno', 'Disciplina', 'Tópico', 'Professor', 'Nota', 'Comentário']].reset_index(drop=True)

        st.markdown('<div class="rh-dash-bloco"><div class="rh-dash-bloco-titulo">⭐ Avaliação e Comentários</div>', unsafe_allow_html=True)
        st.dataframe(tabela_nps, use_container_width=True, hide_index=True, height=250)
        st.markdown('</div>', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# EXPORTAÇÃO
# ════════════════════════════════════════════════════════════
st.markdown('<div class="rh-divider"></div>', unsafe_allow_html=True)

def gerar_excel_aluno(aluno_nome):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        # Aba 1: Acesso ao Canvas
        linha_exp = df_canvas[df_canvas[col_nome].astype(str).str.strip() == aluno_nome]
        if not linha_exp.empty:
            linha_exp.to_excel(writer, sheet_name='Acesso ao Canvas', index=False)
        else:
            pd.DataFrame([{'Info': 'Sem dados de Canvas para este aluno'}]).to_excel(
                writer, sheet_name='Acesso ao Canvas', index=False
            )

        # Aba 2: Módulos e Atividades
        try:
            if df_status is not None:
                col_a = next((c for c in df_status.columns if c.upper() == 'ALUNO'), None)
                mod_exp = df_status[df_status[col_a].astype(str).str.strip() == aluno_nome].copy() if col_a else pd.DataFrame()
            else:
                mod_exp = pd.DataFrame()

            if df_tarefas is not None:
                col_a2   = next((c for c in df_tarefas.columns if c.upper() == 'ALUNO'), None)
                tar_exp  = df_tarefas[df_tarefas[col_a2].astype(str).str.strip() == aluno_nome].copy() if col_a2 else pd.DataFrame()
                col_g    = next((c for c in tar_exp.columns if 'GRUPO' in c.upper()), None)
                col_nt   = next((c for c in tar_exp.columns if 'NOTA'  in c.upper()), None)
                if col_g and col_nt and not tar_exp.empty:
                    prat_e  = tar_exp[tar_exp[col_g] == 'Atividade Pratica'][['DISCIPLINA', col_nt]].rename(columns={col_nt: 'Nota Atividade Prática'}).drop_duplicates('DISCIPLINA')
                    teste_e = tar_exp[tar_exp[col_g] == 'Teste seu conhecimento'][['DISCIPLINA', col_nt]].rename(columns={col_nt: 'Nota Teste seu Conhecimento'}).drop_duplicates('DISCIPLINA')
                    if not mod_exp.empty:
                        mod_exp = mod_exp.merge(prat_e, on='DISCIPLINA', how='left').merge(teste_e, on='DISCIPLINA', how='left')

            if mod_exp.empty:
                mod_exp = pd.DataFrame([{'Info': 'Sem dados de módulos para este aluno'}])
            mod_exp.to_excel(writer, sheet_name='Módulos e Atividades', index=False)
        except Exception:
            pd.DataFrame([{'Info': 'Erro ao gerar dados de módulos'}]).to_excel(
                writer, sheet_name='Módulos e Atividades', index=False
            )

        # Aba 3: Avaliação e Comentários
        try:
            a_norm = aluno_nome.strip().lower()
            aval_e = df_nps_raw[df_nps_raw['_aluno_norm'] == a_norm].copy() if df_nps_raw is not None else pd.DataFrame()
            comt_e = df_coment[df_coment['_aluno_norm']   == a_norm].copy() if df_coment  is not None else pd.DataFrame()

            if not aval_e.empty or not comt_e.empty:
                if not aval_e.empty:
                    base_e = aval_e[['_aluno_norm', '_prof_norm', '_topico_norm',
                                     'Turma', 'Disciplina', 'Tópico', 'Professor', 'Resposta']].copy()
                    base_e = base_e.rename(columns={'Resposta': 'Nota'})
                else:
                    base_e = pd.DataFrame(columns=['_aluno_norm', '_prof_norm', '_topico_norm',
                                                   'Turma', 'Disciplina', 'Tópico', 'Professor', 'Nota'])

                if not comt_e.empty:
                    comt_em = comt_e[['_aluno_norm', '_prof_norm', '_topico_norm',
                                      'Turma', 'Tópico', 'Professor', 'Resposta']].copy()
                    comt_em = comt_em.rename(columns={
                        'Resposta': 'Comentário', 'Turma': '_turma_c',
                        'Tópico': '_topico_c', 'Professor': '_prof_c',
                    })
                else:
                    comt_em = pd.DataFrame(columns=['_aluno_norm', '_prof_norm', '_topico_norm',
                                                    '_turma_c', '_topico_c', '_prof_c', 'Comentário'])

                mg = base_e.merge(comt_em, on=['_aluno_norm', '_prof_norm', '_topico_norm'], how='outer')
                mg['Turma']      = mg['Turma'].fillna(mg.get('_turma_c', '')).fillna('').astype(str).str.strip()
                mg['Tópico']     = mg['Tópico'].fillna(mg.get('_topico_c', '')).fillna('').astype(str).str.strip()
                mg['Professor']  = mg['Professor'].fillna(mg.get('_prof_c', '')).fillna('').astype(str).str.strip()
                mg['Disciplina'] = mg.get('Disciplina', pd.Series(dtype=str)).fillna('').astype(str).str.strip()
                mg['Nota']       = mg.get('Nota', pd.Series(dtype=str)).fillna('').astype(str).str.strip()
                mg['Comentário'] = mg.get('Comentário', pd.Series(dtype=str)).fillna('').astype(str).str.strip()
                mg['Aluno']      = aluno_nome

                mg[['Turma', 'Aluno', 'Disciplina', 'Tópico', 'Professor', 'Nota', 'Comentário']].to_excel(
                    writer, sheet_name='Avaliação e Comentários', index=False
                )
            else:
                pd.DataFrame([{'Info': 'Sem avaliações registradas para este aluno'}]).to_excel(
                    writer, sheet_name='Avaliação e Comentários', index=False
                )
        except Exception:
            pd.DataFrame([{'Info': 'Erro ao gerar dados de avaliação'}]).to_excel(
                writer, sheet_name='Avaliação e Comentários', index=False
            )

    buf.seek(0)
    return buf.read()

nome_arquivo = aluno_sel.replace(' ', '_') + '.xlsx'

st.download_button(
    label="⬇  Exportar para Excel",
    data=gerar_excel_aluno(aluno_sel),
    file_name=f"comportamento_{nome_arquivo}",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=False,
)

st.markdown('</div>', unsafe_allow_html=True)
st.markdown("""
<div class="rh-footer">
  <span>Rehagro</span>
  <span class="rh-footer-dot"></span>
  <span>Customer Success</span>
  <span class="rh-footer-dot"></span>
  <span>Comportamento do Aluno</span>
  <span class="rh-footer-dot"></span>
  <span>© 2026</span>
</div>
""", unsafe_allow_html=True)
