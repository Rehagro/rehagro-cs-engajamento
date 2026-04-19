import streamlit as st
import pandas as pd
import io
import base64, os

# ══════════════════════════════════════════════════════════════
# CSS GLOBAL
# ══════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800;900&family=Inter:wght@400;500;600&display=swap');

:root {
  --g:      #1B3D2A;
  --g2:     #2a5c3f;
  --gold:   #C8A532;
  --cream:  #F2EDE4;
  --c2:     #fdfcf9;
  --c3:     #f8f6f0;
  --txt:    #1c1c1c;
  --txt2:   #4a4a4a;
  --muted:  #888;
  --red:    #dc2626;
  --amber:  #d97706;
  --gok:    #16a34a;
  --blue:   #2563eb;
  --purple: #7c3aed;
  --bd:     #ddd8ce;
  --r:      14px;
}

*, *::before, *::after { box-sizing: border-box; }
html, body, [class*="css"] {
    font-family: 'Inter', sans-serif !important;
    background: var(--cream) !important;
    color: var(--txt) !important;
}
.stApp { background: var(--cream) !important; }
#MainMenu, footer, header { visibility: hidden; }
[data-testid="stAppViewContainer"] { padding-top: 0 !important; }
[data-testid="block-container"]    { padding-top: 0 !important; }
button[aria-label="Close sidebar"],
button[aria-label="Fechar barra lateral"] { display: none !important; }

/* ── Streamlit native overrides ── */
.stButton > button {
    font-family: 'Montserrat', sans-serif !important;
    font-weight: 700 !important; border-radius: 8px !important;
    font-size: 13px !important; transition: all .15s !important;
}
.stButton > button[kind="secondary"] {
    background: #fff !important; border: 1.5px solid var(--bd) !important;
    color: #555 !important; box-shadow: 0 1px 4px rgba(0,0,0,.05) !important;
    padding: 7px 14px !important;
}
.stButton > button[kind="secondary"]:hover {
    border-color: var(--g) !important; color: var(--g) !important;
}
[data-testid="stTextInput"] label {
    font-size: 12px !important; font-weight: 600 !important; color: #555 !important;
}
[data-testid="stTextInput"] input {
    border-radius: 8px !important; border: 1.5px solid #ddd !important;
    font-family: 'Inter', sans-serif !important; font-size: 13px !important;
    padding: 9px 12px !important; background: #fff !important;
}
[data-testid="stTextInput"] input:focus { border-color: var(--g) !important; box-shadow: none !important; }
[data-testid="stSelectbox"] > label { font-size: 12px !important; font-weight: 600 !important; color: #555 !important; }
[data-testid="stSelectbox"] > div > div {
    border-radius: 8px !important; border: 1.5px solid #ddd !important;
    font-family: 'Inter', sans-serif !important; font-size: 13px !important; background: #fff !important;
}
[data-testid="stCheckbox"] label { font-size: 13px !important; color: #555 !important; }
[data-testid="stProgress"] > div > div { background: #f0ede6 !important; border-radius: 3px !important; height: 6px !important; }
[data-testid="stProgress"] > div > div > div {
    background: linear-gradient(90deg, var(--g), var(--gold)) !important;
    border-radius: 3px !important; height: 6px !important;
}
[data-testid="stFileUploader"] {
    background: #fff !important; border: none !important;
    padding: 0 20px 16px !important; margin-top: 0 !important;
    box-shadow: 2px 0 0 rgba(0,0,0,.05), -2px 0 0 rgba(0,0,0,.05) !important;
}
[data-testid="stFileUploader"] section {
    border: 1.5px dashed #d0ccc4 !important; border-radius: 10px !important;
    background: #fafaf8 !important; padding: 14px !important;
}
[data-testid="stFileUploader"] section:hover {
    border-color: var(--gold) !important; background: rgba(200,165,50,.04) !important;
}
[data-testid="stFileUploader"] > label { display: none !important; }
[data-testid="stAlert"] { border-radius: 8px !important; font-size: 13px !important; }
.stDownloadButton > button {
    font-family: 'Montserrat', sans-serif !important; font-weight: 700 !important;
    border-radius: 9px !important; background: var(--g) !important;
    color: #fff !important; border: none !important; padding: 10px 24px !important;
    font-size: 13px !important; box-shadow: 0 4px 14px rgba(27,61,42,.3) !important;
}

/* ── Hero ── */
.rh-hero {
    background: linear-gradient(135deg, var(--g) 0%, var(--g2) 100%);
    padding: 0 40px; margin: -1rem -1rem 0 -1rem;
    position: relative; overflow: hidden;
}
.rh-hero-diag {
    position: absolute; top: 0; right: 0; width: 300px; height: 100%; opacity: .07;
    background-image: repeating-linear-gradient(45deg,#fff 0,#fff 1px,transparent 0,transparent 50%);
    background-size: 12px 12px;
}
.rh-hero-nav {
    height: 52px; display: flex; align-items: center; justify-content: space-between;
    border-bottom: 1px solid rgba(255,255,255,.1);
}
.rh-hero-eyebrow {
    display: flex; align-items: center; gap: 8px;
    color: var(--gold); font-family: 'Montserrat', sans-serif;
    font-weight: 700; font-size: 11px; letter-spacing: 2px; text-transform: uppercase;
}
.rh-hero-eyebrow::before { content: ''; display: block; width: 24px; height: 2px; background: var(--gold); }
.rh-hero-wordmark {
    font-family: 'Montserrat', sans-serif; font-weight: 900;
    font-size: 22px; color: #fff; letter-spacing: -.5px;
    display: flex; align-items: center; gap: 6px;
}
.rh-hero-title-row { padding: 28px 0 24px; }
.rh-hero-h1 {
    font-family: 'Montserrat', sans-serif !important; font-weight: 900; font-size: 34px;
    color: var(--gold) !important; letter-spacing: -.5px; margin: 0 0 6px; text-transform: uppercase;
}
.rh-hero-sub { color: rgba(255,255,255,.65); font-size: 13px; max-width: 480px; line-height: 1.5; margin: 0; }
.rh-hero-pills { display: flex; gap: 8px; margin-top: 16px; flex-wrap: wrap; }
.rh-hero-pill {
    padding: 4px 12px; border: 1.5px solid rgba(255,255,255,.3); border-radius: 20px;
    color: rgba(255,255,255,.8); font-size: 11px; font-weight: 600;
    letter-spacing: .5px; text-transform: uppercase;
}

/* ── Section header ── */
.rh-sec-hdr { display: flex; align-items: center; gap: 10px; margin-bottom: 16px; }
.rh-sec-hdr-lbl {
    font-family: 'Montserrat', sans-serif; font-weight: 800;
    font-size: 11px; letter-spacing: 2px; text-transform: uppercase; color: #555;
}
.rh-sec-hdr-line { flex: 1; height: 1px; background: linear-gradient(to right, #ddd, transparent); }

/* ── Upload card ── */
.rh-uc {
    background: #fff; overflow: hidden;
    border-radius: var(--r) var(--r) 0 0; border-bottom: none;
    box-shadow: 2px 0 0 rgba(0,0,0,.05), -2px 0 0 rgba(0,0,0,.05), 0 -3px 12px rgba(0,0,0,.05);
    margin-bottom: 0 !important;
}
.rh-uc-bar { height: 4px; background: linear-gradient(90deg, var(--g), var(--gold)); }
.rh-uc-body { padding: 18px 20px 14px; }
.rh-uc-hdr { display: flex; align-items: flex-start; gap: 12px; margin-bottom: 14px; }
.rh-uc-num {
    width: 32px; height: 32px; border-radius: 50%; flex-shrink: 0;
    background: rgba(27,61,42,.09);
    display: flex; align-items: center; justify-content: center;
    font-family: 'Montserrat', sans-serif; font-weight: 800; font-size: 13px; color: var(--g);
}
.rh-uc-cat { font-size: 10px; font-weight: 700; color: var(--gold); letter-spacing: 1.2px; text-transform: uppercase; margin-bottom: 2px; }
.rh-uc-title { font-family: 'Montserrat', sans-serif; font-weight: 700; font-size: 14px; color: #1a1a1a; margin-bottom: 2px; }
.rh-uc-sub { font-size: 11px; color: #999; line-height: 1.4; }
.rh-uc-src { background: #f8f6f1; border-radius: 8px; padding: 8px 10px; }
.rh-uc-src-lbl { font-size: 10px; color: #aaa; font-weight: 600; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 4px; }
.rh-uc-src-name { font-size: 11px; color: #555; font-weight: 500; }
.rh-uc-src-extra { font-size: 11px; color: #555; font-weight: 500; margin-top: 2px; }
.rh-uc-src-aval { font-size: 10px; color: var(--gold); font-weight: 600; font-style: italic; margin-top: 5px; }
.rh-uc-bot {
    background: #fff; padding: 10px 20px 18px;
    border-top: 1px solid #f0ede6;
    border-radius: 0 0 var(--r) var(--r);
    box-shadow: 2px 4px 12px rgba(0,0,0,.05), -2px 4px 12px rgba(0,0,0,.04), 0 6px 12px rgba(0,0,0,.04);
    margin-top: 0 !important;
}
.rh-uc-bot-lbl { font-size: 10px; color: #aaa; font-weight: 600; letter-spacing: .8px; text-transform: uppercase; margin-bottom: 7px; }
.rh-tags { display: flex; flex-wrap: wrap; gap: 5px; }
.rh-tag-dk { padding: 2px 8px; border-radius: 20px; font-size: 10px; font-weight: 600; background: var(--g); color: #fff; }
.rh-tag-lt { padding: 2px 8px; border-radius: 20px; font-size: 10px; font-weight: 600; background: #f0ede6; color: #555; }
.rh-tag-var { padding: 2px 8px; border-radius: 20px; font-size: 10px; font-weight: 600; background: rgba(200,165,50,.12); color: #7a5c00; font-style: italic; }

/* ── Info box (slot 6) ── */
.rh-info-box {
    background: rgba(27,61,42,.04); border-radius: var(--r);
    border: 1.5px dashed rgba(27,61,42,.18);
    display: flex; flex-direction: column; align-items: center;
    justify-content: center; gap: 10px; padding: 24px;
    text-align: center; height: 100%; min-height: 200px;
}
.rh-info-box-ico { width: 44px; height: 44px; border-radius: 50%; background: rgba(27,61,42,.1); display: flex; align-items: center; justify-content: center; font-size: 20px; }
.rh-info-box-title { font-family: 'Montserrat', sans-serif; font-weight: 700; font-size: 13px; color: var(--g); }
.rh-info-box-sub { font-size: 12px; color: #888; line-height: 1.5; }

/* ── Progress card ── */
.rh-prog-card { background: #fff; border-radius: var(--r); padding: 22px; box-shadow: 0 2px 12px rgba(0,0,0,.06); text-align: center; min-width: 180px; }
.rh-prog-lbl { font-size: 10px; font-weight: 700; letter-spacing: 1.2px; color: #aaa; text-transform: uppercase; margin-bottom: 12px; }
.rh-prog-nums { display: flex; align-items: baseline; justify-content: center; gap: 4px; margin-bottom: 4px; }
.rh-prog-big { font-family: 'Montserrat', sans-serif; font-weight: 900; font-size: 40px; color: var(--g); line-height: 1; }
.rh-prog-tot { font-size: 18px; color: #aaa; font-weight: 600; }
.rh-prog-sub { font-size: 11px; color: #aaa; margin-bottom: 14px; }

/* ── Student banner ── */
.rh-banner {
    background: linear-gradient(135deg, var(--g) 0%, var(--g2) 100%);
    border-radius: var(--r); padding: 22px 28px; margin-bottom: 28px;
    position: relative; overflow: hidden;
}
.rh-banner-diag {
    position: absolute; top: 0; right: 0; width: 200px; height: 100%; opacity: .06;
    background-image: repeating-linear-gradient(45deg,#fff 0,#fff 1px,transparent 0,transparent 50%);
    background-size: 10px 10px;
}
.rh-banner-eye { font-size: 10px; color: rgba(255,255,255,.5); font-weight: 700; letter-spacing: 1.5px; text-transform: uppercase; margin-bottom: 6px; }
.rh-banner-name { font-family: 'Montserrat', sans-serif; font-weight: 900; font-size: 24px; color: #fff; margin-bottom: 10px; text-transform: uppercase; letter-spacing: -.3px; }
.rh-banner-badges { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
.rh-banner-gld { padding: 4px 12px; border-radius: 20px; background: var(--gold); color: #fff; font-size: 11px; font-weight: 700; letter-spacing: .5px; }
.rh-banner-curso { color: rgba(255,255,255,.4); font-size: 12px; }

/* ── Badges ── */
.rh-badge { display: inline-flex; align-items: center; padding: 3px 9px; border-radius: 20px; font-size: 11px; font-weight: 700; white-space: nowrap; }
.b-ok   { background: #dcfce7; color: #16a34a; }
.b-risk { background: #fef9c3; color: #ca8a04; }
.b-bad  { background: #fee2e2; color: #dc2626; }
.b-prog { background: #dbeafe; color: #2563eb; }
.b-dest { background: #fef3c7; color: #d97706; }
.b-sem  { background: #f1f5f9; color: #94a3b8; }

/* ── KPI cards ── */
.rh-kpi-grid { display: grid; grid-template-columns: repeat(4,1fr); gap: 14px; margin-bottom: 16px; }
.rh-kpi { border-radius: 12px; padding: 16px 18px; }
.rh-kpi-lbl { font-size: 10px; color: #888; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 6px; }
.rh-kpi-val { font-family: 'Montserrat', sans-serif; font-weight: 800; font-size: 18px; line-height: 1.1; }
.rh-kpi-sub { font-size: 11px; font-weight: 600; opacity: .8; margin-top: 4px; }

/* ── Email row ── */
.rh-email-row {
    background: #fff; border-radius: 12px; padding: 14px 18px;
    box-shadow: 0 1px 6px rgba(0,0,0,.05);
    display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 12px;
    margin-bottom: 28px;
}
.rh-email-lbl { font-size: 10px; color: #aaa; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 3px; }
.rh-email-val { font-weight: 600; color: var(--g); font-size: 14px; text-decoration: none; }
.rh-email-alert { display: flex; align-items: center; gap: 8px; padding: 8px 16px; background: #fef2f2; border-radius: 8px; border: 1px solid #fecaca; font-size: 12px; color: #991b1b; font-weight: 600; }

/* ── Summary strip (D2) ── */
.rh-sum-grid { display: grid; grid-template-columns: repeat(4,1fr); gap: 14px; margin-bottom: 16px; }
.rh-sum-card { border-radius: 12px; padding: 14px 18px; }
.rh-sum-lbl { font-size: 10px; color: #888; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 6px; }
.rh-sum-val { font-family: 'Montserrat', sans-serif; font-weight: 900; font-size: 28px; }

/* ── Module table ── */
.rh-mod-wrap { background: #fff; border-radius: var(--r); overflow: hidden; box-shadow: 0 2px 12px rgba(0,0,0,.06); margin-bottom: 28px; }
.rh-mod-hdr { padding: 12px 18px; background: var(--c3); border-bottom: 1px solid #ede9e0; font-family: 'Montserrat', sans-serif; font-weight: 700; font-size: 11px; color: #555; text-transform: uppercase; letter-spacing: 1px; }
.rh-mod-tbl { width: 100%; border-collapse: collapse; font-size: 13px; }
.rh-mod-tbl th { padding: 10px 14px; text-align: left; font-size: 10px; font-weight: 700; color: #888; letter-spacing: .8px; text-transform: uppercase; white-space: nowrap; border-bottom: 2px solid #ede9e0; background: var(--c3); }
.rh-mod-tbl td { padding: 11px 14px; vertical-align: middle; }
.rh-sc-bl { width: 28px; height: 28px; border-radius: 50%; background: #dbeafe; display: inline-flex; align-items: center; justify-content: center; font-family: 'Montserrat', sans-serif; font-weight: 800; font-size: 11px; color: var(--blue); }
.rh-sc-pu { width: 28px; height: 28px; border-radius: 50%; background: #f3e8ff; display: inline-flex; align-items: center; justify-content: center; font-family: 'Montserrat', sans-serif; font-weight: 800; font-size: 11px; color: var(--purple); }
.rh-pt-ok  { padding: 3px 9px; border-radius: 20px; background: #f0fdf4; color: #16a34a; font-size: 11px; font-weight: 600; white-space: nowrap; }
.rh-pt-bad { padding: 3px 9px; border-radius: 20px; background: #fef2f2; color: #dc2626; font-size: 11px; font-weight: 600; white-space: nowrap; }
.rh-pt-neu { padding: 3px 9px; border-radius: 20px; background: var(--c3); color: #888; font-size: 11px; font-weight: 600; white-space: nowrap; }

/* ── Avaliação cards ── */
.rh-aval-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(320px,1fr)); gap: 16px; margin-bottom: 28px; }
.rh-aval-card { background: #fff; border-radius: var(--r); overflow: hidden; box-shadow: 0 2px 12px rgba(0,0,0,.06); border: 1px solid #f0ede6; }
.rh-aval-hdr { padding: 12px 16px; background: linear-gradient(135deg,rgba(27,61,42,.05),rgba(200,165,50,.08)); border-bottom: 1px solid #ede9e0; display: flex; justify-content: space-between; align-items: center; gap: 8px; }
.rh-aval-turma { font-size: 10px; color: var(--gold); font-weight: 700; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 1px; }
.rh-aval-disc { font-family: 'Montserrat', sans-serif; font-weight: 700; font-size: 13px; color: #222; }
.rh-nota-circle { width: 44px; height: 44px; border-radius: 50%; background: var(--g); display: flex; flex-direction: column; align-items: center; justify-content: center; flex-shrink: 0; }
.rh-nota-val { font-family: 'Montserrat', sans-serif; font-weight: 900; font-size: 14px; color: #fff; line-height: 1; }
.rh-nota-lbl { font-size: 8px; color: rgba(255,255,255,.6); letter-spacing: .5px; }
.rh-aval-body { padding: 14px 16px; }
.rh-aval-g2 { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
.rh-f-lbl { font-size: 10px; color: #aaa; font-weight: 600; text-transform: uppercase; letter-spacing: .8px; margin-bottom: 2px; }
.rh-f-val { font-size: 12px; color: #444; line-height: 1.4; }
.rh-comment { background: #faf8f4; border-radius: 8px; padding: 10px 12px; border-left: 3px solid var(--gold); margin-top: 12px; }
.rh-comment-lbl { font-size: 10px; color: #aaa; font-weight: 600; text-transform: uppercase; letter-spacing: .8px; margin-bottom: 4px; }
.rh-comment-txt { font-size: 12px; color: #555; font-style: italic; line-height: 1.5; }

/* ── Empty state ── */
.rh-empty { background: #fff; border-radius: var(--r); padding: 40px; text-align: center; box-shadow: 0 2px 12px rgba(0,0,0,.06); margin-bottom: 28px; }
.rh-empty-ico { width: 48px; height: 48px; border-radius: 50%; background: var(--c3); display: flex; align-items: center; justify-content: center; margin: 0 auto 12px; font-size: 22px; }
.rh-empty-title { font-family: 'Montserrat', sans-serif; font-weight: 700; font-size: 14px; color: #ccc; margin-bottom: 4px; }
.rh-empty-sub { font-size: 12px; color: #bbb; }

/* ── CTA ── */
.rh-cta {
    background: linear-gradient(135deg,rgba(27,61,42,.05),rgba(200,165,50,.1));
    border: 1px solid rgba(200,165,50,.25); border-radius: var(--r);
    padding: 20px 24px; display: flex; align-items: center;
    justify-content: space-between; gap: 16px; flex-wrap: wrap; margin-bottom: 40px;
}
.rh-cta-title { font-family: 'Montserrat', sans-serif; font-weight: 700; font-size: 14px; color: #333; margin-bottom: 3px; }
.rh-cta-sub { font-size: 12px; color: #777; }

/* ── Misc ── */
.rh-div { height: 1px; background: linear-gradient(90deg, var(--bd), transparent); margin: 20px 0; opacity: .5; }
.rh-footer { display: flex; align-items: center; justify-content: center; gap: 16px; padding: 32px 0 16px; border-top: 1px solid var(--bd); margin-top: 40px; color: #888; font-size: 12px; }
.rh-dot { width: 4px; height: 4px; border-radius: 50%; background: var(--gold); display: inline-block; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# CONSTANTS & SESSION STATE
# ══════════════════════════════════════════════════════════════
_MESES = {
    'janeiro': 1, 'fevereiro': 2, 'março': 3, 'abril': 4,
    'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8,
    'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12,
}
if 'ca_vistos' not in st.session_state:
    st.session_state.ca_vistos = set()

# ══════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════
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

# ══════════════════════════════════════════════════════════════
# DATA LOADERS
# ══════════════════════════════════════════════════════════════
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
    df['_aluno_norm']  = df['Aluno'].astype(str).str.strip().str.lower()
    df['_prof_norm']   = df['Professor'].astype(str).str.strip().str.lower()
    df['_topico_norm'] = df['Tópico'].astype(str).str.strip().str.lower()
    return df

@st.cache_data(show_spinner=False)
def carregar_comentarios(arquivo_bytes):
    df = pd.read_excel(io.BytesIO(arquivo_bytes), skiprows=2)
    df.columns = df.columns.str.strip()
    df['_aluno_norm']  = df['Aluno'].astype(str).str.strip().str.lower()
    df['_prof_norm']   = df['Professor'].astype(str).str.strip().str.lower()
    df['_topico_norm'] = df['Tópico'].astype(str).str.strip().str.lower()
    return df

# ══════════════════════════════════════════════════════════════
# UI HELPERS
# ══════════════════════════════════════════════════════════════
def _sec_hdr(icon, label):
    return f'''<div class="rh-sec-hdr">
        <span style="color:var(--gold);font-size:14px">{icon}</span>
        <span class="rh-sec-hdr-lbl">{label}</span>
        <div class="rh-sec-hdr-line"></div>
    </div>'''

def _upload_card(num, cat, title, subtitle, source, avaliar, extra_source=''):
    extra_html = f'<div class="rh-uc-src-extra">{extra_source}</div>' if extra_source else ''
    return (
        f'<div class="rh-uc"><div class="rh-uc-bar"></div><div class="rh-uc-body">'
        f'<div class="rh-uc-hdr"><div class="rh-uc-num">{num}</div>'
        f'<div><div class="rh-uc-cat">{cat}</div><div class="rh-uc-title">{title}</div>'
        f'<div class="rh-uc-sub">{subtitle}</div></div></div>'
        f'<div class="rh-uc-src"><div class="rh-uc-src-lbl">Fonte de dados</div>'
        f'<div class="rh-uc-src-name">{source}</div>{extra_html}'
        f'<div class="rh-uc-src-aval">&#9658; {avaliar}</div>'
        f'</div></div></div>'
    )

def _upload_bot(tags):
    tag_html = ''.join([
        f'<span class="rh-tag-dk">{t["l"]}</span>' if t.get('dk') else
        (f'<span class="rh-tag-var">{t["l"]}</span>' if t.get('var') else
         f'<span class="rh-tag-lt">{t["l"]}</span>')
        for t in tags
    ])
    return f"""<div class="rh-uc-bot">
  <div class="rh-uc-bot-lbl">⚙ Filtros pré-configurados</div>
  <div class="rh-tags">{tag_html}</div>
</div>"""

def _status_badge(status):
    s = str(status)
    if '3 - Finalizado' in s or 'Concluído' in s or 'Finalizado' in s:
        return '<span class="rh-badge b-ok">Concluído</span>'
    if '1 - Em Andamento' in s or 'Em Andamento' in s or 'Em andamento' in s:
        return '<span class="rh-badge b-prog">Em Andamento</span>'
    if '4 - Destravado' in s or 'Destravado' in s:
        return '<span class="rh-badge b-dest">Destravado</span>'
    if 'Sem módulo' in s or 'Sem Registro' in s:
        return '<span class="rh-badge b-sem">Sem Registro</span>'
    return f'<span class="rh-badge b-sem">{s}</span>'

def _pont_badge(val):
    v = str(val) if val else ''
    if not v or v in ('', 'nan'):
        return '<span class="rh-pt-neu">—</span>'
    vl = v.lower()
    if 'prazo' in vl and 'não' not in vl and 'nao' not in vl:
        return f'<span class="rh-pt-ok">{v}</span>'
    if 'não' in vl or 'nao' in vl:
        return f'<span class="rh-pt-bad">{v}</span>'
    return f'<span class="rh-pt-neu">{v}</span>'

def _render_mod_table(tabela):
    if tabela.empty:
        return '<div class="rh-empty"><div class="rh-empty-ico">📚</div><div class="rh-empty-title">Nenhum módulo registrado</div></div>'
    rows = ''
    for i, row in tabela.iterrows():
        disc   = str(row.get('DISCIPLINA', ''))
        status = _status_badge(row.get('STATUS', ''))
        np_v   = str(row.get('Nota Atividade Prática', ''))
        nota_p = f'<div class="rh-sc-bl">{np_v}</div>' if np_v and np_v not in ('', 'nan') else '<span style="color:#d1d5db">—</span>'
        pont   = _pont_badge(row.get('Pontualidade', ''))
        nt_v   = str(row.get('Nota Teste seu Conhecimento', ''))
        nota_t = f'<div class="rh-sc-pu">{nt_v}</div>' if nt_v and nt_v not in ('', 'nan') else '<span style="color:#d1d5db">—</span>'
        bg     = '#fff' if i % 2 == 0 else '#faf9f6'
        rows  += f'<tr style="border-bottom:1px solid #f0ede6;background:{bg}"><td style="padding:11px 14px;font-weight:500;color:#222">{disc}</td><td style="padding:11px 14px">{status}</td><td style="padding:11px 14px">{nota_p}</td><td style="padding:11px 14px">{pont}</td><td style="padding:11px 14px">{nota_t}</td></tr>'
    return f'''<div class="rh-mod-wrap">
  <div class="rh-mod-hdr">📚 Detalhamento por Disciplina</div>
  <div style="overflow-x:auto">
    <table class="rh-mod-tbl">
      <thead><tr>
        <th>Disciplina</th><th>Status</th>
        <th>Nota Atividade Prática</th><th>Pontualidade</th>
        <th>Nota Teste do Conhecimento</th>
      </tr></thead>
      <tbody>{rows}</tbody>
    </table>
  </div>
</div>'''

def _render_aval_cards(tabela_nps):
    if tabela_nps is None or tabela_nps.empty:
        return '''<div class="rh-empty">
  <div class="rh-empty-ico">⭐</div>
  <div class="rh-empty-title">Nenhuma avaliação registrada</div>
  <div class="rh-empty-sub">Este aluno ainda não possui avaliações de aula.</div>
</div>'''
    cards = ''
    for _, row in tabela_nps.iterrows():
        nota_raw = str(row.get('Nota', ''))
        nota_num = nota_raw.replace('Nota ', '').strip() if nota_raw not in ('', 'nan') else ''
        nota_circle = f'<div class="rh-nota-circle"><span class="rh-nota-val">{nota_num}</span><span class="rh-nota-lbl">NOTA</span></div>' if nota_num else ''
        coment = str(row.get('Comentário', ''))
        coment_html = ''
        if coment and coment not in ('', 'nan'):
            coment_html = f'<div class="rh-comment"><div class="rh-comment-lbl">Comentário</div><div class="rh-comment-txt">"{coment}"</div></div>'
        turma = str(row.get('Turma', ''))
        disc  = str(row.get('Disciplina', ''))
        top   = str(row.get('Tópico', ''))
        prof  = str(row.get('Professor', ''))
        cards += f'''<div class="rh-aval-card">
  <div class="rh-aval-hdr">
    <div><div class="rh-aval-turma">{turma}</div><div class="rh-aval-disc">{disc}</div></div>
    {nota_circle}
  </div>
  <div class="rh-aval-body">
    <div class="rh-aval-g2">
      <div><div class="rh-f-lbl">Tópico</div><div class="rh-f-val">{top}</div></div>
      <div><div class="rh-f-lbl">Professor(a)</div><div class="rh-f-val">{prof}</div></div>
    </div>
    {coment_html}
  </div>
</div>'''
    return f'<div class="rh-aval-grid">{cards}</div>'

# ══════════════════════════════════════════════════════════════
# LOGO
# ══════════════════════════════════════════════════════════════
def _logo_b64(path):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return ""

_logo_b64_str = _logo_b64("Logo-Rehagro-chapada-branca.png")
_logo_img = f'<img src="data:image/png;base64,{_logo_b64_str}" style="height:76px;opacity:0.92"/>' if _logo_b64_str else ''

# ══════════════════════════════════════════════════════════════
# HERO
# ══════════════════════════════════════════════════════════════
_logo_slot = f'<div style="flex-shrink:0;padding-left:24px;">{_logo_img}</div>' if _logo_img else ''
_hero_html = (
    '<div class="rh-hero"><div class="rh-hero-diag"></div>'
    '<div class="rh-hero-nav"><div class="rh-hero-eyebrow">Rehagro · Customer Success</div></div>'
    '<div style="display:flex;justify-content:space-between;align-items:center;padding:28px 0 24px;">'
    '<div>'
    '<h1 class="rh-hero-h1">Comportamento do Aluno</h1>'
    '<p class="rh-hero-sub">Panorama completo do aluno para preparar o contato proativo do CS.</p>'
    '<div class="rh-hero-pills">'
    '<span class="rh-hero-pill">Acesso Canvas</span>'
    '<span class="rh-hero-pill">Módulos</span>'
    '<span class="rh-hero-pill">NPS</span>'
    '<span class="rh-hero-pill">Comentários</span>'
    '</div></div>'
    + _logo_slot +
    '</div></div>'
)
st.markdown(_hero_html, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# BACK BUTTON
# ══════════════════════════════════════════════════════════════
st.markdown('<div style="padding:16px 0 0">', unsafe_allow_html=True)
if st.button("← Monitoramento de Alunos", key="nav_mon"):
    st.switch_page("pages/monitoramento.py")
st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# UPLOAD SECTION
# ══════════════════════════════════════════════════════════════
st.markdown(_sec_hdr('↑', 'Carregar Arquivos'), unsafe_allow_html=True)

TAGS_CANVAS = [
    {'l':'Status do aluno','dk':True},{'l':'Ativo'},
    {'l':'Função na disciplina','dk':True},{'l':'Aluno'},
    {'l':'Curso','dk':True},{'l':'seu curso','var':True},
]
TAGS_STATUS = TAGS_CANVAS
TAGS_TAREFAS = TAGS_CANVAS
TAGS_NPS = [
    {'l':'Item','dk':True},{'l':'De 0 a 10 · Nota da aula'},
    {'l':'Status de matrícula','dk':True},{'l':'Matriculado'},
    {'l':'Tipo de aula','dk':True},{'l':'Gravada'},
    {'l':'Curso','dk':True},{'l':'seu curso','var':True},
]
TAGS_COMENT = [
    {'l':'Status de matrícula','dk':True},{'l':'Matriculado'},
    {'l':'Tipo de aula','dk':True},{'l':'Gravada'},
    {'l':'Curso','dk':True},{'l':'seu curso','var':True},
]

# Row 1
col_u1, col_u2, col_u3 = st.columns(3, gap="medium")
with col_u1:
    st.markdown(_upload_card(
        1, 'Engajamento', 'Acesso ao Canvas',
        'Login e atividade na plataforma LMS por aluno.',
        'BI Rehagro Canvas — Resumo de dados dos alunos',
        'Avaliar: Acesso ao Canvas',
    ), unsafe_allow_html=True)
    f_canvas = st.file_uploader("", type=["xlsx"], key="ca_canvas", label_visibility="collapsed")
    st.markdown(_upload_bot(TAGS_CANVAS), unsafe_allow_html=True)

with col_u2:
    st.markdown(_upload_card(
        2, 'Desempenho', 'Envio de Tarefas',
        'Histórico de entrega e avaliação de atividades.',
        'BI Rehagro Canvas — Resumo de dados dos alunos',
        'Avaliar: Entrega de atividades',
    ), unsafe_allow_html=True)
    f_tarefas = st.file_uploader("", type=["xlsx"], key="ca_tarefas", label_visibility="collapsed")
    st.markdown(_upload_bot(TAGS_TAREFAS), unsafe_allow_html=True)

with col_u3:
    st.markdown(_upload_card(
        3, 'Satisfação', 'Comentários',
        'Feedbacks qualitativos dos alunos nas aulas gravadas.',
        'BI Rehagro Educação — Avaliação de Aula',
        'Tabela Comentários',
        'Relatório do Professor · Comentários',
    ), unsafe_allow_html=True)
    f_coment = st.file_uploader("", type=["xlsx"], key="ca_coment", label_visibility="collapsed")
    st.markdown(_upload_bot(TAGS_COMENT), unsafe_allow_html=True)

st.markdown('<div style="height:20px"></div>', unsafe_allow_html=True)

# Row 2
col_u4, col_u5, col_u6 = st.columns(3, gap="medium")
with col_u4:
    st.markdown(_upload_card(
        4, 'Progresso Acadêmico', 'Módulos e Atividades',
        'Status do módulo por aluno e por disciplina.',
        'BI Rehagro Canvas — Resumo de dados dos alunos',
        'Avaliar: Qual o status do módulo para cada aluno e disciplina',
    ), unsafe_allow_html=True)
    f_status = st.file_uploader("", type=["xlsx"], key="ca_status", label_visibility="collapsed")
    st.markdown(_upload_bot(TAGS_STATUS), unsafe_allow_html=True)

with col_u5:
    st.markdown(_upload_card(
        5, 'Avaliação', 'Avaliação de Aula / Etapa',
        'Respostas objetivas e detalhamento de avaliações.',
        'BI Rehagro Educação — Avaliação de Aula',
        'Tabela respostas objetivas',
        'Avaliação de aula/etapa · Detalhamento',
    ), unsafe_allow_html=True)
    f_nps = st.file_uploader("", type=["xlsx"], key="ca_nps", label_visibility="collapsed")
    st.markdown(_upload_bot(TAGS_NPS), unsafe_allow_html=True)

with col_u6:
    st.markdown("""
    <div class="rh-info-box">
      <div class="rh-info-box-ico">ℹ️</div>
      <div class="rh-info-box-title">Todos os arquivos carregados?</div>
      <div class="rh-info-box-sub">Após o upload dos 5 arquivos, selecione um aluno abaixo para ver o painel completo.</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown('<div style="height:20px"></div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# GUARD: Canvas obrigatório
# ══════════════════════════════════════════════════════════════
if not f_canvas:
    st.info("Carregue pelo menos o arquivo **Acesso ao Canvas** para liberar o seletor de aluno.")
    st.stop()

# ══════════════════════════════════════════════════════════════
# LOAD DATA
# ══════════════════════════════════════════════════════════════
with st.spinner("Lendo arquivos..."):
    df_canvas  = carregar_canvas(f_canvas.read())
    df_status  = carregar_status_modulo(f_status.read())  if f_status  else None
    df_tarefas = carregar_envio_tarefas(f_tarefas.read()) if f_tarefas else None
    df_nps_raw = carregar_nps(f_nps.read())               if f_nps     else None
    df_coment  = carregar_comentarios(f_coment.read())    if f_coment  else None

col_nome = next((c for c in df_canvas.columns if 'NOME' in c.upper()), None)
if not col_nome:
    st.error("Coluna de nome não encontrada no arquivo Canvas.")
    st.stop()

nomes_todos = sorted(df_canvas[col_nome].dropna().astype(str).str.strip().unique().tolist())

# ══════════════════════════════════════════════════════════════
# SELECIONAR ALUNO
# ══════════════════════════════════════════════════════════════
st.markdown(_sec_hdr('👤', 'Selecionar Aluno'), unsafe_allow_html=True)

col_sel, col_prog = st.columns([3, 1], gap="large")

with col_sel:
    col_busca, col_filtro = st.columns([2, 1])
    with col_busca:
        busca = st.text_input("Buscar por nome:", placeholder="Digite parte do nome…", key="ca_busca")
    with col_filtro:
        so_nao_vistos = st.checkbox("Apenas não vistos", value=False, key="ca_filtro_vistos",
                                    help="Oculta alunos já visualizados nesta sessão.")

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
        st.stop()

    def _fmt_nome(nome):
        return f"✓  {nome}" if nome in st.session_state.ca_vistos else f"○  {nome}"

    aluno_sel = st.selectbox("Selecionar aluno:", nomes_filtrados, format_func=_fmt_nome, key="ca_aluno_sel")

with col_prog:
    n_vistos = len(st.session_state.ca_vistos)
    n_total  = len(nomes_todos)
    pct      = n_vistos / n_total if n_total > 0 else 0
    st.markdown(f"""
    <div class="rh-prog-card">
      <div class="rh-prog-lbl">Progresso da Sessão</div>
      <div class="rh-prog-nums">
        <span class="rh-prog-big">{n_vistos}</span>
        <span class="rh-prog-tot">/ {n_total}</span>
      </div>
      <div class="rh-prog-sub">alunos visualizados</div>
    </div>
    """, unsafe_allow_html=True)
    st.progress(pct)
    if st.button("↺  Reiniciar progresso", use_container_width=True, key="ca_limpar_vistos"):
        st.session_state.ca_vistos = set()
        st.rerun()

if aluno_sel and aluno_sel not in st.session_state.ca_vistos:
    st.session_state.ca_vistos.add(aluno_sel)
    st.rerun()

# ══════════════════════════════════════════════════════════════
# STUDENT BANNER
# ══════════════════════════════════════════════════════════════
st.markdown('<div style="height:20px"></div>', unsafe_allow_html=True)
st.markdown(f"""
<div class="rh-banner">
  <div class="rh-banner-diag"></div>
  <div class="rh-banner-eye">Aluno Selecionado</div>
  <div class="rh-banner-name">{aluno_sel}</div>
  <div class="rh-banner-badges">
    <span class="rh-banner-gld">COMPORTAMENTO COMPLETO</span>
  </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# DASHBOARD 1 — ACESSO AO CANVAS
# ══════════════════════════════════════════════════════════════
st.markdown(_sec_hdr('📅', 'Dashboard 1 · Acesso ao Canvas'), unsafe_allow_html=True)

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
        ultimo_acesso, cor_acesso, bg_acesso = 'Nunca acessou', '#dc2626', '#fef2f2'
    else:
        ultimo_acesso = _fmt_data(ultima_raw)
        cor_acesso, bg_acesso = '#16a34a', '#f0fdf4'

    dias_raw = row.get(col_dias) if col_dias else None
    dias_val = int(dias_raw) if pd.notna(dias_raw) else None
    dias_str = f"{dias_val} dias" if dias_val is not None else "—"
    if dias_val is None:
        cor_dias, bg_dias, sub_dias = '#888', '#f8f6f0', ''
    elif dias_val > 60:
        cor_dias, bg_dias, sub_dias = '#dc2626', '#fef2f2', 'Ação urgente'
    elif dias_val > 30:
        cor_dias, bg_dias, sub_dias = '#d97706', '#fffbeb', 'Atenção necessária'
    else:
        cor_dias, bg_dias, sub_dias = '#16a34a', '#f0fdf4', 'Dentro do esperado'

    email_val = str(row.get(col_email, '')).strip() if col_email else '—'
    curso_val = str(row.get(col_curso, '')).strip() if col_curso else '—'
    turma_val = str(row.get(col_turma, '')).strip() if col_turma else '—'

    st.markdown(f"""
    <div class="rh-kpi-grid">
      <div class="rh-kpi" style="background:rgba(27,61,42,.08);border:1px solid rgba(27,61,42,.12)">
        <div class="rh-kpi-lbl">Curso</div>
        <div class="rh-kpi-val" style="color:var(--g)">{curso_val}</div>
      </div>
      <div class="rh-kpi" style="background:#eff6ff;border:1px solid rgba(37,99,235,.12)">
        <div class="rh-kpi-lbl">Turma</div>
        <div class="rh-kpi-val" style="color:var(--blue)">{turma_val}</div>
      </div>
      <div class="rh-kpi" style="background:{bg_acesso};border:1px solid {cor_acesso}20">
        <div class="rh-kpi-lbl">Último Acesso</div>
        <div class="rh-kpi-val" style="color:{cor_acesso}">{ultimo_acesso}</div>
      </div>
      <div class="rh-kpi" style="background:{bg_dias};border:1px solid {cor_dias}20">
        <div class="rh-kpi-lbl">Dias sem Acesso</div>
        <div class="rh-kpi-val" style="color:{cor_dias}">{dias_str}</div>
        {f'<div class="rh-kpi-sub" style="color:{cor_dias}">{sub_dias}</div>' if sub_dias else ''}
      </div>
    </div>
    """, unsafe_allow_html=True)

    alerta_html = ''
    if dias_val and dias_val > 30:
        alerta_html = f'<div class="rh-email-alert">⚠️ Sem acesso há <strong style="margin:0 3px">{dias_val} dias</strong> — contato recomendado</div>'

    st.markdown(f"""
    <div class="rh-email-row">
      <div>
        <div class="rh-email-lbl">E-mail do aluno</div>
        <a href="mailto:{email_val}" class="rh-email-val">{email_val}</a>
      </div>
      {alerta_html}
    </div>
    """, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# DASHBOARD 2 — MÓDULOS E ATIVIDADES
# ══════════════════════════════════════════════════════════════
st.markdown(_sec_hdr('📚', 'Dashboard 2 · Módulos e Atividades'), unsafe_allow_html=True)

if df_status is None and df_tarefas is None:
    st.info("Arquivos de módulos/atividades não carregados.")
else:
    if df_status is not None:
        col_aluno_mod = next((c for c in df_status.columns if c.upper() == 'ALUNO'), None)
        df_mod_aluno  = df_status[df_status[col_aluno_mod].astype(str).str.strip() == aluno_sel].copy() if col_aluno_mod else pd.DataFrame()
    else:
        df_mod_aluno = pd.DataFrame()

    if df_tarefas is not None:
        col_aluno_tar = next((c for c in df_tarefas.columns if c.upper() == 'ALUNO'), None)
        df_tar_aluno  = df_tarefas[df_tarefas[col_aluno_tar].astype(str).str.strip() == aluno_sel].copy() if col_aluno_tar else pd.DataFrame()
    else:
        df_tar_aluno = pd.DataFrame()

    if df_mod_aluno.empty and df_tar_aluno.empty:
        st.info("Nenhum módulo ou atividade registrado para este aluno.")
    else:
        if not df_tar_aluno.empty:
            col_grupo = next((c for c in df_tar_aluno.columns if 'GRUPO' in c.upper()), None)
            col_nota  = next((c for c in df_tar_aluno.columns if 'NOTA'  in c.upper()), None)
            col_pont  = next((c for c in df_tar_aluno.columns if 'PONTUALIDADE' in c.upper()), None)
            if col_grupo and col_nota:
                prat_cols   = {col_nota: 'Nota Atividade Prática'}
                if col_pont: prat_cols[col_pont] = 'Pontualidade'
                prat = (df_tar_aluno[df_tar_aluno[col_grupo] == 'Atividade Pratica']
                        [[c for c in ['DISCIPLINA', col_nota] + ([col_pont] if col_pont else []) if c in df_tar_aluno.columns]]
                        .rename(columns=prat_cols).drop_duplicates(subset=['DISCIPLINA']))
                teste = (df_tar_aluno[df_tar_aluno[col_grupo] == 'Teste seu conhecimento']
                         [['DISCIPLINA', col_nota]].rename(columns={col_nota: 'Nota Teste seu Conhecimento'})
                         .drop_duplicates(subset=['DISCIPLINA']))
            else:
                prat  = pd.DataFrame(columns=['DISCIPLINA', 'Nota Atividade Prática'])
                teste = pd.DataFrame(columns=['DISCIPLINA', 'Nota Teste seu Conhecimento'])
        else:
            prat  = pd.DataFrame(columns=['DISCIPLINA', 'Nota Atividade Prática'])
            teste = pd.DataFrame(columns=['DISCIPLINA', 'Nota Teste seu Conhecimento'])

        if not df_mod_aluno.empty:
            tabela = df_mod_aluno[['ALUNO', 'DISCIPLINA', 'STATUS']].copy()
        else:
            discips = pd.concat([prat[['DISCIPLINA']], teste[['DISCIPLINA']]]).drop_duplicates()
            tabela  = discips.copy()
            tabela.insert(0, 'ALUNO', aluno_sel)
            tabela['STATUS'] = 'Sem módulo registrado'

        tabela = tabela.merge(prat,  on='DISCIPLINA', how='left')
        tabela = tabela.merge(teste, on='DISCIPLINA', how='left')

        for col_n in ['Nota Atividade Prática', 'Nota Teste seu Conhecimento']:
            if col_n in tabela.columns:
                tabela[col_n] = tabela[col_n].apply(
                    lambda x: '' if pd.isna(x) else (str(int(x)) if isinstance(x, float) and x == int(x) else str(x))
                )
            else:
                tabela[col_n] = ''
        if 'Pontualidade' not in tabela.columns:
            tabela['Pontualidade'] = ''
        else:
            tabela['Pontualidade'] = tabela['Pontualidade'].fillna('').astype(str).str.strip()

        total      = len(tabela)
        concluidos = tabela['STATUS'].astype(str).str.contains('3 - Finalizado|Concluído|Finalizado', case=False, na=False).sum()
        andamento  = tabela['STATUS'].astype(str).str.contains('1 - Em Andamento|Em andamento', case=False, na=False).sum()
        sem_reg    = tabela['STATUS'].astype(str).str.contains('Sem módulo|Sem Registro', case=False, na=False).sum()

        st.markdown(f"""
        <div class="rh-sum-grid">
          <div class="rh-sum-card" style="background:#f8f6f0;border:1px solid rgba(0,0,0,.06)">
            <div class="rh-sum-lbl">Total de Módulos</div>
            <div class="rh-sum-val" style="color:#555">{total}</div>
          </div>
          <div class="rh-sum-card" style="background:#f0fdf4;border:1px solid rgba(22,163,74,.12)">
            <div class="rh-sum-lbl">Concluídos</div>
            <div class="rh-sum-val" style="color:#16a34a">{concluidos}</div>
          </div>
          <div class="rh-sum-card" style="background:#eff6ff;border:1px solid rgba(37,99,235,.12)">
            <div class="rh-sum-lbl">Em Andamento</div>
            <div class="rh-sum-val" style="color:#2563eb">{andamento}</div>
          </div>
          <div class="rh-sum-card" style="background:#fef2f2;border:1px solid rgba(220,38,38,.12)">
            <div class="rh-sum-lbl">Sem Registro</div>
            <div class="rh-sum-val" style="color:#dc2626">{sem_reg}</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(_render_mod_table(tabela), unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# DASHBOARD 3 — AVALIAÇÃO E COMENTÁRIOS
# ══════════════════════════════════════════════════════════════
st.markdown(_sec_hdr('⭐', 'Dashboard 3 · Avaliação de Aula / Etapa e Comentários'), unsafe_allow_html=True)

if df_nps_raw is None and df_coment is None:
    st.info("Arquivos de avaliação/comentários não carregados.")
else:
    aluno_norm   = aluno_sel.strip().lower()
    aval_aluno   = df_nps_raw[df_nps_raw['_aluno_norm'] == aluno_norm].copy() if df_nps_raw is not None else pd.DataFrame()
    coment_aluno = df_coment[df_coment['_aluno_norm']   == aluno_norm].copy() if df_coment  is not None else pd.DataFrame()

    if aval_aluno.empty and coment_aluno.empty:
        st.markdown(_render_aval_cards(None), unsafe_allow_html=True)
    else:
        if not aval_aluno.empty:
            base = aval_aluno[['_aluno_norm','_prof_norm','_topico_norm','Turma','Disciplina','Tópico','Professor','Resposta']].copy()
            base = base.rename(columns={'Resposta':'Nota'})
        else:
            base = pd.DataFrame(columns=['_aluno_norm','_prof_norm','_topico_norm','Turma','Disciplina','Tópico','Professor','Nota'])

        if not coment_aluno.empty:
            comt = coment_aluno[['_aluno_norm','_prof_norm','_topico_norm','Turma','Tópico','Professor','Resposta']].copy()
            comt = comt.rename(columns={'Resposta':'Comentário','Turma':'_turma_c','Tópico':'_topico_c','Professor':'_prof_c'})
        else:
            comt = pd.DataFrame(columns=['_aluno_norm','_prof_norm','_topico_norm','_turma_c','_topico_c','_prof_c','Comentário'])

        merged = base.merge(comt, on=['_aluno_norm','_prof_norm','_topico_norm'], how='outer')
        merged['Turma']      = merged['Turma'].fillna(merged.get('_turma_c','')).fillna('').astype(str).str.strip()
        merged['Tópico']     = merged['Tópico'].fillna(merged.get('_topico_c','')).fillna('').astype(str).str.strip()
        merged['Professor']  = merged['Professor'].fillna(merged.get('_prof_c','')).fillna('').astype(str).str.strip()
        merged['Disciplina'] = merged.get('Disciplina', pd.Series(dtype=str)).fillna('').astype(str).str.strip()
        merged['Nota']       = merged.get('Nota', pd.Series(dtype=str)).fillna('').astype(str).str.strip()
        merged['Comentário'] = merged.get('Comentário', pd.Series(dtype=str)).fillna('').astype(str).str.strip()
        merged['Aluno']      = aluno_sel

        tabela_nps = merged[['Turma','Aluno','Disciplina','Tópico','Professor','Nota','Comentário']].reset_index(drop=True)
        st.markdown(_render_aval_cards(tabela_nps), unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# CTA
# ══════════════════════════════════════════════════════════════
st.markdown("""
<div class="rh-cta">
  <div>
    <div class="rh-cta-title">Pronto para entrar em contato?</div>
    <div class="rh-cta-sub">Use os dados acima para preparar uma abordagem proativa e personalizada.</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# EXPORTAÇÃO
# ══════════════════════════════════════════════════════════════
st.markdown('<div class="rh-div"></div>', unsafe_allow_html=True)

def gerar_excel_aluno(aluno_nome):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        # Aba 1: Canvas
        linha_exp = df_canvas[df_canvas[col_nome].astype(str).str.strip() == aluno_nome]
        (linha_exp if not linha_exp.empty else pd.DataFrame([{'Info':'Sem dados de Canvas'}])).to_excel(writer, sheet_name='Acesso ao Canvas', index=False)

        # Aba 2: Módulos e Atividades
        try:
            mod_exp = pd.DataFrame()
            if df_status is not None:
                col_a = next((c for c in df_status.columns if c.upper() == 'ALUNO'), None)
                if col_a: mod_exp = df_status[df_status[col_a].astype(str).str.strip() == aluno_nome].copy()
            if df_tarefas is not None:
                col_a2  = next((c for c in df_tarefas.columns if c.upper() == 'ALUNO'), None)
                tar_exp = df_tarefas[df_tarefas[col_a2].astype(str).str.strip() == aluno_nome].copy() if col_a2 else pd.DataFrame()
                col_g   = next((c for c in tar_exp.columns if 'GRUPO' in c.upper()), None)
                col_nt  = next((c for c in tar_exp.columns if 'NOTA'  in c.upper()), None)
                if col_g and col_nt and not tar_exp.empty and not mod_exp.empty:
                    pe = tar_exp[tar_exp[col_g]=='Atividade Pratica'][['DISCIPLINA',col_nt]].rename(columns={col_nt:'Nota Atividade Prática'}).drop_duplicates('DISCIPLINA')
                    te = tar_exp[tar_exp[col_g]=='Teste seu conhecimento'][['DISCIPLINA',col_nt]].rename(columns={col_nt:'Nota Teste seu Conhecimento'}).drop_duplicates('DISCIPLINA')
                    mod_exp = mod_exp.merge(pe,on='DISCIPLINA',how='left').merge(te,on='DISCIPLINA',how='left')
            (mod_exp if not mod_exp.empty else pd.DataFrame([{'Info':'Sem dados de módulos'}])).to_excel(writer, sheet_name='Módulos e Atividades', index=False)
        except Exception:
            pd.DataFrame([{'Info':'Erro ao gerar módulos'}]).to_excel(writer, sheet_name='Módulos e Atividades', index=False)

        # Aba 3: Avaliação e Comentários
        try:
            a_norm  = aluno_nome.strip().lower()
            aval_e  = df_nps_raw[df_nps_raw['_aluno_norm']==a_norm].copy() if df_nps_raw is not None else pd.DataFrame()
            comt_e  = df_coment[df_coment['_aluno_norm']==a_norm].copy()   if df_coment  is not None else pd.DataFrame()
            if not aval_e.empty or not comt_e.empty:
                if not aval_e.empty:
                    base_e = aval_e[['_aluno_norm','_prof_norm','_topico_norm','Turma','Disciplina','Tópico','Professor','Resposta']].copy().rename(columns={'Resposta':'Nota'})
                else:
                    base_e = pd.DataFrame(columns=['_aluno_norm','_prof_norm','_topico_norm','Turma','Disciplina','Tópico','Professor','Nota'])
                if not comt_e.empty:
                    comt_em = comt_e[['_aluno_norm','_prof_norm','_topico_norm','Turma','Tópico','Professor','Resposta']].copy().rename(columns={'Resposta':'Comentário','Turma':'_turma_c','Tópico':'_topico_c','Professor':'_prof_c'})
                else:
                    comt_em = pd.DataFrame(columns=['_aluno_norm','_prof_norm','_topico_norm','_turma_c','_topico_c','_prof_c','Comentário'])
                mg = base_e.merge(comt_em,on=['_aluno_norm','_prof_norm','_topico_norm'],how='outer')
                mg['Turma']      = mg['Turma'].fillna(mg.get('_turma_c','')).fillna('').astype(str).str.strip()
                mg['Tópico']     = mg['Tópico'].fillna(mg.get('_topico_c','')).fillna('').astype(str).str.strip()
                mg['Professor']  = mg['Professor'].fillna(mg.get('_prof_c','')).fillna('').astype(str).str.strip()
                mg['Disciplina'] = mg.get('Disciplina',pd.Series(dtype=str)).fillna('').astype(str).str.strip()
                mg['Nota']       = mg.get('Nota',pd.Series(dtype=str)).fillna('').astype(str).str.strip()
                mg['Comentário'] = mg.get('Comentário',pd.Series(dtype=str)).fillna('').astype(str).str.strip()
                mg['Aluno']      = aluno_nome
                mg[['Turma','Aluno','Disciplina','Tópico','Professor','Nota','Comentário']].to_excel(writer, sheet_name='Avaliação e Comentários', index=False)
            else:
                pd.DataFrame([{'Info':'Sem avaliações'}]).to_excel(writer, sheet_name='Avaliação e Comentários', index=False)
        except Exception:
            pd.DataFrame([{'Info':'Erro ao gerar avaliações'}]).to_excel(writer, sheet_name='Avaliação e Comentários', index=False)

    buf.seek(0)
    return buf.read()

nome_arquivo = aluno_sel.replace(' ', '_') + '.xlsx'
st.download_button(
    label="⬇  Exportar para Excel",
    data=gerar_excel_aluno(aluno_sel),
    file_name=f"comportamento_{nome_arquivo}",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# ══════════════════════════════════════════════════════════════
# FOOTER
# ══════════════════════════════════════════════════════════════
st.markdown("""
<div class="rh-footer">
  <span>Rehagro</span><span class="rh-dot"></span>
  <span>Customer Success</span><span class="rh-dot"></span>
  <span>Comportamento do Aluno</span><span class="rh-dot"></span>
  <span>© 2026</span>
</div>
""", unsafe_allow_html=True)
