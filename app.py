import streamlit as st
import pandas as pd
import re
import io
import base64
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from sendgrid import SendGridAPIClient
    from sendgrid.helpers.mail import (Mail, Attachment, FileContent,
                                        FileName, FileType, Disposition, Cc)
    SENDGRID_OK = True
except ImportError:
    SENDGRID_OK = False

st.set_page_config(
    page_title="CS Rehagro | Engajamento",
    page_icon="🌱",
    layout="wide",
)

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

/* ── HERO ─────────────────────────────── */
.rh-hero {
    background: var(--verde);
    padding: 24px 40px 28px;
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
    font-size: 0.7rem;
    font-weight: 600;
    letter-spacing: 4px;
    text-transform: uppercase;
    color: var(--ouro);
    margin: 0 0 8px 0;
    display: flex;
    align-items: center;
    gap: 10px;
}
.rh-eyebrow::before {
    content: '';
    display: inline-block;
    width: 32px; height: 2px;
    background: var(--ouro);
}
.rh-hero-title {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 2.6rem;
    line-height: 1;
    color: var(--branco);
    margin: 0 0 4px 0;
    letter-spacing: 2px;
    white-space: nowrap;
}
.rh-hero-title span { color: var(--ouro); }
.rh-hero-sub {
    font-size: 0.9rem;
    color: rgba(255,255,255,0.6);
    font-weight: 300;
    margin: 6px 0 0 0;
    max-width: 640px;
    line-height: 1.5;
    white-space: nowrap;
}
.rh-hero-pills {
    display: flex; gap: 8px; margin-top: 14px; flex-wrap: wrap;
}
.rh-pill {
    background: rgba(200,169,81,0.15);
    border: 1px solid rgba(200,169,81,0.35);
    color: var(--ouro2);
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 1px;
    text-transform: uppercase;
    padding: 5px 14px;
    border-radius: 100px;
}

/* ── BODY WRAPPER ─────────────────────── */
.rh-body { padding: 40px 0 0 0; }

/* ── SECTION LABEL ───────────────────── */
.rh-section {
    font-family: 'Outfit', sans-serif;
    font-size: 0.65rem;
    font-weight: 700;
    letter-spacing: 3.5px;
    text-transform: uppercase;
    color: var(--ouro);
    margin: 0 0 20px 0;
    display: flex;
    align-items: center;
    gap: 10px;
}
.rh-section::after {
    content: '';
    flex: 1;
    height: 1px;
    background: linear-gradient(90deg, rgba(200,169,81,0.4), transparent);
}

/* ── DASHBOARD CARDS ─────────────────── */
.rh-dash-card {
    background: var(--branco);
    border-radius: 0;
    padding: 24px 28px;
    margin-bottom: 2px;
    border-left: 3px solid transparent;
    transition: border-color 0.2s;
    position: relative;
}
.rh-dash-card:first-of-type { border-radius: 12px 12px 0 0; }
.rh-dash-card:last-of-type  { border-radius: 0 0 12px 12px; margin-bottom: 0; }
.rh-dash-card:hover { border-left-color: var(--ouro); }

.rh-dash-num {
    font-family: 'Bebas Neue', sans-serif !important;
    font-size: 1.1rem !important;
    letter-spacing: 3px;
    color: var(--ouro);
    margin-bottom: 4px;
}
.rh-dash-title {
    font-size: 1.05rem;
    font-weight: 700;
    color: var(--verde);
    margin-bottom: 12px;
    line-height: 1.2;
}
.rh-dash-desc {
    font-size: 0.88rem;
    color: var(--sub);
    line-height: 1.7;
}
.rh-tag {
    display: inline-block;
    background: #EDF7EE;
    color: var(--verde3);
    border: 1px solid rgba(46,125,50,0.18);
    font-size: 0.75rem;
    font-weight: 500;
    padding: 3px 10px;
    border-radius: 4px;
    margin: 2px 2px 2px 0;
}
.rh-note {
    font-size: 0.76rem;
    color: var(--ouro);
    margin-top: 10px;
    font-style: italic;
}
.rh-opt-badge {
    display: inline-block;
    background: rgba(200,169,81,0.12);
    border: 1px solid rgba(200,169,81,0.3);
    color: var(--ouro);
    font-size: 0.65rem;
    font-weight: 700;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    padding: 2px 8px;
    border-radius: 4px;
    margin-bottom: 6px;
}

/* ── UPLOAD LABELS ───────────────────── */
[data-testid="stFileUploader"] label,
[data-testid="stFileUploader"] > div > label {
    font-size: 1rem !important;
    font-weight: 600 !important;
    color: var(--verde) !important;
    margin-bottom: 6px !important;
}
[data-testid="stFileUploader"] {
    background: var(--branco) !important;
    border-radius: 10px !important;
    border: 1.5px dashed rgba(15,61,32,0.2) !important;
    padding: 6px 10px !important;
    margin-bottom: 12px !important;
}
[data-testid="stFileUploader"]:focus-within {
    border-color: var(--ouro) !important;
    box-shadow: 0 0 0 3px rgba(200,169,81,0.15) !important;
}

/* ── TEXT INPUT ──────────────────────── */
[data-testid="stTextInput"] label {
    font-size: 1rem !important;
    font-weight: 600 !important;
    color: var(--verde) !important;
}
[data-testid="stTextInput"] input {
    border-radius: 8px !important;
    border: 1.5px solid rgba(15,61,32,0.2) !important;
    font-size: 0.95rem !important;
}

/* ── BUTTON ──────────────────────────── */
.stButton > button[kind="primary"] {
    background: var(--verde) !important;
    color: var(--branco) !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'Outfit', sans-serif !important;
    font-size: 1rem !important;
    font-weight: 700 !important;
    letter-spacing: 0.5px !important;
    padding: 14px 28px !important;
    box-shadow: 0 4px 20px rgba(15,61,32,0.25) !important;
    transition: all 0.2s !important;
    position: relative !important;
    overflow: hidden !important;
}
.stButton > button[kind="primary"]::after {
    content: '';
    position: absolute;
    bottom: 0; left: 0; right: 0;
    height: 3px;
    background: var(--ouro);
}
.stButton > button[kind="primary"]:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 28px rgba(15,61,32,0.35) !important;
}

/* ── DOWNLOAD BUTTON ─────────────────── */
.stDownloadButton > button {
    background: transparent !important;
    color: var(--verde) !important;
    border: 2px solid var(--verde) !important;
    border-radius: 8px !important;
    font-family: 'Outfit', sans-serif !important;
    font-size: 0.95rem !important;
    font-weight: 600 !important;
    padding: 12px 24px !important;
}
.stDownloadButton > button:hover {
    background: var(--verde) !important;
    color: var(--branco) !important;
}

/* ── MÉTRICAS ────────────────────────── */
.rh-metrics {
    display: grid;
    grid-template-columns: repeat(4,1fr);
    gap: 10px;
    margin: 24px 0;
}
.rh-metric {
    background: var(--branco);
    border-radius: 10px;
    padding: 20px 14px;
    text-align: center;
    border: 1px solid var(--cinza);
    position: relative;
    overflow: hidden;
}
.rh-metric::after {
    content: '';
    position: absolute;
    bottom: 0; left: 0; right: 0;
    height: 3px;
}
.rh-metric.m-total::after  { background: var(--verde3); }
.rh-metric.m-crit::after   { background: #C62828; }
.rh-metric.m-atenc::after  { background: #E65100; }
.rh-metric.m-mon::after    { background: #F57F17; }
.rh-metric-num {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 3rem;
    line-height: 1;
    margin-bottom: 4px;
}
.rh-metric-label {
    font-size: 0.72rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    color: var(--sub);
}

/* ── DIVIDER ─────────────────────────── */
.rh-divider {
    height: 1px;
    background: linear-gradient(90deg, var(--ouro), transparent);
    margin: 24px 0;
    opacity: 0.35;
}

/* ── EXPANDER ────────────────────────── */
[data-testid="stExpander"] {
    border: 1px solid var(--cinza) !important;
    border-radius: 8px !important;
    background: var(--branco) !important;
}

/* ── DATAFRAME ───────────────────────── */
[data-testid="stDataFrame"] {
    border-radius: 10px !important;
    overflow: hidden !important;
    border: 1px solid var(--cinza) !important;
}

/* ── SUCCESS / INFO / WARNING ────────── */
[data-testid="stAlert"] {
    border-radius: 8px !important;
    font-size: 0.9rem !important;
}

/* ── FOOTER ──────────────────────────── */
.rh-footer {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 16px;
    padding: 32px 0 16px;
    border-top: 1px solid var(--cinza);
    margin-top: 48px;
    color: var(--sub);
    font-size: 0.8rem;
}
.rh-footer-dot {
    width: 4px; height: 4px;
    border-radius: 50%;
    background: var(--ouro);
    display: inline-block;
}
</style>
""", unsafe_allow_html=True)

# ── Configurações de e-mail ──────────────────────────────
DESTINATARIOS_FIXOS = ["rafael.ferraz@rehagro.edu.br"]
REMETENTE = "rafael.ferraz@rehagro.edu.br"


# ── Funções de processamento ─────────────────────────────

def carregar_canvas(arquivo):
    df = pd.read_excel(arquivo, skiprows=2)
    df.columns = df.columns.str.strip()
    col_nome   = next((c for c in df.columns if 'NOME' in c.upper()), None)
    col_email  = next((c for c in df.columns if 'E-MAIL' in c.upper() or 'EMAIL' in c.upper()), None)
    col_dias   = next((c for c in df.columns if 'DIAS' in c.upper()), None)
    col_curso  = next((c for c in df.columns if c.upper() == 'CURSO'), None)
    col_turma  = next((c for c in df.columns if c.upper() == 'TURMA'), None)
    col_funcao = next((c for c in df.columns if 'FUN' in c.upper() and 'DISCIPLINA' in c.upper()), None)
    if not col_nome or not col_dias:
        raise ValueError("Arquivo Canvas: colunas não encontradas.")
    # Filtrar apenas alunos (excluir colaboradores/professores)
    if col_funcao:
        df = df[df[col_funcao].astype(str).str.upper().str.contains('ALUNO', na=False)].copy()
    df['_key']  = (df[col_nome].astype(str).str.strip().str.lower() + '||' +
                   (df[col_turma].astype(str).str.strip().str.lower() if col_turma else ''))
    df['_dias'] = pd.to_numeric(df[col_dias], errors='coerce')
    cols = {'_key': '_key', col_nome: 'Nome', '_dias': 'Dias sem acesso'}
    if col_email: cols[col_email] = 'Email'
    if col_curso: cols[col_curso] = 'Curso'
    if col_turma: cols[col_turma] = 'Turma'
    # Critério: +20 dias sem acesso
    res = df[df['_dias'] > 20][list(cols.keys())].copy().rename(columns=cols)
    if 'Email' not in res.columns: res['Email'] = ''
    if 'Curso' not in res.columns: res['Curso'] = ''
    if 'Turma' not in res.columns: res['Turma'] = ''
    return res[['_key','Nome','Email','Curso','Turma','Dias sem acesso']]

def carregar_nps(arquivo, desistentes_keys=None):
    """
    Retorna dict com alertas NPS por aluno:
    - 'detrator_2': avaliou mal nos últimos 2 encontros
    - 'sem_avaliacao': presente nas últimas 2 aulas mas não avaliou nenhuma
    - 'comentario': fez qualquer comentário (com data, tópico e professor)
    Desistentes são excluídos se desistentes_keys for fornecido.
    """
    df = pd.read_excel(arquivo, skiprows=2)
    df.columns = df.columns.str.strip()

    col_aluno    = next((c for c in df.columns if c.upper() == 'ALUNO'), None)
    col_nps      = next((c for c in df.columns if 'NPS' in c.upper() and 'REA' in c.upper()), None)
    col_turma    = next((c for c in df.columns if c.upper() == 'TURMA'), None)
    col_topico   = next((c for c in df.columns if 'TÓPICO' in c.upper() or 'TOPICO' in c.upper()), None)
    col_prof     = next((c for c in df.columns if 'PROFESSOR' in c.upper() or 'PROF' in c.upper()), None)
    col_coment   = next((c for c in df.columns if 'COMMENT' in c.upper() or 'COMENT' in c.upper() or 'RESPOSTA' in c.upper()), None)

    col_data_unica = next((c for c in df.columns if c.upper() == 'DATA AULA' or (
        'DATA' in c.upper() and 'AULA' in c.upper() and
        'ANO' not in c.upper() and 'MÊS' not in c.upper() and 'DIA' not in c.upper())), None)
    col_dia     = next((c for c in df.columns if c.upper() in ('DIA', 'DIA AULA')), None)
    col_mes_ano = next((c for c in df.columns if 'MÊS' in c.upper() or 'MES' in c.upper()), None)

    if not col_aluno or not col_nps:
        raise ValueError("Arquivo NPS: colunas não encontradas.")

    # Normalizar data
    if col_data_unica:
        df['_data_dt'] = pd.to_datetime(df[col_data_unica], errors='coerce')
    elif col_dia and col_mes_ano:
        df['_data_str'] = (df[col_dia].astype(str).str.strip().str.zfill(2) + '/' +
                           df[col_mes_ano].astype(str).str.strip())
        df['_data_dt'] = pd.to_datetime(df['_data_str'], format='%d/%m/%Y', errors='coerce')
    else:
        df['_data_dt'] = pd.NaT

    df['_key'] = (df[col_aluno].astype(str).str.strip().str.lower() + '||' +
                  (df[col_turma].astype(str).str.strip().str.lower() if col_turma else ''))

    # Remover desistentes
    if desistentes_keys:
        df = df[~df['_key'].isin(desistentes_keys)].copy()

    df = df.sort_values(['_key', '_data_dt'])

    # Helper para formatar contexto de uma aula
    def ctx(row):
        data_str = row['_data_dt'].strftime('%d/%m/%Y') if pd.notna(row['_data_dt']) else 'N/D'
        topico   = str(row[col_topico]).strip() if col_topico and pd.notna(row.get(col_topico)) else ''
        prof     = str(row[col_prof]).strip() if col_prof and pd.notna(row.get(col_prof)) else ''
        partes   = [data_str]
        if topico: partes.append(topico)
        if prof:   partes.append(f"Prof. {prof}")
        return ' | '.join(partes)

    alertas_nps = {}  # _key -> lista de alertas

    for key, group in df.groupby('_key'):
        alertas = []

        # Últimas 2 aulas com NPS respondido
        respondidas = group[group[col_nps].notna()].sort_values('_data_dt')
        ult2_resp   = respondidas.tail(2)

        # Todas as aulas (para checar presença sem avaliação)
        todas_aulas = group.sort_values('_data_dt')
        ult2_todas  = todas_aulas.tail(2)

        # ── Critério 1: Detrator nos últimos 2 encontros avaliados ──
        if len(ult2_resp) == 2 and all(pd.to_numeric(ult2_resp[col_nps], errors='coerce') < 0):
            detalhes = [ctx(r) for _, r in ult2_resp.iterrows()]
            alertas.append({
                'tipo': 'detrator_2',
                'texto': f"Avaliou mal nos últimos 2 encontros: {' · '.join(detalhes)}",
                'acao': 'Retomar feedback negativo da avaliação'
            })

        # ── Critério 2: Presente nas últimas 2 aulas mas não avaliou nenhuma ──
        if len(ult2_todas) == 2 and all(ult2_todas[col_nps].isna()):
            detalhes = [ctx(r) for _, r in ult2_todas.iterrows()]
            alertas.append({
                'tipo': 'sem_avaliacao',
                'texto': f"Presente nas últimas 2 aulas sem avaliação: {' · '.join(detalhes)}",
                'acao': 'Incentivar participação nas avaliações de aula'
            })

        # ── Critério 3: Qualquer comentário ──
        if col_coment:
            com_coment = group[group[col_coment].notna() &
                               (group[col_coment].astype(str).str.strip() != '')].copy()
            if not com_coment.empty:
                itens = []
                for _, row in com_coment.iterrows():
                    c_ctx  = ctx(row)
                    c_text = str(row[col_coment]).strip()
                    itens.append(f"[{c_ctx}]: {c_text}")
                alertas.append({
                    'tipo': 'comentario',
                    'texto': 'Comentário(s) registrado(s): ' + ' | '.join(itens),
                    'acao': 'Analisar comentário e dar retorno ao aluno'
                })

        if alertas:
            alertas_nps[key] = alertas

    return alertas_nps

def carregar_comentarios(arquivo):
    df = pd.read_excel(arquivo, skiprows=2)
    df.columns = df.columns.str.strip()
    col_aluno    = next((c for c in df.columns if 'NOME' in c.upper() or c.upper() in ('NOME ALUNO', 'ALUNO')), None)
    col_turma    = next((c for c in df.columns if c.upper() == 'TURMA'), None)
    col_resposta = next((c for c in df.columns if 'RESPOSTA' in c.upper() or 'COMENT' in c.upper()), None)
    col_data     = next((c for c in df.columns if 'DATA' in c.upper()), None)
    if not col_aluno or not col_resposta:
        raise ValueError("Arquivo Comentários: colunas não encontradas.")
    df = df[df[col_resposta].notna() & (df[col_resposta].astype(str).str.strip() != '')].copy()
    df['_key'] = (df[col_aluno].astype(str).str.strip().str.lower() + '||' +
                  (df[col_turma].astype(str).str.strip().str.lower() if col_turma else ''))
    if col_data:
        df['_data_norm'] = pd.to_datetime(df[col_data], errors='coerce').dt.normalize()
    else:
        df['_data_norm'] = pd.NaT
    return df[['_key', '_data_norm', col_resposta]].rename(columns={col_resposta: 'Comentario'})

def carregar_frequencia(arquivo):
    """
    Retorna:
    - df_ausentes: alunos ausentes nas últimas 2 videoconferências
    - desistentes_keys: set de _keys de alunos desistentes (para excluir do NPS)
    """
    df = pd.read_excel(arquivo, skiprows=2)
    df.columns = df.columns.str.strip()
    col_aluno  = next((c for c in df.columns if c.upper() == 'ALUNO'), None)
    col_status = next((c for c in df.columns if 'STATUS' in c.upper() or 'PRESEN' in c.upper()), None)
    col_data   = next((c for c in df.columns if 'DATA' in c.upper() or 'PARTE' in c.upper()), None)
    col_turma  = next((c for c in df.columns if c.upper() == 'TURMA'), None)
    if not col_aluno or not col_status:
        raise ValueError("Arquivo Frequência: colunas não encontradas.")

    # Identificar desistentes
    df['_key_freq'] = (df[col_aluno].astype(str).str.strip().str.lower() + '||' +
                       (df[col_turma].astype(str).str.strip().str.lower() if col_turma else ''))
    desistentes_keys = set(
        df[df[col_status].str.upper() == 'DESISTENTE']['_key_freq'].unique()
    )

    # Processar apenas ativos (sem '-' e sem desistentes)
    df_a = df[~df[col_status].isin(['-','DESISTENTE'])].copy()

    def extr(s):
        m = re.match(r'(\d{2}/\d{2}/\d{4})', str(s))
        return pd.to_datetime(m.group(1), format='%d/%m/%Y') if m else pd.NaT

    if col_data:
        df_a['_data'] = df_a[col_data].apply(extr)
        df_a = df_a.sort_values([col_aluno, col_turma, '_data'] if col_turma else [col_aluno,'_data'])

    group_cols = [col_aluno, col_turma] if col_turma else [col_aluno]
    ausentes = []
    for keys, g in df_a.groupby(group_cols):
        aluno = keys[0] if col_turma else keys
        turma = keys[1] if col_turma else ''
        if len(g) < 2: continue
        last2 = g.tail(2)
        if all(last2[col_status] == 'AUSENTE'):
            datas = last2[col_data].tolist() if col_data else ['N/D','N/D']
            ausentes.append({
                '_key': f"{str(aluno).strip().lower()}||{str(turma).strip().lower()}",
                'Ultimas_2_aulas': f"{datas[0]} e {datas[1]}"
            })

    df_ausentes = pd.DataFrame(ausentes) if ausentes else pd.DataFrame(columns=['_key','Ultimas_2_aulas'])
    return df_ausentes, desistentes_keys

def gerar_relatorio(df_canvas, alertas_nps, df_freq, desistentes_keys=None):
    """
    Cruza todas as fontes e gera relatório final.
    alertas_nps: dict {_key: [lista de alertas]} retornado por carregar_nps()
    desistentes_keys: set de _keys de desistentes (excluídos do NPS, mas não do Canvas/Freq)
    """
    todos_keys = set(df_canvas['_key']) | set(alertas_nps.keys()) | set(df_freq['_key'])

    info_map = df_canvas.set_index('_key')[['Nome','Email','Curso','Turma']].to_dict('index')
    for key in todos_keys:
        if key not in info_map:
            partes = key.split('||')
            info_map[key] = {'Nome': partes[0].title(), 'Email': '', 'Curso': '',
                             'Turma': partes[1].upper() if len(partes) > 1 else ''}

    relatorio = []
    for key in sorted(todos_keys):
        info = info_map[key]
        alertas, acoes = [], []

        # ── Canvas: +20 dias sem acesso ──────────────────────────
        c = df_canvas[df_canvas['_key'] == key]
        if not c.empty:
            dias = int(c.iloc[0]['Dias sem acesso'])
            alertas.append(f"Sem acesso ao Canvas há {dias} dias")
            acoes.append("Enviar link de acesso à plataforma")

        # ── NPS: múltiplos critérios ──────────────────────────────
        if key in alertas_nps:
            for alerta in alertas_nps[key]:
                alertas.append(alerta['texto'])
                acoes.append(alerta['acao'])

        # ── Frequência: ausente nas últimas 2 videoconferências ──
        f = df_freq[df_freq['_key'] == key]
        if not f.empty:
            alertas.append(f"Ausente nas últimas 2 videoconferências ({f.iloc[0]['Ultimas_2_aulas']})")
            acoes.append("Enviar data da próxima aula ao vivo")

        if alertas:
            relatorio.append({
                'Curso':                 info['Curso'],
                'Turma':                 info['Turma'],
                'Nome':                  info['Nome'],
                'E-mail':                info['Email'],
                'Qtd. Alertas':          len(alertas),
                'Alertas Identificados': ' | '.join(alertas),
                'Ações Recomendadas':    ' | '.join(acoes),
            })

    df = pd.DataFrame(relatorio)
    return df.sort_values(['Curso','Turma','Qtd. Alertas','Nome'],
                          ascending=[True,True,False,True]).reset_index(drop=True)

def exportar_excel_bytes(df):
    wb = Workbook()
    ws = wb.active; ws.title = "Relatório CS"
    VE="0F3D20"; VM="1B5E35"; BR="FFFFFF"; CR="F4F0E6"; DO="C8A951"
    n_cols=8; lc=get_column_letter(n_cols)
    thin=Side(style='thin',color='E0DDD4')
    border=Border(left=thin,right=thin,top=thin,bottom=thin)
    def hcell(cell,txt,bg,fg=BR,sz=10,bold=True,align='center'):
        cell.value=txt
        cell.font=Font(bold=bold,size=sz,color=fg)
        cell.fill=PatternFill("solid",fgColor=bg)
        cell.alignment=Alignment(horizontal=align,vertical='center',wrap_text=True)
    ws.merge_cells(f'A1:{lc}1')
    hcell(ws['A1'],"RELATÓRIO DE ENGAJAMENTO — CUSTOMER SUCCESS REHAGRO",VE,sz=13)
    ws.row_dimensions[1].height=28
    ws.merge_cells(f'A2:{lc}2')
    hcell(ws['A2'],f"Gerado em: {datetime.now().strftime('%d/%m/%Y às %H:%M')}   |   Total desengajados: {len(df)}",VM,sz=9)
    ws.row_dimensions[2].height=18
    ws.append([])
    ws.merge_cells(f'A4:{lc}4')
    hcell(ws['A4'],"LEGENDA DE AÇÕES","EDF7EE",VE,sz=9,align='left')
    for i,txt in enumerate(["⚠️  Sem acesso ao Canvas há +20 dias  →  Enviar link de acesso à plataforma",
        "⚠️  NPS: Detrator nos últimos 2 encontros | Presente sem avaliar | Comentário registrado",
        "⚠️  Ausente nas últimas 2 aulas ao vivo  →  Enviar data da próxima aula"],5):
        ws.merge_cells(f'A{i}:{lc}{i}')
        ws[f'A{i}']=txt
        ws[f'A{i}'].font=Font(size=8,color="5A5A4A")
        ws[f'A{i}'].fill=PatternFill("solid",fgColor=CR)
        ws[f'A{i}'].alignment=Alignment(horizontal='left',indent=2)
    ws.append([])
    headers=['Curso','Turma','Nome do Aluno','E-mail','Qtd.','Alertas Identificados','Ações Recomendadas','Comentários do Aluno']
    ws.append(headers)
    hr=ws.max_row
    for ci,h in enumerate(headers,1):
        c=ws.cell(row=hr,column=ci)
        c.font=Font(bold=True,color=BR,size=9)
        c.fill=PatternFill("solid",fgColor=VE)
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.row_dimensions[hr].height=20
    for _,row in df.iterrows():
        ws.append([row['Curso'],row['Turma'],row['Nome'],row['E-mail'],
                   row['Qtd. Alertas'],row['Alertas Identificados'],
                   row['Ações Recomendadas'],row.get('Comentários do Aluno','')])
        dr=ws.max_row
        fc="FFCDD2" if row['Qtd. Alertas']>=3 else ("FFE0B2" if row['Qtd. Alertas']==2 else "FFFDE7")
        tem_c=bool(str(row.get('Comentários do Aluno','')).strip())
        for ci in range(1,n_cols+1):
            c=ws.cell(row=dr,column=ci)
            if ci<=2: c.fill=PatternFill("solid",fgColor="EDF7EE"); c.font=Font(size=8,color=VM,bold=True)
            elif ci==5: c.fill=PatternFill("solid",fgColor=fc); c.font=Font(bold=True,size=12); c.alignment=Alignment(horizontal='center',vertical='center')
            elif ci==8: c.fill=PatternFill("solid",fgColor="FFF8E1" if tem_c else BR); c.font=Font(size=8,color="5D4037" if tem_c else "AAAAAA",italic=not tem_c)
            else: c.fill=PatternFill("solid",fgColor=fc if ci>4 else BR)
            c.border=border
            if ci!=5: c.alignment=Alignment(vertical='center',wrap_text=True)
        ws.row_dimensions[dr].height=32
    for ci,w in enumerate([16,18,30,28,6,60,45,55],1):
        ws.column_dimensions[get_column_letter(ci)].width=w
    ws2=wb.create_sheet("Resumo por Turma")
    for ci,w in enumerate([20,20,16,12,12,12],1): ws2.column_dimensions[get_column_letter(ci)].width=w
    ws2.append(["RESUMO POR TURMA","","","","",""])
    ws2['A1'].font=Font(bold=True,size=12,color=BR); ws2['A1'].fill=PatternFill("solid",fgColor=VE)
    ws2.merge_cells('A1:F1'); ws2['A1'].alignment=Alignment(horizontal='center')
    ws2.row_dimensions[1].height=24
    ws2.append(["Curso","Turma","Total","🔴 Críticos","🟠 Atenção","🟡 Monitorar"])
    hr2=ws2.max_row
    for ci in range(1,7):
        c=ws2.cell(row=hr2,column=ci); c.font=Font(bold=True,color=BR,size=9)
        c.fill=PatternFill("solid",fgColor=VM); c.alignment=Alignment(horizontal='center')
    ws2.row_dimensions[hr2].height=18
    for (curso,turma),g in df.groupby(['Curso','Turma']):
        ws2.append([curso,turma,len(g),len(g[g['Qtd. Alertas']>=3]),len(g[g['Qtd. Alertas']==2]),len(g[g['Qtd. Alertas']==1])])
        r=ws2.max_row
        for ci in range(1,7):
            c=ws2.cell(row=r,column=ci); c.font=Font(size=9); c.border=border
            c.alignment=Alignment(horizontal='center' if ci>2 else 'left')
        ws2.row_dimensions[r].height=16
    ws2.append(["TOTAL GERAL","",len(df),len(df[df['Qtd. Alertas']>=3]),len(df[df['Qtd. Alertas']==2]),len(df[df['Qtd. Alertas']==1])])
    r=ws2.max_row
    for ci in range(1,7):
        c=ws2.cell(row=r,column=ci); c.font=Font(bold=True,size=9,color=BR)
        c.fill=PatternFill("solid",fgColor=VE); c.alignment=Alignment(horizontal='center' if ci>2 else 'left')
    ws2.row_dimensions[r].height=18
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf.read()

def enviar_email(excel_bytes,destinatarios,data_hoje,total,criticos,atencao,monitorar):
    api_key=st.secrets.get("SENDGRID_API_KEY","")
    if not api_key: return False,"Chave SendGrid não configurada."
    lista=list(set(destinatarios+DESTINATARIOS_FIXOS))
    assunto=f"Relatório CS Rehagro — Engajamento {data_hoje}"
    html=f"""<div style="font-family:Outfit,Arial,sans-serif;max-width:600px;margin:0 auto;background:#F4F0E6;">
      <div style="background:#0F3D20;padding:36px 44px;border-bottom:3px solid #C8A951;">
        <p style="color:#C8A951;font-size:9px;font-weight:700;letter-spacing:4px;text-transform:uppercase;margin:0 0 12px">Customer Success · Rehagro</p>
        <h1 style="color:#fff;font-size:28px;margin:0;font-family:Georgia,serif;letter-spacing:-0.5px;">Relatório de Engajamento</h1>
        <p style="color:rgba(255,255,255,0.5);margin:8px 0 0;font-size:13px;">Gerado em {data_hoje}</p>
      </div>
      <div style="background:#fff;padding:36px 44px;border:1px solid #E0DDD4;">
        <p style="color:#444;font-size:14px;margin:0 0 28px;line-height:1.6;">Segue em anexo o relatório semanal de engajamento com alertas, ações recomendadas e comentários dos alunos.</p>
        <table width="100%" cellpadding="0" cellspacing="8">
          <tr>
            <td style="background:#EDF7EE;border-radius:8px;padding:18px 12px;text-align:center;border-bottom:3px solid #2E7D32">
              <div style="font-size:32px;font-weight:800;color:#0F3D20;font-family:Georgia,serif;">{total}</div>
              <div style="font-size:9px;color:#888;text-transform:uppercase;letter-spacing:1.5px;margin-top:4px;">Total</div>
            </td>
            <td style="background:#FFEBEE;border-radius:8px;padding:18px 12px;text-align:center;border-bottom:3px solid #C62828">
              <div style="font-size:32px;font-weight:800;color:#C62828;font-family:Georgia,serif;">{criticos}</div>
              <div style="font-size:9px;color:#888;text-transform:uppercase;letter-spacing:1.5px;margin-top:4px;">🔴 Críticos</div>
            </td>
            <td style="background:#FFF3E0;border-radius:8px;padding:18px 12px;text-align:center;border-bottom:3px solid #E65100">
              <div style="font-size:32px;font-weight:800;color:#E65100;font-family:Georgia,serif;">{atencao}</div>
              <div style="font-size:9px;color:#888;text-transform:uppercase;letter-spacing:1.5px;margin-top:4px;">🟠 Atenção</div>
            </td>
            <td style="background:#FFFDE7;border-radius:8px;padding:18px 12px;text-align:center;border-bottom:3px solid #F57F17">
              <div style="font-size:32px;font-weight:800;color:#F57F17;font-family:Georgia,serif;">{monitorar}</div>
              <div style="font-size:9px;color:#888;text-transform:uppercase;letter-spacing:1.5px;margin-top:4px;">🟡 Monitorar</div>
            </td>
          </tr>
        </table>
        <p style="color:#999;font-size:11px;margin-top:28px;border-top:1px solid #F0EDE4;padding-top:16px;">O arquivo Excel em anexo contém o relatório completo.</p>
      </div>
      <div style="background:#F4F0E6;padding:16px;text-align:center;border:1px solid #E0DDD4;border-top:none;">
        <p style="color:#aaa;font-size:10px;margin:0;letter-spacing:1px;">REHAGRO · CUSTOMER SUCCESS · AGENTE DE ENGAJAMENTO</p>
      </div>
    </div>"""
    encoded=base64.b64encode(excel_bytes).decode()
    nome_arq=f"relatorio_cs_{data_hoje.replace('/','')}.xlsx"
    message=Mail(from_email=REMETENTE,to_emails=lista[0],subject=assunto,html_content=html)
    if len(lista)>1:
        for dest in lista[1:]: message.add_cc(dest)
    message.attachment=Attachment(FileContent(encoded),FileName(nome_arq),
        FileType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),Disposition("attachment"))
    try:
        sg=SendGridAPIClient(api_key); res=sg.send(message)
        if res.status_code in [200,201,202]: return True,f"E-mail enviado para {len(lista)} destinatário(s)."
        return False,f"Erro SendGrid: status {res.status_code}"
    except Exception as e: return False,str(e)


# ── UI ────────────────────────────────────────────────────

# HERO
st.markdown("""
<div class="rh-hero">
  <div class="rh-hero-bg"></div>
  <div class="rh-hero-grid"></div>
  <div class="rh-hero-inner">
    <p class="rh-eyebrow">Rehagro · Customer Success</p>
    <h1 class="rh-hero-title">MONITORAMENTO <span>DE ALUNOS</span></h1>
    <p class="rh-hero-sub">Identifique alunos desengajados e saiba exatamente qual ação tomar — em segundos.</p>
    <div class="rh-hero-pills">
      <span class="rh-pill">Canvas AVA</span>
      <span class="rh-pill">NPS Avaliações</span>
      <span class="rh-pill">Frequência ao Vivo</span>
      <span class="rh-pill">Comentários</span>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="rh-body">', unsafe_allow_html=True)

col_esq, col_dir = st.columns([1, 1], gap="large")

with col_esq:
    st.markdown('<p class="rh-section">Como exportar do Power BI</p>', unsafe_allow_html=True)
    st.markdown("""
    <div style="background:#FFF8E1;border:1.5px solid #C8A951;border-radius:10px;padding:14px 18px;margin-bottom:16px;display:flex;gap:12px;align-items:flex-start;">
      <span style="font-size:1.2rem;line-height:1.4;">⚠️</span>
      <div>
        <div style="font-size:0.82rem;font-weight:700;color:#5D4037;letter-spacing:0.5px;margin-bottom:4px;text-transform:uppercase;">Atenção antes de exportar</div>
        <div style="font-size:0.85rem;color:#6D4C41;line-height:1.6;">
          Todos os arquivos devem ser exportados no formato <b>Dados Resumidos</b>.<br>
          Os <b>filtros necessários</b> para cada extração estão descritos em cada ponto abaixo — siga as instruções de cada dashboard.
        </div>
      </div>
    </div>
    <div>
      <div class="rh-dash-card" style="border-radius:12px 12px 0 0">
        <div class="rh-dash-num" style="font-size:1.1rem!important">Dashboard 01</div>
        <div class="rh-dash-title">Acesso ao Canvas (AVA)</div>
        <div class="rh-dash-desc">
          Relatório <b>Rehagro - Canvas</b> → página <b>Acesso ao Canvas-Ok</b><br><br>
          <span class="rh-tag">Status Usuário Curso = Ativo</span>
          <span class="rh-tag">Função na disciplina = Aluno</span>
          <span class="rh-tag">Curso = seus cursos</span>
          <span class="rh-tag">Turma = suas turmas</span><br><br>
          Exporte em formato <b>Dados Resumidos</b> → Excel
          <p class="rh-note">💡 Múltiplos cursos e turmas são suportados.</p>
        </div>
      </div>
      <div class="rh-dash-card">
        <div class="rh-dash-num" style="font-size:1.1rem!important">Dashboard 02</div>
        <div class="rh-dash-title">NPS Médio por Aluno</div>
        <div class="rh-dash-desc">
          Relatório <b>Rehagro Educação - Avaliação de Aula</b> → página <b>Avaliações de aula/aluno</b><br><br>
          <span class="rh-tag">Curso = seus cursos</span>
          <span class="rh-tag">Turma = suas turmas</span>
          <span class="rh-tag">Ano resposta = ano atual</span>
          <span class="rh-tag">Ano/mês resposta = período</span>
          <span class="rh-tag">Tipo de aula = On-line ao vivo</span><br><br>
          Exporte a tabela <b>NPS médio/aluno</b> → Excel
          <p class="rh-note">💡 Extraia da página "NPS médio/aluno" do dashboard.</p>
        </div>
      </div>
      <div class="rh-dash-card">
        <div class="rh-dash-num" style="font-size:1.1rem!important">Dashboard 03</div>
        <div class="rh-dash-title">Frequência nas Aulas ao Vivo</div>
        <div class="rh-dash-desc">
          Relatório <b>Rehagro Alunado</b> → página <b>Análise de Frequência e Faltas</b><br><br>
          <span class="rh-tag">Turma = suas turmas</span>
          <span class="rh-tag">Data/Aula = período desejado</span><br><br>
          Exporte a <b>Tabela de frequência</b> → Excel
        </div>
      </div>
      <div class="rh-dash-card" style="border-radius:0 0 12px 12px">
        <div class="rh-opt-badge">Opcional</div>
        <div class="rh-dash-num" style="font-size:1.1rem!important">Dashboard 04</div>
        <div class="rh-dash-title">Comentários das Aulas</div>
        <div class="rh-dash-desc">
          Mesma página do Dashboard 02 → <b>Tabela de Comentários</b> → Excel
          <p class="rh-note">💡 Enriquece o relatório com o contexto do feedback.</p>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

with col_dir:
    st.markdown('<p class="rh-section">Envie os arquivos</p>', unsafe_allow_html=True)

    f_canvas = st.file_uploader("1 — Canvas · Acesso ao AVA",           type=["xlsx"], key="canvas")
    f_nps    = st.file_uploader("2 — NPS · Avaliações de aula",          type=["xlsx"], key="nps")
    f_freq   = st.file_uploader("3 — Frequência · Aulas ao vivo",        type=["xlsx"], key="freq")
    f_coment = st.file_uploader("4 — Comentários · Opcional",            type=["xlsx"], key="coment")

    st.markdown('<p class="rh-section" style="margin-top:28px">Envio por e-mail</p>', unsafe_allow_html=True)
    email_usuario = st.text_input("Seu e-mail — receberá cópia junto à lista fixa:",
                                  placeholder="seuemail@rehagro.com.br")
    with st.expander("Ver destinatários fixos"):
        for d in DESTINATARIOS_FIXOS: st.write(f"• {d}")

    st.markdown('<div class="rh-divider"></div>', unsafe_allow_html=True)

    obrigatorios = f_canvas and f_nps and f_freq
    if not obrigatorios:
        faltando = [n for f,n in [(f_canvas,"Canvas"),(f_nps,"NPS"),(f_freq,"Frequência")] if not f]
        st.info(f"Aguardando: **{', '.join(faltando)}**")
    else:
        st.success("✅ Arquivos prontos — clique para gerar e enviar.")
        if st.button("Gerar e Enviar Relatório →", type="primary", use_container_width=True):
            with st.spinner("Analisando dados..."):
                try:
                    dc                    = carregar_canvas(f_canvas)
                    df_fr, desist_keys    = carregar_frequencia(f_freq)
                    alertas_nps           = carregar_nps(f_nps, desistentes_keys=desist_keys)
                    df_rel                = gerar_relatorio(dc, alertas_nps, df_fr, desist_keys)

                    criticos  = len(df_rel[df_rel['Qtd. Alertas']>=3])
                    atencao   = len(df_rel[df_rel['Qtd. Alertas']==2])
                    monitorar = len(df_rel[df_rel['Qtd. Alertas']==1])

                    st.markdown(f"""
                    <div class="rh-metrics">
                      <div class="rh-metric m-total">
                        <div class="rh-metric-num" style="color:#0F3D20">{len(df_rel)}</div>
                        <div class="rh-metric-label">Total</div>
                      </div>
                      <div class="rh-metric m-crit">
                        <div class="rh-metric-num" style="color:#C62828">{criticos}</div>
                        <div class="rh-metric-label">🔴 Críticos</div>
                      </div>
                      <div class="rh-metric m-atenc">
                        <div class="rh-metric-num" style="color:#E65100">{atencao}</div>
                        <div class="rh-metric-label">🟠 Atenção</div>
                      </div>
                      <div class="rh-metric m-mon">
                        <div class="rh-metric-num" style="color:#F57F17">{monitorar}</div>
                        <div class="rh-metric-label">🟡 Monitorar</div>
                      </div>
                    </div>
                    """, unsafe_allow_html=True)

                    turmas = sorted(df_rel['Turma'].dropna().unique().tolist())
                    df_view = df_rel
                    if len(turmas) > 1:
                        sel = st.multiselect("Filtrar por turma:", turmas, default=turmas)
                        df_view = df_rel[df_rel['Turma'].isin(sel)]

                    st.markdown('<p class="rh-section">Alunos desengajados</p>', unsafe_allow_html=True)
                    cols_view = ['Curso','Turma','Nome','Qtd. Alertas','Alertas Identificados','Ações Recomendadas']
                    st.dataframe(df_view[cols_view], use_container_width=True, hide_index=True, height=320)

                    excel_bytes = exportar_excel_bytes(df_rel)
                    data_hoje   = datetime.now().strftime('%d/%m/%Y')
                    dests = [email_usuario.strip()] if email_usuario and "@" in email_usuario else []
                    if SENDGRID_OK:
                        ok,msg = enviar_email(excel_bytes,dests,data_hoje,len(df_rel),criticos,atencao,monitorar)
                        if ok:
                            st.success(f"📧 {msg}")
                        else:
                            st.warning(f"⚠️ E-mail não enviado: {msg}")
                    else:
                        st.warning("⚠️ Biblioteca SendGrid não instalada.")

                    st.download_button("⬇  Baixar Relatório Excel", data=excel_bytes,
                        file_name=f"relatorio_cs_{datetime.now().strftime('%d%m%Y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)
                except Exception as e:
                    st.error(f"Erro ao processar: {e}")
                    st.info("Verifique se os arquivos corretos foram enviados com os filtros indicados.")

st.markdown('</div>', unsafe_allow_html=True)
st.markdown("""
<div class="rh-footer">
  <span>Rehagro</span>
  <span class="rh-footer-dot"></span>
  <span>Customer Success</span>
  <span class="rh-footer-dot"></span>
  <span>Agente de Engajamento</span>
  <span class="rh-footer-dot"></span>
  <span>© 2026</span>
</div>
""", unsafe_allow_html=True)
