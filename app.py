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

# ── Identidade Visual Rehagro ──────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;800&family=Inter:wght@300;400;500;600&display=swap');

:root {
  --verde:    #1B4D2E;
  --verde2:   #2E7D32;
  --verde3:   #4CAF50;
  --dourado:  #C8A951;
  --dourado2: #F0D080;
  --creme:    #F9F6EF;
  --cinza:    #F2EFE8;
  --texto:    #1a1a1a;
  --sub:      #6B6B5E;
}

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    background: var(--creme) !important;
}
.stApp { background: var(--creme) !important; }

/* Esconde o menu padrão do streamlit */
#MainMenu, footer { visibility: hidden; }

/* ── HERO ── */
.rh-hero {
    background: var(--verde);
    border-radius: 0 0 32px 32px;
    padding: 36px 48px 40px;
    margin: -1rem -1rem 2rem -1rem;
    position: relative;
    overflow: hidden;
}
.rh-hero::before {
    content: '';
    position: absolute; inset: 0;
    background: url("data:image/svg+xml,%3Csvg width='60' height='60' viewBox='0 0 60 60' xmlns='http://www.w3.org/2000/svg'%3E%3Cg fill='none' fill-rule='evenodd'%3E%3Cg fill='%23ffffff' fill-opacity='0.03'%3E%3Cpath d='M36 34v-4h-2v4h-4v2h4v4h2v-4h4v-2h-4zm0-30V0h-2v4h-4v2h4v4h2V6h4V4h-4zM6 34v-4H4v4H0v2h4v4h2v-4h4v-2H6zM6 4V0H4v4H0v2h4v4h2V6h4V4H6z'/%3E%3C/g%3E%3C/g%3E%3C/svg%3E");
}
.rh-hero::after {
    content: '';
    position: absolute;
    right: -80px; top: -80px;
    width: 320px; height: 320px;
    border-radius: 50%;
    background: radial-gradient(circle, rgba(200,169,81,0.2) 0%, transparent 70%);
}
.rh-logo {
    font-family: 'Playfair Display', serif;
    font-size: 1.6rem;
    font-weight: 800;
    color: #ffffff;
    letter-spacing: -0.5px;
    display: flex;
    align-items: center;
    gap: 8px;
    margin-bottom: 20px;
}
.rh-logo-leaf {
    display: inline-block;
    width: 28px; height: 28px;
    background: var(--dourado);
    border-radius: 50% 0 50% 0;
    transform: rotate(-15deg);
}
.rh-badge {
    display: inline-block;
    background: rgba(200,169,81,0.25);
    border: 1px solid rgba(200,169,81,0.5);
    color: var(--dourado2);
    font-size: 0.65rem;
    font-weight: 600;
    letter-spacing: 2px;
    text-transform: uppercase;
    padding: 3px 12px;
    border-radius: 20px;
    margin-bottom: 10px;
}
.rh-title {
    font-family: 'Playfair Display', serif;
    font-size: 2.2rem;
    font-weight: 800;
    color: #ffffff;
    margin: 0 0 8px 0;
    line-height: 1.1;
}
.rh-sub {
    color: rgba(255,255,255,0.65);
    font-size: 0.95rem;
    font-weight: 300;
    margin: 0;
}

/* ── CARDS ── */
.rh-card {
    background: #ffffff;
    border-radius: 16px;
    padding: 22px 26px;
    border: 1px solid rgba(27,77,46,0.08);
    margin-bottom: 14px;
    box-shadow: 0 2px 12px rgba(27,77,46,0.06);
    position: relative;
    overflow: hidden;
}
.rh-card::before {
    content: '';
    position: absolute;
    left: 0; top: 0; bottom: 0;
    width: 4px;
    background: linear-gradient(180deg, var(--dourado), var(--verde2));
    border-radius: 4px 0 0 4px;
}
.rh-step-num {
    font-size: 0.65rem;
    font-weight: 700;
    letter-spacing: 2.5px;
    text-transform: uppercase;
    color: var(--dourado);
    margin-bottom: 4px;
}
.rh-step-title {
    font-family: 'Playfair Display', serif;
    font-size: 1.05rem;
    font-weight: 700;
    color: var(--verde);
    margin-bottom: 10px;
}
.rh-step-desc {
    font-size: 0.82rem;
    color: var(--sub);
    line-height: 1.65;
}
.rh-tag {
    display: inline-block;
    background: #EEF7F0;
    color: var(--verde2);
    border: 1px solid rgba(46,125,50,0.2);
    font-size: 0.72rem;
    font-weight: 500;
    padding: 2px 9px;
    border-radius: 4px;
    margin: 2px 2px 2px 0;
}
.rh-note {
    font-size: 0.75rem;
    color: var(--dourado);
    font-style: italic;
    margin-top: 8px;
    display: flex;
    align-items: center;
    gap: 4px;
}

/* ── SECTION TITLE ── */
.rh-section {
    font-family: 'Playfair Display', serif;
    font-size: 1.05rem;
    font-weight: 700;
    color: var(--verde);
    margin: 20px 0 12px 0;
    padding-bottom: 8px;
    border-bottom: 2px solid var(--dourado);
    display: flex;
    align-items: center;
    gap: 8px;
}

/* ── MÉTRICAS ── */
.rh-metrics {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 12px;
    margin-bottom: 20px;
}
.rh-metric {
    background: #ffffff;
    border-radius: 14px;
    padding: 18px 12px;
    text-align: center;
    border: 1px solid rgba(27,77,46,0.08);
    box-shadow: 0 2px 8px rgba(27,77,46,0.05);
}
.rh-metric-num {
    font-family: 'Playfair Display', serif;
    font-size: 2.4rem;
    font-weight: 800;
    line-height: 1;
    margin-bottom: 4px;
}
.rh-metric-label {
    font-size: 0.7rem;
    color: var(--sub);
    text-transform: uppercase;
    letter-spacing: 1px;
    font-weight: 500;
}

/* ── DIVIDER ── */
.rh-divider {
    height: 1px;
    background: linear-gradient(90deg, var(--dourado), transparent);
    margin: 20px 0;
    opacity: 0.4;
}

/* ── UPLOAD AREA ── */
[data-testid="stFileUploader"] {
    background: #ffffff !important;
    border-radius: 12px !important;
    border: 1.5px dashed rgba(27,77,46,0.2) !important;
    padding: 4px 8px !important;
}

/* ── BUTTON ── */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, var(--verde), var(--verde2)) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    letter-spacing: 0.3px !important;
    padding: 12px 24px !important;
    box-shadow: 0 4px 15px rgba(27,77,46,0.3) !important;
}

/* ── FOOTER ── */
.rh-footer {
    text-align: center;
    color: var(--sub);
    font-size: 0.75rem;
    padding: 24px 0 8px;
    border-top: 1px solid rgba(27,77,46,0.1);
    margin-top: 32px;
}
</style>
""", unsafe_allow_html=True)

# ── Configurações de e-mail ───────────────────────────────
DESTINATARIOS_FIXOS = [
    "rafael.ferraz@rehagro.edu.br",
]
REMETENTE = "rafael.ferraz@rehagro.edu.br"


# ── Funções de processamento ─────────────────────────────

def carregar_canvas(arquivo):
    df = pd.read_excel(arquivo, skiprows=2)
    df.columns = df.columns.str.strip()
    col_nome  = next((c for c in df.columns if 'NOME' in c.upper()), None)
    col_email = next((c for c in df.columns if 'E-MAIL' in c.upper() or 'EMAIL' in c.upper()), None)
    col_dias  = next((c for c in df.columns if 'DIAS' in c.upper()), None)
    col_curso = next((c for c in df.columns if c.upper() == 'CURSO'), None)
    col_turma = next((c for c in df.columns if c.upper() == 'TURMA'), None)
    if not col_nome or not col_dias:
        raise ValueError("Arquivo Canvas: colunas não encontradas.")
    df['_key']  = (df[col_nome].astype(str).str.strip().str.lower() + '||' +
                   (df[col_turma].astype(str).str.strip().str.lower() if col_turma else ''))
    df['_dias'] = pd.to_numeric(df[col_dias], errors='coerce')
    cols = {'_key': '_key', col_nome: 'Nome', '_dias': 'Dias sem acesso'}
    if col_email: cols[col_email] = 'Email'
    if col_curso: cols[col_curso] = 'Curso'
    if col_turma: cols[col_turma] = 'Turma'
    res = df[df['_dias'] > 30][list(cols.keys())].copy().rename(columns=cols)
    if 'Email' not in res.columns: res['Email'] = ''
    if 'Curso' not in res.columns: res['Curso'] = ''
    if 'Turma' not in res.columns: res['Turma'] = ''
    return res[['_key', 'Nome', 'Email', 'Curso', 'Turma', 'Dias sem acesso']]


def carregar_nps(arquivo):
    df = pd.read_excel(arquivo, skiprows=2)
    df.columns = df.columns.str.strip()
    col_aluno = next((c for c in df.columns if c.upper() == 'ALUNO'), None)
    col_nps   = next((c for c in df.columns if 'NPS' in c.upper() and 'REA' in c.upper()), None)
    col_turma = next((c for c in df.columns if c.upper() == 'TURMA'), None)
    col_data  = next((c for c in df.columns if c.upper() == 'DATA AULA' or (
        'DATA' in c.upper() and 'AULA' in c.upper() and
        'ANO' not in c.upper() and 'MÊS' not in c.upper() and 'DIA' not in c.upper())), None)
    if not col_aluno or not col_nps:
        raise ValueError("Arquivo NPS: colunas não encontradas.")
    df_v = df[df[col_nps].notna()].copy()
    if col_data:
        df_v['_data_dt'] = pd.to_datetime(df_v[col_data], errors='coerce')
        df_v = df_v.sort_values([col_aluno, '_data_dt'])
    df_v['_key'] = (df_v[col_aluno].astype(str).str.strip().str.lower() + '||' +
                    (df_v[col_turma].astype(str).str.strip().str.lower() if col_turma else ''))
    ultima = df_v.groupby('_key').last().reset_index()
    det = ultima[pd.to_numeric(ultima[col_nps], errors='coerce') < 0].copy()
    det = det[['_key', col_nps] + ([col_data] if col_data else [])].copy()
    det.columns = ['_key', 'NPS_valor'] + (['NPS_data'] if col_data else [])
    if 'NPS_data' not in det.columns:
        det['NPS_data'] = ''
    return det


def carregar_comentarios(arquivo):
    """Carrega comentários de aula — opcional."""
    df = pd.read_excel(arquivo, skiprows=2)
    df.columns = df.columns.str.strip()
    col_aluno    = next((c for c in df.columns if 'NOME' in c.upper() or c.upper() == 'NOME ALUNO'), None)
    col_turma    = next((c for c in df.columns if c.upper() == 'TURMA'), None)
    col_resposta = next((c for c in df.columns if 'RESPOSTA' in c.upper() or 'COMENTÁRIO' in c.upper() or 'COMENT' in c.upper()), None)
    col_data     = next((c for c in df.columns if 'DATA' in c.upper()), None)
    col_aula     = next((c for c in df.columns if 'COMENTÁRIO' in c.upper() or 'AULA' in c.upper() or 'COMMENT' in c.upper()), None)
    if not col_aluno or not col_resposta:
        raise ValueError("Arquivo Comentários: colunas não encontradas.")
    df = df[df[col_resposta].notna() & (df[col_resposta].astype(str).str.strip() != '')].copy()
    df['_key'] = (df[col_aluno].astype(str).str.strip().str.lower() + '||' +
                  (df[col_turma].astype(str).str.strip().str.lower() if col_turma else ''))
    # Agrupa múltiplos comentários do mesmo aluno
    def agrupar(group):
        comentarios = []
        for _, row in group.iterrows():
            data_str = ''
            if col_data and pd.notna(row.get(col_data)):
                try:
                    data_str = pd.to_datetime(row[col_data]).strftime('%d/%m/%Y')
                except:
                    data_str = str(row[col_data])[:10]
            aula_str = str(row[col_aula])[:60] + '...' if col_aula and pd.notna(row.get(col_aula)) and len(str(row[col_aula])) > 60 else (str(row[col_aula]) if col_aula and pd.notna(row.get(col_aula)) else '')
            linha = f"[{data_str}] {aula_str}: {str(row[col_resposta]).strip()}" if data_str else str(row[col_resposta]).strip()
            comentarios.append(linha)
        return ' | '.join(comentarios)
    resultado = df.groupby('_key').apply(agrupar).reset_index()
    resultado.columns = ['_key', 'Comentarios']
    return resultado


def carregar_frequencia(arquivo):
    df = pd.read_excel(arquivo, skiprows=2)
    df.columns = df.columns.str.strip()
    col_aluno  = next((c for c in df.columns if c.upper() == 'ALUNO'), None)
    col_status = next((c for c in df.columns if 'STATUS' in c.upper() or 'PRESEN' in c.upper()), None)
    col_data   = next((c for c in df.columns if 'DATA' in c.upper() or 'PARTE' in c.upper()), None)
    col_turma  = next((c for c in df.columns if c.upper() == 'TURMA'), None)
    if not col_aluno or not col_status:
        raise ValueError("Arquivo Frequência: colunas não encontradas.")
    df_a = df[~df[col_status].isin(['-', 'DESISTENTE'])].copy()
    def extr(s):
        m = re.match(r'(\d{2}/\d{2}/\d{4})', str(s))
        return pd.to_datetime(m.group(1), format='%d/%m/%Y') if m else pd.NaT
    if col_data:
        df_a['_data'] = df_a[col_data].apply(extr)
        df_a = df_a.sort_values([col_aluno, col_turma, '_data'] if col_turma else [col_aluno, '_data'])
    group_cols = [col_aluno, col_turma] if col_turma else [col_aluno]
    ausentes = []
    for keys, g in df_a.groupby(group_cols):
        aluno = keys[0] if col_turma else keys
        turma = keys[1] if col_turma else ''
        if len(g) < 2: continue
        last2 = g.tail(2)
        if all(last2[col_status] == 'AUSENTE'):
            datas = last2[col_data].tolist() if col_data else ['N/D', 'N/D']
            ausentes.append({
                '_key': f"{str(aluno).strip().lower()}||{str(turma).strip().lower()}",
                'Ultimas_2_aulas': f"{datas[0]} e {datas[1]}"
            })
    return pd.DataFrame(ausentes) if ausentes else pd.DataFrame(columns=['_key', 'Ultimas_2_aulas'])


def gerar_relatorio(df_canvas, df_nps, df_freq, df_coment=None):
    todos_keys = set(df_canvas['_key']) | set(df_nps['_key']) | set(df_freq['_key'])
    info_map = df_canvas.set_index('_key')[['Nome', 'Email', 'Curso', 'Turma']].to_dict('index')
    for key in todos_keys:
        if key not in info_map:
            partes = key.split('||')
            info_map[key] = {'Nome': partes[0].title(), 'Email': '',
                             'Curso': '', 'Turma': partes[1].upper() if len(partes) > 1 else ''}
    coment_map = {}
    if df_coment is not None and not df_coment.empty:
        coment_map = df_coment.set_index('_key')['Comentarios'].to_dict()

    relatorio = []
    for key in sorted(todos_keys):
        info = info_map[key]
        alertas, acoes = [], []

        c = df_canvas[df_canvas['_key'] == key]
        if not c.empty:
            dias = int(c.iloc[0]['Dias sem acesso'])
            alertas.append(f"Sem acesso ao Canvas há {dias} dias")
            acoes.append("Enviar link de acesso à plataforma")

        n = df_nps[df_nps['_key'] == key]
        if not n.empty:
            nps_val  = int(n.iloc[0]['NPS_valor'])
            nps_data = ''
            if 'NPS_data' in n.columns and pd.notna(n.iloc[0]['NPS_data']):
                try:
                    nps_data = pd.to_datetime(n.iloc[0]['NPS_data']).strftime('%d/%m/%Y')
                except:
                    nps_data = str(n.iloc[0]['NPS_data'])[:10]
            label = f"Detrator (NPS {nps_val})"
            if nps_data:
                label += f" — avaliação em {nps_data}"
            alertas.append(f"Última avaliação: {label}")
            acoes.append("Retomar feedback negativo da avaliação")

        f = df_freq[df_freq['_key'] == key]
        if not f.empty:
            alertas.append(f"Ausente nas últimas 2 aulas ao vivo ({f.iloc[0]['Ultimas_2_aulas']})")
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
                'Comentários do Aluno':  coment_map.get(key, ''),
            })

    df = pd.DataFrame(relatorio)
    return df.sort_values(['Curso', 'Turma', 'Qtd. Alertas', 'Nome'],
                          ascending=[True, True, False, True]).reset_index(drop=True)


def exportar_excel_bytes(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Relatório CS"
    VE = "1B4D2E"; VM = "2E7D32"; BR = "FFFFFF"; CR = "F9F6EF"
    DO = "C8A951"; n_cols = 8
    lc = get_column_letter(n_cols)

    def hr_cell(cell, txt, bg, fg=BR, sz=11, bold=True):
        cell.value = txt
        cell.font = Font(bold=bold, size=sz, color=fg)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws.merge_cells(f'A1:{lc}1')
    hr_cell(ws['A1'], "RELATÓRIO DE ENGAJAMENTO — CUSTOMER SUCCESS REHAGRO", VE, sz=13)
    ws.row_dimensions[1].height = 28

    ws.merge_cells(f'A2:{lc}2')
    hr_cell(ws['A2'],
            f"Gerado em: {datetime.now().strftime('%d/%m/%Y às %H:%M')}   |   Total de alunos desengajados: {len(df)}",
            VM, sz=9)
    ws.row_dimensions[2].height = 18

    ws.append([])
    ws.merge_cells(f'A4:{lc}4')
    hr_cell(ws['A4'], "LEGENDA DE AÇÕES", "EEF7F0", VE, sz=9)
    ws['A4'].alignment = Alignment(horizontal='left', indent=1)

    thin = Side(style='thin', color='DEDAD2')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for i, txt in enumerate([
        "⚠️  Sem acesso ao Canvas há +30 dias  →  Enviar link de acesso à plataforma",
        "⚠️  Última avaliação com NPS negativo (nota e data indicadas)  →  Retomar feedback negativo",
        "⚠️  Ausente nas últimas 2 aulas ao vivo  →  Enviar data da próxima aula ao vivo",
    ], 5):
        ws.merge_cells(f'A{i}:{lc}{i}')
        ws[f'A{i}'] = txt
        ws[f'A{i}'].font = Font(size=8, color="6B6B5E")
        ws[f'A{i}'].fill = PatternFill("solid", fgColor=CR)
        ws[f'A{i}'].alignment = Alignment(horizontal='left', indent=2)

    ws.append([])
    headers = ['Curso', 'Turma', 'Nome do Aluno', 'E-mail',
               'Qtd.', 'Alertas Identificados', 'Ações Recomendadas', 'Comentários do Aluno']
    ws.append(headers)
    hr = ws.max_row
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=hr, column=ci)
        c.font = Font(bold=True, color=BR, size=9)
        c.fill = PatternFill("solid", fgColor=VE)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[hr].height = 20

    for _, row in df.iterrows():
        ws.append([row['Curso'], row['Turma'], row['Nome'], row['E-mail'],
                   row['Qtd. Alertas'], row['Alertas Identificados'],
                   row['Ações Recomendadas'], row.get('Comentários do Aluno', '')])
        dr = ws.max_row
        fc = "FFCDD2" if row['Qtd. Alertas'] >= 3 else ("FFE0B2" if row['Qtd. Alertas'] == 2 else "FFFDE7")
        tem_coment = bool(str(row.get('Comentários do Aluno', '')).strip())
        for ci in range(1, n_cols + 1):
            c = ws.cell(row=dr, column=ci)
            if ci <= 2:
                c.fill = PatternFill("solid", fgColor="EEF7F0")
                c.font = Font(size=8, color=VM, bold=True)
            elif ci == 5:
                c.fill = PatternFill("solid", fgColor=fc)
                c.font = Font(bold=True, size=12)
                c.alignment = Alignment(horizontal='center', vertical='center')
            elif ci == 8:
                # Comentários — destaque dourado se houver conteúdo
                c.fill = PatternFill("solid", fgColor="FFF8E1" if tem_coment else BR)
                c.font = Font(size=8, color="5D4037" if tem_coment else "AAAAAA",
                              italic=not tem_coment)
            else:
                c.fill = PatternFill("solid", fgColor=fc if ci > 4 else BR)
            c.border = border
            if ci not in [5]:
                c.alignment = Alignment(vertical='center', wrap_text=True)
        ws.row_dimensions[dr].height = 32

    for ci, w in enumerate([16, 18, 30, 28, 6, 60, 45, 55], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    # Aba Resumo
    ws2 = wb.create_sheet("Resumo por Turma")
    for ci, w in enumerate([20, 20, 16, 12, 12, 12], 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w
    ws2.append(["RESUMO POR TURMA", "", "", "", "", ""])
    ws2['A1'].font = Font(bold=True, size=12, color=BR)
    ws2['A1'].fill = PatternFill("solid", fgColor=VE)
    ws2.merge_cells('A1:F1')
    ws2['A1'].alignment = Alignment(horizontal='center')
    ws2.row_dimensions[1].height = 24
    ws2.append(["Curso", "Turma", "Total", "🔴 Críticos", "🟠 Atenção", "🟡 Monitorar"])
    hr2 = ws2.max_row
    for ci in range(1, 7):
        c = ws2.cell(row=hr2, column=ci)
        c.font = Font(bold=True, color=BR, size=9)
        c.fill = PatternFill("solid", fgColor=VM)
        c.alignment = Alignment(horizontal='center')
    ws2.row_dimensions[hr2].height = 18
    for (curso, turma), g in df.groupby(['Curso', 'Turma']):
        ws2.append([curso, turma, len(g),
                    len(g[g['Qtd. Alertas'] >= 3]),
                    len(g[g['Qtd. Alertas'] == 2]),
                    len(g[g['Qtd. Alertas'] == 1])])
        r = ws2.max_row
        for ci in range(1, 7):
            c = ws2.cell(row=r, column=ci)
            c.font = Font(size=9)
            c.alignment = Alignment(horizontal='center' if ci > 2 else 'left')
            c.border = border
        ws2.row_dimensions[r].height = 16
    ws2.append(["TOTAL GERAL", "", len(df),
                len(df[df['Qtd. Alertas'] >= 3]),
                len(df[df['Qtd. Alertas'] == 2]),
                len(df[df['Qtd. Alertas'] == 1])])
    r = ws2.max_row
    for ci in range(1, 7):
        c = ws2.cell(row=r, column=ci)
        c.font = Font(bold=True, size=9, color=BR)
        c.fill = PatternFill("solid", fgColor=VE)
        c.alignment = Alignment(horizontal='center' if ci > 2 else 'left')
    ws2.row_dimensions[r].height = 18

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def enviar_email(excel_bytes, destinatarios, data_hoje, total, criticos, atencao, monitorar):
    api_key = st.secrets.get("SENDGRID_API_KEY", "")
    if not api_key:
        return False, "Chave SendGrid não configurada."
    lista = list(set(destinatarios + DESTINATARIOS_FIXOS))
    assunto = f"Relatório CS Rehagro — Engajamento {data_hoje}"
    html = f"""
    <div style="font-family:Inter,Arial,sans-serif;max-width:600px;margin:0 auto;background:#F9F6EF;">
      <div style="background:#1B4D2E;padding:32px 40px;border-radius:12px 12px 0 0;">
        <p style="color:#C8A951;font-size:10px;font-weight:700;letter-spacing:3px;text-transform:uppercase;margin:0 0 8px">Customer Success · Rehagro</p>
        <h1 style="color:#fff;font-size:22px;margin:0;font-family:Georgia,serif;">Relatório de Engajamento</h1>
        <p style="color:rgba(255,255,255,0.6);margin:6px 0 0;font-size:13px;">Gerado em {data_hoje}</p>
      </div>
      <div style="background:#fff;padding:32px 40px;border:1px solid #e8e2d4;">
        <p style="color:#444;font-size:14px;margin:0 0 24px;">Segue em anexo o relatório semanal de engajamento dos alunos.</p>
        <table width="100%" cellpadding="0" cellspacing="8">
          <tr>
            <td style="background:#EEF7F0;border-radius:10px;padding:16px;text-align:center;width:25%">
              <div style="font-size:30px;font-weight:800;color:#1B4D2E;font-family:Georgia,serif;">{total}</div>
              <div style="font-size:10px;color:#888;text-transform:uppercase;letter-spacing:1px;">Total</div>
            </td>
            <td style="background:#FFEBEE;border-radius:10px;padding:16px;text-align:center;width:25%">
              <div style="font-size:30px;font-weight:800;color:#C62828;font-family:Georgia,serif;">{criticos}</div>
              <div style="font-size:10px;color:#888;text-transform:uppercase;letter-spacing:1px;">🔴 Críticos</div>
            </td>
            <td style="background:#FFF3E0;border-radius:10px;padding:16px;text-align:center;width:25%">
              <div style="font-size:30px;font-weight:800;color:#E65100;font-family:Georgia,serif;">{atencao}</div>
              <div style="font-size:10px;color:#888;text-transform:uppercase;letter-spacing:1px;">🟠 Atenção</div>
            </td>
            <td style="background:#FFFDE7;border-radius:10px;padding:16px;text-align:center;width:25%">
              <div style="font-size:30px;font-weight:800;color:#F57F17;font-family:Georgia,serif;">{monitorar}</div>
              <div style="font-size:10px;color:#888;text-transform:uppercase;letter-spacing:1px;">🟡 Monitorar</div>
            </td>
          </tr>
        </table>
        <p style="color:#999;font-size:11px;margin-top:24px;">O arquivo Excel em anexo contém o relatório completo com alertas, ações recomendadas e comentários dos alunos.</p>
      </div>
      <div style="background:#F9F6EF;padding:16px;border-radius:0 0 12px 12px;text-align:center;border:1px solid #e8e2d4;border-top:none;">
        <p style="color:#aaa;font-size:10px;margin:0;">Rehagro · Customer Success · Agente de Engajamento</p>
      </div>
    </div>"""

    encoded  = base64.b64encode(excel_bytes).decode()
    nome_arq = f"relatorio_cs_{data_hoje.replace('/', '')}.xlsx"
    message  = Mail(from_email=REMETENTE, to_emails=lista[0],
                    subject=assunto, html_content=html)
    if len(lista) > 1:
        for dest in lista[1:]:
            message.add_cc(dest)
    message.attachment = Attachment(
        FileContent(encoded), FileName(nome_arq),
        FileType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        Disposition("attachment"))
    try:
        sg  = SendGridAPIClient(api_key)
        res = sg.send(message)
        if res.status_code in [200, 201, 202]:
            return True, f"E-mail enviado para {len(lista)} destinatário(s)."
        return False, f"Erro SendGrid: status {res.status_code}"
    except Exception as e:
        return False, str(e)


# ── UI ───────────────────────────────────────────────────

# Hero
st.markdown("""
<div class="rh-hero">
  <div class="rh-logo">
    <span class="rh-logo-leaf"></span> Rehagro
  </div>
  <div class="rh-badge">Customer Success · Agente de Engajamento</div>
  <p class="rh-title">Monitoramento<br>de Alunos</p>
  <p class="rh-sub">Identifique alunos desengajados e saiba exatamente qual ação tomar — em segundos.</p>
</div>
""", unsafe_allow_html=True)

col_esq, col_dir = st.columns([1, 1], gap="large")

with col_esq:
    st.markdown('<p class="rh-section">📋 Como exportar do Power BI</p>', unsafe_allow_html=True)
    st.markdown("""
    <div class="rh-card">
      <div class="rh-step-num">Dashboard 1</div>
      <div class="rh-step-title">Acesso ao Canvas (AVA)</div>
      <div class="rh-step-desc">
        Relatório <b>Rehagro - Canvas</b> → página <b>Acesso ao Canvas-Ok</b><br><br>
        <span class="rh-tag">Colaborador = Não</span>
        <span class="rh-tag">Status Usuário Curso = Ativo</span>
        <span class="rh-tag">Curso = seus cursos</span>
        <span class="rh-tag">Turma = suas turmas</span><br><br>
        Exporte a <b>Tabela de dados</b> → Excel
        <p class="rh-note">💡 Múltiplos cursos e turmas são suportados.</p>
      </div>
    </div>
    <div class="rh-card">
      <div class="rh-step-num">Dashboard 2</div>
      <div class="rh-step-title">NPS Médio por Aluno</div>
      <div class="rh-step-desc">
        Relatório <b>Rehagro Educação - Avaliação de Aula</b> → página <b>Avaliações de aula/aluno</b><br><br>
        <span class="rh-tag">Formato Curso = Online</span>
        <span class="rh-tag">Tipo de aula = On-line ao vivo</span>
        <span class="rh-tag">Ano_aula = ano atual</span>
        <span class="rh-tag">Questão = "...você indicaria o Rehagro..."</span>
        <span class="rh-tag">Área = sua área</span>
        <span class="rh-tag">Curso = seus cursos</span><br><br>
        Exporte a <b>Tabela NPS médio/aluno</b> → Excel
      </div>
    </div>
    <div class="rh-card">
      <div class="rh-step-num">Dashboard 3</div>
      <div class="rh-step-title">Frequência nas Aulas ao Vivo</div>
      <div class="rh-step-desc">
        Relatório <b>Rehagro Alunado</b> → página <b>Análise de Frequência e Faltas</b><br><br>
        <span class="rh-tag">Área = sua área</span>
        <span class="rh-tag">Formato do curso = Online</span>
        <span class="rh-tag">Turma = suas turmas</span><br><br>
        Exporte a <b>Tabela de frequência</b> → Excel
      </div>
    </div>
    <div class="rh-card">
      <div class="rh-step-num">Dashboard 4 — Opcional</div>
      <div class="rh-step-title">Comentários das Aulas</div>
      <div class="rh-step-desc">
        Relatório <b>Rehagro Educação - Avaliação de Aula</b> → página <b>Avaliações de aula/aluno</b><br><br>
        <span class="rh-tag">Mesmos filtros do Dashboard 2</span><br><br>
        Exporte a <b>Tabela de Comentários</b> → Excel
        <p class="rh-note">💡 Enriquece o relatório com o contexto do feedback do aluno.</p>
      </div>
    </div>
    """, unsafe_allow_html=True)

with col_dir:
    st.markdown('<p class="rh-section">📁 Envie os arquivos</p>', unsafe_allow_html=True)

    f_canvas  = st.file_uploader("1️⃣  Canvas — acesso ao AVA",           type=["xlsx"], key="canvas")
    f_nps     = st.file_uploader("2️⃣  NPS — avaliações de aula",          type=["xlsx"], key="nps")
    f_freq    = st.file_uploader("3️⃣  Frequência — aulas ao vivo",        type=["xlsx"], key="freq")
    f_coment  = st.file_uploader("4️⃣  Comentários — opcional",            type=["xlsx"], key="coment")

    st.markdown('<p class="rh-section">✉️ Envio por e-mail</p>', unsafe_allow_html=True)
    email_usuario = st.text_input("Seu e-mail (receberá cópia junto à lista fixa):",
                                  placeholder="seuemail@rehagro.com.br")
    with st.expander("👁️ Ver destinatários fixos"):
        for d in DESTINATARIOS_FIXOS:
            st.write(f"• {d}")

    st.markdown('<div class="rh-divider"></div>', unsafe_allow_html=True)

    obrigatorios = f_canvas and f_nps and f_freq
    if not obrigatorios:
        faltando = [n for f, n in [(f_canvas,"Canvas"),(f_nps,"NPS"),(f_freq,"Frequência")] if not f]
        st.info(f"Aguardando arquivos obrigatórios: **{', '.join(faltando)}**")
    else:
        st.success("✅ Arquivos recebidos! Clique para gerar e enviar o relatório.")
        if st.button("🔍 Gerar e Enviar Relatório", type="primary", use_container_width=True):
            with st.spinner("Analisando dados..."):
                try:
                    dc = carregar_canvas(f_canvas)
                    dn = carregar_nps(f_nps)
                    df_fr = carregar_frequencia(f_freq)
                    df_co = carregar_comentarios(f_coment) if f_coment else None
                    df_rel = gerar_relatorio(dc, dn, df_fr, df_co)

                    criticos  = len(df_rel[df_rel['Qtd. Alertas'] >= 3])
                    atencao   = len(df_rel[df_rel['Qtd. Alertas'] == 2])
                    monitorar = len(df_rel[df_rel['Qtd. Alertas'] == 1])

                    st.markdown(f"""
                    <div class="rh-metrics">
                      <div class="rh-metric">
                        <div class="rh-metric-num" style="color:#1B4D2E">{len(df_rel)}</div>
                        <div class="rh-metric-label">Total</div>
                      </div>
                      <div class="rh-metric">
                        <div class="rh-metric-num" style="color:#C62828">{criticos}</div>
                        <div class="rh-metric-label">🔴 Críticos</div>
                      </div>
                      <div class="rh-metric">
                        <div class="rh-metric-num" style="color:#E65100">{atencao}</div>
                        <div class="rh-metric-label">🟠 Atenção</div>
                      </div>
                      <div class="rh-metric">
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

                    st.markdown('<p class="rh-section">👥 Alunos desengajados</p>', unsafe_allow_html=True)
                    colunas_view = ['Curso', 'Turma', 'Nome', 'Qtd. Alertas',
                                    'Alertas Identificados', 'Ações Recomendadas']
                    if df_co is not None:
                        colunas_view.append('Comentários do Aluno')
                    st.dataframe(df_view[colunas_view],
                                 use_container_width=True, hide_index=True, height=320)

                    excel_bytes = exportar_excel_bytes(df_rel)
                    data_hoje   = datetime.now().strftime('%d/%m/%Y')

                    dests = [email_usuario.strip()] if email_usuario and "@" in email_usuario else []
                    if SENDGRID_OK:
                        ok, msg = enviar_email(excel_bytes, dests, data_hoje,
                                               len(df_rel), criticos, atencao, monitorar)
                        if ok:
                            st.success(f"📧 {msg}")
                        else:
                            st.warning(f"⚠️ E-mail não enviado: {msg}")
                    else:
                        st.warning("⚠️ Biblioteca SendGrid não instalada.")

                    st.download_button(
                        "⬇️  Baixar Relatório Excel",
                        data=excel_bytes,
                        file_name=f"relatorio_cs_{datetime.now().strftime('%d%m%Y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                except Exception as e:
                    st.error(f"Erro ao processar: {e}")
                    st.info("Verifique se os arquivos corretos foram enviados com os filtros indicados.")

st.markdown("""
<div class="rh-footer">
    Rehagro &copy; 2026 · Customer Success · Agente de Engajamento
</div>
""", unsafe_allow_html=True)
