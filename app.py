import streamlit as st
import pandas as pd
import re
import io
import base64
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# SendGrid
try:
    from sendgrid import SendGridAPIClient
    from sendgrid.helpers.mail import (
        Mail, Attachment, FileContent, FileName,
        FileType, Disposition
    )
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
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.stApp { background: #f7f5f0; }
.hero {
    background: linear-gradient(135deg, #1B4D2E 0%, #2E7D32 60%, #388E3C 100%);
    border-radius: 16px; padding: 40px 48px; margin-bottom: 32px;
    position: relative; overflow: hidden;
}
.hero::before {
    content: ''; position: absolute; top: -60px; right: -60px;
    width: 220px; height: 220px; border-radius: 50%;
    background: rgba(249,168,37,0.15);
}
.hero-title {
    font-family: 'Syne', sans-serif; font-size: 2rem; font-weight: 800;
    color: #ffffff; margin: 0 0 6px 0; letter-spacing: -0.5px;
}
.hero-sub { font-size: 1rem; color: rgba(255,255,255,0.75); margin: 0; font-weight: 300; }
.hero-badge {
    display: inline-block; background: #F9A825; color: #1B4D2E;
    font-family: 'Syne', sans-serif; font-weight: 700; font-size: 0.7rem;
    letter-spacing: 1.5px; text-transform: uppercase; padding: 4px 12px;
    border-radius: 20px; margin-bottom: 12px;
}
.step-card {
    background: #ffffff; border-radius: 12px; padding: 20px 24px;
    border-left: 4px solid #2E7D32; margin-bottom: 12px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}
.step-num { font-family: 'Syne', sans-serif; font-size: 0.7rem; font-weight: 700; letter-spacing: 2px; text-transform: uppercase; color: #2E7D32; margin-bottom: 4px; }
.step-title { font-family: 'Syne', sans-serif; font-size: 1rem; font-weight: 700; color: #1a1a1a; margin-bottom: 6px; }
.step-desc { font-size: 0.85rem; color: #666; line-height: 1.6; }
.filter-tag { display: inline-block; background: #E8F5E9; color: #2E7D32; font-size: 0.75rem; font-weight: 500; padding: 2px 8px; border-radius: 4px; margin: 2px 2px 2px 0; }
.filter-note { font-size: 0.78rem; color: #888; font-style: italic; margin-top: 6px; }
.metric-row { display: flex; gap: 12px; margin-bottom: 24px; }
.metric-card { flex: 1; background: #ffffff; border-radius: 12px; padding: 20px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
.metric-num { font-family: 'Syne', sans-serif; font-size: 2.2rem; font-weight: 800; line-height: 1; margin-bottom: 4px; }
.metric-label { font-size: 0.78rem; color: #888; text-transform: uppercase; letter-spacing: 0.8px; }
.section-title { font-family: 'Syne', sans-serif; font-size: 1.1rem; font-weight: 700; color: #1B4D2E; margin: 24px 0 12px 0; padding-bottom: 6px; border-bottom: 2px solid #E8F5E9; }
.divider { height: 1px; background: linear-gradient(90deg, #2E7D32, transparent); margin: 24px 0; opacity: 0.3; }
[data-testid="stFileUploader"] { background: #ffffff; border-radius: 12px; padding: 8px; border: 1px solid #e0e0e0; }
</style>
""", unsafe_allow_html=True)


# ─── Configurações fixas de e-mail ───────────────────────────
# Lista fixa de destinatários que SEMPRE recebem o relatório.
# Edite esta lista diretamente aqui para adicionar ou remover pessoas.
DESTINATARIOS_FIXOS = [
    "rafael.ferraz@rehagro.edu.br",
]

REMETENTE = "rafael.ferraz@rehagro.edu.br"   # E-mail verificado no SendGrid


# ─── Funções de processamento ────────────────────────────────

def carregar_canvas(arquivo):
    df = pd.read_excel(arquivo, skiprows=2)
    df.columns = df.columns.str.strip()
    col_nome   = next((c for c in df.columns if 'NOME' in c.upper()), None)
    col_email  = next((c for c in df.columns if 'E-MAIL' in c.upper() or 'EMAIL' in c.upper()), None)
    col_dias   = next((c for c in df.columns if 'DIAS' in c.upper()), None)
    col_curso  = next((c for c in df.columns if c.upper() == 'CURSO'), None)
    col_turma  = next((c for c in df.columns if c.upper() == 'TURMA'), None)
    if not col_nome or not col_dias:
        raise ValueError("Arquivo Canvas: colunas esperadas não encontradas. Verifique os filtros.")
    df['_key']  = (df[col_nome].astype(str).str.strip().str.lower() + '||' +
                   (df[col_turma].astype(str).str.strip().str.lower() if col_turma else ''))
    df['_dias'] = pd.to_numeric(df[col_dias], errors='coerce')
    cols = {'_key': '_key', col_nome: 'Nome', '_dias': 'Dias sem acesso'}
    if col_email: cols[col_email] = 'Email'
    if col_curso: cols[col_curso] = 'Curso'
    if col_turma: cols[col_turma] = 'Turma'
    res = df[df['_dias'] > 30][list(cols.keys())].copy().rename(columns=cols)
    if 'Email'  not in res.columns: res['Email']  = ''
    if 'Curso'  not in res.columns: res['Curso']  = ''
    if 'Turma'  not in res.columns: res['Turma']  = ''
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
        raise ValueError("Arquivo NPS: colunas esperadas não encontradas. Verifique os filtros.")
    df_v = df[df[col_nps].notna()].copy()
    if col_data:
        df_v['_data'] = pd.to_datetime(df_v[col_data], errors='coerce')
        df_v = df_v.sort_values([col_aluno, '_data'])
    df_v['_key'] = (df_v[col_aluno].astype(str).str.strip().str.lower() + '||' +
                    (df_v[col_turma].astype(str).str.strip().str.lower() if col_turma else ''))
    ultima = df_v.groupby('_key').last().reset_index()
    return ultima[pd.to_numeric(ultima[col_nps], errors='coerce') < 0][['_key']].copy()


def carregar_frequencia(arquivo):
    df = pd.read_excel(arquivo, skiprows=2)
    df.columns = df.columns.str.strip()
    col_aluno  = next((c for c in df.columns if c.upper() == 'ALUNO'), None)
    col_status = next((c for c in df.columns if 'STATUS' in c.upper() or 'PRESEN' in c.upper()), None)
    col_data   = next((c for c in df.columns if 'DATA' in c.upper() or 'PARTE' in c.upper()), None)
    col_turma  = next((c for c in df.columns if c.upper() == 'TURMA'), None)
    if not col_aluno or not col_status:
        raise ValueError("Arquivo Frequência: colunas esperadas não encontradas. Verifique os filtros.")
    df_a = df[~df[col_status].isin(['-', 'DESISTENTE'])].copy()
    def extr(s):
        m = re.match(r'(\d{2}/\d{2}/\d{4})', str(s))
        return pd.to_datetime(m.group(1), format='%d/%m/%Y') if m else pd.NaT
    if col_data:
        df_a['_data'] = df_a[col_data].apply(extr)
        sort_cols = [col_aluno, col_turma, '_data'] if col_turma else [col_aluno, '_data']
        df_a = df_a.sort_values(sort_cols)
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


def gerar_relatorio(df_canvas, df_nps, df_freq):
    todos_keys = set(df_canvas['_key']) | set(df_nps['_key']) | set(df_freq['_key'])
    info_map = df_canvas.set_index('_key')[['Nome', 'Email', 'Curso', 'Turma']].to_dict('index')
    for key in todos_keys:
        if key not in info_map:
            partes = key.split('||')
            info_map[key] = {'Nome': partes[0].title(), 'Email': '',
                             'Curso': '', 'Turma': partes[1].upper() if len(partes) > 1 else ''}
    relatorio = []
    for key in sorted(todos_keys):
        info = info_map[key]
        alertas, acoes = [], []
        c = df_canvas[df_canvas['_key'] == key]
        if not c.empty:
            dias = int(c.iloc[0]['Dias sem acesso'])
            alertas.append(f"Sem acesso ao Canvas há {dias} dias")
            acoes.append("Enviar link de acesso à plataforma")
        if not df_nps[df_nps['_key'] == key].empty:
            alertas.append("Última avaliação: detrator (NPS negativo)")
            acoes.append("Retomar feedback negativo da avaliação")
        f = df_freq[df_freq['_key'] == key]
        if not f.empty:
            alertas.append(f"Ausente nas últimas 2 aulas ao vivo ({f.iloc[0]['Ultimas_2_aulas']})")
            acoes.append("Enviar data da próxima aula ao vivo")
        if alertas:
            relatorio.append({
                'Curso': info['Curso'], 'Turma': info['Turma'],
                'Nome': info['Nome'], 'E-mail': info['Email'],
                'Qtd. Alertas': len(alertas),
                'Alertas Identificados': ' | '.join(alertas),
                'Ações Recomendadas': ' | '.join(acoes),
            })
    df = pd.DataFrame(relatorio)
    return df.sort_values(['Curso', 'Turma', 'Qtd. Alertas', 'Nome'],
                          ascending=[True, True, False, True]).reset_index(drop=True)


def exportar_excel_bytes(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Relatório CS"
    verde_escuro = "1B4D2E"; verde_medio = "2E7D32"; branco = "FFFFFF"; cinza = "F5F5F5"
    n_cols = 7
    last_col = get_column_letter(n_cols)

    ws.merge_cells(f'A1:{last_col}1')
    ws['A1'] = "RELATÓRIO DE ENGAJAMENTO — CUSTOMER SUCCESS REHAGRO"
    ws['A1'].font = Font(bold=True, size=14, color=branco)
    ws['A1'].fill = PatternFill("solid", fgColor=verde_escuro)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    ws.merge_cells(f'A2:{last_col}2')
    ws['A2'] = f"Gerado em: {datetime.now().strftime('%d/%m/%Y às %H:%M')}   |   Total de alunos desengajados: {len(df)}"
    ws['A2'].font = Font(size=10, color=branco)
    ws['A2'].fill = PatternFill("solid", fgColor=verde_medio)
    ws['A2'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 20

    ws.append([])
    ws.merge_cells(f'A4:{last_col}4')
    ws['A4'] = "LEGENDA DE AÇÕES"
    ws['A4'].font = Font(bold=True, size=10, color=verde_escuro)
    ws['A4'].fill = PatternFill("solid", fgColor="E8F5E9")
    ws['A4'].alignment = Alignment(horizontal='left')

    for i, txt in enumerate([
        "⚠️  Sem acesso ao Canvas há +30 dias  →  Enviar link de acesso à plataforma",
        "⚠️  Última avaliação com NPS negativo  →  Retomar feedback negativo da avaliação",
        "⚠️  Ausente nas últimas 2 aulas ao vivo  →  Enviar data da próxima aula ao vivo",
    ], 5):
        ws.merge_cells(f'A{i}:{last_col}{i}')
        ws[f'A{i}'] = txt
        ws[f'A{i}'].font = Font(size=9, color="424242")
        ws[f'A{i}'].fill = PatternFill("solid", fgColor=cinza)
        ws[f'A{i}'].alignment = Alignment(horizontal='left', indent=1)

    ws.append([])
    headers = ['Curso', 'Turma', 'Nome do Aluno', 'E-mail', 'Qtd. Alertas', 'Alertas Identificados', 'Ações Recomendadas']
    ws.append(headers)
    hr = ws.max_row
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=hr, column=ci)
        c.font = Font(bold=True, color=branco, size=10)
        c.fill = PatternFill("solid", fgColor=verde_escuro)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[hr].height = 22

    thin = Side(style='thin', color='BDBDBD')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for _, row in df.iterrows():
        ws.append([row['Curso'], row['Turma'], row['Nome'], row['E-mail'],
                   row['Qtd. Alertas'], row['Alertas Identificados'], row['Ações Recomendadas']])
        dr = ws.max_row
        fc = "FFCDD2" if row['Qtd. Alertas'] >= 3 else ("FFE0B2" if row['Qtd. Alertas'] == 2 else "FFF9C4")
        for ci in range(1, n_cols + 1):
            c = ws.cell(row=dr, column=ci)
            if ci <= 2:
                c.fill = PatternFill("solid", fgColor="E8F5E9")
                c.font = Font(size=9, color="2E7D32", bold=True)
            elif ci == 5:
                c.fill = PatternFill("solid", fgColor=fc)
                c.font = Font(bold=True, size=11)
                c.alignment = Alignment(horizontal='center', vertical='center')
            else:
                c.fill = PatternFill("solid", fgColor=fc if ci > 4 else "FFFFFF")
            c.border = border
            if ci not in [2, 5]:
                c.alignment = Alignment(vertical='center', wrap_text=True)
        ws.row_dimensions[dr].height = 30

    for ci, w in enumerate([18, 20, 32, 30, 12, 65, 50], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    # Aba resumo por turma
    ws2 = wb.create_sheet("Resumo por Turma")
    for ci, w in enumerate([22, 22, 18, 14, 14, 14], 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w
    ws2.append(["RESUMO POR TURMA", "", "", "", "", ""])
    ws2['A1'].font = Font(bold=True, size=12, color=branco)
    ws2['A1'].fill = PatternFill("solid", fgColor=verde_escuro)
    ws2.merge_cells('A1:F1')
    ws2['A1'].alignment = Alignment(horizontal='center')
    ws2.row_dimensions[1].height = 25
    ws2.append(["Curso", "Turma", "Total Desengajados", "🔴 Críticos", "🟠 Atenção", "🟡 Monitorar"])
    hr2 = ws2.max_row
    for ci in range(1, 7):
        c = ws2.cell(row=hr2, column=ci)
        c.font = Font(bold=True, color=branco, size=9)
        c.fill = PatternFill("solid", fgColor="2E7D32")
        c.alignment = Alignment(horizontal='center', wrap_text=True)
    ws2.row_dimensions[hr2].height = 20
    for (curso, turma), g in df.groupby(['Curso', 'Turma']):
        ws2.append([curso, turma, len(g),
                    len(g[g['Qtd. Alertas'] >= 3]),
                    len(g[g['Qtd. Alertas'] == 2]),
                    len(g[g['Qtd. Alertas'] == 1])])
        r = ws2.max_row
        for ci in range(1, 7):
            c = ws2.cell(row=r, column=ci)
            c.font = Font(size=10)
            c.alignment = Alignment(horizontal='center' if ci > 2 else 'left')
            c.border = border
        ws2.row_dimensions[r].height = 18
    ws2.append(["TOTAL GERAL", "", len(df),
                len(df[df['Qtd. Alertas'] >= 3]),
                len(df[df['Qtd. Alertas'] == 2]),
                len(df[df['Qtd. Alertas'] == 1])])
    r = ws2.max_row
    for ci in range(1, 7):
        c = ws2.cell(row=r, column=ci)
        c.font = Font(bold=True, size=10, color=branco)
        c.fill = PatternFill("solid", fgColor=verde_escuro)
        c.alignment = Alignment(horizontal='center' if ci > 2 else 'left')
    ws2.row_dimensions[r].height = 20

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def enviar_email(excel_bytes, destinatarios, data_hoje, total, criticos, atencao, monitorar):
    """Envia o relatório Excel por e-mail via SendGrid."""
    api_key = st.secrets.get("SENDGRID_API_KEY", "")
    if not api_key:
        return False, "Chave SendGrid não configurada."

    lista_dest = list(set(destinatarios + DESTINATARIOS_FIXOS))

    assunto = f"Relatório CS Rehagro — Engajamento {data_hoje}"

    corpo_html = f"""
    <div style="font-family: 'DM Sans', Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #1B4D2E, #2E7D32); padding: 32px; border-radius: 12px 12px 0 0;">
            <p style="color:#F9A825; font-size:11px; font-weight:700; letter-spacing:2px; text-transform:uppercase; margin:0 0 8px 0;">Customer Success · Rehagro</p>
            <h1 style="color:#ffffff; font-size:22px; margin:0;">🌱 Relatório de Engajamento</h1>
            <p style="color:rgba(255,255,255,0.7); margin:6px 0 0 0; font-size:14px;">Gerado em {data_hoje}</p>
        </div>
        <div style="background:#ffffff; padding: 32px; border: 1px solid #e0e0e0;">
            <p style="color:#444; font-size:15px;">Segue em anexo o relatório semanal de engajamento dos alunos.</p>
            <div style="display:flex; gap:12px; margin: 24px 0;">
                <div style="flex:1; background:#f7f5f0; border-radius:8px; padding:16px; text-align:center;">
                    <div style="font-size:28px; font-weight:800; color:#1B4D2E;">{total}</div>
                    <div style="font-size:11px; color:#888; text-transform:uppercase; letter-spacing:1px;">Total</div>
                </div>
                <div style="flex:1; background:#FFEBEE; border-radius:8px; padding:16px; text-align:center;">
                    <div style="font-size:28px; font-weight:800; color:#C62828;">{criticos}</div>
                    <div style="font-size:11px; color:#888; text-transform:uppercase; letter-spacing:1px;">🔴 Críticos</div>
                </div>
                <div style="flex:1; background:#FFF3E0; border-radius:8px; padding:16px; text-align:center;">
                    <div style="font-size:28px; font-weight:800; color:#E65100;">{atencao}</div>
                    <div style="font-size:11px; color:#888; text-transform:uppercase; letter-spacing:1px;">🟠 Atenção</div>
                </div>
                <div style="flex:1; background:#FFFDE7; border-radius:8px; padding:16px; text-align:center;">
                    <div style="font-size:28px; font-weight:800; color:#F57F17;">{monitorar}</div>
                    <div style="font-size:11px; color:#888; text-transform:uppercase; letter-spacing:1px;">🟡 Monitorar</div>
                </div>
            </div>
            <p style="color:#888; font-size:12px; margin-top:24px;">
                O arquivo Excel em anexo contém o relatório completo com os alertas e ações recomendadas para cada aluno.
            </p>
        </div>
        <div style="background:#f7f5f0; padding:16px; border-radius: 0 0 12px 12px; text-align:center;">
            <p style="color:#aaa; font-size:11px; margin:0;">Rehagro · Customer Success · Agente de Engajamento</p>
        </div>
    </div>
    """

    encoded = base64.b64encode(excel_bytes).decode()
    nome_arquivo = f"relatorio_cs_{data_hoje.replace('/', '')}.xlsx"

    message = Mail(
        from_email=REMETENTE,
        to_emails=lista_dest[0],
        subject=assunto,
        html_content=corpo_html,
    )

    # Adicionar demais destinatários em CC
    if len(lista_dest) > 1:
        from sendgrid.helpers.mail import Cc
        for dest in lista_dest[1:]:
            message.add_cc(dest)

    attachment = Attachment(
        FileContent(encoded),
        FileName(nome_arquivo),
        FileType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        Disposition("attachment")
    )
    message.attachment = attachment

    try:
        sg = SendGridAPIClient(api_key)
        response = sg.send(message)
        if response.status_code in [200, 201, 202]:
            return True, f"E-mail enviado para {len(lista_dest)} destinatário(s)."
        else:
            return False, f"Erro SendGrid: status {response.status_code}"
    except Exception as e:
        return False, str(e)


# ─── UI ─────────────────────────────────────────────────────

st.markdown("""
<div class="hero">
    <div class="hero-badge">Customer Success · Rehagro</div>
    <p class="hero-title">🌱 Agente de Engajamento</p>
    <p class="hero-sub">Identifique alunos desengajados e saiba exatamente qual ação tomar — em segundos.</p>
</div>
""", unsafe_allow_html=True)

col_esq, col_dir = st.columns([1, 1], gap="large")

with col_esq:
    st.markdown('<p class="section-title">📋 Como exportar do Power BI</p>', unsafe_allow_html=True)
    st.markdown("""
    <div class="step-card">
        <div class="step-num">Dashboard 1</div>
        <div class="step-title">Acesso ao Canvas (AVA)</div>
        <div class="step-desc">
            Relatório <b>Rehagro - Canvas</b> → página <b>Acesso ao Canvas-Ok</b><br><br>
            Filtros obrigatórios:<br>
            <span class="filter-tag">Colaborador = Não</span>
            <span class="filter-tag">Status Usuário Curso = Ativo</span><br>
            <span class="filter-tag">Curso = selecione os cursos da sua área</span>
            <span class="filter-tag">Turma = selecione as turmas desejadas</span><br><br>
            Exporte a <b>Tabela de dados</b> → Excel
            <p class="filter-note">💡 Você pode selecionar múltiplos cursos e turmas.</p>
        </div>
    </div>
    <div class="step-card">
        <div class="step-num">Dashboard 2</div>
        <div class="step-title">NPS Médio por Aluno</div>
        <div class="step-desc">
            Relatório <b>Rehagro Educação - Avaliação de Aula</b> → página <b>Avaliações de aula/aluno</b><br><br>
            Filtros obrigatórios:<br>
            <span class="filter-tag">Formato Curso = Online</span>
            <span class="filter-tag">Tipo de aula = On-line ao vivo</span>
            <span class="filter-tag">Ano_aula = ano atual</span>
            <span class="filter-tag">Questão = "...quanto...você indicaria o Rehagro..."</span><br>
            <span class="filter-tag">Área = sua área</span>
            <span class="filter-tag">Curso = seus cursos</span><br><br>
            Exporte a <b>Tabela NPS médio/aluno</b> → Excel
        </div>
    </div>
    <div class="step-card">
        <div class="step-num">Dashboard 3</div>
        <div class="step-title">Frequência nas Aulas ao Vivo</div>
        <div class="step-desc">
            Relatório <b>Rehagro Alunado</b> → página <b>Análise de Frequência e Faltas</b><br><br>
            Filtros obrigatórios:<br>
            <span class="filter-tag">Área = sua área</span>
            <span class="filter-tag">Formato do curso = Online</span>
            <span class="filter-tag">Turma = selecione as turmas da sua área</span><br><br>
            Exporte a <b>Tabela de frequência</b> → Excel
            <p class="filter-note">💡 Você pode selecionar todas as turmas da área de uma vez.</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

with col_dir:
    st.markdown('<p class="section-title">📁 Envie os arquivos exportados</p>', unsafe_allow_html=True)

    f_canvas = st.file_uploader("1️⃣  Arquivo Canvas (acesso ao AVA)",    type=["xlsx"], key="canvas")
    f_nps    = st.file_uploader("2️⃣  Arquivo NPS (avaliações de aula)",   type=["xlsx"], key="nps")
    f_freq   = st.file_uploader("3️⃣  Arquivo Frequência (aulas ao vivo)", type=["xlsx"], key="freq")

    st.markdown('<p class="section-title">✉️ Envio por e-mail</p>', unsafe_allow_html=True)

    email_usuario = st.text_input(
        "Seu e-mail (receberá uma cópia junto com a lista fixa):",
        placeholder="seuemail@rehagro.com.br"
    )

    with st.expander("👁️ Ver destinatários fixos"):
        st.write("O relatório sempre será enviado para:")
        for d in DESTINATARIOS_FIXOS:
            st.write(f"• {d}")

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    todos_prontos = f_canvas and f_nps and f_freq

    if not todos_prontos:
        faltando = []
        if not f_canvas: faltando.append("Canvas")
        if not f_nps:    faltando.append("NPS")
        if not f_freq:   faltando.append("Frequência")
        st.info(f"Aguardando arquivos: **{', '.join(faltando)}**")
    else:
        st.success("✅ Todos os arquivos recebidos! Clique para gerar e enviar o relatório.")

        if st.button("🔍 Gerar e Enviar Relatório", type="primary", use_container_width=True):
            with st.spinner("Analisando dados e enviando e-mail..."):
                try:
                    df_canvas_data = carregar_canvas(f_canvas)
                    df_nps_data    = carregar_nps(f_nps)
                    df_freq_data   = carregar_frequencia(f_freq)
                    df_rel         = gerar_relatorio(df_canvas_data, df_nps_data, df_freq_data)

                    criticos  = len(df_rel[df_rel['Qtd. Alertas'] >= 3])
                    atencao   = len(df_rel[df_rel['Qtd. Alertas'] == 2])
                    monitorar = len(df_rel[df_rel['Qtd. Alertas'] == 1])

                    st.markdown(f"""
                    <div class="metric-row">
                        <div class="metric-card">
                            <div class="metric-num" style="color:#1B4D2E">{len(df_rel)}</div>
                            <div class="metric-label">Total desengajados</div>
                        </div>
                        <div class="metric-card">
                            <div class="metric-num" style="color:#C62828">{criticos}</div>
                            <div class="metric-label">🔴 Críticos</div>
                        </div>
                        <div class="metric-card">
                            <div class="metric-num" style="color:#E65100">{atencao}</div>
                            <div class="metric-label">🟠 Atenção</div>
                        </div>
                        <div class="metric-card">
                            <div class="metric-num" style="color:#F57F17">{monitorar}</div>
                            <div class="metric-label">🟡 Monitorar</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                    turmas_disp = sorted(df_rel['Turma'].dropna().unique().tolist())
                    if len(turmas_disp) > 1:
                        turma_sel = st.multiselect("Filtrar por turma:", options=turmas_disp, default=turmas_disp)
                        df_view = df_rel[df_rel['Turma'].isin(turma_sel)]
                    else:
                        df_view = df_rel

                    st.markdown('<p class="section-title">👥 Alunos desengajados</p>', unsafe_allow_html=True)
                    st.dataframe(
                        df_view[['Curso', 'Turma', 'Nome', 'Qtd. Alertas', 'Alertas Identificados', 'Ações Recomendadas']],
                        use_container_width=True, hide_index=True, height=300
                    )

                    excel_bytes = exportar_excel_bytes(df_rel)
                    data_hoje   = datetime.now().strftime('%d/%m/%Y')

                    # Envio por e-mail
                    destinatarios_envio = []
                    if email_usuario and "@" in email_usuario:
                        destinatarios_envio.append(email_usuario.strip())

                    if SENDGRID_OK:
                        ok, msg = enviar_email(
                            excel_bytes, destinatarios_envio,
                            data_hoje, len(df_rel), criticos, atencao, monitorar
                        )
                        if ok:
                            st.success(f"📧 {msg}")
                        else:
                            st.warning(f"⚠️ Relatório gerado, mas e-mail não enviado: {msg}")
                    else:
                        st.warning("⚠️ Biblioteca SendGrid não instalada. E-mail não enviado.")

                    # Download manual sempre disponível
                    st.download_button(
                        label="⬇️  Baixar Relatório Excel",
                        data=excel_bytes,
                        file_name=f"relatorio_cs_{datetime.now().strftime('%d%m%Y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                except Exception as e:
                    st.error(f"Erro ao processar: {e}")
                    st.info("Verifique se os arquivos corretos foram enviados com os filtros indicados.")

st.markdown("""
<br>
<div style="text-align:center; color:#aaa; font-size:0.78rem; padding: 16px 0;">
    Rehagro · Customer Success · Agente de Engajamento
</div>
""", unsafe_allow_html=True)
