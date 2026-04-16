import streamlit as st

st.set_page_config(
    page_title="CS Rehagro",
    page_icon="🌱",
    layout="wide",
    initial_sidebar_state="expanded",
)

pg = st.navigation([
    st.Page("pages/monitoramento.py",        title="Monitoramento de Alunos", icon="📊", default=True),
    st.Page("pages/comportamento_aluno.py",  title="Comportamento do Aluno",  icon="🔍"),
])
pg.run()
