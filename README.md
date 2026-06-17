# Agente CS Rehagro — Engajamento de Alunos

Aplicação web para identificar alunos desengajados no GPL a partir de três fontes do Power BI.

## Como publicar no Streamlit Cloud (gratuito)

### 1. Crie uma conta no GitHub
Acesse https://github.com e crie uma conta gratuita.

### 2. Crie um repositório
- Clique em "New repository"
- Nome sugerido: `rehagro-cs-engajamento`
- Deixe como **Public**
- Clique em "Create repository"

### 3. Suba os arquivos
Faça upload dos três arquivos desta pasta:
- `app.py`
- `requirements.txt`
- `README.md`

### 4. Publique no Streamlit Cloud
- Acesse https://share.streamlit.io
- Faça login com sua conta GitHub
- Clique em "New app"
- Selecione o repositório criado
- Em "Main file path" coloque: `app.py`
- Clique em "Deploy!"

### 5. Pronto!
Em ~2 minutos você terá um link público como:
`https://rehagro-cs-engajamento.streamlit.app`

Compartilhe com qualquer pessoa da Rehagro.

## Critérios de alerta

Um aluno entra no relatório quando dispara pelo menos um destes critérios:

- 🟡 **Sem acesso ao Canvas** — não acessa a plataforma há mais de **20, 40 ou 60 dias** (limite selecionável na tela) → enviar link de acesso.
- 🟠 **Ausente nas aulas ao vivo** — faltou às 2 últimas videoconferências seguidas (apenas alunos ativos; aprovados, reprovados e desistentes ficam de fora) → enviar a data da próxima aula.
- 🔴 **Última avaliação detratora** — a avaliação de aula mais recente foi negativa (NPS detrator) → retomar o feedback negativo.
- 🔵 **Presente, mas não avaliou** — presente nas 2 últimas aulas, mas não respondeu à avaliação → incentivar a participação nas avaliações.
- 🔴 **Detratou e depois faltou** — detrator na penúltima aula e ausente na última → retomar o feedback e verificar a ausência.
- 🟢 **Comentário na avaliação** — escreveu um comentário na avaliação de aula → ler e dar retorno ao aluno.

**Severidade:** cada alerta vale 1 ponto — 🔴 Crítico (4+) · 🟠 Atenção (2–3) · 🟡 Monitorar (1).

## Relatório Excel gerado
- **Aba 1 — Relatório CS:** alertas e ações por aluno, com legenda dos critérios.
- **Aba 2 — Resumo por Turma:** contagem de alertas por turma.
- **Aba 3 — Frequência ao Vivo:** % de presença de cada aluno nas videoconferências a que foi exposto (após entrar no curso), com KPIs e gráficos. Exclui aprovados, reprovados e desistentes.
