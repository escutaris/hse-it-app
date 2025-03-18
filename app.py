import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
from fpdf import FPDF
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime

# Configuração da página
st.set_page_config(
    page_title="Avaliação de Fatores Psicossociais - HSE-IT",
    page_icon="📊",
    layout="wide"
)

# Função para verificar senha (autenticação básica)
def check_password():
    """Retorna `True` se a senha estiver correta."""
    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False

    if st.session_state.password_correct:
        return True

    # Formulário de login
    with st.form("login_form"):
        st.text_input("Usuário", key="username")
        st.text_input("Senha", type="password", key="password")
        submitted = st.form_submit_button("Login")

    if submitted:
        if st.session_state.password == "senha123" and st.session_state.username == "admin":
            st.session_state.password_correct = True
            return True
        else:
            st.error("Usuário ou senha incorretos")
            return False
    return False

# Definir as questões invertidas do HSE-IT
QUESTOES_INVERTIDAS = [3, 5, 6, 9, 12, 14, 16, 18, 20, 21, 22, 34]

# Definir os fatores do HSE-IT (dimensões originais)
DIMENSOES_HSE = {
    "Demanda": [3, 6, 9, 12, 16, 18, 20, 22],
    "Controle": [2, 10, 15, 19, 25, 30],
    "Apoio da Chefia": [8, 23, 29, 33, 35],
    "Apoio dos Colegas": [7, 24, 27, 31],
    "Relacionamentos": [5, 14, 21, 34],
    "Função": [1, 4, 11, 13, 17],
    "Mudança": [26, 28, 32]
}

# Descrições das dimensões para informação do usuário
DESCRICOES_DIMENSOES = {
    "Demanda": "Inclui aspectos como carga de trabalho, padrões e ambiente de trabalho.",
    "Controle": "Refere-se a quanto controle a pessoa tem sobre como realiza seu trabalho.",
    "Apoio da Chefia": "O incentivo, patrocínio e recursos fornecidos pela organização e liderança.",
    "Apoio dos Colegas": "O incentivo, patrocínio e recursos fornecidos pelos colegas.",
    "Relacionamentos": "Inclui a promoção de trabalho positivo para evitar conflitos e lidar com comportamentos inaceitáveis.",
    "Função": "Se as pessoas entendem seu papel na organização e se a organização garante que não tenham papéis conflitantes.",
    "Mudança": "Como as mudanças organizacionais são gerenciadas e comunicadas."
}

# Função para classificar os riscos com base na pontuação média
@st.cache_data
def classificar_risco(media):
    if media <= 1:
        return "Risco Muito Alto 🔴", "red"
    elif media > 1 and media <= 2:
        return "Risco Alto 🟠", "orange"
    elif media > 2 and media <= 3:
        return "Risco Moderado 🟡", "yellow"
    elif media > 3 and media <= 4:
        return "Risco Baixo 🟢", "green"
    else:
        return "Risco Muito Baixo 🟣", "purple"

# Função para carregar e processar dados
@st.cache_data
def carregar_dados(uploaded_file):
    # Determinar o tipo de arquivo e carregá-lo adequadamente
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file, sep=';' if ';' in uploaded_file.getvalue().decode('utf-8', errors='replace')[:1000] else ',')
    else:
        df = pd.read_excel(uploaded_file)
    
    # Separar os dados de filtro e as perguntas
    colunas_filtro = list(df.columns[:7])
    if "Carimbo de data/hora" in colunas_filtro:
        colunas_filtro.remove("Carimbo de data/hora")
    
    # Garantir que as colunas das perguntas são corretamente identificadas
    colunas_perguntas = [col for col in df.columns if str(col).strip() and str(col).strip()[0].isdigit()]
    
    # Converter valores para numéricos, tratando erros
    df_perguntas = df[colunas_perguntas]
    for col in df_perguntas.columns:
        df_perguntas[col] = pd.to_numeric(df_perguntas[col], errors='coerce')
    
    return df, df_perguntas, colunas_filtro, colunas_perguntas

# Função para processar os dados (incluindo inversão de questões)
def processar_dados_hse(df_perguntas, colunas_perguntas):
    # Copiar o dataframe para não modificar o original
    df_processado = df_perguntas.copy()
    
    # Inverter a pontuação das questões invertidas
    for col in colunas_perguntas:
        # Extrair o número da questão (considerar que o formato é "1. Pergunta")
        try:
            numero_questao = int(str(col).strip().split('.')[0])
            if numero_questao in QUESTOES_INVERTIDAS:
                df_processado[col] = 6 - df_processado[col]  # Inverte a escala (1->5, 2->4, 3->3, 4->2, 5->1)
        except (ValueError, IndexError):
            continue
    
    return df_processado

# Função para calcular resultados por dimensão
@st.cache_data
def calcular_resultados_dimensoes(df, df_perguntas_filtradas, colunas_perguntas):
    # Processar dados (inverter questões necessárias)
    df_processado = processar_dados_hse(df_perguntas_filtradas, colunas_perguntas)
    
    resultados = []
    cores = []
    
    # Número total de respostas após filtragem
    num_total_respostas = len(df_processado)
    
    # Calcular resultados para cada dimensão
    for dimensao, numeros_questoes in DIMENSOES_HSE.items():
        # Converter números de pergunta para índices de coluna
        indices_validos = []
        for q in numeros_questoes:
            for col in colunas_perguntas:
                if str(col).strip().startswith(str(q) + ".") or str(col).strip().startswith(str(q) + " "):
                    indices_validos.append(col)
                    break
        
        if indices_validos:
            # Calcular média para esta dimensão
            valores = df_processado[indices_validos].values.flatten()
            valores = valores[~pd.isna(valores)]  # Remove NaN
            
            if len(valores) > 0:
                media = valores.mean()
                risco, cor = classificar_risco(media)
                
                resultados.append({
                    "Dimensão": dimensao,
                    "Descrição": DESCRICOES_DIMENSOES[dimensao],
                    "Média": round(media, 2),
                    "Risco": risco,
                    "Número de Respostas": num_total_respostas,
                    "Questões": numeros_questoes
                })
                cores.append(cor)
    
    return resultados

# Função para gerar sugestões de ações com base no nível de risco
def gerar_sugestoes_acoes(df_resultados):
    # Dicionário com sugestões de ações para cada dimensão por nível de risco
    sugestoes_por_dimensao = {
        "Demanda": {
            "Risco Muito Alto": [
                "Realizar auditoria completa da distribuição de carga de trabalho",
                "Implementar sistema de gestão de tarefas e priorização",
                "Reavaliar prazos e expectativas de produtividade",
                "Contratar pessoal adicional para áreas sobrecarregadas",
                "Implementar pausas obrigatórias durante o dia de trabalho"
            ],
            "Risco Alto": [
                "Mapear atividades e identificar gargalos de processo",
                "Implementar ferramentas para melhor organização e planejamento do trabalho",
                "Revisar e ajustar prazos de entregas e metas",
                "Capacitar gestores em gerenciamento de carga de trabalho das equipes"
            ],
            "Risco Moderado": [
                "Promover treinamentos de gestão do tempo e priorização",
                "Revisar distribuição de tarefas entre membros da equipe",
                "Estabelecer momentos regulares para feedback sobre carga de trabalho"
            ],
            "Risco Baixo": [
                "Manter monitoramento regular das demandas de trabalho",
                "Realizar check-ins periódicos sobre volume de trabalho",
                "Oferecer recursos de apoio para períodos de pico de trabalho"
            ],
            "Risco Muito Baixo": [
                "Documentar boas práticas atuais de gestão de demandas",
                "Compartilhar casos de sucesso no gerenciamento de carga de trabalho",
                "Manter práticas de gestão de demandas e continuar monitorando"
            ]
        },
        "Controle": {
            "Risco Muito Alto": [
                "Redesenhar processos para aumentar a autonomia dos trabalhadores",
                "Implementar esquemas de trabalho flexível",
                "Revisar políticas de microgerenciamento",
                "Treinar gestores em delegação efetiva",
                "Criar espaços para participação em decisões estratégicas"
            ],
            "Risco Alto": [
                "Identificar áreas específicas onde os trabalhadores podem ter mais controle",
                "Envolver colaboradores no planejamento de metas e métodos de trabalho",
                "Implementar sistema de sugestões para melhorias nos processos",
                "Oferecer opções de horários flexíveis"
            ],
            "Risco Moderado": [
                "Aumentar gradualmente a autonomia nas decisões rotineiras",
                "Solicitar feedback regular sobre nível de controle no trabalho",
                "Implementar projetos-piloto para testar maior autonomia"
            ],
            "Risco Baixo": [
                "Manter boas práticas de autonomia",
                "Revisar periodicamente áreas onde o controle pode ser ampliado",
                "Reconhecer e celebrar iniciativas independentes"
            ],
            "Risco Muito Baixo": [
                "Documentar práticas bem-sucedidas de autonomia",
                "Compartilhar casos de sucesso com outras áreas da organização",
                "Manter cultura de confiança e delegação"
            ]
        },
        "Apoio da Chefia": {
            "Risco Muito Alto": [
                "Implementar programa estruturado de desenvolvimento de lideranças",
                "Realizar avaliação 360° para gestores",
                "Estabelecer canais de comunicação direta entre colaboradores e alta direção",
                "Revisar políticas de promoção para garantir que bons líderes sejam reconhecidos",
                "Oferecer coaching individual para gestores com desafios específicos"
            ],
            "Risco Alto": [
                "Treinar gestores em habilidades de suporte e feedback",
                "Implementar reuniões regulares one-on-one entre líderes e liderados",
                "Estabelecer expectativas claras para comportamentos de liderança",
                "Criar fóruns para líderes compartilharem desafios e soluções"
            ],
            "Risco Moderado": [
                "Revisar e melhorar as práticas de feedback das lideranças",
                "Oferecer recursos e ferramentas para líderes apoiarem suas equipes",
                "Promover workshops sobre comunicação efetiva"
            ],
            "Risco Baixo": [
                "Manter programas de desenvolvimento de lideranças",
                "Reconhecer e celebrar boas práticas de liderança",
                "Implementar sistema de mentoria entre líderes"
            ],
            "Risco Muito Baixo": [
                "Documentar e compartilhar práticas exemplares de liderança",
                "Utilizar líderes como mentores para novos gestores",
                "Manter cultura de apoio e desenvolvimento contínuo"
            ]
        },
        "Apoio dos Colegas": {
            "Risco Muito Alto": [
                "Implementar programas estruturados de team building",
                "Revisar a composição e dinâmica das equipes",
                "Estabelecer facilitadores de equipe para melhorar integração",
                "Criar espaços físicos e virtuais para colaboração",
                "Implementar sistema de reconhecimento por comportamentos colaborativos"
            ],
            "Risco Alto": [
                "Promover atividades regulares de integração de equipes",
                "Treinar em habilidades de trabalho em equipe e comunicação",
                "Estabelecer objetivos compartilhados que incentivem a colaboração",
                "Revisar processos que possam estar criando competição indesejada"
            ],
            "Risco Moderado": [
                "Implementar reuniões regulares de equipe para compartilhamento",
                "Criar projetos colaborativos entre diferentes membros",
                "Oferecer oportunidades para pessoas se conhecerem melhor"
            ],
            "Risco Baixo": [
                "Manter momentos regulares de integração",
                "Monitorar dinâmicas de equipe, especialmente com novos membros",
                "Reconhecer comportamentos de apoio entre colegas"
            ],
            "Risco Muito Baixo": [
                "Documentar boas práticas de colaboração",
                "Utilizar a cultura de apoio como exemplo para novos funcionários",
                "Manter ambiente de confiança e colaboração"
            ]
        },
        "Relacionamentos": {
            "Risco Muito Alto": [
                "Implementar política de tolerância zero para assédio e comportamentos inadequados",
                "Criar canais confidenciais para denúncias",
                "Treinar todos os colaboradores em respeito e diversidade",
                "Estabelecer mediação de conflitos com profissionais externos",
                "Auditar clima organizacional e relacionamentos interpessoais"
            ],
            "Risco Alto": [
                "Desenvolver e comunicar política clara sobre comportamentos aceitáveis",
                "Treinar gestores na identificação e gestão de conflitos",
                "Implementar processos estruturados para resolução de conflitos",
                "Promover diálogo aberto sobre relacionamentos saudáveis no trabalho"
            ],
            "Risco Moderado": [
                "Realizar workshops sobre comunicação não-violenta",
                "Estabelecer acordos de equipe sobre comportamentos esperados",
                "Promover atividades que construam confiança entre colegas"
            ],
            "Risco Baixo": [
                "Manter comunicação regular sobre respeito no ambiente de trabalho",
                "Incorporar avaliação de relacionamentos nas pesquisas de clima",
                "Reconhecer exemplos positivos de resolução de conflitos"
            ],
            "Risco Muito Baixo": [
                "Documentar boas práticas de relacionamentos saudáveis",
                "Utilizar a cultura positiva como diferencial da organização",
                "Manter monitoramento contínuo do clima relacional"
            ]
        },
        "Função": {
            "Risco Muito Alto": [
                "Realizar revisão completa de descrições de cargos e responsabilidades",
                "Implementar processo de clarificação de funções para toda a organização",
                "Treinar gestores em delegação clara e definição de expectativas",
                "Estabelecer processos para resolver conflitos de papéis e responsabilidades",
                "Revisar estrutura organizacional para eliminar ambiguidades"
            ],
            "Risco Alto": [
                "Revisar e atualizar descrições de cargo",
                "Implementar reuniões regulares para esclarecer expectativas",
                "Criar matriz RACI (Responsável, Aprovador, Consultado, Informado) para projetos",
                "Treinar equipes em comunicação sobre papéis e responsabilidades"
            ],
            "Risco Moderado": [
                "Revisar processos onde ocorrem conflitos de funções",
                "Promover workshops para clarificar interfaces entre áreas",
                "Estabelecer fóruns para discutir e esclarecer papéis em projetos"
            ],
            "Risco Baixo": [
                "Manter atualizações periódicas de responsabilidades",
                "Incluir clareza de papéis nas avaliações de desempenho",
                "Promover comunicação contínua sobre expectativas"
            ],
            "Risco Muito Baixo": [
                "Documentar boas práticas de clareza de funções",
                "Utilizar como modelo para novos departamentos ou projetos",
                "Manter cultura de transparência sobre papéis e responsabilidades"
            ]
        },
        "Mudança": {
            "Risco Muito Alto": [
                "Implementar metodologia estruturada de gestão de mudanças",
                "Criar comitê representativo para planejamento de mudanças",
                "Estabelecer múltiplos canais de comunicação sobre processos de mudança",
                "Treinar gestores em liderança durante transformações",
                "Avaliar impacto de mudanças anteriores e implementar lições aprendidas"
            ],
            "Risco Alto": [
                "Desenvolver plano de comunicação para mudanças organizacionais",
                "Envolver representantes de diferentes níveis no planejamento",
                "Implementar feedbacks regulares durante processos de mudança",
                "Oferecer suporte adicional para equipes mais afetadas"
            ],
            "Risco Moderado": [
                "Melhorar a transparência sobre razões das mudanças",
                "Criar fóruns para esclarecer dúvidas sobre transformações",
                "Celebrar pequenas vitórias durante processos de mudança"
            ],
            "Risco Baixo": [
                "Manter comunicação proativa sobre possíveis mudanças",
                "Oferecer oportunidades regulares para feedback durante transformações",
                "Reconhecer contribuições positivas em processos de mudança"
            ],
            "Risco Muito Baixo": [
                "Documentar práticas bem-sucedidas de gestão de mudanças",
                "Utilizar abordagem participativa como padrão",
                "Manter cultura de adaptabilidade e melhoria contínua"
            ]
        }
    }

    plano_acao = []

    # Para cada dimensão no resultado
    for _, row in df_resultados.iterrows():
        dimensao = row['Dimensão']
        risco = row['Risco']
        nivel_risco = risco.split()[0] + " " + risco.split()[1]  # Ex: "Risco Alto"

        # Obter sugestões para esta dimensão e nível de risco
        if dimensao in sugestoes_por_dimensao and nivel_risco in sugestoes_por_dimensao[dimensao]:
            sugestoes = sugestoes_por_dimensao[dimensao][nivel_risco]

            # Adicionar ao plano de ação
            for sugestao in sugestoes:
                plano_acao.append({
                    "Dimensão": dimensao,
                    "Nível de Risco": nivel_risco,
                    "Média": row['Média'],
                    "Sugestão de Ação": sugestao,
                    "Responsável": "",
                    "Prazo": "",
                    "Status": "Não iniciada"
                })

    # Criar DataFrame com o plano de ação
    df_plano_acao = pd.DataFrame(plano_acao)
    return df_plano_acao

# Função para formatar aba do plano de ação
def formatar_aba_plano_acao(workbook, worksheet, df):
    # Definir formatos
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D7E4BC',
        'border': 1
    })

    risco_format = {
        'Risco Muito Alto': workbook.add_format({'bg_color': '#FF6B6B', 'font_color': 'white'}),
        'Risco Alto': workbook.add_format({'bg_color': '#FFA500'}),
        'Risco Moderado': workbook.add_format({'bg_color': '#FFFF00'}),
        'Risco Baixo': workbook.add_format({'bg_color': '#90EE90'}),
        'Risco Muito Baixo': workbook.add_format({'bg_color': '#BB8FCE'})
    }

    # Configurar largura das colunas
    worksheet.set_column('A:A', 25)  # Dimensão
    worksheet.set_column('B:B', 15)  # Nível de Risco
    worksheet.set_column('C:C', 10)  # Média
    worksheet.set_column('D:D', 50)  # Sugestão de Ação
    worksheet.set_column('E:E', 15)  # Responsável
    worksheet.set_column('F:F', 15)  # Prazo
    worksheet.set_column('G:G', 15)  # Status

    # Adicionar cabeçalhos formatados
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    # Aplicar formatação condicional baseada no nível de risco
    for row_num, (_, row) in enumerate(df.iterrows(), 1):
        nivel_risco = row["Nível de Risco"]
        if nivel_risco in risco_format:
            worksheet.write(row_num, 1, nivel_risco, risco_format[nivel_risco])

    # Adicionar validação de dados para a coluna Status
    status_options = ['Não iniciada', 'Em andamento', 'Concluída', 'Cancelada']
    worksheet.data_validation('G2:G1000', {'validate': 'list',
                                         'source': status_options,
                                         'input_title': 'Selecione o status:',
                                         'input_message': 'Escolha um status da lista'})

    # Adicionar filtros
    worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)

    # Congelar painel para manter cabeçalhos visíveis durante rolagem
    worksheet.freeze_panes(1, 0)

# Função para formatar a aba Excel
def formatar_aba_excel(workbook, worksheet, df, is_pivot=False):
    # Definir formatos
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D7E4BC',
        'border': 1
    })

    # Configurar largura das colunas
    if not is_pivot:
        worksheet.set_column('A:A', 5)  # Índice
        worksheet.set_column('B:B', 25)  # Dimensão
        worksheet.set_column('C:C', 40)  # Descrição
        worksheet.set_column('D:D', 10)  # Média
        worksheet.set_column('E:E', 20)  # Risco
        worksheet.set_column('F:F', 15)  # Número de Respostas
    else:
        # Para tabelas pivotadas
        worksheet.set_column('A:A', 25)  # Dimensão
        for col in range(1, len(df.columns) + 1):
            worksheet.set_column(col, col, 12)  # Colunas de valores

    # Adicionar cabeçalhos formatados
    if not is_pivot:
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
    
    # Adicionar filtros para tabelas não pivotadas
    if not is_pivot:
        worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
        
    # Congelar painel para manter cabeçalhos visíveis
    worksheet.freeze_panes(1, 0)

# Função para aplicar formatação condicional em tabelas pivotadas
def aplicar_formatacao_condicional(workbook, worksheet, df_pivot):
    # Definir formatos para níveis de risco
    risco_muito_alto = workbook.add_format({'bg_color': '#FF6B6B', 'font_color': 'white'})
    risco_alto = workbook.add_format({'bg_color': '#FFA500'})
    risco_moderado = workbook.add_format({'bg_color': '#FFFF00'})
    risco_baixo = workbook.add_format({'bg_color': '#90EE90'})
    risco_muito_baixo = workbook.add_format({'bg_color': '#BB8FCE'})
    
    # Aplicar formatação condicional para todas as células de dados
    for col in range(1, len(df_pivot.columns) + 1):
        for row in range(1, len(df_pivot) + 1):
            worksheet.conditional_format(row, col, row, col, {
                'type': 'cell',
                'criteria': '<=',
                'value': 1,
                'format': risco_muito_alto
            })
            worksheet.conditional_format(row, col, row, col, {
                'type': 'cell',
                'criteria': 'between',
                'minimum': 1.01,
                'maximum': 2,
                'format': risco_alto
            })
            worksheet.conditional_format(row, col, row, col, {
                'type': 'cell',
                'criteria': 'between',
                'minimum': 2.01,
                'maximum': 3,
                'format': risco_moderado
            })
            worksheet.conditional_format(row, col, row, col, {
                'type': 'cell',
                'criteria': 'between',
                'minimum': 3.01,
                'maximum': 4,
                'format': risco_baixo
            })
            worksheet.conditional_format(row, col, row, col, {
                'type': 'cell',
                'criteria': '>',
                'value': 4,
                'format': risco_muito_baixo
            })

# Função para gerar o arquivo PDF do Plano de Ação
def gerar_pdf_plano_acao(df_plano_acao):
    try:
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        
        # Configurando fonte para suporte básico
        pdf.set_font("Arial", style='B', size=16)
        
        # Título
        pdf.cell(200, 10, "Plano de Acao - HSE-IT: Fatores Psicossociais", ln=True, align='C')
        pdf.ln(10)
        
        # Data do relatório
        pdf.set_font("Arial", size=10)
        data_atual = datetime.now().strftime("%d/%m/%Y")
        pdf.cell(0, 10, f"Data do relatorio: {data_atual}", ln=True)
        pdf.ln(5)
        
        # Agrupar por dimensão
        for dimensao in df_plano_acao["Dimensão"].unique():
            df_dimensao = df_plano_acao[df_plano_acao["Dimensão"] == dimensao]
            nivel_risco = df_dimensao["Nível de Risco"].iloc[0]
            media = df_dimensao["Média"].iloc[0]
            
            # Título da dimensão
            pdf.set_font("Arial", style='B', size=12)
            pdf.cell(0, 10, f"{dimensao}", ln=True)
            
            # Adicionar descrição da dimensão
            pdf.set_font("Arial", style='I', size=10)
            pdf.multi_cell(0, 6, DESCRICOES_DIMENSOES[dimensao], 0)
            
            # Informações de risco
            pdf.set_font("Arial", style='I', size=10)
            pdf.cell(0, 8, f"Media: {media} - Nivel de Risco: {nivel_risco}", ln=True)
            
            # Ações sugeridas
            pdf.set_font("Arial", size=10)
            pdf.cell(0, 6, "Acoes Sugeridas:", ln=True)
            
            for _, row in df_dimensao.iterrows():
                pdf.set_x(15)  # Recuo para lista
                pdf.cell(0, 6, f"- {row['Sugestão de Ação'].encode('ascii', 'replace').decode('ascii')}", ln=True)
            
            # Espaço para tabela de responsáveis e prazos
            pdf.ln(5)
            pdf.cell(0, 6, "Implementacao:", ln=True)
            
            # Criar tabela
            col_width = [100, 40, 40]
            pdf.set_font("Arial", style='B', size=8)
            
            # Cabeçalho da tabela
            pdf.set_x(15)
            pdf.cell(col_width[0], 7, "Acao", border=1)
            pdf.cell(col_width[1], 7, "Responsavel", border=1)
            pdf.cell(col_width[2], 7, "Prazo", border=1)
            pdf.ln()
            
            # Linhas da tabela
            pdf.set_font("Arial", size=8)
            for _, row in df_dimensao.iterrows():
                pdf.set_x(15)
                acao = row['Sugestão de Ação'].encode('ascii', 'replace').decode('ascii')
                
                # Verificar se o texto é muito longo
                if len(acao) > 60:
                    acao = acao[:57] + "..."
                
                pdf.cell(col_width[0], 7, acao, border=1)
                pdf.cell(col_width[1], 7, "", border=1)
                pdf.cell(col_width[2], 7, "", border=1)
                pdf.ln()
            
            pdf.ln(10)
        
        # Adicionar informações sobre a interpretação dos riscos
        pdf.ln(10)
        pdf.set_font("Arial", style='B', size=11)
        pdf.cell(0, 8, "Legenda de Classificacao de Riscos:", ln=True)
        pdf.set_font("Arial", size=9)
        pdf.cell(0, 6, "Risco Muito Alto: Media <= 1", ln=True)
        pdf.cell(0, 6, "Risco Alto: 1 < Media <= 2", ln=True)
        pdf.cell(0, 6, "Risco Moderado: 2 < Media <= 3", ln=True)
        pdf.cell(0, 6, "Risco Baixo: 3 < Media <= 4", ln=True)
        pdf.cell(0, 6, "Risco Muito Baixo: Media > 4", ln=True)
        
        # Adicionar observações sobre o plano de ação
        pdf.ln(10)
        pdf.set_font("Arial", style='B', size=11)
        pdf.cell(0, 8, "Observacoes:", ln=True)
        pdf.set_font("Arial", size=9)
        pdf.multi_cell(0, 6, "Este plano de acao foi gerado automaticamente com base nos resultados da avaliacao HSE-IT. As acoes sugeridas devem ser analisadas e adaptadas ao contexto especifico da organizacao. Recomenda-se definir responsaveis, prazos e indicadores de monitoramento para cada acao implementada.", 0)
        
        pdf.multi_cell(0, 6, "O questionario HSE-IT avalia 7 dimensoes de fatores psicossociais no trabalho: Demanda, Controle, Apoio da Chefia, Apoio dos Colegas, Relacionamentos, Funcao e Mudanca. Priorize as acoes nas dimensoes com maior risco.", 0)
        
        # Corrigindo o problema de BytesIO
        temp_file = "temp_plano_acao.pdf"
        pdf.output(temp_file)
        
        with open(temp_file, "rb") as file:
            output = io.BytesIO(file.read())
        
        import os
        if os.path.exists(temp_file):
            os.remove(temp_file)
        
        return output
    
    except Exception as e:
        st.error(f"Erro ao gerar o PDF do Plano de Ação: {str(e)}")
        return None

# Função para gerar Excel completo com múltiplas abas
def gerar_excel_completo(df, df_perguntas, colunas_filtro, colunas_perguntas):
    output = io.BytesIO()
    
    try:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Aba 1: Empresa Toda - Dimensões
            resultados_empresa = calcular_resultados_dimensoes(df, df_perguntas, colunas_perguntas)
            df_resultados_empresa = pd.DataFrame(resultados_empresa)
            
            # Remover a coluna de questões para o Excel (mantém a visualização mais limpa)
            df_resultados_excel = df_resultados_empresa.copy()
            if 'Questões' in df_resultados_excel.columns:
                df_resultados_excel = df_resultados_excel.drop(columns=['Questões'])
            
            df_resultados_excel.to_excel(writer, sheet_name='Empresa Toda', index=False)
            
            # Formatar a planilha
            worksheet = writer.sheets['Empresa Toda']
            formatar_aba_excel(workbook, worksheet, df_resultados_excel)
            
            # Aba 2: Plano de Ação
            df_plano_acao = gerar_sugestoes_acoes(df_resultados_empresa)
            df_plano_acao.to_excel(writer, sheet_name='Plano de Ação', index=False)
            
            # Formatar a aba de plano de ação
            worksheet_plano = writer.sheets['Plano de Ação']
            formatar_aba_plano_acao(workbook, worksheet_plano, df_plano_acao)
            
            # Aba 3: Detalhes das Questões (explicação da metodologia)
            worksheet_detalhes = workbook.add_worksheet('Detalhes das Dimensões')
            
            # Cabeçalhos
            headers = ["Dimensão", "Questões", "Descrição"]
            for col, header in enumerate(headers):
                worksheet_detalhes.write(0, col, header, workbook.add_format({'bold': True, 'bg_color': '#D7E4BC'}))
            
            # Configurar largura das colunas
            worksheet_detalhes.set_column('A:A', 20)  # Dimensão
            worksheet_detalhes.set_column('B:B', 30)  # Questões
            worksheet_detalhes.set_column('C:C', 50)  # Descrição
            
            # Preencher dados
            row = 1
            for dimensao, questoes in DIMENSOES_HSE.items():
                worksheet_detalhes.write(row, 0, dimensao)
                worksheet_detalhes.write(row, 1, str(questoes))
                worksheet_detalhes.write(row, 2, DESCRICOES_DIMENSOES[dimensao])
                row += 1
            
            # Adicionar abas para cada tipo de filtro
            for filtro in colunas_filtro:
                if filtro != "Carimbo de data/hora":
                    valores_unicos = df[filtro].dropna().unique()
                    
                    # Criar uma aba resumo para este filtro
                    sheet_name = f'Por {filtro}'
                    if len(sheet_name) > 31:  # Excel limita nomes de abas a 31 caracteres
                        sheet_name = sheet_name[:31]
                    
                    # Criar DataFrame para o resumo deste filtro
                    resultados_resumo = []
                    
                    # Para cada valor único do filtro, calcular resultados
                    for valor in valores_unicos:
                        if pd.notna(valor):
                            df_filtrado = df[df[filtro] == valor]
                            indices_filtrados = df.index[df[filtro] == valor].tolist()
                            df_perguntas_filtradas = df_perguntas.loc[indices_filtrados]
                            
                            # Calcular resultados para este filtro
                            resultados_filtro = calcular_resultados_dimensoes(df_filtrado, df_perguntas_filtradas, colunas_perguntas)
                            
                            # Adicionar coluna com o valor do filtro
                            for res in resultados_filtro:
                                res[filtro] = valor
                                resultados_resumo.append(res)
                    
                    if resultados_resumo:
                        df_resumo = pd.DataFrame(resultados_resumo)
                        
                        # Remover a coluna de questões para o Excel
                        if 'Questões' in df_resumo.columns:
                            df_resumo = df_resumo.drop(columns=['Questões'])
                        
                        # Pivotear para melhor visualização
                        if len(resultados_resumo) > 0:
                            try:
                                df_pivot = df_resumo.pivot(index='Dimensão', columns=filtro, values='Média')
                                df_pivot.to_excel(writer, sheet_name=sheet_name)
                                
                                # Formatar a planilha pivotada
                                worksheet = writer.sheets[sheet_name]
                                formatar_aba_excel(workbook, worksheet, df_pivot, is_pivot=True)
                                
                                # Adicionar informação de risco usando formatação condicional
                                aplicar_formatacao_condicional(workbook, worksheet, df_pivot)
                            except Exception as e:
                                # Fallback caso pivot falhe
                                df_resumo.to_excel(writer, sheet_name=sheet_name, index=False)
                                worksheet = writer.sheets[sheet_name]
                                formatar_aba_excel(workbook, worksheet, df_resumo)
            
            # Adicionar aba com gráfico de resultados gerais
            adicionar_aba_grafico(writer, workbook, df_resultados_empresa)
        
        output.seek(0)
        return output
    
    except Exception as e:
        st.error(f"Erro ao gerar o arquivo Excel: {str(e)}")
        return None

# Função para adicionar uma aba com gráfico dos resultados
def adicionar_aba_grafico(writer, workbook, df_resultados):
    # Criar uma nova aba para o gráfico
    worksheet = workbook.add_worksheet('Gráfico de Riscos')
    
    # Adicionar os dados para o gráfico
    worksheet.write_column('A1', ['Dimensão'] + list(df_resultados['Dimensão']))
    worksheet.write_column('B1', ['Média'] + list(df_resultados['Média']))
    worksheet.write_column('C1', ['Risco'] + list(df_resultados['Risco']))
    
    # Criar o gráfico
    chart = workbook.add_chart({'type': 'bar'})
    
    # Configurar o gráfico
    chart.add_series({
        'name': 'Média',
        'categories': ['Gráfico de Riscos', 1, 0, len(df_resultados), 0],
        'values': ['Gráfico de Riscos', 1, 1, len(df_resultados), 1],
        'data_labels': {'value': True},
        'fill': {'color': '#4472C4'}
    })
    
    # Configurar aparência do gráfico
    chart.set_title({'name': 'HSE-IT: Fatores Psicossociais - Avaliação de Riscos'})
    chart.set_x_axis({'name': 'Dimensão'})
    chart.set_y_axis({'name': 'Média (1-5)', 'min': 0, 'max': 5})
    chart.set_legend({'position': 'bottom'})
    chart.set_size({'width': 720, 'height': 576})
    
    # Adicionar linhas de referência para os níveis de risco
    chart.set_y_axis({
        'name': 'Média (1-5)',
        'min': 0,
        'max': 5,
        'major_gridlines': {'visible': True},
        'minor_gridlines': {'visible': False},
        'major_unit': 1
    })
    
    # Inserir o gráfico na planilha
    worksheet.insert_chart('E1', chart, {'x_scale': 1.5, 'y_scale': 1.5})
    
    # Adicionar legenda de interpretação dos riscos
    worksheet.write('A20', 'Interpretação dos Riscos:', workbook.add_format({'bold': True}))
    worksheet.write('A21', 'Média ≤ 1: Risco Muito Alto')
    worksheet.write('A22', '1 < Média ≤ 2: Risco Alto')
    worksheet.write('A23', '2 < Média ≤ 3: Risco Moderado')
    worksheet.write('A24', '3 < Média ≤ 4: Risco Baixo')
    worksheet.write('A25', 'Média > 4: Risco Muito Baixo')

# Função para gerar PDF do relatório
def gerar_pdf(df_resultados):
    try:
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        
        # Configurando fonte para suporte básico
        pdf.set_font("Arial", style='B', size=14)
        
        # Usando strings sem acentos para evitar problemas de codificação
        pdf.cell(200, 10, "Relatorio de Fatores Psicossociais - HSE-IT", ln=True, align='C')
        pdf.ln(10)
        
        # Adicionar explicação da metodologia
        pdf.set_font("Arial", style='I', size=10)
        pdf.multi_cell(0, 5, "O questionario HSE-IT avalia 7 dimensoes de fatores psicossociais no trabalho. Os resultados sao apresentados em uma escala de 1 a 5, onde valores mais altos indicam melhores resultados.", 0)
        pdf.ln(5)
        
        # Tabela de resultados
        pdf.set_font("Arial", style='B', size=10)
        pdf.cell(80, 7, "Dimensao", 1)
        pdf.cell(25, 7, "Media", 1)
        pdf.cell(60, 7, "Nivel de Risco", 1)
        pdf.ln()
        
        pdf.set_font("Arial", size=10)
        for index, row in df_resultados.iterrows():
            # Remover acentos e caracteres especiais
            dimensao = row['Dimensão'].encode('ascii', 'replace').decode('ascii')
            risco = row['Risco'].split(' ')[0] + ' ' + row['Risco'].split(' ')[1]  # Remove emoji
            
            # Determinar cor de fundo baseada no risco
            if "Muito Alto" in risco:
                pdf.set_fill_color(255, 107, 107)  # Vermelho claro
                text_color = 255  # Branco
            elif "Alto" in risco:
                pdf.set_fill_color(255, 165, 0)  # Laranja
                text_color = 0  # Preto
            elif "Moderado" in risco:
                pdf.set_fill_color(255, 255, 0)  # Amarelo
                text_color = 0  # Preto
            elif "Baixo" in risco:
                pdf.set_fill_color(144, 238, 144)  # Verde claro
                text_color = 0  # Preto
            elif "Muito Baixo" in risco:
                pdf.set_fill_color(187, 143, 206)  # Roxo claro
                text_color = 0  # Preto
            else:
                pdf.set_fill_color(255, 255, 255)  # Branco
                text_color = 0  # Preto
                
            pdf.set_text_color(text_color)
            pdf.cell(80, 7, dimensao, 1, 0, 'L', 1)
            pdf.cell(25, 7, f"{row['Média']:.2f}", 1, 0, 'C', 1)
            pdf.cell(60, 7, risco, 1, 0, 'C', 1)
            pdf.ln()
        
        # Resetar cor do texto
        pdf.set_text_color(0)
        
        # Adicionar descrições das dimensões
        pdf.ln(10)
        pdf.set_font("Arial", style='B', size=12)
        pdf.cell(0, 10, "Descricao das Dimensoes:", ln=True)
        
        for dimensao, descricao in DESCRICOES_DIMENSOES.items():
            pdf.set_font("Arial", style='B', size=10)
            pdf.cell(0, 7, dimensao.encode('ascii', 'replace').decode('ascii'), ln=True)
            pdf.set_font("Arial", size=10)
            pdf.multi_cell(0, 5, descricao.encode('ascii', 'replace').decode('ascii'), 0)
            pdf.ln(3)
        
        # Adicionar informações sobre classificação de risco
        pdf.ln(10)
        pdf.set_font("Arial", style='B', size=12)
        pdf.cell(0, 10, "Legenda de Classificacao:", ln=True)
        pdf.set_font("Arial", size=10)
        pdf.cell(0, 8, "Risco Muito Alto: Media <= 1", ln=True)
        pdf.cell(0, 8, "Risco Alto: 1 < Media <= 2", ln=True)
        pdf.cell(0, 8, "Risco Moderado: 2 < Media <= 3", ln=True)
        pdf.cell(0, 8, "Risco Baixo: 3 < Media <= 4", ln=True)
        pdf.cell(0, 8, "Risco Muito Baixo: Media > 4", ln=True)
        
        # Adicionar recomendações gerais
        pdf.ln(10)
        pdf.set_font("Arial", style='B', size=12)
        pdf.cell(0, 10, "Recomendacoes Gerais:", ln=True)
        pdf.set_font("Arial", size=10)
        pdf.multi_cell(0, 5, "Para resultados mais detalhados e um plano de acao personalizado, consulte o arquivo Excel completo ou o documento do Plano de Acao fornecido. Priorize intervencoes nas dimensoes com maior nivel de risco.", 0)
        
        # Corrigindo o problema de BytesIO
        temp_file = "temp_report.pdf"
        pdf.output(temp_file)
        
        with open(temp_file, "rb") as file:
            output = io.BytesIO(file.read())
        
        import os
        if os.path.exists(temp_file):
            os.remove(temp_file)
        
        return output
    
    except Exception as e:
        st.error(f"Erro ao gerar o PDF: {str(e)}")
        return None

# Função para criar dashboard melhorado com visualizações mais ricas
def criar_dashboard(df_resultados, filtro_opcao, filtro_valor):
    st.markdown("## Dashboard de Riscos Psicossociais HSE-IT")
    
    # Explicação sobre a metodologia HSE-IT
    with st.expander("Sobre a Metodologia HSE-IT"):
        st.write("""
        O HSE-IT (Health and Safety Executive - Indicator Tool) é um questionário desenvolvido pelo órgão britânico 
        de saúde e segurança ocupacional para avaliar os riscos psicossociais no ambiente de trabalho.
        
        O questionário avalia 7 dimensões:
        
        1. **Demanda**: Aspectos como carga de trabalho, padrões e ambiente de trabalho
        2. **Controle**: Quanto controle a pessoa tem sobre como realiza seu trabalho
        3. **Apoio da Chefia**: Incentivo, suporte e recursos fornecidos pela liderança
        4. **Apoio dos Colegas**: Incentivo e suporte fornecidos pelos colegas
        5. **Relacionamentos**: Promoção de trabalho positivo para evitar conflitos
        6. **Função**: Compreensão do papel na organização e ausência de conflitos de papel
        7. **Mudança**: Como mudanças organizacionais são gerenciadas e comunicadas
        
        A pontuação varia de 1 a 5, onde valores mais altos indicam melhores resultados.
        """)
    
    # Layout com 3 métricas principais
    col1, col2, col3 = st.columns(3)
    
    # Calcular métricas
    media_geral = df_resultados["Média"].mean()
    risco_geral, cor_geral = classificar_risco(media_geral)
    dimensao_mais_critica = df_resultados.loc[df_resultados["Média"].idxmin()]
    dimensao_melhor = df_resultados.loc[df_resultados["Média"].idxmax()]
    
    # Exibir métricas com formatação visual melhorada
    with col1:
        st.metric(
            label="Média Geral",
            value=f"{media_geral:.2f}",
            delta=risco_geral.split()[0],
            delta_color="inverse"
        )
        st.markdown(f"<div style='text-align: center; color: {cor_geral};'><b>{risco_geral}</b></div>", unsafe_allow_html=True)
    
    with col2:
        st.metric(
            label="Dimensão Mais Crítica",
            value=dimensao_mais_critica["Dimensão"],
            delta=f"{dimensao_mais_critica['Média']:.2f}",
            delta_color="off"
        )
        risco, cor = classificar_risco(dimensao_mais_critica["Média"])
        st.markdown(f"<div style='text-align: center; color: {cor};'><b>{risco}</b></div>", unsafe_allow_html=True)
    
    with col3:
        st.metric(
            label="Dimensão Melhor Avaliada",
            value=dimensao_melhor["Dimensão"],
            delta=f"{dimensao_melhor['Média']:.2f}",
            delta_color="off"
        )
        risco, cor = classificar_risco(dimensao_melhor["Média"])
        st.markdown(f"<div style='text-align: center; color: {cor};'><b>{risco}</b></div>", unsafe_allow_html=True)
    
    # Criar gráfico de barras para visualização dos riscos
    fig = criar_grafico_barras(df_resultados)
    st.plotly_chart(fig, use_container_width=True)
    
    # Mostrar informações detalhadas sobre cada dimensão
    st.subheader("Detalhes por Dimensão")
    
    for idx, row in df_resultados.iterrows():
        dimensao = row["Dimensão"]
        media = row["Média"]
        risco = row["Risco"]
        _, cor = classificar_risco(media)
        descricao = row["Descrição"]
        
        with st.expander(f"{dimensao} - {risco}"):
            st.markdown(f"**Média**: {media:.2f}")
            st.markdown(f"**Nível de Risco**: <span style='color:{cor};'>{risco}</span>", unsafe_allow_html=True)
            st.markdown(f"**Descrição**: {descricao}")
            st.markdown(f"**Questões HSE-IT relacionadas**: {row['Questões']}")
            
            # Adicionar sugestões de melhoria
            if "Muito Alto" in risco or "Alto" in risco:
                st.markdown("**Sugestões de Melhoria Prioritárias:**")
                plano_temp = gerar_sugestoes_acoes(pd.DataFrame([row]))
                for _, sugestao in plano_temp.iterrows():
                    st.markdown(f"- {sugestao['Sugestão de Ação']}")
    
    return

# Função para criar gráfico de barras melhorado usando Plotly
def criar_grafico_barras(df_resultados):
    # Ordenar resultados do menor para o maior (pior para melhor)
    df_sorted = df_resultados.sort_values(by="Média")
    
    # Preparar dados para o gráfico
    cores = []
    hover_texts = []
    
    for media in df_sorted["Média"]:
        _, cor = classificar_risco(media)
        cores.append(cor)
    
    for _, row in df_sorted.iterrows():
        hover_texts.append(f"Dimensão: {row['Dimensão']}<br>" +
                          f"Média: {row['Média']:.2f}<br>" +
                          f"Classificação: {row['Risco']}<br>" +
                          f"Descrição: {row['Descrição']}")
    
    # Criar gráfico
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        x=df_sorted["Média"],
        y=df_sorted["Dimensão"],
        orientation='h',
        marker_color=cores,
        text=[f"{v:.2f}" for v in df_sorted["Média"]],
        textposition='outside',
        hovertext=hover_texts,
        hoverinfo='text'
    ))
    
    # Adicionar linhas verticais para níveis de risco
    fig.add_shape(
        type="line",
        x0=1, y0=-0.5, x1=1, y1=len(df_resultados)-0.5,
        line=dict(color="Red", width=2, dash="dash")
    )
    
    fig.add_shape(
        type="line",
        x0=2, y0=-0.5, x1=2, y1=len(df_resultados)-0.5,
        line=dict(color="Orange", width=2, dash="dash")
    )
    
    fig.add_shape(
        type="line",
        x0=3, y0=-0.5, x1=3, y1=len(df_resultados)-0.5,
        line=dict(color="Yellow", width=2, dash="dash")
    )
    
    fig.add_shape(
        type="line",
        x0=4, y0=-0.5, x1=4, y1=len(df_resultados)-0.5,
        line=dict(color="Green", width=2, dash="dash")
    )
    
    # Adicionar anotações para níveis de risco
    fig.add_annotation(x=0.5, y=len(df_resultados), text="Risco Muito Alto", 
                      showarrow=False, font=dict(color="red"))
    fig.add_annotation(x=1.5, y=len(df_resultados), text="Risco Alto", 
                      showarrow=False, font=dict(color="orange"))
    fig.add_annotation(x=2.5, y=len(df_resultados), text="Risco Moderado", 
                      showarrow=False, font=dict(color="black"))
    fig.add_annotation(x=3.5, y=len(df_resultados), text="Risco Baixo", 
                      showarrow=False, font=dict(color="green"))
    fig.add_annotation(x=4.5, y=len(df_resultados), text="Risco Muito Baixo", 
                      showarrow=False, font=dict(color="purple"))
    
    # Configurar layout
    fig.update_layout(
        title=f"Classificação de Riscos Psicossociais HSE-IT",
        xaxis_title="Média (Escala 1-5)",
        yaxis_title="Dimensão",
        xaxis=dict(range=[0, 5]),
        height=500,
        margin=dict(l=20, r=20, t=50, b=80)
    )
    
    return fig

# Função para o plano de ação editável
def plano_acao_editavel(df_plano_acao):
    st.header("Plano de Ação Personalizado")
    st.write("Personalize o plano de ação sugerido ou adicione suas próprias ações para cada dimensão HSE-IT.")
    
    # Inicializar plano de ação no state se não existir
    if "plano_acao_personalizado" not in st.session_state:
        st.session_state.plano_acao_personalizado = df_plano_acao.copy()
    
    # Criar tabs para cada dimensão
    dimensoes_unicas = df_plano_acao["Dimensão"].unique()
    dimensao_tabs = st.tabs(dimensoes_unicas)
    
    # Para cada dimensão, criar um editor de ações
    for i, dimensao in enumerate(dimensoes_unicas):
        with dimensao_tabs[i]:
            df_dimensao = st.session_state.plano_acao_personalizado[
                st.session_state.plano_acao_personalizado["Dimensão"] == dimensao
            ].copy()
            
            # Mostrar informações da dimensão
            nivel_risco = df_dimensao["Nível de Risco"].iloc[0]
            media = df_dimensao["Média"].iloc[0]
            
            # Definir cor com base no nível de risco
            cor = {
                "Risco Muito Alto": "red",
                "Risco Alto": "orange",
                "Risco Moderado": "yellow",
                "Risco Baixo": "green",
                "Risco Muito Baixo": "purple"
            }.get(nivel_risco, "gray")
            
            # Exibir descrição da dimensão
            st.markdown(f"**Descrição da Dimensão:** {DESCRICOES_DIMENSOES[dimensao]}")
            st.markdown(f"**Média:** {media} - **Nível de Risco:** :{cor}[{nivel_risco}]")
            
            # Adicionar nova ação
            st.subheader("Adicionar Nova Ação:")
            col1, col2 = st.columns([3, 1])
            
            with col1:
                nova_acao = st.text_area("Descrição da ação", key=f"nova_acao_{dimensao}")
            
            with col2:
                st.write("&nbsp;")  # Espaçamento
                adicionar = st.button("Adicionar", key=f"add_{dimensao}")
                
                if adicionar and nova_acao.strip():
                    # Criar nova linha para o DataFrame
                    nova_linha = {
                        "Dimensão": dimensao,
                        "Nível de Risco": nivel_risco,
                        "Média": media,
                        "Sugestão de Ação": nova_acao,
                        "Responsável": "",
                        "Prazo": "",
                        "Status": "Não iniciada",
                        "Personalizada": True  # Marcar como ação personalizada
                    }
                    
                    # Adicionar ao DataFrame
                    st.session_state.plano_acao_personalizado = pd.concat([
                        st.session_state.plano_acao_personalizado, 
                        pd.DataFrame([nova_linha])
                    ], ignore_index=True)
                    
                    # Limpar campo de texto
                    st.session_state[f"nova_acao_{dimensao}"] = ""
                    st.experimental_rerun()
            
            # Mostrar ações existentes para editar
            st.subheader("Ações Sugeridas:")
            for j, (index, row) in enumerate(df_dimensao.iterrows()):
                with st.expander(f"Ação {j+1}: {row['Sugestão de Ação'][:50]}...", expanded=False):
                    col1, col2 = st.columns([3, 1])
                    
                    with col1:
                        # Editor de texto para a ação
                        acao_editada = st.text_area(
                            "Descrição da ação", 
                            row["Sugestão de Ação"], 
                            key=f"acao_{dimensao}_{j}"
                        )
                        if acao_editada != row["Sugestão de Ação"]:
                            st.session_state.plano_acao_personalizado.at[index, "Sugestão de Ação"] = acao_editada
                    
                    with col2:
                        # Campos para responsável, prazo e status
                        responsavel = st.text_input(
                            "Responsável", 
                            row.get("Responsável", ""), 
                            key=f"resp_{dimensao}_{j}"
                        )
                        if responsavel != row.get("Responsável", ""):
                            st.session_state.plano_acao_personalizado.at[index, "Responsável"] = responsavel
                        
                        # Campo de data para prazo
                        try:
                            data_padrao = None
                            if row.get("Prazo") and row.get("Prazo") != "":
                                try:
                                    data_padrao = datetime.strptime(row.get("Prazo"), "%d/%m/%Y")
                                except:
                                    data_padrao = None
                            
                            prazo = st.date_input(
                                "Prazo", 
                                value=data_padrao,
                                key=f"prazo_{dimensao}_{j}"
                            )
                            if prazo:
                                st.session_state.plano_acao_personalizado.at[index, "Prazo"] = prazo.strftime("%d/%m/%Y")
                        except Exception as e:
                            st.warning(f"Erro ao processar data: {e}")
                        
                        # Seletor de status
                        status = st.selectbox(
                            "Status",
                            options=["Não iniciada", "Em andamento", "Concluída", "Cancelada"],
                            index=["Não iniciada", "Em andamento", "Concluída", "Cancelada"].index(row.get("Status", "Não iniciada")),
                            key=f"status_{dimensao}_{j}"
                        )
                        if status != row.get("Status", "Não iniciada"):
                            st.session_state.plano_acao_personalizado.at[index, "Status"] = status
                        
                        # Botão para remover (apenas para ações personalizadas)
                        if row.get("Personalizada", False):
                            if st.button("🗑️ Remover", key=f"del_{dimensao}_{j}"):
                                st.session_state.plano_acao_personalizado = st.session_state.plano_acao_personalizado.drop(index)
                                st.experimental_rerun()
    
    # Botão para exportar plano personalizado
    if st.button("Exportar Plano de Ação Personalizado"):
        # Gerar Excel com o plano personalizado
        output = gerar_excel_plano_personalizado(st.session_state.plano_acao_personalizado)
        if output:
            st.success("Plano de Ação gerado com sucesso!")
            st.download_button(
                label="Baixar Plano de Ação Personalizado",
                data=output,
                file_name="plano_acao_personalizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# Função para gerar excel do plano personalizado
def gerar_excel_plano_personalizado(df_plano_personalizado):
    output = io.BytesIO()
    
    try:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Aba principal com o plano de ação
            df_plano_copia = df_plano_personalizado.copy()
            
            # Remover coluna Personalizada antes de exportar
            if "Personalizada" in df_plano_copia.columns:
                df_plano_copia = df_plano_copia.drop(columns=["Personalizada"])
                
            df_plano_copia.to_excel(writer, sheet_name='Plano de Ação', index=False)
            
            # Formatar a aba
            worksheet = writer.sheets['Plano de Ação']
            
            # Definir formatos
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })

            risco_format = {
                'Risco Muito Alto': workbook.add_format({'bg_color': '#FF6B6B', 'font_color': 'white'}),
                'Risco Alto': workbook.add_format({'bg_color': '#FFA500'}),
                'Risco Moderado': workbook.add_format({'bg_color': '#FFFF00'}),
                'Risco Baixo': workbook.add_format({'bg_color': '#90EE90'}),
                'Risco Muito Baixo': workbook.add_format({'bg_color': '#BB8FCE'})
            }
            
            # Configurar largura das colunas
            worksheet.set_column('A:A', 25)  # Dimensão
            worksheet.set_column('B:B', 15)  # Nível de Risco
            worksheet.set_column('C:C', 10)  # Média
            worksheet.set_column('D:D', 50)  # Sugestão de Ação
            worksheet.set_column('E:E', 15)  # Responsável
            worksheet.set_column('F:F', 15)  # Prazo
            worksheet.set_column('G:G', 15)  # Status
            
            # Adicionar cabeçalhos formatados
            for col_num, value in enumerate(df_plano_copia.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Aplicar formatação condicional baseada no nível de risco
            for row_num, (_, row) in enumerate(df_plano_copia.iterrows(), 1):
                nivel_risco = row["Nível de Risco"]
                if nivel_risco in risco_format:
                    worksheet.write(row_num, 1, nivel_risco, risco_format[nivel_risco])
            
            # Adicionar validação de dados para a coluna Status
            status_options = ['Não iniciada', 'Em andamento', 'Concluída', 'Cancelada']
            worksheet.data_validation('G2:G1000', {'validate': 'list',
                                                'source': status_options,
                                                'input_title': 'Selecione o status:',
                                                'input_message': 'Escolha um status da lista'})
            
            # Adicionar filtros
            worksheet.autofilter(0, 0, len(df_plano_copia), len(df_plano_copia.columns) - 1)
            
            # Congelar painel para manter cabeçalhos visíveis durante rolagem
            worksheet.freeze_panes(1, 0)
            
            # Adicionar aba com explicação das dimensões HSE-IT
            worksheet_dimensoes = workbook.add_worksheet('Informações HSE-IT')
            
            # Título
            title_format = workbook.add_format({'bold': True, 'size': 14})
            worksheet_dimensoes.write('A1', 'Questionário HSE-IT: Dimensões e Significados', title_format)
            worksheet_dimensoes.set_column('A:A', 20)
            worksheet_dimensoes.set_column('B:B', 60)
            
            # Cabeçalhos
            header_format = workbook.add_format({'bold': True, 'fg_color': '#D7E4BC', 'border': 1})
            worksheet_dimensoes.write('A3', 'Dimensão', header_format)
            worksheet_dimensoes.write('B3', 'Descrição', header_format)
            
            # Conteúdo
            row = 4
            for dimensao, descricao in DESCRICOES_DIMENSOES.items():
                worksheet_dimensoes.write(row, 0, dimensao)
                worksheet_dimensoes.write(row, 1, descricao)
                row += 1
                
            # Adicionar informações sobre a interpretação dos riscos
            row += 2
            worksheet_dimensoes.write(row, 0, 'Interpretação dos Riscos:', workbook.add_format({'bold': True}))
            row += 1
            worksheet_dimensoes.write(row, 0, 'Risco Muito Alto:')
            worksheet_dimensoes.write(row, 1, 'Média ≤ 1 (Prioridade máxima para intervenção)')
            row += 1
            worksheet_dimensoes.write(row, 0, 'Risco Alto:')
            worksheet_dimensoes.write(row, 1, '1 < Média ≤ 2 (Alta prioridade para intervenção)')
            row += 1
            worksheet_dimensoes.write(row, 0, 'Risco Moderado:')
            worksheet_dimensoes.write(row, 1, '2 < Média ≤ 3 (Prioridade média para intervenção)')
            row += 1
            worksheet_dimensoes.write(row, 0, 'Risco Baixo:')
            worksheet_dimensoes.write(row, 1, '3 < Média ≤ 4 (Baixa prioridade, mas ainda há espaço para melhorias)')
            row += 1
            worksheet_dimensoes.write(row, 0, 'Risco Muito Baixo:')
            worksheet_dimensoes.write(row, 1, 'Média > 4 (Manter as boas práticas atuais)')
        
        output.seek(0)
        return output
    
    except Exception as e:
        st.error(f"Erro ao gerar o arquivo Excel: {str(e)}")
        return None

# Função para gerar template para coleta de dados HSE-IT
def gerar_template_excel():
    output = io.BytesIO()
    
    try:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Criar DataFrame com a estrutura esperada
            colunas = [
                "Setor", "Cargo", "Tempo_Empresa", "Genero", "Faixa_Etaria", "Escolaridade", "Regime_Trabalho",
            ]
            
            # As questões HSE-IT
            questoes_hse = [
                "1. Sei claramente o que é esperado de mim no trabalho",
                "2. Posso decidir quando fazer uma pausa",
                "3. Grupos de trabalho diferentes pedem-me coisas difíceis de conjugar",
                "4. Sei do que necessito para fazer o meu trabalho",
                "5. Sou sujeito a assédio pessoal sob a forma de palavras ou comportamentos incorretos",
                "6. Tenho prazos impossíveis de cumprir",
                "7. Se o trabalho se torna difícil, os colegas ajudam-me",
                "8. Recebo feedback de apoio sobre o trabalho que faço",
                "9. Tenho que trabalhar muito intensivamente",
                "10. Tenho capacidade de decisão sobre a minha rapidez de trabalho",
                "11. Sei claramente os meus deveres e responsabilidades",
                "12. Tenho que negligenciar tarefas porque tenho uma carga elevada para cumprir",
                "13. Sei claramente as metas e objetivos do meu departamento",
                "14. Há fricção ou animosidade entre os colegas",
                "15. Posso decidir como fazer o meu trabalho",
                "16. Não consigo fazer pausas suficientes",
                "17. Compreendo como o meu trabalho se integra no objetivo geral da organização",
                "18. Sou pressionado a trabalhar durante horários longos",
                "19. Tenho poder de escolha para decidir o que faço no trabalho",
                "20. Tenho que trabalhar muito depressa",
                "21. Sou sujeito a intimidação/perseguição no trabalho",
                "22. Tenho pressões de tempo irrealistas",
                "23. Posso estar seguro de que o meu chefe imediato me ajuda num problema de trabalho",
                "24. Tenho ajuda e apoio necessários dos colegas",
                "25. Tenho algum poder de decisão sobre a minha forma de trabalho",
                "26. Tenho oportunidades suficientes para questionar os chefes sobre mudanças no trabalho",
                "27. Sou respeitado como mereço pelos colegas de trabalho",
                "28. O pessoal é sempre consultado sobre mudança no trabalho",
                "29. Posso falar com o meu chefe imediato sobre algo no trabalho que me transtornou ou irritou",
                "30. O meu horário pode ser flexível",
                "31. Os meus colegas estão dispostos a ouvir os meus problemas relacionados com o trabalho",
                "32. Quando são efetuadas mudanças no trabalho, sei claramente como resultarão na prática",
                "33. Recebo apoio durante trabalho que pode ser emocionalmente exigente",
                "34. Os relacionamentos no trabalho estão sob pressão",
                "35. O meu chefe imediato encoraja-me no trabalho"
            ]
            
            # Adicionar questões ao template
            for q in questoes_hse:
                colunas.append(q)
            
            df_template = pd.DataFrame(columns=colunas)
            
            # Adicionar linha de exemplo
            valores_exemplo = ["TI", "Analista", "1-3 anos", "Masculino", "25-35", "Superior", "CLT"]
            # Valores para as questões (gerando valores aleatórios entre 1 e 5 para exemplo)
            import random
            random.seed(42)  # Para consistência
            valores_exemplo.extend([random.randint(1, 5) for _ in range(len(questoes_hse))])
            
            df_template.loc[0] = valores_exemplo
            
            # Exportar para Excel
            df_template.to_excel(writer, sheet_name="Questionário HSE-IT", index=False)
            
            # Formatar a planilha
            worksheet = writer.sheets["Questionário HSE-IT"]
            
            # Configurar largura das colunas
            for i, col in enumerate(df_template.columns):
                # Calcular largura baseada no comprimento da coluna
                col_width = max(len(str(col)), 15)
                worksheet.set_column(i, i, col_width)
            
            # Criar folha de instruções
            worksheet_instrucoes = workbook.add_worksheet("Instruções")
            
            # Formatar cabeçalho das instruções
            header_format = workbook.add_format({'bold': True, 'size': 14})
            worksheet_instrucoes.write('A1', 'Instruções para Aplicação do Questionário HSE-IT', header_format)
            worksheet_instrucoes.set_column('A:A', 100)
            
            # Adicionar instruções
            instrucoes = [
                "",
                "Sobre o HSE-IT:",
                "O HSE-IT (Health and Safety Executive - Indicator Tool) é um questionário validado para avaliação de fatores psicossociais no trabalho, desenvolvido pela instituição britânica de saúde e segurança.",
                "",
                "Instruções de uso:",
                "",
                "1. As primeiras 7 colunas contêm informações demográficas que podem ser adaptadas conforme necessário:",
                "   - Setor: Área/departamento do colaborador",
                "   - Cargo: Função desempenhada",
                "   - Tempo_Empresa: Tempo de permanência na organização",
                "   - Gênero: Identificação de gênero",
                "   - Faixa_Etária: Grupo etário",
                "   - Escolaridade: Nível de formação",
                "   - Regime_Trabalho: Tipo de contrato/regime",
                "",
                "2. As demais colunas (35) contêm as perguntas do questionário HSE-IT. Não altere o conteúdo destas perguntas para preservar a validade do instrumento.",
                "",
                "3. As respostas devem ser preenchidas com valores de 1 a 5, seguindo esta escala:",
                "   - 1: Nunca / Discordo totalmente",
                "   - 2: Raramente / Discordo parcialmente",
                "   - 3: Às vezes / Nem concordo nem discordo",
                "   - 4: Frequentemente / Concordo parcialmente",
                "   - 5: Sempre / Concordo totalmente",
                "",
                "4. IMPORTANTE: Para algumas questões específicas, a escala é invertida devido à formulação negativa. Estas questões são: 3, 5, 6, 9, 12, 14, 16, 18, 20, 21, 22 e 34.",
                "   Exemplo: Para a questão 'Tenho prazos impossíveis de cumprir', responder '1' (Nunca) é positivo.",
                "",
                "5. Mantenha a estrutura exata deste template, incluindo numeração e texto completo das questões.",
                "",
                "6. Aplique o questionário garantindo o anonimato das respostas para obter dados mais honestos e confiáveis.",
                "",
                "7. Recomenda-se aplicar o questionário a pelo menos 60% dos colaboradores para obter um panorama representativo da organização.",
                "",
                "8. Os resultados são agrupados em 7 dimensões: Demanda, Controle, Apoio da Chefia, Apoio dos Colegas, Relacionamentos, Função e Mudança.",
                "",
                "Após a coleta, utilize a plataforma HSE-IT Analytics para processar os dados e gerar relatórios de riscos psicossociais."
            ]
            
            # Escrever instruções
            for i, texto in enumerate(instrucoes, 2):
                worksheet_instrucoes.write(i, 0, texto)
                
            # Adicionar aba com detalhes sobre as dimensões
            worksheet_dimensoes = workbook.add_worksheet("Dimensões HSE-IT")
            
            # Título
            worksheet_dimensoes.write('A1', 'Dimensões do Questionário HSE-IT', header_format)
            worksheet_dimensoes.set_column('A:A', 15)
            worksheet_dimensoes.set_column('B:B', 40)
            worksheet_dimensoes.set_column('C:C', 30)
            
            # Cabeçalhos
            dim_header = workbook.add_format({'bold': True, 'fg_color': '#D7E4BC', 'border': 1})
            worksheet_dimensoes.write('A3', 'Dimensão', dim_header)
            worksheet_dimensoes.write('B3', 'Descrição', dim_header)
            worksheet_dimensoes.write('C3', 'Questões', dim_header)
            
            # Conteúdo
            row = 4
            for dimensao, questoes in DIMENSOES_HSE.items():
                worksheet_dimensoes.write(row, 0, dimensao)
                worksheet_dimensoes.write(row, 1, DESCRICOES_DIMENSOES[dimensao])
                worksheet_dimensoes.write(row, 2, str(questoes))
                row += 1
            
            # Adicionar informações sobre questões invertidas
            row += 2
            worksheet_dimensoes.write(row, 0, 'IMPORTANTE:', workbook.add_format({'bold': True}))
            row += 1
            worksheet_dimensoes.write(row, 0, 'Questões com escala invertida:')
            worksheet_dimensoes.write(row, 1, str(QUESTOES_INVERTIDAS))
            row += 1
            worksheet_dimensoes.write(row, 0, 'Nestas questões, uma pontuação mais baixa é positiva devido à formulação negativa da pergunta.')
        
        output.seek(0)
        return output
    
    except Exception as e:
        st.error(f"Erro ao gerar o template Excel: {str(e)}")
        return None

# Aplicativo principal
def main():
    # Verificar autenticação antes de mostrar o conteúdo
    if check_password():
        # Criar sistema de abas para melhor organização
        tabs = st.tabs(["Upload de Dados", "Resultados", "Plano de Ação", "Relatórios", "Informações HSE-IT"])

        # Primeira aba - Upload e configuração
        with tabs[0]:
            st.title("Avaliação de Fatores Psicossociais - HSE-IT")
            st.write("Faça upload do arquivo Excel contendo os resultados do questionário HSE-IT.")
            
            # Adiciona uma explicação sobre o formato esperado do arquivo
            with st.expander("Informações sobre o formato do arquivo"):
                st.write("""
                O arquivo deve conter:
                - Colunas de filtro (Setor, Cargo, etc.) nas primeiras 7 colunas
                - Colunas com as perguntas numeradas do HSE-IT (começando com números seguidos de ponto)
                - As respostas devem ser valores numéricos de 1 a 5
                
                Você pode baixar um template na aba "Informações HSE-IT".
                """)
            
            uploaded_file = st.file_uploader("Escolha um arquivo Excel ou CSV", type=["xlsx", "csv"])
            
            # Barra lateral para filtros e configurações
            with st.sidebar:
                st.header("Filtros e Configurações")
                
                # Inicializar variáveis para evitar erros
                df_resultados = None
                df_plano_acao = None
                filtro_opcao = "Empresa Toda"
                filtro_valor = "Geral"
                
                if uploaded_file is not None:
                    try:
                        # Carregar dados
                        df, df_perguntas, colunas_filtro, colunas_perguntas = carregar_dados(uploaded_file)
                        
                        # Criar opções de filtro
                        opcoes_filtro = ["Empresa Toda"] + [col for col in colunas_filtro if col != "Carimbo de data/hora"]
                        
                        filtro_opcao = st.selectbox("Filtrar por", opcoes_filtro)
                        
                        filtro_valor = "Geral"
                        df_filtrado = df
                        
                        if filtro_opcao != "Empresa Toda":
                            valores_unicos = df[filtro_opcao].dropna().unique()
                            if len(valores_unicos) > 0:
                                filtro_valor = st.selectbox(f"Escolha um {filtro_opcao}", valores_unicos)
                                df_filtrado = df[df[filtro_opcao] == filtro_valor]
                                
                                # Filtrar as perguntas de acordo com o filtro
                                indices_filtrados = df.index[df[filtro_opcao] == filtro_valor].tolist()
                                df_perguntas_filtradas = df_perguntas.loc[indices_filtrados]
                            else:
                                st.warning(f"Não há valores únicos para o filtro '{filtro_opcao}'.")
                        else:
                            df_perguntas_filtradas = df_perguntas
                        
                        # Calcular resultados
                        resultados = calcular_resultados_dimensoes(df_filtrado, df_perguntas_filtradas, colunas_perguntas)
                        df_resultados = pd.DataFrame(resultados)
                        
                        # Armazenar no session_state para acesso em outras abas
                        st.session_state.df_resultados = df_resultados
                        st.session_state.df = df
                        st.session_state.df_perguntas = df_perguntas
                        st.session_state.colunas_filtro = colunas_filtro
                        st.session_state.colunas_perguntas = colunas_perguntas
                        st.session_state.filtro_opcao = filtro_opcao
                        st.session_state.filtro_valor = filtro_valor
                        
                        # Gerar plano de ação
                        df_plano_acao = gerar_sugestoes_acoes(df_resultados)
                        st.session_state.df_plano_acao = df_plano_acao
                        
                        st.success(f"Dados carregados com sucesso! Foram encontrados {len(df)} registros e {len(colunas_perguntas)} perguntas.")
                        
                    except Exception as e:
                        st.error(f"Ocorreu um erro inesperado: {str(e)}")
                        st.write("Detalhes do erro para debug:")
                        st.exception(e)
                
                # Informações sobre a metodologia HSE-IT
                with st.expander("Sobre a Metodologia HSE-IT"):
                    st.write("""
                    O HSE-IT é um questionário de 35 perguntas dividido em 7 dimensões:
                    
                    1. **Demanda**: Carga de trabalho, ritmo e ambiente
                    2. **Controle**: Autonomia no trabalho
                    3. **Apoio da Chefia**: Suporte dos superiores
                    4. **Apoio dos Colegas**: Suporte dos pares
                    5. **Relacionamentos**: Promoção de trabalho positivo
                    6. **Função**: Clareza de papéis e responsabilidades
                    7. **Mudança**: Gestão de transformações organizacionais
                    """)
            
            if uploaded_file is not None and "df_resultados" in st.session_state and st.session_state.df_resultados is not None:
                st.success("Dados processados com sucesso. Navegue para a aba 'Resultados' para visualizar a análise.")
                
                # Mostrar resumo dos dados
                st.subheader("Resumo dos Dados Carregados")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("Total de Registros", len(st.session_state.df))
                with col2:
                    st.metric("Dimensões Avaliadas", "7")
                with col3:
                    st.metric("Questões do HSE-IT", "35")
                
                # Preview dos dados
                with st.expander("Visualizar primeiras linhas dos dados"):
                    st.dataframe(st.session_state.df.head())
        
        # Segunda aba - Visualização de resultados
        with tabs[1]:
            if "df_resultados" in st.session_state and st.session_state.df_resultados is not None:
                df_resultados = st.session_state.df_resultados
                filtro_opcao = st.session_state.filtro_opcao
                filtro_valor = st.session_state.filtro_valor
                
                # Mostrar dashboard melhorado
                criar_dashboard(df_resultados, filtro_opcao, filtro_valor)
                
                st.write("### Resultados Detalhados da Avaliação")
                
                # Remover a coluna de questões para visualização mais limpa
                df_visualizacao = df_resultados.copy()
                if 'Questões' in df_visualizacao.columns:
                    df_visualizacao = df_visualizacao.drop(columns=['Questões'])
                    
                st.dataframe(df_visualizacao)
                
                # Adicionar um botão para exportar os resultados detalhados
                csv = df_visualizacao.to_csv(index=False)
                st.download_button(
                    label="Baixar Resultados CSV",
                    data=csv,
                    file_name=f"resultados_hse_it_{filtro_opcao}_{filtro_valor}.csv",
                    mime="text/csv",
                )
            else:
                st.info("Primeiro faça o upload dos dados na aba 'Upload de Dados'.")
        
        # Terceira aba - Plano de ação
        with tabs[2]:
            if "df_plano_acao" in st.session_state and st.session_state.df_plano_acao is not None:
                # Mostrar o plano de ação editável
                plano_acao_editavel(st.session_state.df_plano_acao)
            else:
                st.info("Primeiro faça o upload dos dados na aba 'Upload de Dados'.")
        
        # Quarta aba - Relatórios
        with tabs[3]:
            if "df_resultados" in st.session_state and st.session_state.df_resultados is not None:
                st.header("Download de Relatórios")
                
                # Recuperar dados do session_state
                df = st.session_state.df
                df_perguntas = st.session_state.df_perguntas
                colunas_filtro = st.session_state.colunas_filtro
                colunas_perguntas = st.session_state.colunas_perguntas
                df_resultados = st.session_state.df_resultados
                df_plano_acao = st.session_state.df_plano_acao
                filtro_opcao = st.session_state.filtro_opcao
                filtro_valor = st.session_state.filtro_valor
                
                st.write("""
                Escolha abaixo o tipo de relatório que deseja gerar. Todos os relatórios são baseados nos dados carregados e# Escolha abaixo o tipo de relatório que deseja gerar. Todos os relatórios são baseados nos dados carregados e 
                # filtros selecionados.
                """)
                
                col1, col2 = st.columns(2)
                
                # Relatório Excel completo
                with col1:
                    st.subheader("Relatório Excel Completo")
                    st.write("Relatório detalhado com múltiplas abas, incluindo análises por filtros, gráficos e plano de ação.")
                    
                    # Usar a função para gerar Excel com múltiplas abas e plano de ação
                    excel_data = gerar_excel_completo(df, df_perguntas, colunas_filtro, colunas_perguntas)
                    if excel_data:
                        st.download_button(
                            label="Baixar Relatório Excel Completo",
                            data=excel_data,
                            file_name=f"relatorio_completo_hse_it_{filtro_opcao}_{filtro_valor}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            help="Baixe um relatório Excel com múltiplas abas contendo análises por empresa, setor, cargo, etc. e o plano de ação sugerido."
                        )
                
                # Relatórios PDF
                with col2:
                    st.subheader("Relatórios PDF")
                    st.write("Relatórios simplificados para compartilhamento e apresentação.")
                    
                    # Gerar PDF da análise
                    pdf_data = gerar_pdf(df_resultados)
                    if pdf_data:
                        st.download_button(
                            label="Baixar Relatório PDF de Resultados",
                            data=pdf_data,
                            file_name=f"resultados_hse_it_{filtro_opcao}_{filtro_valor}.pdf",
                            mime="application/pdf",
                            help="Relatório resumido com os resultados da avaliação HSE-IT."
                        )
                    
                    # Adicionar opção para baixar o PDF do Plano de Ação
                    pdf_plano_acao = gerar_pdf_plano_acao(df_plano_acao)
                    if pdf_plano_acao:
                        st.download_button(
                            label="Baixar Plano de Ação PDF",
                            data=pdf_plano_acao,
                            file_name=f"plano_acao_hse_it_{filtro_opcao}_{filtro_valor}.pdf",
                            mime="application/pdf",
                            help="Baixe um PDF específico com o plano de ação sugerido, incluindo campos para preenchimento de responsáveis e prazos."
                        )
                
                # Adicionar informações sobre como usar os relatórios
                with st.expander("Como usar os relatórios"):
                    st.write("""
                    ### Relatório Excel Completo
                    Contém múltiplas abas com análises detalhadas:
                    - **Empresa Toda**: Visão geral dos fatores psicossociais nas 7 dimensões do HSE-IT
                    - **Plano de Ação**: Sugestões de ações para cada dimensão com maior risco
                    - **Detalhes das Dimensões**: Explicação sobre cada dimensão e suas questões
                    - **Por Setor/Cargo/etc.**: Análises segmentadas por filtros demográficos
                    - **Gráfico de Riscos**: Visualização gráfica das dimensões
                    
                    ### Relatório PDF de Resultados
                    Versão simplificada com os principais resultados, ideal para compartilhamento com a liderança e stakeholders. Contém:
                    - Tabela de resultados por dimensão
                    - Descrição das dimensões do HSE-IT
                    - Interpretação dos níveis de risco
                    
                    ### Plano de Ação PDF
                    Documento específico com as ações sugeridas, contendo espaços para preenchimento manual de responsáveis e prazos. Ideal para:
                    - Discussão em reuniões de planejamento
                    - Acompanhamento de ações
                    - Documentação do programa de saúde psicossocial
                    """)
            else:
                st.info("Primeiro faça o upload dos dados na aba 'Upload de Dados'.")
        
        # Quinta aba - Informações HSE-IT
        with tabs[4]:
            st.header("Informações sobre o Questionário HSE-IT")
            
            st.write("""
            O HSE-IT (Health and Safety Executive - Indicator Tool) é um questionário validado para avaliação de 
            fatores psicossociais no ambiente de trabalho, desenvolvido pela instituição britânica de saúde e segurança ocupacional.
            
            O questionário consiste em 35 perguntas que avaliam 7 dimensões de fatores psicossociais, permitindo identificar
            áreas de risco que precisam de intervenção.
            """)
            
            # Mostrar dimensões do HSE-IT
            st.subheader("Dimensões Avaliadas")
            
            for dimensao, questoes in DIMENSOES_HSE.items():
                with st.expander(f"{dimensao} - Questões {questoes}"):
                    st.write(f"**Descrição**: {DESCRICOES_DIMENSOES[dimensao]}")
                    st.write("**Questões relacionadas**:")
                    for q in questoes:
                        # Buscar o texto das questões se disponível
                        texto_questao = f"Questão {q}"
                        for i in range(1, 36):
                            if i == q:
                                if q == 1:
                                    texto_questao = "Sei claramente o que é esperado de mim no trabalho"
                                elif q == 2:
                                    texto_questao = "Posso decidir quando fazer uma pausa"
                                elif q == 3:
                                    texto_questao = "Grupos de trabalho diferentes pedem-me coisas difíceis de conjugar"
                                # Continue com as demais questões se necessário
                                break
                        
                        if q in QUESTOES_INVERTIDAS:
                            st.write(f"- {q}: {texto_questao} *(escala invertida)*")
                        else:
                            st.write(f"- {q}: {texto_questao}")
            
            # Explicação sobre a interpretação dos resultados
            st.subheader("Interpretação dos Resultados")
            
            st.write("""
            Os resultados são apresentados em uma escala de 1 a 5, onde valores mais altos geralmente indicam
            melhores condições psicossociais (exceto em questões com escala invertida, onde a lógica é oposta).
            
            A classificação de risco é feita com base na média de cada dimensão:
            """)
            
            # Tabela de classificação de risco
            risco_data = {
                "Classificação": ["Risco Muito Alto 🔴", "Risco Alto 🟠", "Risco Moderado 🟡", "Risco Baixo 🟢", "Risco Muito Baixo 🟣"],
                "Pontuação Média": ["≤ 1", "> 1 e ≤ 2", "> 2 e ≤ 3", "> 3 e ≤ 4", "> 4"],
                "Interpretação": [
                    "Situação crítica, requer intervenção imediata", 
                    "Condição preocupante, intervenção necessária em curto prazo", 
                    "Condição de alerta, melhorias necessárias", 
                    "Condição favorável, com oportunidades de melhoria", 
                    "Condição excelente, manter as boas práticas"
                ]
            }
            df_risco = pd.DataFrame(risco_data)
            st.table(df_risco)
            
            # Notas sobre questões invertidas
            st.subheader("Notas sobre Questões Invertidas")
            
            st.write(f"""
            Algumas questões do HSE-IT possuem escala invertida devido à sua formulação negativa. São elas: {QUESTOES_INVERTIDAS}
            
            Nestas questões, uma resposta de valor mais baixo é considerada positiva. Por exemplo, para a questão 
            "Tenho prazos impossíveis de cumprir", responder "1 - Nunca" representa uma boa condição de trabalho.
            
            Este aplicativo já realiza automaticamente a inversão destas questões durante o processamento dos dados.
            """)
            
            # Template para coleta de dados
            st.subheader("Template para Coleta de Dados")
            
            st.write("""
            Você pode baixar um template Excel para aplicação do questionário HSE-IT. Este template contém:
            - As 35 questões originais do HSE-IT
            - Campos para informações demográficas (Setor, Cargo, etc.)
            - Instruções detalhadas para aplicação
            - Informações sobre as dimensões e interpretação
            """)
            
            # Gerar e oferecer o template para download
            template_excel = gerar_template_excel()
            if template_excel:
                st.download_button(
                    label="Baixar Template HSE-IT",
                    data=template_excel,
                    file_name="template_questionario_hse_it.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Template Excel com as 35 questões do HSE-IT e instruções para aplicação."
                )
            
            # Referências
            st.subheader("Referências")
            
            st.write("""
            1. Health and Safety Executive (HSE). "Managing the causes of work-related stress: A step-by-step approach using the Management Standards." HSE Books, 2007.
            
            2. Cousins, R., et al. "Management Standards and work-related stress in the UK: Practical development." Work & Stress, 18(2), 113-136, 2004.
            
            3. Mackay, C. J., et al. "Management Standards and work-related stress in the UK: Policy background and science." Work & Stress, 18(2), 91-112, 2004.
            
            Para mais informações, visite o site do HSE: [www.hse.gov.uk/stress](https://www.hse.gov.uk/stress/)
            """)
    else:
        st.stop()  # Não mostrar nada abaixo deste ponto se a autenticação falhar

# Executar o aplicativo
if __name__ == "__main__":
    main()
