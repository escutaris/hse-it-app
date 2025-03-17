import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
from fpdf import FPDF

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

# Função para calcular resultados
@st.cache_data
def calcular_resultados(df, df_perguntas_filtradas, fatores, colunas_perguntas):
    resultados = []
    cores = []
    
    # Número total de respostas após filtragem
    num_total_respostas = len(df_perguntas_filtradas)
    
    for fator, perguntas in fatores.items():
        # Converter números de pergunta para índices de coluna
        indices_validos = []
        for p in perguntas:
            for col in colunas_perguntas:
                if str(col).strip().startswith(str(p)):
                    indices_validos.append(col)
                    break
        
        if indices_validos:
            # Calcular média diretamente de todas as respostas para este fator
            valores = df_perguntas_filtradas[indices_validos].values.flatten()
            valores = valores[~pd.isna(valores)] # Remove NaN
            
            if len(valores) > 0:
                media = valores.mean()
                risco, cor = classificar_risco(media)
                resultados.append({
                    "Fator Psicossocial": fator,
                    "Média": round(media, 2),
                    "Risco": risco,
                    "Número de Respostas": num_total_respostas
                })
                cores.append(cor)
    
    return resultados

# Função para gerar sugestões de ações com base no nível de risco
def gerar_sugestoes_acoes(df_resultados):
    # Dicionário com sugestões de ações para cada fator psicossocial por nível de risco
    sugestoes_por_fator = {
        "Gestão organizacional": {
            "Risco Muito Alto": [
                "Realizar auditoria completa dos processos de gestão organizacional",
                "Implementar plano de reestruturação da gestão com consultoria especializada",
                "Estabelecer canal direto e anônimo para feedback sobre gestão",
                "Promover treinamento intensivo para líderes em gestão de pessoas"
            ],
            "Risco Alto": [
                "Revisar processos decisórios com participação dos colaboradores",
                "Implementar reuniões regulares de feedback entre gestores e equipes",
                "Capacitar líderes em comunicação efetiva e gestão participativa",
                "Criar comitê de desenvolvimento organizacional"
            ],
            "Risco Moderado": [
                "Mapear processos organizacionais para identificar melhorias",
                "Oferecer workshops de liderança para gestores intermediários",
                "Estabelecer métricas de desempenho claras e transparentes"
            ],
            "Risco Baixo": [
                "Manter avaliação periódica de clima organizacional",
                "Promover encontros regulares para compartilhamento de boas práticas",
                "Revisar anualmente políticas e procedimentos organizacionais"
            ],
            "Risco Muito Baixo": [
                "Documentar as boas práticas de gestão para referência futura",
                "Avaliar periodicamente a manutenção das boas práticas de gestão"
            ]
        },
        "Violência e assédio moral/sexual no trabalho": {
            "Risco Muito Alto": [
                "Implementar imediatamente canal de denúncias anônimo e seguro",
                "Contratar consultoria especializada em prevenção de assédio",
                "Realizar treinamento obrigatório sobre assédio para todos os níveis hierárquicos",
                "Estabelecer comitê de ética com participação de representantes dos trabalhadores",
                "Revisar processos disciplinares para casos de assédio ou violência"
            ],
            "Risco Alto": [
                "Desenvolver política específica contra assédio moral e sexual",
                "Realizar treinamentos sobre respeito no ambiente de trabalho",
                "Estabelecer procedimentos claros para denúncias",
                "Promover campanhas de conscientização sobre assédio"
            ],
            "Risco Moderado": [
                "Realizar palestras sobre diversidade e respeito",
                "Revisar código de conduta da empresa",
                "Implementar programa de mediação de conflitos"
            ],
            "Risco Baixo": [
                "Manter comunicação periódica sobre canais de denúncia",
                "Incluir temas de respeito e diversidade em treinamentos regulares"
            ],
            "Risco Muito Baixo": [
                "Avaliar periodicamente a eficácia das políticas contra assédio",
                "Compartilhar práticas bem-sucedidas com outras organizações"
            ]
        },
        "Contexto da organização do trabalho": {
            "Risco Muito Alto": [
                "Realizar diagnóstico completo da organização do trabalho com ajuda externa",
                "Redesenhar processos de trabalho com participação dos colaboradores",
                "Implementar ferramentas de gestão visual para clarificar fluxos de trabalho",
                "Revisar distribuição de carga de trabalho entre equipes"
            ],
            "Risco Alto": [
                "Mapear e otimizar processos de trabalho críticos",
                "Estabelecer reuniões estruturadas para clarificar papéis e responsabilidades",
                "Implementar metodologias ágeis para melhor organização do trabalho",
                "Revisar políticas de metas e prazos"
            ],
            "Risco Moderado": [
                "Promover workshops para identificar gargalos nos processos",
                "Implementar sistemas de gestão de tarefas",
                "Estabelecer canais de comunicação eficientes entre departamentos"
            ],
            "Risco Baixo": [
                "Revisar periodicamente fluxos de trabalho",
                "Solicitar feedback regular sobre processos organizacionais"
            ],
            "Risco Muito Baixo": [
                "Documentar boas práticas de organização do trabalho",
                "Compartilhar métodos bem-sucedidos entre departamentos"
            ]
        },
        "Características das relações sociais no trabalho": {
            "Risco Muito Alto": [
                "Implementar programa de desenvolvimento de relações interpessoais",
                "Realizar intervenções de team building com facilitadores externos",
                "Estabelecer política de comunicação transparente e respeitosa",
                "Oferecer coaching para lideranças com foco em relacionamentos"
            ],
            "Risco Alto": [
                "Promover atividades de integração entre equipes",
                "Implementar práticas de feedback construtivo",
                "Capacitar gestores em gestão de conflitos",
                "Estabelecer espaços de diálogo regulares"
            ],
            "Risco Moderado": [
                "Realizar eventos de reconhecimento e celebração",
                "Promover trabalhos em equipe multidisciplinares",
                "Implementar rituais de integração para novos colaboradores"
            ],
            "Risco Baixo": [
                "Manter calendário de atividades sociais",
                "Incentivar comunicação interdepartamental"
            ],
            "Risco Muito Baixo": [
                "Documentar práticas bem-sucedidas de relacionamento interpessoal",
                "Promover o compartilhamento de conhecimento entre equipes"
            ]
        },
        "Conteúdo das tarefas do trabalho": {
            "Risco Muito Alto": [
                "Realizar análise ergonômica do trabalho (AET) completa",
                "Redesenhar funções com participação dos trabalhadores",
                "Implementar rotação de tarefas para reduzir monotonia",
                "Estabelecer metas realistas com autonomia parcial"
            ],
            "Risco Alto": [
                "Revisar descrições de cargos e funções",
                "Implementar reuniões para discutir melhorias nos processos",
                "Oferecer treinamentos para ampliar competências",
                "Permitir personalização parcial de métodos de trabalho"
            ],
            "Risco Moderado": [
                "Promover participação em projetos multidisciplinares",
                "Incentivar sugestões de melhorias nos processos",
                "Estabelecer momentos de feedback sobre conteúdo do trabalho"
            ],
            "Risco Baixo": [
                "Avaliar periodicamente satisfação com tarefas desempenhadas",
                "Oferecer oportunidades de desenvolvimento profissional"
            ],
            "Risco Muito Baixo": [
                "Compartilhar boas práticas de desenho de função",
                "Documentar casos de sucesso de enriquecimento de cargos"
            ]
        },
        "Discriminação no trabalho": {
            "Risco Muito Alto": [
                "Implementar política de diversidade e inclusão com metas mensuráveis",
                "Realizar treinamento obrigatório sobre vieses inconscientes",
                "Estabelecer comitê de diversidade com representatividade",
                "Contratar consultoria especializada em equidade no trabalho",
                "Revisar processos de RH para eliminar potenciais fontes de discriminação"
            ],
            "Risco Alto": [
                "Desenvolver programa de conscientização sobre diversidade",
                "Revisar políticas de contratação e promoção",
                "Implementar canal específico para denúncias de discriminação",
                "Realizar censo de diversidade para identificar gaps"
            ],
            "Risco Moderado": [
                "Promover palestras sobre diversidade e inclusão",
                "Celebrar datas comemorativas ligadas à diversidade",
                "Incentivar grupos de afinidade"
            ],
            "Risco Baixo": [
                "Manter comunicação regular sobre políticas de diversidade",
                "Monitorar indicadores de diversidade e inclusão"
            ],
            "Risco Muito Baixo": [
                "Documentar e compartilhar boas práticas de inclusão",
                "Participar de fóruns externos sobre diversidade"
            ]
        },
        "Condições do ambiente de trabalho": {
            "Risco Muito Alto": [
                "Realizar avaliação completa do ambiente físico de trabalho",
                "Implementar melhorias imediatas nas condições ambientais críticas",
                "Estabelecer cronograma de adequações ergonômicas",
                "Oferecer opções de trabalho remoto quando possível",
                "Revisar layout dos espaços de trabalho com participação dos trabalhadores"
            ],
            "Risco Alto": [
                "Realizar medições de ruído, iluminação e temperatura",
                "Implementar programa de ergonomia",
                "Revisar mobiliário e equipamentos de trabalho",
                "Estabelecer momentos de pausa programada"
            ],
            "Risco Moderado": [
                "Promover campanhas de organização e limpeza",
                "Implementar melhorias em áreas de descanso",
                "Oferecer orientações sobre ergonomia"
            ],
            "Risco Baixo": [
                "Realizar check-ups periódicos das condições ambientais",
                "Incentivar sugestões de melhoria do ambiente"
            ],
            "Risco Muito Baixo": [
                "Documentar boas práticas ambientais",
                "Monitorar satisfação com o ambiente de trabalho"
            ]
        },
        "Interação pessoa-tarefa": {
            "Risco Muito Alto": [
                "Realizar análise detalhada da adequação pessoa-função",
                "Implementar programa de adequação de perfis às funções",
                "Oferecer treinamentos específicos para desenvolvimento de competências críticas",
                "Estabelecer mentoria para funções complexas",
                "Revisar processos de seleção interna e alocação de pessoal"
            ],
            "Risco Alto": [
                "Mapear competências necessárias vs. disponíveis",
                "Implementar programa de desenvolvimento individualizado",
                "Revisar sistemas de suporte às tarefas",
                "Estabelecer reuniões regulares de alinhamento"
            ],
            "Risco Moderado": [
                "Promover treinamentos para aprimoramento técnico",
                "Implementar sistema de documentação de processos",
                "Estabelecer grupos de compartilhamento de práticas"
            ],
            "Risco Baixo": [
                "Manter programas de educação continuada",
                "Solicitar feedback sobre adequação pessoa-tarefa"
            ],
            "Risco Muito Baixo": [
                "Documentar casos de sucesso de adequação pessoa-tarefa",
                "Incentivar inovações nos métodos de trabalho"
            ]
        },
        "Jornada de trabalho": {
            "Risco Muito Alto": [
                "Realizar diagnóstico completo de carga horária e distribuição do trabalho",
                "Implementar políticas de limitação de horas extras",
                "Revisar escalas de trabalho com participação dos trabalhadores",
                "Estabelecer sistema de monitoramento de banco de horas",
                "Implementar medidas para evitar trabalho em horários de descanso"
            ],
            "Risco Alto": [
                "Revisar políticas de jornada flexível",
                "Implementar sistema de compensação de horas",
                "Estabelecer limites claros para comunicação fora do expediente",
                "Promover distribuição equilibrada de tarefas"
            ],
            "Risco Moderado": [
                "Oferecer orientações sobre gestão do tempo",
                "Implementar pausas programadas",
                "Monitorar horas extras recorrentes"
            ],
            "Risco Baixo": [
                "Avaliar periodicamente satisfação com a jornada",
                "Manter opções de flexibilidade horária"
            ],
            "Risco Muito Baixo": [
                "Documentar boas práticas de gestão de jornada",
                "Compartilhar casos de sucesso em equilíbrio trabalho-vida"
            ]
        }
    }

    plano_acao = []

    # Para cada fator psicossocial no resultado
    for _, row in df_resultados.iterrows():
        fator = row['Fator Psicossocial']
        risco = row['Risco']
        nivel_risco = risco.split()[0] + " " + risco.split()[1]  # Ex: "Risco Alto"

        # Obter sugestões para este fator e nível de risco
        if fator in sugestoes_por_fator and nivel_risco in sugestoes_por_fator[fator]:
            sugestoes = sugestoes_por_fator[fator][nivel_risco]

            # Adicionar ao plano de ação
            for sugestao in sugestoes:
                plano_acao.append({
                    "Fator Psicossocial": fator,
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

# Função para formatar a aba de plano de ação
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
    worksheet.set_column('A:A', 25)  # Fator Psicossocial
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
    for row_num, row in enumerate(df.itertuples(), 1):
        nivel_risco = getattr(row, 'Nível_de_Risco')
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
        worksheet.set_column('B:B', 30)  # Fator Psicossocial
        worksheet.set_column('C:C', 10)  # Média
        worksheet.set_column('D:D', 20)  # Risco
        worksheet.set_column('E:E', 15)  # Número de Respostas
    else:
        # Para tabelas pivotadas
        worksheet.set_column('A:A', 30)  # Fator Psicossocial
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
        pdf.cell(200, 10, "Plano de Acao - Fatores Psicossociais (HSE-IT)", ln=True, align='C')
        pdf.ln(10)
        
        # Data do relatório
        pdf.set_font("Arial", size=10)
        from datetime import datetime
        data_atual = datetime.now().strftime("%d/%m/%Y")
        pdf.cell(0, 10, f"Data do relatorio: {data_atual}", ln=True)
        pdf.ln(5)
        
        # Agrupar por fator psicossocial
        for fator in df_plano_acao["Fator Psicossocial"].unique():
            df_fator = df_plano_acao[df_plano_acao["Fator Psicossocial"] == fator]
            nivel_risco = df_fator["Nível de Risco"].iloc[0]
            media = df_fator["Média"].iloc[0]
            
            # Título do fator
            pdf.set_font("Arial", style='B', size=12)
            pdf.cell(0, 10, f"{fator}", ln=True)
            
            # Informações de risco
            pdf.set_font("Arial", style='I', size=10)
            pdf.cell(0, 8, f"Media: {media} - Nivel de Risco: {nivel_risco}", ln=True)
            
            # Ações sugeridas
            pdf.set_font("Arial", size=10)
            pdf.cell(0, 6, "Acoes Sugeridas:", ln=True)
            
            for _, row in df_fator.iterrows():
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
            for _, row in df_fator.iterrows():
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
def gerar_excel_completo(df, df_perguntas, colunas_filtro, fatores, colunas_perguntas):
    output = io.BytesIO()
    
    try:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Aba 1: Empresa Toda
            resultados_empresa = calcular_resultados(df, df_perguntas, fatores, colunas_perguntas)
            df_resultados_empresa = pd.DataFrame(resultados_empresa)
            df_resultados_empresa.to_excel(writer, sheet_name='Empresa Toda', index=False)
            
            # Formatar a planilha
            worksheet = writer.sheets['Empresa Toda']
            formatar_aba_excel(workbook, worksheet, df_resultados_empresa)
            
            # Aba 2: Plano de Ação
            df_plano_acao = gerar_sugestoes_acoes(df_resultados_empresa)
            df_plano_acao.to_excel(writer, sheet_name='Plano de Ação', index=False)
            
            # Formatar a aba de plano de ação
            worksheet_plano = writer.sheets['Plano de Ação']
            formatar_aba_plano_acao(workbook, worksheet_plano, df_plano_acao)
            
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
                            resultados_filtro = calcular_resultados(df_filtrado, df_perguntas_filtradas, fatores, colunas_perguntas)
                            
                            # Adicionar coluna com o valor do filtro
                            for res in resultados_filtro:
                                res[filtro] = valor
                                resultados_resumo.append(res)
                    
                    if resultados_resumo:
                        df_resumo = pd.DataFrame(resultados_resumo)
                        
                        # Pivotear para melhor visualização
                        if len(resultados_resumo) > 0:
                            try:
                                df_pivot = df_resumo.pivot(index='Fator Psicossocial', columns=filtro, values='Média')
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
    worksheet.write_column('A1', ['Fator Psicossocial'] + list(df_resultados['Fator Psicossocial']))
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
    chart.set_title({'name': 'Fatores Psicossociais - Avaliação de Riscos'})
    chart.set_x_axis({'name': 'Fator Psicossocial'})
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
# Verificar autenticação antes de mostrar o conteúdo
if check_password():
    # Criar sistema de abas para melhor organização
    tabs = st.tabs(["Upload de Dados", "Resultados", "Plano de Ação", "Relatórios"])

    # Primeira aba - Upload e configuração
    with tabs[0]:
        st.title("Avaliação de Fatores Psicossociais - HSE-IT")
        st.write("Faça upload do arquivo Excel contendo os resultados do questionário.")
        
        # Adiciona uma explicação sobre o formato esperado do arquivo
        with st.expander("Informações sobre o formato do arquivo"):
            st.write("""
            O arquivo deve conter:
            - Colunas de filtro (Setor, Cargo, etc.) nas primeiras 7 colunas
            - Colunas com as perguntas numeradas (começando com números)
            - As respostas devem ser valores numéricos de 1 a 5
            """)
        
        uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=["xlsx", "csv"])
        
        # Barra lateral para filtros e configurações
        with st.sidebar:
            st.header("Filtros e Configurações")
            
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
                    
                    # Definir fatores psicossociais
                    fatores = {
                        "Gestão organizacional": [15, 19, 25, 26, 28, 32],
                        "Violência e assédio moral/sexual no trabalho": [5, 14, 21, 34],
                        "Contexto da organização do trabalho": [3, 6, 9, 1, 4, 11, 13, 8, 27, 31, 35],
                        "Características das relações sociais no trabalho": [7, 24, 27, 31, 5, 14, 21, 34],
                        "Conteúdo das tarefas do trabalho": [6, 9, 12, 18, 20, 22],
                        "Discriminação no trabalho": [14, 21],
                        "Condições do ambiente de trabalho": [5, 6, 14, 21, 34],
                        "Interação pessoa-tarefa": [10, 19, 6, 9, 12, 20],
                        "Jornada de trabalho": [16, 18, 30]
                    }
                    
                    # Calcular resultados
                    resultados = calcular_resultados(df_filtrado, df_perguntas_filtradas, fatores, colunas_perguntas)
                    df_resultados = pd.DataFrame(resultados)
                    
                    # Gerar plano de ação
                    df_plano_acao = gerar_sugestoes_acoes(df_resultados)
                    
                except Exception as e:
                    st.error(f"Ocorreu um erro inesperado: {str(e)}")
                    st.write("Detalhes do erro para debug:")
                    st.exception(e)
    
    # Segunda aba - Visualização de resultados
    with tabs[1]:
        if uploaded_file is not None:
            st.header("Resultados da Avaliação")
            
            # Mostrar estatísticas gerais
            st.write("### Resumo da Avaliação")
            media_geral = df_resultados["Média"].mean()
            st.metric(
                label="Média Geral",
                value=f"{media_geral:.2f}",
                delta=f"{classificar_risco(media_geral)[0]}"
            )
            
            st.write("### Resultados Detalhados da Avaliação")
            st.dataframe(df_resultados)
            
            # Criar gráfico de barras para visualização dos riscos
            fig, ax = plt.subplots(figsize=(10, 6))
            cores = [classificar_risco(media)[1] for media in df_resultados["Média"]]
            bars = ax.barh(df_resultados["Fator Psicossocial"], df_resultados["Média"], color=cores)
            
            # Adicionar rótulos de dados nas barras
            for i, bar in enumerate(bars):
                ax.text(
                    bar.get_width() + 0.1,
                    bar.get_y() + bar.get_height()/2,
                    f"{df_resultados['Média'].iloc[i]:.2f} ({df_resultados['Risco'].iloc[i].split()[0]} {df_resultados['Risco'].iloc[i].split()[1]})",
                    va='center'
                )
            
            # Melhorar o layout e aparência do gráfico
            ax.set_xlabel("Média das Respostas")
            ax.set_title(f"Classificação de Riscos Psicossociais ({filtro_opcao}: {filtro_valor})")
            ax.grid(axis='x', linestyle='--', alpha=0.7)
            ax.set_xlim(0, 5) # Define o limite dos eixos para escala de 1-5
            
            # Adicionar legenda de cores
            from matplotlib.lines import Line2D
            legend_elements = [
                Line2D([0], [0], color='red', lw=4, label='Risco Muito Alto (≤1)'),
                Line2D([0], [0], color='orange', lw=4, label='Risco Alto (1-2)'),
                Line2D([0], [0], color='yellow', lw=4, label='Risco Moderado (2-3)'),
                Line2D([0], [0], color='green', lw=4, label='Risco Baixo (3-4)'),
                Line2D([0], [0], color='purple', lw=4, label='Risco Muito Baixo (>4)'),
            ]
            ax.legend(handles=legend_elements, loc='lower right')
            plt.tight_layout()
            st.pyplot(fig)
    
    # Terceira aba - Plano de ação
    with tabs[2]:
        if uploaded_file is not None:
            st.header("Plano de Ação Sugerido")
            st.write("Com base nos resultados da avaliação, aqui estão sugestões de ações para seu PGR:")
            
            # Criar tabs para organizar a visualização do plano de ação
            subtabs = st.tabs(["Todos os Fatores"] + list(df_plano_acao["Fator Psicossocial"].unique()))
            
            # Tab com todas as ações
            with subtabs[0]:
                # Adicionar filtros interativos
                col1, col2 = st.columns(2)
                with col1:
                    niveis_risco = ["Todos"] + sorted(df_plano_acao["Nível de Risco"].unique().tolist(),
                        key=lambda x: ["Risco Muito Alto", "Risco Alto", "Risco Moderado",
                        "Risco Baixo", "Risco Muito Baixo"].index(x)
                        if x in ["Risco Muito Alto", "Risco Alto", "Risco Moderado",
                        "Risco Baixo", "Risco Muito Baixo"] else 999)
                    filtro_risco = st.selectbox("Filtrar por nível de risco:", niveis_risco)
                
                # Aplicar filtros
                df_filtrado = df_plano_acao
                if filtro_risco != "Todos":
                    df_filtrado = df_filtrado[df_filtrado["Nível de Risco"] == filtro_risco]
                
                # Mostrar o plano de ação filtrado
                st.dataframe(
                    df_filtrado[["Fator Psicossocial", "Nível de Risco", "Média", "Sugestão de Ação"]],
                    use_container_width=True,
                    height=400
                )
                
                # Indicar o total de ações sugeridas
                st.info(f"Total de {len(df_filtrado)} ações sugeridas")
            
            # Tabs para cada fator psicossocial
            for i, fator in enumerate(df_plano_acao["Fator Psicossocial"].unique(), 1):
                with subtabs[i]:
                    df_fator = df_plano_acao[df_plano_acao["Fator Psicossocial"] == fator]
                    
                    # Mostrar média e nível de risco
                    nivel_risco = df_fator["Nível de Risco"].iloc[0]
                    media = df_fator["Média"].iloc[0]
                    
                    # Definir cor com base no nível de risco
                    cor = {
                        "Risco Muito Alto": "red",
                        "Risco Alto": "orange",
                        "Risco Moderado": "yellow",
                        "Risco Baixo": "green",
                        "Risco Muito Baixo": "purple"
                    }.get(nivel_risco, "gray")
                    
                    st.markdown(f"**Média:** {media} - **Nível de Risco:** :{cor}[{nivel_risco}]")
                    
                    # Mostrar as ações sugeridas
                    st.subheader("Ações Sugeridas:")
                    for i, row in df_fator.iterrows():
                        st.markdown(f"- {row['Sugestão de Ação']}")
    
    # Quarta aba - Relatórios
    with tabs[3]:
        if uploaded_file is not None:
            st.header("Download de Relatórios")
            
            col1, col2 = st.columns(2)
            
            # Usar a função para gerar Excel com múltiplas abas e plano de ação
            excel_data = gerar_excel_completo(df, df_perguntas, colunas_filtro, fatores, colunas_perguntas)
            if excel_data:
                with col1:
                    st.download_button(
                        label="Baixar Relatório Excel Completo",
                        data=excel_data,
                        file_name=f"relatorio_completo_hse.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Baixe um relatório Excel com múltiplas abas contendo análises por empresa, setor, cargo, etc. e o plano de ação sugerido."
                    )
            
            # Gerar PDF da análise
            pdf_data = gerar_pdf(df_resultados)
            if pdf_data:
                with col2:
                    st.download_button(
                        label="Baixar Relatório PDF",
                        data=pdf_data,
                        file_name=f"relatorio_hse_{filtro_opcao}_{filtro_valor}.pdf",
                        mime="application/pdf"
                    )
            
            # Adicionar opção para baixar o PDF do Plano de Ação
            pdf_plano_acao = gerar_pdf_plano_acao(df_plano_acao)
            if pdf_plano_acao:
                st.download_button(
                    label="Baixar Plano de Ação PDF",
                    data=pdf_plano_acao,
                    file_name=f"plano_acao_hse_{filtro_opcao}_{filtro_valor}.pdf",
                    mime="application/pdf",
                    help="Baixe um PDF específico com o plano de ação sugerido, incluindo campos para preenchimento de responsáveis e prazos."
                )
else:
    st.stop()  # Não mostrar nada abaixo deste ponto se a autenticação falhar
