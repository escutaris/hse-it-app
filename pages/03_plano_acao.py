import streamlit as st
import pandas as pd
import io
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta
from utils.processamento import classificar_risco, padronizar_formato_data

# Aplicar estilo consistente da Escutaris
def aplicar_estilo_escutaris():
    st.markdown("""
    <style>
    /* Cores principais */
    :root {
        --escutaris-verde: #5A713D;
        --escutaris-cinza: #2E2F2F;
        --escutaris-bege: #F5F0EB;
        --escutaris-laranja: #FF5722;
    }
    
    /* Títulos */
    h1, h2, h3 {
        color: var(--escutaris-verde) !important;
        font-family: 'Helvetica Neue', Arial, sans-serif !important;
    }
    
    /* Subtítulos */
    h4, h5, h6 {
        color: var(--escutaris-cinza) !important;
        font-family: 'Helvetica Neue', Arial, sans-serif !important;
    }
    
    /* Botões */
    .stButton>button {
        background-color: var(--escutaris-verde) !important;
        color: white !important;
        border-radius: 5px !important;
    }
    
    /* Métricas */
    div.css-1xarl3l.e16fv1kl1 {
        background-color: white;
        border-radius: 10px;
        padding: 10px 15px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    
    /* Cards */
    .card {
        background-color: white;
        border-radius: 10px;
        padding: 15px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        margin-bottom: 15px;
    }
    
    /* Abas */
    .stTabs [data-baseweb="tab-list"] {
        gap: 2px;
    }
    
    .stTabs [data-baseweb="tab"] {
        background-color: #f0f0f0;
        border-radius: 4px 4px 0px 0px;
        padding: 10px 20px;
        border: none;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: var(--escutaris-verde) !important;
        color: white !important;
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        background-color: #f5f5f5;
        border-radius: 5px;
    }
    
    /* Progress bar */
    .stProgress > div > div > div > div {
        background-color: var(--escutaris-verde);
    }
    
    /* Cor dos links */
    a {
        color: var(--escutaris-verde) !important;
    }
    
    /* Formulário de ação */
    .action-form {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 8px;
        margin-bottom: 10px;
        border-left: 4px solid var(--escutaris-verde);
    }
    
    /* Status de risco */
    .risco-muito-alto {
        background-color: #f8d7da;
        color: #721c24;
        padding: 3px 8px;
        border-radius: 4px;
        font-size: 0.9em;
        font-weight: bold;
    }
    
    .risco-alto {
        background-color: #fff3cd;
        color: #856404;
        padding: 3px 8px;
        border-radius: 4px;
        font-size: 0.9em;
        font-weight: bold;
    }
    
    .risco-moderado {
        background-color: #fff3cd;
        color: #856404;
        padding: 3px 8px;
        border-radius: 4px;
        font-size: 0.9em;
    }
    
    .risco-baixo {
        background-color: #d4edda;
        color: #155724;
        padding: 3px 8px;
        border-radius: 4px;
        font-size: 0.9em;
    }
    
    .risco-muito-baixo {
        background-color: #d1ecf1;
        color: #0c5460;
        padding: 3px 8px;
        border-radius: 4px;
        font-size: 0.9em;
    }
    
    /* Tabela de plano de ação */
    .tabela-plano-acao {
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 20px;
    }
    
    .tabela-plano-acao th {
        background-color: var(--escutaris-verde);
        color: white;
        text-align: left;
        padding: 12px;
        font-size: 0.9em;
    }
    
    .tabela-plano-acao td {
        padding: 10px;
        border-bottom: 1px solid #ddd;
        font-size: 0.9em;
    }
    
    .tabela-plano-acao tr:nth-child(even) {
        background-color: #f8f9fa;
    }
    
    .tabela-plano-acao tr:hover {
        background-color: #f1f1f1;
    }
    
    /* Tooltip personalizado */
    .tooltip {
        position: relative;
        display: inline-block;
        border-bottom: 1px dotted #ccc;
        cursor: help;
    }
    
    .tooltip .tooltiptext {
        visibility: hidden;
        width: 200px;
        background-color: #555;
        color: #fff;
        text-align: center;
        border-radius: 6px;
        padding: 5px;
        position: absolute;
        z-index: 1;
        bottom: 125%;
        left: 50%;
        margin-left: -100px;
        opacity: 0;
        transition: opacity 0.3s;
    }
    
    .tooltip:hover .tooltiptext {
        visibility: visible;
        opacity: 1;
    }
    
    /* Melhoria visual para áreas de texto */
    textarea {
        border-radius: 5px;
    }
    
    /* Caixa de informação destacada */
    .info-box {
        background-color: #F5F7F0;
        border-left: 4px solid var(--escutaris-verde);
        padding: 15px;
        margin: 15px 0;
        border-radius: 5px;
    }
    
    /* Badge de nível de risco */
    .badge-risk {
        display: inline-block;
        padding: 3px 10px;
        border-radius: 15px;
        font-size: 0.8em;
        font-weight: bold;
        margin: 5px 0;
    }
    
    /* Ajustes para visualização mobile */
    @media (max-width: 768px) {
        .tabela-plano-acao th, .tabela-plano-acao td {
            padding: 8px 5px;
            font-size: 0.8em;
        }
    }
    </style>
    """, unsafe_allow_html=True)

# Aplicar o estilo
aplicar_estilo_escutaris()

# Título da página
st.title("Plano de Ação - HSE-IT")

# Verificar se há dados para exibir - CORRIGIDO: Verificação segura
if "df_resultados" not in st.session_state or st.session_state.get("df_resultados") is None:
    st.warning("Nenhum resultado disponível. Por favor, faça upload de um arquivo na página 'Upload de Dados'.")
    st.stop()

# Texto introdutório explicando o propósito do plano de ação
st.markdown("""
<div class="info-box">
<p>Este plano de ação contém sugestões de medidas baseadas nos resultados da avaliação HSE-IT. 
As sugestões devem ser adaptadas à realidade da sua organização e servir como ponto de partida
para discussões com a gestão e equipe de saúde e segurança.</p> 

<p>Utilize este conteúdo para subsidiar o Programa de Gerenciamento de Riscos (PGR) e demais 
documentos pertinentes à saúde e segurança no trabalho.</p>
</div>
""", unsafe_allow_html=True)

# Filtros para o plano de ação
st.subheader("Filtros")

# Recuperar os resultados de forma segura - CORRIGIDO
try:
    df_resultados = st.session_state.get("df_resultados")
    df_plano_acao = st.session_state.get("df_plano_acao")
    
    if df_resultados is None or df_plano_acao is None:
        st.error("Dados do plano de ação não disponíveis. Por favor, retorne à página de upload.")
        st.stop()
except Exception as e:
    st.error(f"Erro ao recuperar dados da sessão: {str(e)}")
    st.info("Por favor, retorne à página de upload e carregue seus dados novamente.")
    st.stop()

# Definir níveis de risco para filtrar
niveis_risco = ["Risco Muito Alto 🔴", "Risco Alto 🟠", "Risco Moderado 🟡", "Risco Baixo 🟢", "Risco Muito Baixo 🟣"]

# Definir filtros para o plano de ação
col1, col2 = st.columns(2)

with col1:
    # Filtro por nível de risco
    niveis_selecionados = st.multiselect(
        "Filtrar por nível de risco:",
        niveis_risco,
        default=["Risco Muito Alto 🔴", "Risco Alto 🟠", "Risco Moderado 🟡"],
        help="Selecione os níveis de risco que deseja incluir no plano de ação"
    )
    
    # Adicionar nota sobre níveis opcionais
    st.caption("Nota: Os níveis 'Risco Baixo 🟢' e 'Risco Muito Baixo 🟣' são opcionais e podem ser incluídos conforme necessidade da organização.")

with col2:
    # Filtro por dimensão - CORRIGIDO: Verificação segura de "Dimensão"
    dimensoes = []
    try:
        if "Dimensão" in df_resultados.columns:
            dimensoes = df_resultados["Dimensão"].unique()
        else:
            st.warning("Não foi possível identificar as dimensões nos resultados.")
    except Exception as e:
        st.error(f"Erro ao obter dimensões: {str(e)}")
        
    dimensoes_selecionadas = st.multiselect(
        "Filtrar por dimensão:",
        dimensoes,
        default=list(dimensoes),
        help="Selecione as dimensões que deseja incluir no plano de ação"
    )

# Criar dataframe para conter os dados do plano de ação
def gerar_plano_acao_tabular():
    try:
        # Filtrar resultados conforme seleções
        df_filtrado = df_resultados[
            (df_resultados["Dimensão"].isin(dimensoes_selecionadas)) &
            (df_resultados["Risco"].isin(niveis_selecionados))
        ]
        
        # Se não há dados após filtro, retornar DataFrame vazio com as colunas corretas
        if df_filtrado.empty:
            return pd.DataFrame(columns=[
                "Dimensão", 
                "Média", 
                "Nível de Risco",
                "Riscos Potenciais", 
                "Sugestões de Ações Mitigantes", 
                "Outras Soluções", 
                "Responsável", 
                "Prazo"
            ])
        
        # Criar o plano de ação em formato tabular
        plano_acao_rows = []
        
        # Mapeamento de riscos potenciais para cada dimensão
        riscos_potenciais = {
            "Demanda": "Sobrecarga de trabalho levando a estresse, burnout, fadiga, erros operacionais e potencial adoecimento físico e mental.",
            "Controle": "Falta de autonomia e participação nas decisões, gerando desmotivação, alienação, insatisfação e menor engajamento.",
            "Apoio da Chefia": "Liderança inadequada causando desmotivação, conflitos, falta de direcionamento e baixo desempenho da equipe.",
            "Apoio dos Colegas": "Ambiente de trabalho não colaborativo, levando a isolamento, dificuldades de integração e baixa produtividade em equipe.",
            "Relacionamentos": "Conflitos interpessoais, assédio moral/sexual, violência e deterioração do clima organizacional.",
            "Função": "Falta de clareza sobre papéis e responsabilidades, causando retrabalho, conflitos, ineficiência e estresse.",
            "Mudança": "Resistência, ansiedade e insegurança frente a mudanças organizacionais, prejudicando adaptação e implementação."
        }
        
        # Definir ações mitigantes para cada nível de risco e dimensão
        acoes_mitigantes = {
            "Demanda": {
                "Risco Muito Alto 🔴": [
                    "Realizar auditoria completa da distribuição de carga de trabalho",
                    "Implementar sistema de gestão de tarefas e priorização",
                    "Reavaliar prazos e expectativas de produtividade",
                    "Contratar pessoal adicional para áreas sobrecarregadas"
                ],
                "Risco Alto 🟠": [
                    "Mapear atividades e identificar gargalos de processo",
                    "Implementar ferramentas para melhor organização do trabalho",
                    "Revisar e ajustar prazos de entregas e metas",
                    "Capacitar gestores em gerenciamento de carga de trabalho"
                ],
                "Risco Moderado 🟡": [
                    "Promover treinamentos de gestão do tempo e priorização",
                    "Revisar distribuição de tarefas entre membros da equipe",
                    "Implementar pausas regulares durante a jornada de trabalho"
                ],
                "Risco Baixo 🟢": [
                    "Manter monitoramento regular das demandas de trabalho",
                    "Realizar check-ins periódicos sobre volume de trabalho"
                ],
                "Risco Muito Baixo 🟣": [
                    "Documentar boas práticas atuais de gestão de demandas",
                    "Manter práticas de gestão de demandas e continuar monitorando"
                ]
            },
            "Controle": {
                "Risco Muito Alto 🔴": [
                    "Redesenhar processos para aumentar a autonomia dos trabalhadores",
                    "Implementar esquemas de trabalho flexível",
                    "Revisar políticas de microgerenciamento",
                    "Treinar gestores em delegação efetiva"
                ],
                "Risco Alto 🟠": [
                    "Identificar áreas onde os trabalhadores podem ter mais controle",
                    "Envolver colaboradores no planejamento de metas e métodos",
                    "Implementar sistema de sugestões para melhorias nos processos",
                    "Oferecer opções de horários flexíveis"
                ],
                "Risco Moderado 🟡": [
                    "Aumentar gradualmente a autonomia nas decisões rotineiras",
                    "Solicitar feedback regular sobre nível de controle no trabalho",
                    "Implementar projetos-piloto para testar maior autonomia"
                ],
                "Risco Baixo 🟢": [
                    "Manter boas práticas de autonomia",
                    "Revisar periodicamente áreas onde o controle pode ser ampliado"
                ],
                "Risco Muito Baixo 🟣": [
                    "Documentar práticas bem-sucedidas de autonomia",
                    "Manter cultura de confiança e delegação"
                ]
            },
            "Apoio da Chefia": {
                "Risco Muito Alto 🔴": [
                    "Implementar programa estruturado de desenvolvimento de lideranças",
                    "Realizar avaliação 360° para gestores",
                    "Estabelecer canais de comunicação direta entre colaboradores e direção",
                    "Revisar políticas de promoção para valorizar bons líderes"
                ],
                "Risco Alto 🟠": [
                    "Treinar gestores em habilidades de suporte e feedback",
                    "Implementar reuniões regulares one-on-one com liderados",
                    "Estabelecer expectativas claras para comportamentos de liderança",
                    "Criar fóruns para líderes compartilharem desafios e soluções"
                ],
                "Risco Moderado 🟡": [
                    "Revisar e melhorar as práticas de feedback das lideranças",
                    "Promover workshops sobre comunicação efetiva",
                    "Oferecer recursos para líderes apoiarem suas equipes"
                ],
                "Risco Baixo 🟢": [
                    "Manter programas de desenvolvimento de lideranças",
                    "Reconhecer e celebrar boas práticas de liderança"
                ],
                "Risco Muito Baixo 🟣": [
                    "Documentar práticas exemplares de liderança",
                    "Utilizar líderes como mentores para novos gestores"
                ]
            },
            "Apoio dos Colegas": {
                "Risco Muito Alto 🔴": [
                    "Implementar programas estruturados de team building",
                    "Revisar a composição e dinâmica das equipes",
                    "Estabelecer facilitadores de equipe para melhorar integração",
                    "Criar espaços físicos e virtuais para colaboração"
                ],
                "Risco Alto 🟠": [
                    "Promover atividades regulares de integração de equipes",
                    "Treinar em habilidades de trabalho em equipe e comunicação",
                    "Estabelecer objetivos compartilhados que incentivem a colaboração",
                    "Revisar processos que possam estar criando competição indesejada"
                ],
                "Risco Moderado 🟡": [
                    "Implementar reuniões regulares de equipe para compartilhamento",
                    "Criar projetos colaborativos entre diferentes membros",
                    "Oferecer oportunidades para pessoas se conhecerem melhor"
                ],
                "Risco Baixo 🟢": [
                    "Manter momentos regulares de integração",
                    "Monitorar dinâmicas de equipe, especialmente com novos membros"
                ],
                "Risco Muito Baixo 🟣": [
                    "Documentar boas práticas de colaboração",
                    "Manter ambiente de confiança e colaboração"
                ]
            },
            "Relacionamentos": {
                "Risco Muito Alto 🔴": [
                    "Implementar política de tolerância zero para assédio e comportamentos inadequados",
                    "Criar canais confidenciais para denúncias",
                    "Treinar todos os colaboradores em respeito e diversidade",
                    "Estabelecer mediação de conflitos com profissionais externos"
                ],
                "Risco Alto 🟠": [
                    "Desenvolver política clara sobre comportamentos aceitáveis",
                    "Treinar gestores na identificação e gestão de conflitos",
                    "Implementar processos estruturados para resolução de conflitos",
                    "Promover diálogo sobre relacionamentos saudáveis no trabalho"
                ],
                "Risco Moderado 🟡": [
                    "Realizar workshops sobre comunicação não-violenta",
                    "Estabelecer acordos de equipe sobre comportamentos esperados",
                    "Promover atividades que construam confiança entre colegas"
                ],
                "Risco Baixo 🟢": [
                    "Manter comunicação regular sobre respeito no ambiente de trabalho",
                    "Incorporar avaliação de relacionamentos nas pesquisas de clima"
                ],
                "Risco Muito Baixo 🟣": [
                    "Documentar boas práticas de relacionamentos saudáveis",
                    "Manter monitoramento contínuo do clima relacional"
                ]
            },
            "Função": {
                "Risco Muito Alto 🔴": [
                    "Realizar revisão completa de descrições de cargos e responsabilidades",
                    "Implementar processo de clarificação de funções para toda a organização",
                    "Treinar gestores em delegação clara e definição de expectativas",
                    "Estabelecer processos para resolver conflitos de papéis"
                ],
                "Risco Alto 🟠": [
                    "Revisar e atualizar descrições de cargo",
                    "Implementar reuniões para esclarecer expectativas",
                    "Criar matriz RACI para projetos e processos",
                    "Treinar equipes em comunicação sobre papéis e responsabilidades"
                ],
                "Risco Moderado 🟡": [
                    "Revisar processos onde ocorrem conflitos de funções",
                    "Promover workshops para clarificar interfaces entre áreas",
                    "Estabelecer fóruns para discutir e esclarecer papéis em projetos"
                ],
                "Risco Baixo 🟢": [
                    "Manter atualizações periódicas de responsabilidades",
                    "Incluir clareza de papéis nas avaliações de desempenho"
                ],
                "Risco Muito Baixo 🟣": [
                    "Documentar boas práticas de clareza de funções",
                    "Manter cultura de transparência sobre papéis e responsabilidades"
                ]
            },
            "Mudança": {
                "Risco Muito Alto 🔴": [
                    "Implementar metodologia estruturada de gestão de mudanças",
                    "Criar comitê representativo para planejamento de mudanças",
                    "Estabelecer múltiplos canais de comunicação sobre processos de mudança",
                    "Treinar gestores em liderança durante transformações"
                ],
                "Risco Alto 🟠": [
                    "Desenvolver plano de comunicação para mudanças organizacionais",
                    "Envolver representantes de diferentes níveis no planejamento",
                    "Implementar feedbacks regulares durante processos de mudança",
                    "Oferecer suporte adicional para equipes mais afetadas"
                ],
                "Risco Moderado 🟡": [
                    "Melhorar a transparência sobre razões das mudanças",
                    "Criar fóruns para esclarecer dúvidas sobre transformações",
                    "Celebrar pequenas vitórias durante processos de mudança"
                ],
                "Risco Baixo 🟢": [
                    "Manter comunicação proativa sobre possíveis mudanças",
                    "Oferecer oportunidades regulares para feedback durante transformações"
                ],
                "Risco Muito Baixo 🟣": [
                    "Documentar práticas bem-sucedidas de gestão de mudanças",
                    "Manter cultura de adaptabilidade e melhoria contínua"
                ]
            }
        }
        
        # Para cada dimensão nos resultados filtrados, gerar linhas do plano de ação
        for _, row in df_filtrado.iterrows():
            dimensao = row["Dimensão"]
            media = row["Média"]
            nivel_risco = row["Risco"]
            
            # Obter riscos potenciais para a dimensão com fallback seguro
            risco_potencial = riscos_potenciais.get(dimensao, "Riscos não especificados para esta dimensão.")
            
            # Obter ações mitigantes para o nível de risco e dimensão
            acoes = []
            acoes_por_dimensao = acoes_mitigantes.get(dimensao, {})
            if nivel_risco in acoes_por_dimensao:
                acoes = acoes_por_dimensao[nivel_risco]
            else:
                # Tentar correspondência parcial se não houver correspondência exata
                for nivel_key in acoes_por_dimensao.keys():
                    if nivel_key.split()[0] in nivel_risco:
                        acoes = acoes_por_dimensao[nivel_key]
                        break
                
                # Se ainda não encontrou, usar mensagem padrão
                if not acoes:
                    acoes = ["Não há sugestões específicas para este nível de risco."]
                
            acoes_formatadas = "\n".join([f"• {acao}" for acao in acoes])
            
            # Adicionar à lista de linhas
            plano_acao_rows.append({
                "Dimensão": dimensao,
                "Média": media,
                "Nível de Risco": nivel_risco,
                "Riscos Potenciais": risco_potencial,
                "Sugestões de Ações Mitigantes": acoes_formatadas,
                "Outras Soluções": "",
                "Responsável": "",
                "Prazo": ""
            })
        
        return pd.DataFrame(plano_acao_rows)
    except Exception as e:
        st.error(f"Erro ao gerar plano de ação: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return pd.DataFrame()

# Gerar o plano de ação
df_plano = gerar_plano_acao_tabular()

# Verificar se há resultados após a filtragem
if df_plano.empty:
    st.warning("Não foram encontrados resultados para os filtros selecionados. Por favor, ajuste os filtros.")
else:
    # Exibir o plano de ação em formato tabular editável
    st.subheader("Plano de Ação - HSE-IT")
    
    # Inicializar ou recuperar o plano editável da sessão - CORRIGIDO: Uso seguro
    if "plano_editavel" not in st.session_state:
        st.session_state["plano_editavel"] = df_plano.copy()
    elif len(st.session_state.get("plano_editavel", pd.DataFrame())) != len(df_plano):
        # Se os filtros mudaram e o tamanho mudou, atualizar o plano
        st.session_state["plano_editavel"] = df_plano.copy()
    
    # Criar uma visualização editável do plano
    try:
        # CORREÇÃO: Verificar se o atributo TextAreaColumn existe
        # Define configurações de colunas com verificação de compatibilidade para diferentes versões do Streamlit
        
        # Configuração para Dimensão
        dimensao_config = st.column_config.TextColumn(
            "Dimensão",
            help="Dimensão do HSE-IT avaliada",
            disabled=True,
            width="medium"
        )
        
        # Configuração para Média
        media_config = st.column_config.NumberColumn(
            "Média",
            help="Pontuação média obtida na avaliação (1-5)",
            format="%.2f",
            disabled=True,
            width="small"
        )
        
        # Configuração para Nível de Risco
        nivel_risco_config = st.column_config.TextColumn(
            "Nível de Risco",
            help="Classificação do nível de risco baseada na média",
            disabled=True,
            width="medium"
        )
        
        # Configuração para Riscos Potenciais
        riscos_config = st.column_config.TextColumn(
            "Riscos Potenciais",
            help="Consequências potenciais relacionadas a este fator psicossocial",
            width="large"
        )
        
        # CORREÇÃO: Verificar se TextAreaColumn está disponível para Sugestões de Ações
        if hasattr(st.column_config, "TextAreaColumn"):
            # Usar TextAreaColumn se disponível (versões mais recentes)
            sugestoes_config = st.column_config.TextAreaColumn(
                "Sugestões de Ações Mitigantes",
                help="Ações sugeridas para mitigar os riscos identificados",
                width="large",
                height="medium"
            )
        else:
            # Fallback para versões mais antigas
            sugestoes_config = st.column_config.Column(
                "Sugestões de Ações Mitigantes",
                help="Ações sugeridas para mitigar os riscos identificados",
                width="large"
            )
        
        # CORREÇÃO: Verificar se TextAreaColumn está disponível para Outras Soluções
        if hasattr(st.column_config, "TextAreaColumn"):
            # Usar TextAreaColumn se disponível (versões mais recentes)
            outras_solucoes_config = st.column_config.TextAreaColumn(
                "Outras Soluções",
                help="Adicione outras soluções específicas para sua organização",
                width="large",
                height="medium"
            )
        else:
            # Fallback para versões mais antigas
            outras_solucoes_config = st.column_config.Column(
                "Outras Soluções",
                help="Adicione outras soluções específicas para sua organização",
                width="large"
            )
        
        # Configuração para Responsável
        responsavel_config = st.column_config.TextColumn(
            "Responsável",
            help="Pessoa ou equipe responsável pela implementação",
            width="medium"
        )
        
        # Configuração para Prazo
        prazo_config = st.column_config.DateColumn(
            "Prazo",
            help="Prazo para implementação das ações",
            min_value=datetime.now().date(),
            format="DD/MM/YYYY",
            width="medium"
        )
        
        # Usar as configurações na definição do data_editor
        edited_df = st.data_editor(
            st.session_state.get("plano_editavel", df_plano),
            column_config={
                "Dimensão": dimensao_config,
                "Média": media_config,
                "Nível de Risco": nivel_risco_config,
                "Riscos Potenciais": riscos_config,
                "Sugestões de Ações Mitigantes": sugestoes_config,
                "Outras Soluções": outras_solucoes_config,
                "Responsável": responsavel_config,
                "Prazo": prazo_config
            },
            use_container_width=True,
            num_rows="dynamic",
            hide_index=True
        )
        
        # Atualizar o estado da sessão com as edições
        st.session_state["plano_editavel"] = edited_df
    except Exception as e:
        st.error(f"Erro ao exibir editor de dados: {str(e)}")
        st.info("Tente ajustar os filtros ou recarregar a página.")
        
        # Exibir detalhes técnicos do erro dentro de um expander para ajudar na depuração
        with st.expander("Detalhes técnicos do erro", expanded=False):
            st.code(f"{str(e)}")
            import traceback
            st.code(traceback.format_exc())
    
    # Informações adicionais
    with st.expander("Informações sobre o Plano de Ação"):
        st.markdown("""
        ### Como utilizar este plano de ação:
        
        1. **Revise os riscos potenciais**: Verifique se os riscos descritos se aplicam à sua realidade organizacional.
        
        2. **Ajuste as sugestões de ações**: Modifique as ações sugeridas conforme necessário para adequá-las ao seu contexto.
        
        3. **Adicione soluções específicas**: Use o campo "Outras Soluções" para incluir medidas personalizadas para sua organização.
        
        4. **Defina responsáveis e prazos**: Atribua claramente cada ação a um responsável e estabeleça prazos realistas.
        
        5. **Exporte o plano**: Use o botão abaixo para exportar o plano completo para Excel e compartilhar com a equipe de SST.
        
        6. **Integre ao PGR**: Incorpore estas ações ao Programa de Gerenciamento de Riscos da empresa.
        
        ### Níveis de prioridade recomendados:
        
        - **Risco Muito Alto** 🔴: Prioridade imediata (1-3 meses)
        - **Risco Alto** 🟠: Prioridade alta (3-6 meses)
        - **Risco Moderado** 🟡: Prioridade média (6-12 meses)
        - **Risco Baixo** 🟢: Prioridade baixa (12-18 meses)
        - **Risco Muito Baixo** 🟣: Manutenção de boas práticas (revisão anual)
        """)
    
    # Função para exportar o plano para Excel
    def exportar_para_excel(df):
        output = io.BytesIO()
        
        try:
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # Formatar e exportar o plano de ação
                df.to_excel(writer, sheet_name='Plano de Ação HSE-IT', index=False)
                worksheet = writer.sheets['Plano de Ação HSE-IT']
                
                # Definir formatos
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#5A713D',
                    'font_color': 'white',
                    'border': 1
                })
                
                wrap_format = workbook.add_format({
                    'text_wrap': True,
                    'valign': 'top',
                    'border': 1
                })
                
                # Formatos para níveis de risco
                risco_formats = {
                    'Risco Muito Alto 🔴': workbook.add_format({'bg_color': '#FF6B6B', 'bold': True, 'border': 1}),
                    'Risco Alto 🟠': workbook.add_format({'bg_color': '#FFA500', 'bold': True, 'border': 1}),
                    'Risco Moderado 🟡': workbook.add_format({'bg_color': '#FFFF00', 'border': 1}),
                    'Risco Baixo 🟢': workbook.add_format({'bg_color': '#90EE90', 'border': 1}),
                    'Risco Muito Baixo 🟣': workbook.add_format({'bg_color': '#B0E0E6', 'border': 1})
                }
                
                # Configurar largura das colunas
                worksheet.set_column('A:A', 25)  # Dimensão
                worksheet.set_column('B:B', 10)  # Média
                worksheet.set_column('C:C', 15)  # Nível de Risco
                worksheet.set_column('D:D', 40)  # Riscos Potenciais
                worksheet.set_column('E:E', 50)  # Sugestões de Ações
                worksheet.set_column('F:F', 40)  # Outras Soluções
                worksheet.set_column('G:G', 15)  # Responsável
                worksheet.set_column('H:H', 15)  # Prazo
                
                # Adicionar cabeçalhos formatados
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # Formatar células com quebra de linha
                for row_num, (_, row) in enumerate(df.itertuples(), 1):
                    # Acesso seguro aos atributos
                    try:
                        # Obter valores com verificação de existência
                        dimensao = getattr(row, "_1", "") if hasattr(row, "_1") else ""
                        media = getattr(row, "_2", 0) if hasattr(row, "_2") else 0
                        nivel_risco = getattr(row, "_3", "") if hasattr(row, "_3") else ""
                        riscos = getattr(row, "_4", "") if hasattr(row, "_4") else ""
                        sugestoes = getattr(row, "_5", "") if hasattr(row, "_5") else ""
                        outras = getattr(row, "_6", "") if hasattr(row, "_6") else ""
                        responsavel = getattr(row, "_7", "") if hasattr(row, "_7") else ""
                        prazo = getattr(row, "_8", "") if hasattr(row, "_8") else ""
                        
                        worksheet.write(row_num, 0, dimensao, wrap_format)  # Dimensão
                        worksheet.write(row_num, 1, media, wrap_format)  # Média
                        
                        # Aplicar formatos de nível de risco
                        if nivel_risco in risco_formats:
                            worksheet.write(row_num, 2, nivel_risco, risco_formats[nivel_risco])
                        else:
                            # Verificar correspondência parcial
                            formato_encontrado = False
                            for key, formato in risco_formats.items():
                                if key.split()[0] in nivel_risco:
                                    worksheet.write(row_num, 2, nivel_risco, formato)
                                    formato_encontrado = True
                                    break
                            
                            if not formato_encontrado:
                                worksheet.write(row_num, 2, nivel_risco, wrap_format)
                        
                        worksheet.write(row_num, 3, riscos, wrap_format)  # Riscos Potenciais
                        worksheet.write(row_num, 4, sugestoes, wrap_format)  # Sugestões
                        worksheet.write(row_num, 5, outras, wrap_format)  # Outras Soluções
                        worksheet.write(row_num, 6, responsavel, wrap_format)  # Responsável
                        
                        # Tratamento especial para datas
                        if isinstance(prazo, (datetime, pd.Timestamp)):
                            prazo_str = prazo.strftime("%d/%m/%Y")
                            worksheet.write(row_num, 7, prazo_str, wrap_format)
                        else:
                            # Tentar converter string para data, se não for None
                            if prazo:
                                prazo_str = padronizar_formato_data(prazo) or prazo
                                worksheet.write(row_num, 7, prazo_str, wrap_format)
                            else:
                                worksheet.write(row_num, 7, "", wrap_format)
                    except Exception as e:
                        print(f"Erro ao formatar linha {row_num}: {str(e)}")
                        # Continue com próxima linha mesmo se houver erro
                
                # Ajustar altura das linhas para acomodar conteúdo
                for i in range(1, len(df) + 1):
                    worksheet.set_row(i, 80)  # Altura em pontos
                
                # Adicionar filtros
                worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
                
                # Congelar a primeira linha
                worksheet.freeze_panes(1, 0)
                
                # Adicionar aba de instruções
                worksheet_info = workbook.add_worksheet('Instruções')
                title_format = workbook.add_format({
                    'bold': True,
                    'font_size': 14,
                    'align': 'center',
                    'valign': 'vcenter',
                    'fg_color': '#5A713D',
                    'font_color': 'white',
                    'border': 1
                })
                
                subtitle_format = workbook.add_format({
                    'bold': True,
                    'font_size': 12,
                    'fg_color': '#E8EDDF'
                })
                
                worksheet_info.set_column('A:A', 15)
                worksheet_info.set_column('B:B', 60)
                
                # Título
                worksheet_info.merge_range('A1:B1', 'PLANO DE AÇÃO - HSE-IT: INSTRUÇÕES', title_format)
                
                # Conteúdo
                linha = 3
                worksheet_info.write(linha, 0, 'Objetivo:', subtitle_format)
                linha += 1
                worksheet_info.write(linha, 0, '')
                worksheet_info.write(linha, 1, 'Este plano de ação contém sugestões de medidas baseadas nos resultados da avaliação HSE-IT. As sugestões devem ser adaptadas à realidade da sua organização e servir como ponto de partida para discussões com a gestão e equipe de saúde e segurança.')
                
                linha += 2
                worksheet_info.write(linha, 0, 'Como utilizar:', subtitle_format)
                linha += 1
                instrucoes = [
                    '1. Revise os riscos potenciais para verificar se se aplicam à sua realidade organizacional.',
                    '2. Ajuste as sugestões de ações conforme necessário para adequá-las ao seu contexto.',
                    '3. Adicione soluções específicas no campo "Outras Soluções".',
                    '4. Defina responsáveis e prazos para cada ação.',
                    '5. Integre as ações ao Programa de Gerenciamento de Riscos (PGR) da empresa.',
                    '6. Monitore a implementação e eficácia das ações ao longo do tempo.'
                ]
                
                for instrucao in instrucoes:
                    worksheet_info.write(linha, 1, instrucao)
                    linha += 1
                
                linha += 1
                worksheet_info.write(linha, 0, 'Níveis de risco:', subtitle_format)
                linha += 1
                
                niveis = [
                    {'nivel': 'Risco Muito Alto 🔴', 'descricao': 'Prioridade imediata. Requer intervenção em 1-3 meses.', 'formato': risco_formats['Risco Muito Alto 🔴']},
                    {'nivel': 'Risco Alto 🟠', 'descricao': 'Prioridade alta. Implementar ações em 3-6 meses.', 'formato': risco_formats['Risco Alto 🟠']},
                    {'nivel': 'Risco Moderado 🟡', 'descricao': 'Prioridade média. Implementar ações em 6-12 meses.', 'formato': risco_formats['Risco Moderado 🟡']},
                    {'nivel': 'Risco Baixo 🟢', 'descricao': 'Prioridade baixa. Implementar ações em 12-18 meses.', 'formato': risco_formats['Risco Baixo 🟢']},
                    {'nivel': 'Risco Muito Baixo 🟣', 'descricao': 'Manutenção. Manter boas práticas e revisar anualmente.', 'formato': risco_formats['Risco Muito Baixo 🟣']}
                ]
                
                for n in niveis:
                    worksheet_info.write(linha, 0, n['nivel'], n['formato'])
                    worksheet_info.write(linha, 1, n['descricao'])
                    linha += 1
                
                linha += 1
                worksheet_info.write(linha, 0, 'Observações:', subtitle_format)
                linha += 1
                worksheet_info.write(linha, 1, 'O plano de ação deve ser revisado e aprovado pela equipe de Saúde e Segurança do Trabalho e pela gestão da organização.')
                linha += 1
                worksheet_info.write(linha, 1, 'As ações devem ser implementadas conforme a prioridade definida, mas também considerando a viabilidade prática e os recursos disponíveis.')
                linha += 1
                worksheet_info.write(linha, 1, 'Recomenda-se revisar periodicamente o progresso das ações e ajustar o plano conforme necessário.')
                
                # Adicionar aba com dados de diagnóstico - CORRIGIDO: Verificação segura
                df_resultados = st.session_state.get("df_resultados")
                if df_resultados is not None and not df_resultados.empty:
                    try:
                        df_resultados.to_excel(writer, sheet_name='Diagnóstico HSE-IT', index=False)
                        worksheet_diag = writer.sheets['Diagnóstico HSE-IT']
                        
                        # Formatar cabeçalhos
                        for col_num, value in enumerate(df_resultados.columns.values):
                            worksheet_diag.write(0, col_num, value, header_format)
                        
                        # Configurar largura das colunas
                        worksheet_diag.set_column('A:A', 25)  # Dimensão
                        worksheet_diag.set_column('B:B', 40)  # Descrição
                        worksheet_diag.set_column('C:C', 10)  # Média
                        worksheet_diag.set_column('D:D', 15)  # Risco
                        worksheet_diag.set_column('E:E', 15)  # Número de Respostas
                        
                        # Formatar células de risco
                        for row_num, row in enumerate(df_resultados.itertuples(), 1):
                            if hasattr(row, "Risco"):
                                nivel_risco = row.Risco
                                if nivel_risco in risco_formats:
                                    worksheet_diag.write(row_num, 3, nivel_risco, risco_formats[nivel_risco])
                    except Exception as e:
                        print(f"Erro ao adicionar aba de diagnóstico: {str(e)}")
                        # Não interromper por erro na aba de diagnóstico
            
            output.seek(0)
            return output
        
        except Exception as e:
            st.error(f"Erro ao gerar o arquivo Excel: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
            return None
    
    # Botão para exportar o plano em Excel
    col1, col2 = st.columns([1, 2])
    
    with col1:
        if st.button("Exportar para Excel", type="primary", use_container_width=True):
            # CORRIGIDO: Verificação segura do plano editável
            plano_a_exportar = st.session_state.get("plano_editavel")
            if plano_a_exportar is not None and not plano_a_exportar.empty:
                with st.spinner("Gerando Excel... Por favor, aguarde."):
                    data_atual = datetime.now().strftime("%d%m%Y")
                    excel_data = exportar_para_excel(plano_a_exportar)
                    
                    if excel_data:
                        st.success("Plano de ação gerado com sucesso!")
                        # Salvar no estado da sessão para o botão de download
                        st.session_state["excel_plano"] = excel_data
                        st.session_state["excel_plano_ready"] = True
                        st.balloons()  # Efeito visual para confirmar sucesso
            else:
                st.error("Não há plano de ação para exportar. Verifique os filtros selecionados.")
    
    with col2:
        # CORRIGIDO: Verificação segura do estado do Excel
        if st.session_state.get("excel_plano_ready", False) and st.session_state.get("excel_plano") is not None:
            data_geracao = datetime.now().strftime("%d%m%Y")
            st.download_button(
                label="Baixar Plano de Ação HSE-IT",
                data=st.session_state["excel_plano"],
                file_name=f"plano_acao_hse_it_{data_geracao}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help="Baixe o arquivo Excel contendo o plano de ação completo"
            )
        else:
            st.info("Clique em 'Exportar para Excel' para gerar o arquivo para download")
