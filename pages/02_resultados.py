import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np
from utils.processamento import classificar_risco
from utils.constantes import DESCRICOES_DIMENSOES

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
    
    /* Tabelas */
    .dataframe {
        border-collapse: collapse;
    }
    
    .dataframe th {
        background-color: var(--escutaris-verde);
        color: white;
        padding: 8px;
    }
    
    .dataframe td {
        padding: 8px;
        border: 1px solid #ddd;
    }
    
    .dataframe tr:nth-child(even) {
        background-color: #f2f2f2;
    }
    </style>
    """, unsafe_allow_html=True)

# Aplicar o estilo
aplicar_estilo_escutaris()

# Título da página
st.title("Resultados da Avaliação HSE-IT")

# Verificar se há dados para exibir - CORRIGIDO: Uso seguro de session_state
if "df_resultados" not in st.session_state or st.session_state.get("df_resultados") is None:
    st.warning("Nenhum dado carregado ainda. Por favor, faça upload de um arquivo na página 'Upload de Dados'.")
    st.stop()

# Recuperar dados da sessão de forma segura - CORRIGIDO
try:
    df_resultados = st.session_state.get("df_resultados")
    filtro_opcao = st.session_state.get("filtro_opcao", "Empresa Toda")
    filtro_valor = st.session_state.get("filtro_valor", "Geral")
    
    # Garantir que df_resultados não é None
    if df_resultados is None:
        st.warning("Resultados não disponíveis. Por favor, carregue seus dados na página 'Upload de Dados'.")
        st.stop()
except Exception as e:
    st.error(f"Erro ao recuperar dados da sessão: {str(e)}")
    st.info("Por favor, retorne à página de upload e carregue seus dados novamente.")
    st.stop()

# Barra lateral para filtros adicionais
with st.sidebar:
    st.header("Filtros e Opções")
    
    # Adicionar opções de visualização
    view_mode = st.radio(
        "Modo de Visualização",
        ["Dashboard Resumido", "Análise Detalhada", "Análise Demográfica"]
    )
    
    st.divider()
    
    # Adicionar filtros demográficos - CORRIGIDO: Verificação segura
    df = st.session_state.get("df")
    colunas_filtro = st.session_state.get("colunas_filtro")
    
    if df is not None and colunas_filtro is not None:
        st.subheader("Filtrar Resultados")
        
        # Criar seletores para cada coluna de filtro
        filtro_selecionado = st.selectbox(
            "Filtrar por:",
            ["Empresa Toda"] + [col for col in colunas_filtro if col != "Carimbo de data/hora"]
        )
        
        if filtro_selecionado != "Empresa Toda":
            valores_filtro = df[filtro_selecionado].dropna().unique()
            valor_selecionado = st.selectbox(f"Selecione {filtro_selecionado}:", valores_filtro)
            
            if st.button("Aplicar Filtro"):
                try:
                    # Atualizar dados com base no filtro
                    df_perguntas = st.session_state.get("df_perguntas")
                    colunas_perguntas = st.session_state.get("colunas_perguntas")
                    
                    if None in [df_perguntas, colunas_perguntas]:
                        st.error("Dados necessários para filtragem não estão disponíveis.")
                    else:
                        indices_filtrados = df.index[df[filtro_selecionado] == valor_selecionado].tolist()
                        df_perguntas_filtradas = df_perguntas.loc[indices_filtrados]
                        
                        # Importar função necessária
                        from utils.processamento import calcular_resultados_dimensoes
                        
                        # Calcular novos resultados
                        resultados_filtrados = calcular_resultados_dimensoes(
                            df[df[filtro_selecionado] == valor_selecionado],
                            df_perguntas_filtradas,
                            colunas_perguntas
                        )
                        
                        if resultados_filtrados:
                            df_resultados = pd.DataFrame(resultados_filtrados)
                            st.session_state["df_resultados_filtrados"] = df_resultados
                            st.session_state["filtro_aplicado"] = True
                            st.session_state["filtro_opcao"] = filtro_selecionado
                            st.session_state["filtro_valor"] = valor_selecionado
                            st.success(f"Filtro aplicado: {filtro_selecionado} = {valor_selecionado}")
                            st.experimental_rerun()
                except Exception as e:
                    st.error(f"Erro ao aplicar filtro: {str(e)}")
        
        # Botão para limpar filtros
        if st.session_state.get("filtro_aplicado", False):
            if st.button("Limpar Filtros"):
                if st.session_state.get("df_resultados_original") is not None:
                    st.session_state["df_resultados"] = st.session_state["df_resultados_original"].copy()
                    st.session_state["filtro_aplicado"] = False
                    st.session_state["filtro_opcao"] = "Empresa Toda"
                    st.session_state["filtro_valor"] = "Geral"
                    if "df_resultados_filtrados" in st.session_state:
                        del st.session_state["df_resultados_filtrados"]
                    st.experimental_rerun()

# Se for a primeira vez carregando a página, salvar os resultados originais
if "df_resultados_original" not in st.session_state:
    st.session_state["df_resultados_original"] = df_resultados.copy()

# Se houver resultados filtrados, usar esses
if st.session_state.get("filtro_aplicado", False) and st.session_state.get("df_resultados_filtrados") is not None:
    df_resultados = st.session_state["df_resultados_filtrados"]
    filtro_opcao = st.session_state.get("filtro_opcao", "Empresa Toda")
    filtro_valor = st.session_state.get("filtro_valor", "Geral")

# Função para criar o dashboard principal
def criar_dashboard(df_resultados, filtro_opcao, filtro_valor):
    st.markdown(f"## Dashboard de Riscos Psicossociais - {filtro_opcao}: {filtro_valor}")
    
    # Layout com 3 métricas principais
    col1, col2, col3 = st.columns(3)
    
    # Calcular métricas de forma segura
    try:
        media_geral = df_resultados["Média"].mean()
        risco_geral, cor_geral = classificar_risco(media_geral)
        
        if len(df_resultados) > 0:
            dimensao_mais_critica = df_resultados.loc[df_resultados["Média"].idxmin()]
            dimensao_melhor = df_resultados.loc[df_resultados["Média"].idxmax()]
        else:
            # Valores padrão se não houver dados
            dimensao_mais_critica = {"Dimensão": "N/A", "Média": 0}
            dimensao_melhor = {"Dimensão": "N/A", "Média": 0}
        
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
                value=dimensao_mais_critica.get("Dimensão", "N/A"),
                delta=f"{dimensao_mais_critica.get('Média', 0):.2f}",
                delta_color="off"
            )
            risco, cor = classificar_risco(dimensao_mais_critica.get("Média", 0))
            st.markdown(f"<div style='text-align: center; color: {cor};'><b>{risco}</b></div>", unsafe_allow_html=True)
        
        with col3:
            st.metric(
                label="Dimensão Melhor Avaliada",
                value=dimensao_melhor.get("Dimensão", "N/A"),
                delta=f"{dimensao_melhor.get('Média', 0):.2f}",
                delta_color="off"
            )
            risco, cor = classificar_risco(dimensao_melhor.get("Média", 0))
            st.markdown(f"<div style='text-align: center; color: {cor};'><b>{risco}</b></div>", unsafe_allow_html=True)
    except Exception as e:
        st.error(f"Erro ao calcular métricas: {str(e)}")
        st.info("Verifique os dados carregados e tente novamente.")
    
    # Criar gráfico de barras para visualização dos riscos
    try:
        fig = criar_grafico_barras(df_resultados)
        st.plotly_chart(fig, use_container_width=True)
        
        # Adicionar gráfico de radar para visão geral das dimensões
        st.subheader("Visão Geral das Dimensões")
        fig_radar = criar_grafico_radar(df_resultados)
        st.plotly_chart(fig_radar, use_container_width=True)
    except Exception as e:
        st.error(f"Erro ao criar gráficos: {str(e)}")
    
    return

# Função para criar gráfico de barras usando Plotly
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
        # Acesso seguro para a descrição da dimensão
        descricao = row.get("Descrição", "Sem descrição disponível")
        hover_texts.append(f"Dimensão: {row['Dimensão']}<br>" +
                          f"Média: {row['Média']:.2f}<br>" +
                          f"Classificação: {row['Risco']}<br>" +
                          f"Descrição: {descricao}")
    
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
        title=f"Classificação de Riscos Psicossociais",
        xaxis_title="Média (Escala 1-5)",
        yaxis_title="Dimensão",
        xaxis=dict(range=[0, 5]),
        height=500,
        margin=dict(l=20, r=20, t=50, b=80),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family="Helvetica Neue, Arial", color="#2E2F2F")
    )
    
    return fig

# Função para criar gráfico de radar
def criar_grafico_radar(df_resultados):
    # Preparar dados para o gráfico de radar
    categorias = df_resultados["Dimensão"].tolist()
    valores = df_resultados["Média"].tolist()
    
    # Criar gráfico de radar
    fig = go.Figure()
    
    fig.add_trace(go.Scatterpolar(
        r=valores,
        theta=categorias,
        fill='toself',
        name='Média',
        line=dict(color='#5A713D'),  # Cor da Escutaris
        fillcolor='rgba(90, 113, 61, 0.3)'  # Verde Escutaris com transparência
    ))
    
    # Adicionar linhas de referência para níveis de risco
    for nivel, valor in [(1, "Risco Muito Alto"), (2, "Risco Alto"), 
                         (3, "Risco Moderado"), (4, "Risco Baixo")]:
        fig.add_trace(go.Scatterpolar(
            r=[nivel] * len(categorias),
            theta=categorias,
            mode='lines',
            line=dict(color='gray', width=1, dash='dash'),
            name=valor,
            hoverinfo='skip'
        ))
    
    # Configurar layout
    fig.update_layout(
        polar=dict(
            radialaxis=dict(
                visible=True,
                range=[0, 5],
                nticks=5,
                tickfont=dict(size=10)
            ),
            angularaxis=dict(
                tickfont=dict(size=10)
            )
        ),
        showlegend=True,
        legend=dict(orientation='h', y=-0.1),
        height=500,
        margin=dict(l=40, r=40, t=20, b=40)
    )
    
    return fig

# Função para mostrar detalhes por dimensão
def mostrar_detalhes_dimensao(df_resultados):
    st.subheader("Detalhes por Dimensão")
    
    # Criar abas para cada dimensão
    dimensoes = df_resultados["Dimensão"].tolist()
    tabs = st.tabs(dimensoes)
    
    for i, dimensao in enumerate(dimensoes):
        with tabs[i]:
            try:
                row = df_resultados[df_resultados["Dimensão"] == dimensao].iloc[0]
                media = row["Média"]
                risco = row["Risco"]
                _, cor = classificar_risco(media)
                # Acesso seguro à descrição
                descricao = row.get("Descrição", DESCRICOES_DIMENSOES.get(dimensao, "Sem descrição disponível"))
                
                # Layout de duas colunas
                col1, col2 = st.columns([1, 2])
                
                with col1:
                    # Card com informações principais
                    st.markdown('<div class="card">', unsafe_allow_html=True)
                    st.markdown(f"### {dimensao}")
                    st.markdown(f"**Média**: {media:.2f}")
                    st.markdown(f"**Nível de Risco**: <span style='color:{cor};'>{risco}</span>", unsafe_allow_html=True)
                    
                    # CORREÇÃO: Substituir 'lightpurple' por '#B0A8E3' (roxo claro)
                    # Adicionar medidor visual
                    fig = go.Figure(go.Indicator(
                        mode = "gauge+number",
                        value = media,
                        domain = {'x': [0, 1], 'y': [0, 1]},
                        gauge = {
                            'axis': {'range': [0, 5], 'tickwidth': 1},
                            'bar': {'color': cor},
                            'steps': [
                                {'range': [0, 1], 'color': 'red'},
                                {'range': [1, 2], 'color': 'orange'},
                                {'range': [2, 3], 'color': 'yellow'},
                                {'range': [3, 4], 'color': 'lightgreen'},
                                {'range': [4, 5], 'color': '#B0A8E3'}  # CORRIGIDO: Usar código hex para roxo claro
                            ]
                        }
                    ))
                    
                    fig.update_layout(height=200, margin=dict(l=20, r=20, t=30, b=20))
                    st.plotly_chart(fig, use_container_width=True)
                    st.markdown('</div>', unsafe_allow_html=True)
                
                with col2:
                    # Card com descrição e questões
                    st.markdown('<div class="card">', unsafe_allow_html=True)
                    st.markdown("### Descrição")
                    st.markdown(descricao)
                    
                    # Acesso seguro às questões
                    if "Questões" in row:
                        st.markdown("### Questões relacionadas")
                        # Tratar diferentes formatos de questões
                        questoes = None
                        if isinstance(row["Questões"], str):
                            try:
                                questoes = eval(row["Questões"])
                            except Exception:
                                # Se falhar em eval, tentar outras abordagens
                                try:
                                    import ast
                                    questoes = ast.literal_eval(row["Questões"])
                                except Exception:
                                    # Se ainda falhar, usar a string como está
                                    questoes = [row["Questões"]]
                        else:
                            questoes = row["Questões"]
                            
                        # Garantir que questoes é iterável
                        if questoes:
                            for q in questoes:
                                # Tentar buscar o texto da questão
                                q_text = f"Questão {q}"
                                colunas_perguntas = st.session_state.get("colunas_perguntas", [])
                                for col in colunas_perguntas:
                                    if str(col).strip().startswith(str(q)):
                                        q_text = str(col).strip()
                                        break
                                st.markdown(f"- **{q}**: {q_text}")
                    st.markdown('</div>', unsafe_allow_html=True)
                
                # Sugestões de ação
                st.markdown("### Sugestões de Ação")
                
                # Acesso seguro às sugestões
                df_plano_acao = st.session_state.get("df_plano_acao")
                if df_plano_acao is not None:
                    # Verificar se "Dimensão" está no DataFrame
                    if "Dimensão" in df_plano_acao.columns:
                        sugestoes = df_plano_acao[
                            df_plano_acao["Dimensão"] == dimensao
                        ]["Sugestão de Ação"].tolist()
                        
                        if sugestoes:
                            for j, sugestao in enumerate(sugestoes):
                                st.markdown(f"{j+1}. {sugestao}")
                        else:
                            st.info("Não há sugestões específicas para esta dimensão.")
                    else:
                        st.info("O formato do plano de ação não é compatível.")
                else:
                    st.info("O plano de ação não está disponível. Visite a página 'Plano de Ação' para mais detalhes.")
            except Exception as e:
                st.error(f"Erro ao mostrar detalhes da dimensão {dimensao}: {str(e)}")

# Função para criar análise demográfica
def criar_analise_demografica(df, df_perguntas, colunas_perguntas):
    st.subheader("Análise por Características Demográficas")
    st.write("Explore os resultados por diferentes características demográficas como setor, cargo, gênero, etc.")
    
    # Obter colunas demográficas disponíveis
    demograficas = [col for col in st.session_state.get("colunas_filtro", []) if col != "Carimbo de data/hora"]
    
    if not demograficas:
        st.warning("Não foram encontradas colunas demográficas nos dados.")
        return
    
    # Seletor para a coluna demográfica
    col_demo = st.selectbox("Selecione uma característica para análise:", demograficas)
    
    # Criar função para análise demográfica
    def analisar_por_demografia(df, col_demo):
        resultados = []
        try:
            valores_unicos = df[col_demo].dropna().unique()
            
            for valor in valores_unicos:
                if pd.notna(valor):
                    df_filtrado = df[df[col_demo] == valor]
                    indices_filtrados = df.index[df[col_demo] == valor].tolist()
                    df_perguntas_filtradas = df_perguntas.loc[indices_filtrados]
                    
                    # Importar função necessária
                    from utils.processamento import calcular_resultados_dimensoes
                    
                    # Calcular resultados para este filtro
                    resultados_filtro = calcular_resultados_dimensoes(
                        df_filtrado, 
                        df_perguntas_filtradas, 
                        colunas_perguntas
                    )
                    
                    # Adicionar valor do filtro
                    for res in resultados_filtro:
                        res[col_demo] = valor
                        resultados.append(res)
        except Exception as e:
            st.error(f"Erro ao analisar dados por {col_demo}: {str(e)}")
            
        return pd.DataFrame(resultados) if resultados else None
    
    # Obter e mostrar resultados
    with st.spinner(f"Analisando dados por {col_demo}..."):
        df_demografico = analisar_por_demografia(df, col_demo)
    
    if df_demografico is not None and not df_demografico.empty:  # CORRIGIDO: Verificação de DataFrame vazio
        # Verificar a quantidade de valores únicos para escolher a melhor visualização
        valores_unicos = df_demografico[col_demo].unique()
        
        # Mostrar tabela pivotada
        st.subheader(f"Médias por Dimensão e {col_demo}")
        try:
            df_pivot = df_demografico.pivot(index="Dimensão", columns=col_demo, values="Média")
            
            # Adicionar heatmap
            st.write("Clique em uma célula para ver os detalhes:")
            
            # Aplicar formatação de cores
            def color_scale(val):
                if pd.isna(val):
                    return ""
                    
                if val <= 1:
                    return f'background-color: #FF6B6B; color: white'
                elif val <= 2:
                    return f'background-color: #FFA500'
                elif val <= 3:
                    return f'background-color: #FFFF00'
                elif val <= 4:
                    return f'background-color: #90EE90'
                else:
                    return f'background-color: #BB8FCE'
            
            # Aplicar o estilo
            df_styled = df_pivot.style.applymap(color_scale).format("{:.2f}")
            
            st.dataframe(df_styled)
            
            # Criar tabs para diferentes visualizações
            chart_tab, compare_tab = st.tabs(["Gráfico de Barras", "Comparação de Dimensões"])
            
            with chart_tab:
                try:
                    # Criar gráfico de barras agrupadas
                    fig = go.Figure()
                    
                    for col in df_pivot.columns:
                        fig.add_trace(go.Bar(
                            x=df_pivot.index,
                            y=df_pivot[col],
                            name=str(col),
                            text=[f"{v:.2f}" if not pd.isna(v) else "N/A" for v in df_pivot[col]],
                            textposition="auto"
                        ))
                    
                    fig.update_layout(
                        title=f"Comparação de Dimensões por {col_demo}",
                        xaxis_title="Dimensão",
                        yaxis_title="Média",
                        yaxis=dict(range=[0, 5]),
                        barmode="group",
                        height=500,
                        plot_bgcolor='rgba(0,0,0,0)',
                        paper_bgcolor='rgba(0,0,0,0)',
                        font=dict(family="Helvetica Neue, Arial", color="#2E2F2F")
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)
                except Exception as e:
                    st.error(f"Erro ao criar gráfico de barras: {str(e)}")
            
            with compare_tab:
                try:
                    # Criar gráfico de radar para comparação
                    st.subheader(f"Comparar: Criar gráfico de radar para comparação")
                    st.subheader(f"Comparação por {col_demo}")
                    
                    # Seleção de grupos para comparar
                    if len(valores_unicos) > 6:
                        grupos_para_comparar = st.multiselect(
                            f"Selecione até 6 grupos de {col_demo} para comparar:",
                            valores_unicos,
                            default=list(valores_unicos)[:3] if len(valores_unicos) > 3 else list(valores_unicos)
                        )
                    else:
                        grupos_para_comparar = valores_unicos
                    
                    if grupos_para_comparar:
                        # Filtrar o dataframe pivotado
                        df_radar = df_pivot[grupos_para_comparar]
                        
                        # Criar gráfico de radar
                        fig = go.Figure()
                        
                        for grupo in grupos_para_comparar:
                            # Tratar valores NaN
                            valores = df_radar[grupo].values
                            valores = np.nan_to_num(valores, nan=0)  # Substituir NaN por 0
                            
                            fig.add_trace(go.Scatterpolar(
                                r=valores,
                                theta=df_radar.index,
                                fill='toself',
                                name=str(grupo)
                            ))
                        
                        fig.update_layout(
                            polar=dict(
                                radialaxis=dict(
                                    visible=True,
                                    range=[0, 5]
                                )
                            ),
                            showlegend=True,
                            height=600
                        )
                        
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info(f"Selecione pelo menos um grupo de {col_demo} para visualizar.")
                except Exception as e:
                    st.error(f"Erro ao criar gráfico de radar: {str(e)}")
            
            # Adicionar análise de diferenças significativas
            st.subheader("Diferenças Significativas")
            
            try:
                # Calcular as maiores diferenças entre grupos
                df_diff = df_pivot.max(axis=1) - df_pivot.min(axis=1)
                df_diff = df_diff.sort_values(ascending=False)
                
                # Mostrar as dimensões com maiores diferenças
                st.write("Dimensões com maiores variações entre grupos:")
                
                for dimensao, diff in df_diff.items():
                    if pd.isna(diff):
                        continue
                        
                    max_grupo = df_pivot.loc[dimensao].idxmax()
                    min_grupo = df_pivot.loc[dimensao].idxmin()
                    max_valor = df_pivot.loc[dimensao, max_grupo]
                    min_valor = df_pivot.loc[dimensao, min_grupo]
                    
                    if diff > 0.5:  # Apenas mostrar diferenças relevantes
                        st.markdown(f"**{dimensao}**: Diferença de {diff:.2f} pontos")
                        st.markdown(f"- Maior média: {max_valor:.2f} ({max_grupo})")
                        st.markdown(f"- Menor média: {min_valor:.2f} ({min_grupo})")
            except Exception as e:
                st.error(f"Erro ao calcular diferenças significativas: {str(e)}")
        except Exception as e:
            st.error(f"Erro ao criar tabela pivotada: {str(e)}")
            st.dataframe(df_demografico)
    else:
        st.warning(f"Não foi possível gerar análise para {col_demo}. Verifique se há dados suficientes.")

# Renderizar a página com base no modo de visualização selecionado
try:
    if view_mode == "Dashboard Resumido":
        criar_dashboard(df_resultados, filtro_opcao, filtro_valor)
        
        # Adicionar botão para ver detalhes
        if st.button("Ver Análise Detalhada"):
            st.session_state["view_mode"] = "Análise Detalhada"
            st.experimental_rerun()

    elif view_mode == "Análise Detalhada":
        # Mostrar dashboard resumido
        criar_dashboard(df_resultados, filtro_opcao, filtro_valor)
        
        # Mostrar detalhes por dimensão
        mostrar_detalhes_dimensao(df_resultados)
        
        # Botão para voltar ao dashboard
        if st.button("Voltar ao Dashboard Resumido"):
            st.session_state["view_mode"] = "Dashboard Resumido"
            st.experimental_rerun()

    elif view_mode == "Análise Demográfica":
        # CORREÇÃO: Verificação adequada dos dados demográficos
        df = st.session_state.get("df")
        df_perguntas = st.session_state.get("df_perguntas")
        colunas_perguntas = st.session_state.get("colunas_perguntas")
        
        if None not in [df, df_perguntas, colunas_perguntas]:
            criar_analise_demografica(df, df_perguntas, colunas_perguntas)
        else:
            st.error("Dados necessários para análise demográfica não estão disponíveis.")
        
        # Botão para voltar ao dashboard
        if st.button("Voltar ao Dashboard Resumido"):
            st.session_state["view_mode"] = "Dashboard Resumido"
            st.experimental_rerun()
except Exception as e:
    st.error(f"Erro ao renderizar o modo de visualização '{view_mode}': {str(e)}")
    import traceback
    st.code(traceback.format_exc())

# Salvar o modo de visualização na sessão
st.session_state["view_mode"] = view_mode

# Exibir tabela de resultados detalhados (opcional, expandível)
with st.expander("Tabela de Resultados Detalhados", expanded=False):
    st.dataframe(df_resultados)
    
    # Botão para download dos resultados
    try:
        csv = df_resultados.to_csv(index=False)
        st.download_button(
            label="Baixar Resultados CSV",
            data=csv,
            file_name=f"resultados_{filtro_opcao}_{filtro_valor}.csv",
            mime="text/csv",
            help="Baixe os resultados para análise em outras ferramentas"
        )
    except Exception as e:
        st.error(f"Erro ao preparar download: {str(e)}")
