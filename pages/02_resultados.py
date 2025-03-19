import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from utils.processamento import classificar_risco
from utils.constantes import DESCRICOES_DIMENSOES

# Título da página
st.title("Resultados da Avaliação HSE-IT")

# Verificar se há dados para exibir
if "df_resultados" not in st.session_state or st.session_state.df_resultados is None:
    st.warning("Nenhum dado carregado ainda. Por favor, faça upload de um arquivo na página 'Upload de Dados'.")
    st.stop()

# Recuperar dados da sessão
df_resultados = st.session_state.df_resultados
filtro_opcao = st.session_state.get("filtro_opcao", "Empresa Toda")
filtro_valor = st.session_state.get("filtro_valor", "Geral")

# Função para criar o dashboard
def criar_dashboard(df_resultados, filtro_opcao, filtro_valor):
    st.markdown(f"## Dashboard de Riscos Psicossociais - {filtro_opcao}: {filtro_valor}")
    
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
            if "Questões" in row:
                st.markdown(f"**Questões HSE-IT relacionadas**: {row['Questões']}")
            
            # Adicionar sugestões de melhoria para dimensões com maior risco
            if "Alto" in risco or "Muito Alto" in risco:
                st.markdown("**Sugestões de Melhoria Prioritárias:**")
                if "df_plano_acao" in st.session_state:
                    sugestoes = st.session_state.df_plano_acao[
                        st.session_state.df_plano_acao["Dimensão"] == dimensao
                    ]["Sugestão de Ação"].head(3).tolist()
                    
                    for sugestao in sugestoes:
                        st.markdown(f"- {sugestao}")
                    
                    st.markdown("*Confira mais sugestões na página 'Plano de Ação'*")

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
        title=f"Classificação de Riscos Psicossociais",
        xaxis_title="Média (Escala 1-5)",
        yaxis_title="Dimensão",
        xaxis=dict(range=[0, 5]),
        height=500,
        margin=dict(l=20, r=20, t=50, b=80)
    )
    
    return fig

# Chamar a função do dashboard para exibir os resultados
criar_dashboard(df_resultados, filtro_opcao, filtro_valor)

# Exibir tabela de resultados detalhados
st.subheader("Tabela de Resultados Detalhados")
st.dataframe(df_resultados)

# Adicionar botão para download dos resultados
csv = df_resultados.to_csv(index=False)
st.download_button(
    label="Baixar Resultados CSV",
    data=csv,
    file_name=f"resultados_{filtro_opcao}_{filtro_valor}.csv",
    mime="text/csv",
)
