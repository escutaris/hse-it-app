import streamlit as st
import pandas as pd
from utils.processamento import carregar_dados, calcular_resultados_dimensoes
from utils.constantes import QUESTOES_INVERTIDAS, DIMENSOES_HSE
from utils.sugestoes import gerar_sugestoes_acoes

st.title("Upload de Dados - HSE-IT")
st.write("Faça upload do arquivo Excel contendo os resultados do questionário HSE-IT.")

# Adiciona uma explicação sobre o formato esperado do arquivo
with st.expander("Informações sobre o formato do arquivo"):
    st.write("""
    O arquivo deve conter:
    - Colunas de filtro (Setor, Cargo, etc.) nas primeiras 7 colunas
    - Colunas com as perguntas numeradas do HSE-IT (começando com números seguidos de ponto)
    - As respostas devem ser valores numéricos de 1 a 5
    
    Você pode baixar um template na página 'Informações HSE-IT'.
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
            
            # Armazenar no session_state para acesso em outras páginas
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

if uploaded_file is not None and "df_resultados" in st.session_state and st.session_state.df_resultados is not None:
    st.success("Dados processados com sucesso. Navegue para a página 'Resultados' para visualizar a análise.")
    
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
