import streamlit as st
import pandas as pd
import io
from fpdf import FPDF
from datetime import datetime
import plotly.graph_objects as go
from utils.processamento import classificar_risco

# Título da página
st.title("Relatórios - HSE-IT")

# Verificar se há dados para exibir
if "df_resultados" not in st.session_state or st.session_state.df_resultados is None:
    st.warning("Nenhum dado carregado ainda. Por favor, faça upload de um arquivo na página 'Upload de Dados'.")
    st.stop()

# Recuperar dados da sessão
df = st.session_state.df
df_perguntas = st.session_state.df_perguntas
colunas_filtro = st.session_state.colunas_filtro
colunas_perguntas = st.session_state.colunas_perguntas
df_resultados = st.session_state.df_resultados
df_plano_acao = st.session_state.df_plano_acao
filtro_opcao = st.session_state.filtro_opcao
filtro_valor = st.session_state.filtro_valor

st.header("Download de Relatórios")
st.write("Escolha abaixo o tipo de relatório que deseja gerar.")

# Layout de duas colunas para os tipos de relatório
col1, col2 = st.columns(2)

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
        pdf.cell(0,
