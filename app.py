import streamlit as st
from utils.autenticacao import verificar_senha

# Configuração da página
st.set_page_config(
    page_title="Avaliação de Fatores Psicossociais - HSE-IT",
    page_icon="📊",
    layout="wide"
)

# Carregar CSS personalizado
def load_css():
    with open('styles/main.css', 'r') as f:
        st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

try:
    load_css()
except Exception as e:
    st.warning("Não foi possível carregar o CSS personalizado. Usando estilo padrão.")

# Cabeçalho personalizado com as cores da Escutaris
def exibir_cabecalho():
    st.markdown("""
    <div style="background-color: #F5F7F0; padding: 1.5rem; border-radius: 10px; margin-bottom: 1.5rem; border-left: 8px solid #5A713D;">
        <h1 style="color: #5A713D; margin-bottom: 0.5rem;">HSE-IT Analytics</h1>
        <p style="color: #BC582C; font-size: 1.2rem;">Avaliação de Fatores Psicossociais no Trabalho</p>
        <p style="font-size: 0.9rem;">Transformando ambientes de trabalho através de soluções personalizadas em saúde mental</p>
    </div>
    """, unsafe_allow_html=True)

# Autenticação
if verificar_senha():
    # Exibir cabeçalho
    exibir_cabecalho()
    
    # Conteúdo da página inicial
    st.write("""
    ### Bem-vindo à Plataforma HSE-IT Analytics
    
    Esta plataforma permite analisar os fatores psicossociais no ambiente de trabalho usando a 
    metodologia HSE-IT (Health and Safety Executive - Indicator Tool).
    
    Navegue através das páginas disponíveis no menu lateral para:
    """)
    
    # Cartões para as principais funcionalidades
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div style="background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); height: 200px;">
            <h3 style="color: #5A713D;">📤 Upload de Dados</h3>
            <p>Carregue os resultados do questionário HSE-IT para análise.</p>
            <p>Suporta arquivos Excel (.xlsx) e CSV, permitindo filtrar por diferentes segmentos.</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        <div style="background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-top: 20px; height: 200px;">
            <h3 style="color: #5A713D;">📋 Plano de Ação</h3>
            <p>Crie e personalize um plano de ação baseado nos resultados.</p>
            <p>Adicione responsáveis, prazos e acompanhe o status de cada ação.</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div style="background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); height: 200px;">
            <h3 style="color: #5A713D;">📊 Resultados</h3>
            <p>Visualize os resultados da avaliação em um dashboard interativo.</p>
            <p>Identifique as dimensões que necessitam de maior atenção.</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        <div style="background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-top: 20px; height: 200px;">
            <h3 style="color: #5A713D;">📁 Relatórios</h3>
            <p>Gere relatórios em Excel ou PDF para compartilhamento.</p>
            <p>Exporte o plano de ação para implementação na sua organização.</p>
        </div>
        """, unsafe_allow_html=True)
    
    # Rodapé com informações adicionais
    st.markdown("""
    <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd;">
        <p style="text-align: center; color: #666;">
            HSE-IT Analytics by <a href="https://escutaris.com.br" style="color: #5A713D;">Escutaris</a> | © 2023
        </p>
    </div>
    """, unsafe_allow_html=True)

else:
    st.stop()  # Não mostrar nada abaixo deste ponto se a autenticação falhar
