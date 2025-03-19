import streamlit as st
import os

def aplicar_estilo_escutaris():
    """Aplica o estilo visual da Escutaris ao aplicativo Streamlit."""
    
    # Tenta carregar o CSS do arquivo existente styles/main.css
    try:
        # Verificar o caminho do arquivo CSS
        script_dir = os.path.dirname(os.path.abspath(__file__))
        css_path = os.path.join(os.path.dirname(script_dir), "styles", "main.css")
        
        # Verificar se o arquivo existe
        if os.path.exists(css_path):
            with open(css_path, "r", encoding="utf-8") as f:
                st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
                return
    except Exception as e:
        print(f"Erro ao carregar o arquivo CSS: {e}")
    
    # Se não conseguir carregar do arquivo, use a versão incorporada
    st.markdown("""
    <style>
    /* Cores e estilos para a identidade visual da Escutaris */
    :root {
        --primary-color: #5A713D;       /* Verde oliva da Escutaris */
        --secondary-color: #BC582C;     /* Laranja terracota para destaque */
        --accent-color: #DEB887;        /* Tom bege para elementos secundários */
        --text-color: #333333;          /* Cinza escuro para textos */
        --background-color: #F5F7F0;    /* Fundo levemente esverdeado */
        --card-background: #FFFFFF;     /* Branco para cartões */
        --success-color: #5A713D;       /* Verde da marca para sucesso */
        --warning-color: #DEB887;       /* Bege para alertas */
        --danger-color: #BC582C;        /* Terracota para erros/perigos */
    }

    /* Estilo para o corpo do aplicativo */
    .reportview-container {
        background-color: var(--background-color);
        color: var(--text-color);
    }

    /* Estilo para cabeçalhos */
    h1, h2, h3 {
        color: var(--primary-color);
        font-weight: 600;
    }

    /* Estilo para cartões/containers */
    .stCard, .card, .info-card, .report-card {
        background-color: var(--card-background);
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        margin-bottom: 20px;
    }

    /* Estilo para botões */
    .stButton > button {
        background-color: var(--primary-color);
        color: white;
        border-radius: 5px;
        border: none;
        padding: 0.5rem 1rem;
        transition: all 0.3s ease;
    }

    .stButton > button:hover {
        opacity: 0.9;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }

    /* Estilo para métricas */
    div.css-1xarl3l.e16fv1kl1 {
        background-color: var(--card-background);
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }

    /* Personalização da barra lateral */
    .sidebar .sidebar-content {
        background-color: var(--card-background);
    }

    /* Estilo para tabelas */
    .dataframe {
        border-collapse: collapse;
        width: 100%;
    }
    .dataframe th {
        background-color: var(--primary-color);
        color: white;
        text-align: left;
        padding: 8px;
    }
    .dataframe td {
        padding: 8px;
        border-bottom: 1px solid #ddd;
    }
    .dataframe tr:nth-child(even) {
        background-color: rgba(90, 113, 61, 0.1);
    }

    /* Estilo para expansores */
    .streamlit-expanderHeader {
        background-color: rgba(90, 113, 61, 0.1);
        border-radius: 5px;
    }
    .streamlit-expanderContent {
        border-left: 1px solid var(--primary-color);
        padding-left: 15px;
    }

    /* Estilo para barra de progresso */
    .stProgress > div > div > div > div {
        background-color: var(--primary-color);
    }
    </style>
    """, unsafe_allow_html=True)
