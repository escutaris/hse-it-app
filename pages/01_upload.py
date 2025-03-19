import streamlit as st
import pandas as pd
import numpy as np
import io
from utils.processamento import carregar_dados, calcular_resultados_dimensoes
from utils.constantes import DIMENSOES_HSE, DESCRICOES_DIMENSOES

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
    
    /* Botões */
    .stButton>button {
        background-color: var(--escutaris-verde) !important;
        color: white !important;
        border-radius: 5px !important;
    }
    
    /* Cards para conteúdo */
    .card {
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        background-color: white;
        margin-bottom: 20px;
    }
    
    /* Upload box */
    .uploadbox {
        border: 2px dashed var(--escutaris-verde);
        border-radius: 10px;
        padding: 20px;
        text-align: center;
        margin-bottom: 20px;
    }
    
    /* Success message */
    .success {
        padding: 10px;
        background-color: #d4edda;
        color: #155724;
        border-radius: 5px;
        margin-bottom: 15px;
    }
    
    /* Error message */
    .error {
        padding: 10px;
        background-color: #f8d7da;
        color: #721c24;
        border-radius: 5px;
        margin-bottom: 15px;
    }
    </style>
    """, unsafe_allow_html=True)

# Aplicar o estilo
aplicar_estilo_escutaris()

# Título da página
st.title("Upload de Dados - HSE-IT")
st.markdown("### Carregue seu arquivo de respostas do questionário HSE-IT para análise")

# Descrição da página com instruções
with st.expander("ℹ️ Instruções para upload", expanded=False):
    st.markdown("""
    ### Formatos suportados:
    - **Excel (.xlsx)**: Planilha Excel com respostas individuais
    - **CSV (.csv)**: Arquivo CSV exportado do Google Forms ou outro sistema
    
    ### Estrutura necessária:
    - As primeiras colunas devem conter informações demográficas (Setor, Cargo, etc.)
    - As colunas das perguntas devem estar numeradas (começando com números)
    - As respostas devem ser valores numéricos de 1 a 5
    
    ### Como obter os dados:
    1. Aplique o questionário HSE-IT via formulário Google Forms ou planilha Excel
    2. Exporte as respostas em formato CSV ou Excel
    3. Faça o upload do arquivo nesta página
    
    Para um modelo pronto, acesse a página "Informações" e baixe o template HSE-IT.
    """)

# Container para o upload
with st.container():
    st.markdown('<div class="uploadbox">', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Escolha um arquivo Excel ou CSV", type=["xlsx", "csv"])
    st.markdown('</div>', unsafe_allow_html=True)

# Processar o arquivo quando enviado
if uploaded_file is not None:
    try:
        # Informação de progresso
        with st.spinner("Carregando e processando dados..."):
            # Carregar dados
            df, df_perguntas, colunas_filtro, colunas_perguntas = carregar_dados(uploaded_file)
            
            # Verificar se os dados foram carregados corretamente
            if df is None or len(df) == 0:
                st.error("Não foi possível carregar dados do arquivo. Verifique o formato e tente novamente.")
                st.stop()
            
            # Contagens e estatísticas básicas
            total_respostas = len(df)
            colunas_demograficas = [col for col in colunas_filtro if col != "Carimbo de data/hora"]
            
            # Avaliar qualidade dos dados
            dados_validos = True
            mensagens_validacao = []
            
            # Verificar se existem respostas
            if total_respostas == 0:
                dados_validos = False
                mensagens_validacao.append("O arquivo não contém nenhuma resposta.")
            
            # Verificar se as perguntas do HSE-IT estão presentes
            if len(colunas_perguntas) < 30:  # Esperamos pelo menos 30 das 35 perguntas
                dados_validos = False
                mensagens_validacao.append(f"Foram encontradas apenas {len(colunas_perguntas)} perguntas. O HSE-IT contém 35 perguntas.")
            
            # Verificar valores ausentes nas perguntas
            valores_ausentes = df_perguntas.isna().sum().sum()
            percentual_ausente = (valores_ausentes / (total_respostas * len(colunas_perguntas))) * 100
            if percentual_ausente > 20:  # Se mais de 20% dos dados estão ausentes
                dados_validos = False
                mensagens_validacao.append(f"Há {percentual_ausente:.1f}% de respostas ausentes, o que pode comprometer a análise.")
            
            # Calcular resultados preliminares
            if dados_validos:
                # Calcular resultados para validação
                resultados = calcular_resultados_dimensoes(df, df_perguntas, colunas_perguntas)
                df_resultados = pd.DataFrame(resultados)
                
                # Gerar plano de ação baseado nos resultados
                from utils.processamento import gerar_sugestoes_acoes
                df_plano_acao = gerar_sugestoes_acoes(df_resultados)
                
                # Armazenar no session_state para acesso em outras páginas
                st.session_state.df = df
                st.session_state.df_perguntas = df_perguntas
                st.session_state.colunas_filtro = colunas_filtro
                st.session_state.colunas_perguntas = colunas_perguntas
                st.session_state.df_resultados = df_resultados
                st.session_state.df_plano_acao = df_plano_acao
                st.session_state.filtro_opcao = "Empresa Toda"
                st.session_state.filtro_valor = "Geral"
                st.session_state.DIMENSOES_HSE = DIMENSOES_HSE
                st.session_state.DESCRICOES_DIMENSOES = DESCRICOES_DIMENSOES
                
                # Mostrar informações do arquivo
                st.markdown('<div class="success">', unsafe_allow_html=True)
                st.success(f"✅ Arquivo carregado com sucesso! {total_respostas} respostas processadas.")
                st.markdown('</div>', unsafe_allow_html=True)
                
                # Exibir resumo dos dados
                col1, col2 = st.columns(2)
                
                with col1:
                    st.subheader("Resumo dos Dados")
                    st.markdown(f"**Total de respostas:** {total_respostas}")
                    st.markdown(f"**Dados demográficos:** {', '.join(colunas_demograficas)}")
                    st.markdown(f"**Perguntas identificadas:** {len(colunas_perguntas)}/35")
                
                with col2:
                    st.subheader("Dimensões Avaliadas")
                    # Mostrar resumo das dimensões encontradas
                    for dimensao, _ in DIMENSOES_HSE.items():
                        resultado_dimensao = df_resultados[df_resultados["Dimensão"] == dimensao]
                        if not resultado_dimensao.empty:
                            media = resultado_dimensao.iloc[0]["Média"]
                            risco = resultado_dimensao.iloc[0]["Risco"]
                            st.markdown(f"- **{dimensao}**: {media:.2f} - {risco}")
                
                # Mostrar próximos passos
                st.subheader("Próximos Passos")
                st.markdown("""
                Seus dados foram processados e estão prontos para análise. Navegue para as outras páginas para explorar os resultados:
                
                1. **Resultados**: Visualize gráficos detalhados das dimensões avaliadas
                2. **Plano de Ação**: Veja sugestões de ações para melhorar cada dimensão
                3. **Relatórios**: Gere relatórios em Excel ou PDF para compartilhar com sua equipe
                """)
                
                # Botões para navegação direta
                col1, col2, col3 = st.columns(3)
                with col1:
                    if st.button("Ver Resultados Detalhados"):
                        st.switch_page("pages/02_resultados.py")
                with col2:
                    if st.button("Ir para Plano de Ação"):
                        st.switch_page("pages/03_plano_acao.py")
                with col3:
                    if st.button("Gerar Relatórios"):
                        st.switch_page("pages/04_relatorios.py")
            else:
                # Mostrar mensagens de validação se houver problemas
                st.markdown('<div class="error">', unsafe_allow_html=True)
                st.error("⚠️ Foram encontrados problemas com o arquivo enviado:")
                for mensagem in mensagens_validacao:
                    st.markdown(f"- {mensagem}")
                st.markdown('</div>', unsafe_allow_html=True)
                
                # Sugestões para resolução dos problemas
                st.subheader("Sugestões para resolução")
                st.markdown("""
                - Verifique se o arquivo está no formato correto (CSV ou Excel)
                - Confirme se as perguntas estão identificadas com números no início
                - Certifique-se de que as respostas estão em formato numérico (1 a 5)
                - Verifique se há muitos valores ausentes nas respostas
                - Baixe o template na página de Informações e utilize-o para formatar seus dados
                """)
    except Exception as e:
        # Tratamento de erro genérico
        st.markdown('<div class="error">', unsafe_allow_html=True)
        st.error(f"❌ Ocorreu um erro ao processar o arquivo: {str(e)}")
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Informações de depuração (opcional - pode ser removido em produção)
        with st.expander("Detalhes técnicos do erro", expanded=False):
            st.write(e)
            import traceback
            st.code(traceback.format_exc())
else:
    # Instruções quando não há arquivo carregado
    st.info("Carregue um arquivo para começar a análise. Se você não tem um arquivo com dados, visite a página 'Informações' para baixar um template.")
    
    # Exibir dados de exemplo (opcional)
    with st.expander("Mostrar exemplo dos dados esperados", expanded=False):
        # Criar dados de exemplo para visualização
        colunas_exemplo = ["Setor", "Cargo", "Tempo_Empresa", "1. Pergunta", "2. Pergunta", "3. Pergunta", "4. Pergunta", "5. Pergunta"]
        dados_exemplo = [
            ["TI", "Analista", "1-3 anos", 4, 5, 3, 4, 2],
            ["RH", "Gerente", "3-5 anos", 5, 4, 2, 3, 3],
            ["Financeiro", "Assistente", "<1 ano", 3, 4, 4, 5, 2]
        ]
        df_exemplo = pd.DataFrame(dados_exemplo, columns=colunas_exemplo)
        st.dataframe(df_exemplo)
        
        st.markdown("""
        **Observações sobre os dados de exemplo:**
        - As primeiras colunas contêm informações demográficas (Setor, Cargo, Tempo_Empresa)
        - As perguntas do HSE-IT são identificadas por números (1, 2, 3...)
        - As respostas são valores numéricos entre 1 e 5
        """)
