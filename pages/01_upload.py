import streamlit as st
import pandas as pd
import numpy as np
import io
import os
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
    
    /* Progress indicator */
    .progress-indicator {
        display: flex;
        justify-content: space-between;
        margin-bottom: 20px;
    }
    
    .progress-step {
        display: flex;
        flex-direction: column;
        align-items: center;
        width: 20%;
    }
    
    .step-number {
        width: 30px;
        height: 30px;
        border-radius: 50%;
        background-color: #e0e0e0;
        display: flex;
        justify-content: center;
        align-items: center;
        margin-bottom: 5px;
    }
    
    .step-number.active {
        background-color: var(--escutaris-verde);
        color: white;
    }
    
    .step-line {
        height: 3px;
        background-color: #e0e0e0;
        flex-grow: 1;
    }
    
    /* Informações adicionais */
    .info-box {
        background-color: #f8f9fa;
        border-left: 4px solid var(--escutaris-verde);
        padding: 15px;
        margin: 15px 0;
        border-radius: 5px;
    }
    
    /* Lista de verificação */
    .checklist {
        list-style-type: none;
        padding-left: 0;
    }
    
    .checklist li {
        margin-bottom: 10px;
        padding-left: 30px;
        position: relative;
    }
    
    .checklist li:before {
        content: '✓';
        position: absolute;
        left: 0;
        color: var(--escutaris-verde);
        font-weight: bold;
    }
    
    .checklist-item-error:before {
        content: '✗' !important;
        color: #dc3545 !important;
    }
    
    .checklist-item-warning:before {
        content: '⚠' !important;
        color: #ffc107 !important;
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
    - **Excel (.xlsx, .xls)**: Planilha Excel com respostas individuais
    - **CSV (.csv)**: Arquivo CSV exportado do Google Forms ou outro sistema
    
    ### Estrutura necessária:
    - As primeiras colunas devem conter informações demográficas (Setor, Cargo, etc.)
    - As colunas das perguntas devem estar numeradas (começando com números)
    - As respostas devem ser valores numéricos de 1 a 5
    
    ### Como obter os dados:
    1. Aplique o questionário HSE-IT via formulário Google Forms ou planilha Excel
    2. Exporte as respostas em formato CSV ou Excel
    3. Faça o upload do arquivo nesta página
    
    ### Verificações realizadas:
    - Formato do arquivo (.csv, .xlsx ou .xls)
    - Tamanho do arquivo (máximo 10 MB)
    - Presença das perguntas do HSE-IT (espera-se pelo menos 30 das 35 perguntas)
    - Quantidade de dados ausentes (idealmente menos de 20%)
    - Intervalo dos valores (deve estar entre 1 e 5)
    
    ### Empresas pequenas:
    O HSE-IT é válido para empresas de qualquer tamanho. Mesmo com poucas respostas, 
    o diagnóstico ajuda no cumprimento da NR-1 para avaliação de riscos psicossociais.
    
    Para um modelo pronto, acesse a página "Informações" e baixe o template HSE-IT.
    """)

# Container para o upload
with st.container():
    st.markdown('<div class="uploadbox">', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Escolha um arquivo Excel ou CSV", type=["xlsx", "xls", "csv"])
    st.caption("Tamanho máximo: 10 MB. Formatos aceitos: Excel (.xlsx, .xls) e CSV (.csv)")
    st.markdown('</div>', unsafe_allow_html=True)

# Processar o arquivo quando enviado
if uploaded_file is not None:
    try:
        # Informação de progresso
        with st.spinner("Carregando e processando dados..."):
            # Verificar extensão do arquivo
            extension = uploaded_file.name.split('.')[-1].lower()
            if extension not in ['csv', 'xlsx', 'xls']:
                st.error(f"Formato de arquivo não suportado: .{extension}. Por favor, use arquivos CSV ou Excel (XLSX/XLS).")
                st.stop()
            
            # Verificar tamanho do arquivo (limite de 10MB)
            file_size = len(uploaded_file.getvalue()) / (1024 * 1024)  # em MB
            if file_size > 10:
                st.error(f"O arquivo é muito grande ({file_size:.1f} MB). O tamanho máximo permitido é 10 MB.")
                st.stop()
            
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
            mensagens_aviso = []
            
            # Verificar se existem respostas
            if total_respostas == 0:
                dados_validos = False
                mensagens_validacao.append("O arquivo não contém nenhuma resposta.")
            elif total_respostas < 5:  # Reduzido para apenas alertar com menos de 5 respostas
                mensagens_aviso.append(f"O arquivo contém apenas {total_respostas} respostas. Embora os resultados sejam válidos, a representatividade estatística pode ser limitada.")
            
            # Verificar se as perguntas do HSE-IT estão presentes
            if len(colunas_perguntas) < 30:  # Esperamos pelo menos 30 das 35 perguntas
                dados_validos = False
                mensagens_validacao.append(f"Foram encontradas apenas {len(colunas_perguntas)} perguntas. O HSE-IT contém 35 perguntas.")
            
            # Verificar valores ausentes nas perguntas
            valores_ausentes = df_perguntas.isna().sum().sum()
            total_valores = df_perguntas.size
            percentual_ausente = (valores_ausentes / total_valores) * 100 if total_valores > 0 else 0
            
            if percentual_ausente > 30:  # Se mais de 30% dos dados estão ausentes
                dados_validos = False
                mensagens_validacao.append(f"Há {percentual_ausente:.1f}% de respostas ausentes, o que compromete a análise.")
            elif percentual_ausente > 20:
                mensagens_aviso.append(f"Há {percentual_ausente:.1f}% de respostas ausentes. Os resultados podem ser afetados.")
            
            # Verificar se os valores estão no intervalo correto (1-5)
            valores_min = df_perguntas.min().min()
            valores_max = df_perguntas.max().max()
            
            if pd.notna(valores_min) and pd.notna(valores_max):
                if valores_min < 1 or valores_max > 5:
                    mensagens_aviso.append(f"Foram encontrados valores fora do intervalo esperado (1-5): mínimo={valores_min}, máximo={valores_max}. Isso pode afetar os resultados.")
            
            # Adicionar alerta para empresas pequenas
            if total_respostas < 10:
                st.info("""
                **Nota para empresas pequenas:** O HSE-IT pode ser usado por empresas de qualquer tamanho. 
                Embora um número maior de respostas proporcione resultados estatisticamente mais robustos, 
                o diagnóstico ainda é válido para empresas menores e ajudará no cumprimento da NR-1 em relação 
                à avaliação de riscos psicossociais. Recomendamos que todas as pessoas da empresa respondam para 
                obter os melhores resultados possíveis.
                """)
            
            # Exibir avisos se houver
            for aviso in mensagens_aviso:
                st.warning(aviso)
            
            # Calcular resultados se os dados forem válidos
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
                    
                    # Lista de verificação de qualidade
                    st.markdown("### Verificações de qualidade:")
                    checklist_html = "<ul class='checklist'>"
                    
                    # Verificações passadas
                    checklist_html += f"<li>Formato do arquivo válido (.{extension})</li>"
                    checklist_html += f"<li>Tamanho do arquivo adequado ({file_size:.2f} MB)</li>"
                    checklist_html += f"<li>Estrutura de colunas correta</li>"
                    
                    # Verificações com condições
                    if percentual_ausente <= 20:
                        checklist_html += f"<li>Dados completos ({100-percentual_ausente:.1f}% preenchido)</li>"
                    elif percentual_ausente <= 30:
                        checklist_html += f"<li class='checklist-item-warning'>Dados parcialmente completos ({100-percentual_ausente:.1f}% preenchido)</li>"
                    else:
                        checklist_html += f"<li class='checklist-item-error'>Muitos dados ausentes ({percentual_ausente:.1f}% ausente)</li>"
                    
                    if total_respostas >= 10:
                        checklist_html += f"<li>Número adequado de respostas ({total_respostas})</li>"
                    elif total_respostas >= 5:
                        checklist_html += f"<li class='checklist-item-warning'>Número limitado de respostas ({total_respostas}), mas aceitável para empresas pequenas</li>"
                    else:
                        checklist_html += f"<li class='checklist-item-warning'>Número muito pequeno de respostas ({total_respostas}), análise será simplificada</li>"
                    
                    if pd.notna(valores_min) and pd.notna(valores_max) and valores_min >= 1 and valores_max <= 5:
                        checklist_html += f"<li>Valores dentro do intervalo esperado (1-5)</li>"
                    else:
                        checklist_html += f"<li class='checklist-item-warning'>Valores fora do intervalo esperado (mín={valores_min}, máx={valores_max})</li>"
                    
                    checklist_html += "</ul>"
                    st.markdown(checklist_html, unsafe_allow_html=True)
                
                with col2:
                    st.subheader("Dimensões Avaliadas")
                    # Mostrar resumo das dimensões encontradas
                    for dimensao, _ in DIMENSOES_HSE.items():
                        resultado_dimensao = df_resultados[df_resultados["Dimensão"] == dimensao]
                        if not resultado_dimensao.empty:
                            media = resultado_dimensao.iloc[0]["Média"]
                            risco = resultado_dimensao.iloc[0]["Risco"]
                            
                            # Obter a cor com base no nível de risco
                            cor = "red" if "Muito Alto" in risco else \
                                  "orange" if "Alto" in risco else \
                                  "yellow" if "Moderado" in risco else \
                                  "green" if "Baixo" in risco else \
                                  "purple"
                            
                            st.markdown(f"- **{dimensao}**: {media:.2f} - <span style='color:{cor};'>{risco}</span>", unsafe_allow_html=True)
                
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
                    if st.button("Ver Resultados Detalhados", use_container_width=True):
                        st.switch_page("pages/02_resultados.py")
                with col2:
                    if st.button("Ir para Plano de Ação", use_container_width=True):
                        st.switch_page("pages/03_plano_acao.py")
                with col3:
                    if st.button("Gerar Relatórios", use_container_width=True):
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
                
                # Exibir detalhes adicionais para debug
                with st.expander("Detalhes do arquivo", expanded=False):
                    if df is not None:
                        st.write("Primeiras linhas do arquivo:")
                        st.dataframe(df.head())
                        
                        st.write("Colunas encontradas:")
                        st.write(df.columns.tolist())
                        
                        st.write("Colunas interpretadas como perguntas:")
                        st.write(colunas_perguntas)
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
    
    # Mostrar progresso do fluxo de trabalho
    st.markdown("""
    <div class="info-box">
    <h4>Fluxo de Trabalho da Análise HSE-IT</h4>
    <p>Você está na <b>Etapa 1</b> de 4 do processo de análise HSE-IT</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    <div class="progress-indicator">
        <div class="progress-step">
            <div class="step-number active">1</div>
            <div>Upload</div>
        </div>
        <div class="step-line"></div>
        <div class="progress-step">
            <div class="step-number">2</div>
            <div>Resultados</div>
        </div>
        <div class="step-line"></div>
        <div class="progress-step">
            <div class="step-number">3</div>
            <div>Plano de Ação</div>
        </div>
        <div class="step-line"></div>
        <div class="progress-step">
            <div class="step-number">4</div>
            <div>Relatórios</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Exibir dados de exemplo para referência
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
        
        # Oferecer download de modelo de exemplo
        st.markdown("### Precisa de um modelo para começar?")
        
        if st.button("Baixar Modelo de Exemplo (CSV)"):
            # Criar um CSV de exemplo mais completo
            colunas_completas = ["Setor", "Cargo", "Tempo_Empresa", "Genero", "Faixa_Etaria", "Escolaridade", "Regime_Trabalho"]
            # Adicionar as 35 perguntas
            for i in range(1, 36):
                colunas_completas.append(f"{i}. Pergunta exemplo")
            
            # Criar alguns dados de exemplo
            dados_completos = []
            setores = ["TI", "RH", "Financeiro", "Marketing", "Operações"]
            cargos = ["Analista", "Gerente", "Assistente", "Coordenador", "Diretor"]
            tempo = ["<1 ano", "1-3 anos", "3-5 anos", "5-10 anos", ">10 anos"]
            
            # Gerar 10 linhas de exemplo
            for _ in range(10):
                linha = [
                    setores[np.random.randint(0, len(setores))],
                    cargos[np.random.randint(0, len(cargos))],
                    tempo[np.random.randint(0, len(tempo))],
                    "Masculino" if np.random.random() > 0.5 else "Feminino",
                    f"{20 + np.random.randint(0, 4)*10}-{29 + np.random.randint(0, 4)*10}",
                    np.random.choice(["Médio", "Superior", "Pós-graduação"]),
                    np.random.choice(["CLT", "PJ", "Autônomo"])
                ]
                
                # Adicionar respostas para as 35 perguntas
                for _ in range(35):
                    linha.append(np.random.randint(1, 6))
                
                dados_completos.append(linha)
            
            df_completo = pd.DataFrame(dados_completos, columns=colunas_completas)
            
            # Converter para CSV
            csv = df_completo.to_csv(index=False)
            
            # Oferecer para download
            st.download_button(
                label="Confirmar Download",
                data=csv,
                file_name="modelo_hse_it.csv",
                mime="text/csv"
            )
