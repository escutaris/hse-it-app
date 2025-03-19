import streamlit as st
import pandas as pd
import io
import plotly.graph_objects as go
from datetime import datetime
from utils.processamento import classificar_risco

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
    
    /* Status tags */
    .status-not-started {
        background-color: #f8d7da;
        color: #721c24;
        padding: 3px 8px;
        border-radius: 4px;
        font-size: 0.8em;
    }
    
    .status-in-progress {
        background-color: #fff3cd;
        color: #856404;
        padding: 3px 8px;
        border-radius: 4px;
        font-size: 0.8em;
    }
    
    .status-completed {
        background-color: #d4edda;
        color: #155724;
        padding: 3px 8px;
        border-radius: 4px;
        font-size: 0.8em;
    }
    
    .status-canceled {
        background-color: #e2e3e5;
        color: #383d41;
        padding: 3px 8px;
        border-radius: 4px;
        font-size: 0.8em;
    }
    </style>
    """, unsafe_allow_html=True)

# Aplicar o estilo
aplicar_estilo_escutaris()

# Título da página
st.title("Plano de Ação - HSE-IT")

# Verificar se há dados para exibir
if "df_plano_acao" not in st.session_state or st.session_state.df_plano_acao is None:
    st.warning("Nenhum plano de ação disponível. Por favor, faça upload de um arquivo na página 'Upload de Dados'.")
    st.stop()

# Recuperar plano de ação
df_plano_acao = st.session_state.df_plano_acao

# Função para plano de ação editável - com verificações adicionais para evitar o ValueError
def plano_acao_editavel(df_plano_acao):
    st.header("Plano de Ação Personalizado")
    st.write("Personalize o plano de ação sugerido ou adicione suas próprias ações para cada dimensão HSE-IT.")
    
    # Inicializar plano de ação no state se não existir
    if "plano_acao_personalizado" not in st.session_state:
        st.session_state.plano_acao_personalizado = df_plano_acao.copy()
    
    # Verificar se há dimensões disponíveis
    if "Dimensão" not in df_plano_acao.columns:
        st.error("O plano de ação não contém a coluna 'Dimensão'. Verifique o formato dos dados.")
        return
    
    dimensoes_unicas = df_plano_acao["Dimensão"].unique()
    
    # Verificar se há dimensões únicas
    if len(dimensoes_unicas) == 0:
        st.error("Não foram encontradas dimensões para criar o plano de ação. Verifique os dados carregados.")
        return
    
    # Criar tabs para cada dimensão
    try:
        dimensao_tabs = st.tabs(dimensoes_unicas)
    except Exception as e:
        st.error(f"Erro ao criar abas para dimensões: {str(e)}")
        st.write("Detalhes das dimensões encontradas:")
        st.write(dimensoes_unicas)
        return
    
    # Para cada dimensão, criar um editor de ações
    for i, dimensao in enumerate(dimensoes_unicas):
        with dimensao_tabs[i]:
            df_dimensao = st.session_state.plano_acao_personalizado[
                st.session_state.plano_acao_personalizado["Dimensão"] == dimensao
            ].copy()
            
            # Verificar se há dados para esta dimensão
            if df_dimensao.empty:
                st.warning(f"Não há ações definidas para a dimensão {dimensao}.")
                continue
            
            # Mostrar informações da dimensão
            try:
                nivel_risco = df_dimensao["Nível de Risco"].iloc[0]
                media = df_dimensao["Média"].iloc[0]
                
                # Definir cor com base no nível de risco
                cor = {
                    "Risco Muito Alto": "red",
                    "Risco Alto": "orange",
                    "Risco Moderado": "yellow",
                    "Risco Baixo": "green",
                    "Risco Muito Baixo": "purple"
                }.get(nivel_risco, "gray")
                
                # Exibir status com cor
                st.markdown(f"**Média:** {media} - **Nível de Risco:** :{cor}[{nivel_risco}]")
                
                # Mostrar descrição da dimensão se disponível
                if "DESCRICOES_DIMENSOES" in st.session_state and dimensao in st.session_state.DESCRICOES_DIMENSOES:
                    with st.expander("Sobre esta dimensão"):
                        st.write(st.session_state.DESCRICOES_DIMENSOES[dimensao])
            except Exception as e:
                st.error(f"Erro ao exibir informações para a dimensão {dimensao}: {str(e)}")
                continue
            
            # Adicionar nova ação
            st.subheader("Adicionar Nova Ação:")
            st.markdown('<div class="action-form">', unsafe_allow_html=True)
            col1, col2 = st.columns([3, 1])
            
            with col1:
                nova_acao = st.text_area("Descrição da ação", key=f"nova_acao_{dimensao}")
            
            with col2:
                st.write("&nbsp;")  # Espaçamento
                adicionar = st.button("Adicionar", key=f"add_{dimensao}")
                
                if adicionar and nova_acao.strip():
                    # Criar nova linha para o DataFrame
                    nova_linha = {
                        "Dimensão": dimensao,
                        "Nível de Risco": nivel_risco,
                        "Média": media,
                        "Sugestão de Ação": nova_acao,
                        "Responsável": "",
                        "Prazo": "",
                        "Status": "Não iniciada",
                        "Personalizada": True  # Marcar como ação personalizada
                    }
                    
                    # Adicionar ao DataFrame
                    st.session_state.plano_acao_personalizado = pd.concat([
                        st.session_state.plano_acao_personalizado, 
                        pd.DataFrame([nova_linha])
                    ], ignore_index=True)
                    
                    # Limpar campo de texto
                    st.session_state[f"nova_acao_{dimensao}"] = ""
                    st.experimental_rerun()
            st.markdown('</div>', unsafe_allow_html=True)
            
            # Mostrar ações existentes para editar
            st.subheader("Ações Sugeridas:")
            
            # Agrupar por status para melhor organização
            status_order = ["Em andamento", "Não iniciada", "Concluída", "Cancelada"]
            df_dimensao_sorted = df_dimensao.copy()
            df_dimensao_sorted["Status_Order"] = df_dimensao_sorted["Status"].apply(lambda x: status_order.index(x) if x in status_order else 999)
            df_dimensao_sorted = df_dimensao_sorted.sort_values(by=["Status_Order", "Sugestão de Ação"])
            
            # Contagem de ações por status
            status_counts = df_dimensao["Status"].value_counts()
            total_acoes = len(df_dimensao)
            
            # Mostrar progresso para esta dimensão
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total de Ações", total_acoes)
            col2.metric("Não Iniciadas", status_counts.get("Não iniciada", 0))
            col3.metric("Em Andamento", status_counts.get("Em andamento", 0))
            col4.metric("Concluídas", status_counts.get("Concluída", 0))
            
            # Barra de progresso
            progresso = status_counts.get("Concluída", 0) / total_acoes if total_acoes > 0 else 0
            st.progress(progresso)
            st.write(f"Progresso: {int(progresso*100)}% completo")
            
            # Filtro de status
            status_filter = st.multiselect(
                "Filtrar por status:",
                status_order,
                default=status_order
            )
            
            if status_filter:
                df_dimensao_sorted = df_dimensao_sorted[df_dimensao_sorted["Status"].isin(status_filter)]
            
            # Mostrar ações
            for j, (index, row) in enumerate(df_dimensao_sorted.iterrows()):
                # Definir estilo baseado no status
                status_class = {
                    "Não iniciada": "status-not-started",
                    "Em andamento": "status-in-progress",
                    "Concluída": "status-completed",
                    "Cancelada": "status-canceled"
                }.get(row.get("Status", "Não iniciada"), "")
                
                # Criar card para cada ação
                st.markdown(f'<div class="card">', unsafe_allow_html=True)
                
                # Título da ação com status
                st.markdown(f"""
                <h4 style="margin-bottom: 10px;">Ação {j+1} 
                <span class="{status_class}">{row.get("Status", "Não iniciada")}</span>
                </h4>
                """, unsafe_allow_html=True)
                
                # Layout para os campos editáveis
                col1, col2 = st.columns([3, 1])
                
                with col1:
                    # Editor de texto para a ação
                    acao_editada = st.text_area(
                        "Descrição da ação", 
                        row["Sugestão de Ação"], 
                        key=f"acao_{dimensao}_{j}"
                    )
                    if acao_editada != row["Sugestão de Ação"]:
                        st.session_state.plano_acao_personalizado.at[index, "Sugestão de Ação"] = acao_editada
                
                with col2:
                    # Campos para responsável, prazo e status
                    responsavel = st.text_input(
                        "Responsável", 
                        row.get("Responsável", ""), 
                        key=f"resp_{dimensao}_{j}"
                    )
                    if responsavel != row.get("Responsável", ""):
                        st.session_state.plano_acao_personalizado.at[index, "Responsável"] = responsavel
                    
                    # Campo de data para prazo
                    try:
                        data_padrao = None
                        if row.get("Prazo") and row.get("Prazo") != "":
                            try:
                                data_padrao = datetime.strptime(row.get("Prazo"), "%d/%m/%Y")
                            except Exception:
                                # Tentar outros formatos comuns
                                try:
                                    data_padrao = datetime.strptime(row.get("Prazo"), "%Y-%m-%d")
                                except Exception:
                                    data_padrao = None
                        
                        prazo = st.date_input(
                            "Prazo", 
                            value=data_padrao,
                            key=f"prazo_{dimensao}_{j}"
                        )
                        if prazo:
                            st.session_state.plano_acao_personalizado.at[index, "Prazo"] = prazo.strftime("%d/%m/%Y")
                    except Exception as e:
                        st.warning(f"Erro ao processar data: {e}")
                    
                    # Seletor de status
                    status = st.selectbox(
                        "Status",
                        options=["Não iniciada", "Em andamento", "Concluída", "Cancelada"],
                        index=["Não iniciada", "Em andamento", "Concluída", "Cancelada"].index(row.get("Status", "Não iniciada")),
                        key=f"status_{dimensao}_{j}"
                    )
                    if status != row.get("Status", "Não iniciada"):
                        st.session_state.plano_acao_personalizado.at[index, "Status"] = status
                    
                    # Botão para remover (apenas para ações personalizadas)
                    if row.get("Personalizada", False):
                        if st.button("🗑️ Remover", key=f"del_{dimensao}_{j}"):
                            st.session_state.plano_acao_personalizado = st.session_state.plano_acao_personalizado.drop(index)
                            st.experimental_rerun()
                
                st.markdown('</div>', unsafe_allow_html=True)
    
    # Função para gerar excel do plano personalizado
    def gerar_excel_plano_personalizado(df_plano_personalizado):
        import io
        import pandas as pd
        
        output = io.BytesIO()
        
        try:
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # Aba principal com o plano de ação
                df_plano_copia = df_plano_personalizado.copy()
                
                # Remover coluna Personalizada antes de exportar
                if "Personalizada" in df_plano_copia.columns:
                    df_plano_copia = df_plano_copia.drop(columns=["Personalizada"])
                
                # Remover colunas adicionais usadas para ordenação
                if "Status_Order" in df_plano_copia.columns:
                    df_plano_copia = df_plano_copia.drop(columns=["Status_Order"])
                    
                df_plano_copia.to_excel(writer, sheet_name='Plano de Ação', index=False)
                
                # Formatar a planilha
                worksheet = writer.sheets['Plano de Ação']
                
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
                
                status_format = {
                    'Não iniciada': workbook.add_format({'bg_color': '#f8d7da'}),
                    'Em andamento': workbook.add_format({'bg_color': '#fff3cd'}),
                    'Concluída': workbook.add_format({'bg_color': '#d4edda'}),
                    'Cancelada': workbook.add_format({'bg_color': '#e2e3e5'})
                }
                
                # Configurar largura das colunas
                worksheet.set_column('A:A', 25)  # Dimensão
                worksheet.set_column('B:B', 15)  # Nível de Risco
                worksheet.set_column('C:C', 10)  # Média
                worksheet.set_column('D:D', 50)  # Sugestão de Ação
                worksheet.set_column('E:E', 15)  # Responsável
                worksheet.set_column('F:F', 15)  # Prazo
                worksheet.set_column('G:G', 15)  # Status
                
                # Adicionar cabeçalhos formatados
                for col_num, value in enumerate(df_plano_copia.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # Aplicar formatação condicional baseada no nível de risco e status
                for row_num, (_, row) in enumerate(df_plano_copia.iterrows(), 1):
                    nivel_risco = row["Nível de Risco"]
                    if nivel_risco in risco_format:
                        worksheet.write(row_num, 1, nivel_risco, risco_format[nivel_risco])
                    
                    status = row["Status"]
                    if status in status_format:
                        worksheet.write(row_num, 6, status, status_format[status])
                
                # Adicionar validação de dados para a coluna Status
                status_options = ['Não iniciada', 'Em andamento', 'Concluída', 'Cancelada']
                worksheet.data_validation('G2:G1000', {'validate': 'list',
                                                    'source': status_options,
                                                    'input_title': 'Selecione o status:',
                                                    'input_message': 'Escolha um status da lista'})
                
                # Adicionar filtros
                worksheet.autofilter(0, 0, len(df_plano_copia), len(df_plano_copia.columns) - 1)
                
                # Congelar painel para manter cabeçalhos visíveis durante rolagem
                worksheet.freeze_panes(1, 0)
                
                # Adicionar uma aba com gráfico de progresso
                worksheet_resumo = workbook.add_worksheet('Resumo de Progresso')
                
                # Adicionar cabeçalhos
                worksheet_resumo.write(0, 0, 'Dimensão', header_format)
                worksheet_resumo.write(0, 1, 'Total de Ações', header_format)
                worksheet_resumo.write(0, 2, 'Não Iniciadas', header_format)
                worksheet_resumo.write(0, 3, 'Em Andamento', header_format)
                worksheet_resumo.write(0, 4, 'Concluídas', header_format)
                worksheet_resumo.write(0, 5, 'Canceladas', header_format)
                worksheet_resumo.write(0, 6, 'Progresso (%)', header_format)
                
                # Configurar largura das colunas
                worksheet_resumo.set_column('A:A', 25)
                worksheet_resumo.set_column('B:F', 15)
                worksheet_resumo.set_column('G:G', 15)
                
                # Calcular estatísticas por dimensão
                dimensoes = df_plano_copia["Dimensão"].unique()
                dados_resumo = []
                
                for i, dimensao in enumerate(dimensoes, 1):
                    df_dim = df_plano_copia[df_plano_copia["Dimensão"] == dimensao]
                    total = len(df_dim)
                    nao_iniciadas = sum(df_dim["Status"] == "Não iniciada")
                    em_andamento = sum(df_dim["Status"] == "Em andamento")
                    concluidas = sum(df_dim["Status"] == "Concluída")
                    canceladas = sum(df_dim["Status"] == "Cancelada")
                    progresso = (concluidas / total) * 100 if total > 0 else 0
                    
                    worksheet_resumo.write(i, 0, dimensao)
                    worksheet_resumo.write(i, 1, total)
                    worksheet_resumo.write(i, 2, nao_iniciadas)
                    worksheet_resumo.write(i, 3, em_andamento)
                    worksheet_resumo.write(i, 4, concluidas)
                    worksheet_resumo.write(i, 5, canceladas)
                    worksheet_resumo.write(i, 6, progresso)
                    
                    dados_resumo.append([dimensao, total, nao_iniciadas, em_andamento, concluidas, canceladas, progresso])
                
                # Adicionar gráfico de barras para visualizar progresso por dimensão
                chart = workbook.add_chart({'type': 'bar', 'subtype': 'stacked'})
                
                # Adicionar séries para cada status
                for idx, (status, col) in enumerate([
                    ('Concluídas', 4), 
                    ('Em Andamento', 3), 
                    ('Não Iniciadas', 2),
                    ('Canceladas', 5)
                ]):
                    chart.add_series({
                        'name': status,
                        'categories': ['Resumo de Progresso', 1, 0, len(dimensoes), 0],
                        'values': ['Resumo de Progresso', 1, col, len(dimensoes), col],
                        'fill': {'color': ['#d4edda', '#fff3cd', '#f8d7da', '#e2e3e5'][idx]}
                    })
                
                chart.set_title({'name': 'Progresso do Plano de Ação por Dimensão'})
                chart.set_x_axis({'name': 'Número de Ações'})
                chart.set_y_axis({'name': 'Dimensão'})
                chart.set_legend({'position': 'bottom'})
                
                worksheet_resumo.insert_chart('I1', chart, {'x_scale': 1.5, 'y_scale': 2})
            
            output.seek(0)
            return output
        
        except Exception as e:
            st.error(f"Erro ao gerar o arquivo Excel: {str(e)}")
            import traceback
            st.write(traceback.format_exc())
            return None
    
    # Layout para os botões de exportação e métricas gerais
    st.divider()
    
    # Métricas gerais do plano de ação
    st.subheader("Resumo Geral do Plano de Ação")
    
    # Calcular estatísticas gerais
    plano = st.session_state.plano_acao_personalizado
    total_acoes = len(plano)
    acoes_por_status = plano["Status"].value_counts()
    acoes_por_dimensao = plano["Dimensão"].value_counts()
    
    # Primeira linha de métricas
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total de Ações", total_acoes)
    
    with col2:
        st.metric("Não Iniciadas", acoes_por_status.get("Não iniciada", 0))
    
    with col3:
        st.metric("Em Andamento", acoes_por_status.get("Em andamento", 0))
    
    with col4:
        st.metric("Concluídas", acoes_por_status.get("Concluída", 0))
    
    # Adicionar gráfico de progresso geral
    progresso_geral = acoes_por_status.get("Concluída", 0) / total_acoes if total_acoes > 0 else 0
    st.progress(progresso_geral)
    st.write(f"Progresso geral: {progresso_geral*100:.1f}%")
    
    # Adicionar visualizações de progresso
    col1, col2 = st.columns(2)
    
    with col1:
        # Gráfico de pizza para status
        fig_pie = go.Figure(data=[go.Pie(
            labels=acoes_por_status.index,
            values=acoes_por_status.values,
            hole=.3,
            marker_colors=['#f8d7da', '#fff3cd', '#d4edda', '#e2e3e5']
        )])
        fig_pie.update_layout(
            title="Distribuição por Status",
            height=350,
            margin=dict(l=20, r=20, t=30, b=20)
        )
        st.plotly_chart(fig_pie, use_container_width=True)
    
    with col2:
        # Gráfico de barras para dimensões
        fig_bar = go.Figure()
        
        fig_bar.add_trace(go.Bar(
            x=acoes_por_dimensao.values,
            y=acoes_por_dimensao.index,
            orientation='h',
            marker_color='#5A713D'
        ))
        
        fig_bar.update_layout(
            title="Ações por Dimensão",
            xaxis_title="Número de Ações",
            yaxis_title="Dimensão",
            height=350,
            margin=dict(l=20, r=20, t=30, b=20)
        )
        
        st.plotly_chart(fig_bar, use_container_width=True)
    
    # Botões para exportação
    st.subheader("Exportar Plano de Ação")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("Gerar Plano de Ação Excel", use_container_width=True):
            with st.spinner("Gerando arquivo Excel..."):
                # Gerar Excel com o plano personalizado
                output = gerar_excel_plano_personalizado(st.session_state.plano_acao_personalizado)
                if output:
                    st.success("Plano de Ação gerado com sucesso!")
                    st.session_state.excel_plano = output
                    st.session_state.excel_plano_ready = True
    
    with col2:
        # Botão para baixar Excel
        if st.session_state.get("excel_plano_ready", False) and st.session_state.get("excel_plano", None) is not None:
            st.download_button(
                label="Baixar Plano de Ação Excel",
                data=st.session_state.excel_plano,
                file_name="plano_acao_personalizado_hse_it.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.write("Clique em 'Gerar Plano de Ação Excel' para criar o arquivo")

# Executar a função de plano de ação editável com tratamento de erro
try:
    plano_acao_editavel(df_plano_acao)
except Exception as e:
    st.error(f"Ocorreu um erro ao processar o plano de ação: {str(e)}")
    
    # Informações detalhadas para depuração
    with st.expander("Detalhes técnicos do erro"):
        st.write("Detalhes do erro:")
        st.write(e)
        import traceback
        st.code(traceback.format_exc())
        
        st.write("Informações sobre os dados:")
        st.write(f"Tipo do objeto df_plano_acao: {type(df_plano_acao)}")
        st.write(f"Colunas disponíveis: {df_plano_acao.columns.tolist() if isinstance(df_plano_acao, pd.DataFrame) else 'N/A'}")
        st.write(f"Número de linhas: {len(df_plano_acao) if isinstance(df_plano_acao, pd.DataFrame) else 'N/A'}")
        
        if isinstance(df_plano_acao, pd.DataFrame) and "Dimensão" in df_plano_acao.columns:
            st.write(f"Valores únicos na coluna 'Dimensão': {df_plano_acao['Dimensão'].unique()}")
