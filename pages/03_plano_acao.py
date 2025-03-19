import streamlit as st
import pandas as pd
import io
from datetime import datetime
from utils.processamento import classificar_risco

# Título da página
st.title("Plano de Ação - HSE-IT")

# Verificar se há dados para exibir
if "df_plano_acao" not in st.session_state or st.session_state.df_plano_acao is None:
    st.warning("Nenhum plano de ação disponível. Por favor, faça upload de um arquivo na página 'Upload de Dados'.")
    st.stop()

# Recuperar plano de ação
df_plano_acao = st.session_state.df_plano_acao

def plano_acao_editavel(df_plano_acao):
    st.header("Plano de Ação Personalizado")
    st.write("Personalize o plano de ação sugerido ou adicione suas próprias ações para cada dimensão HSE-IT.")
    
    # Inicializar plano de ação no state se não existir
    if "plano_acao_personalizado" not in st.session_state:
        st.session_state.plano_acao_personalizado = df_plano_acao.copy()
    
    # Criar tabs para cada dimensão
    dimensoes_unicas = df_plano_acao["Dimensão"].unique()
    dimensao_tabs = st.tabs(dimensoes_unicas)
    
    # Para cada dimensão, criar um editor de ações
    for i, dimensao in enumerate(dimensoes_unicas):
        with dimensao_tabs[i]:
            df_dimensao = st.session_state.plano_acao_personalizado[
                st.session_state.plano_acao_personalizado["Dimensão"] == dimensao
            ].copy()
            
            # Mostrar informações da dimensão
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
            
            # Adicionar nova ação
            st.subheader("Adicionar Nova Ação:")
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
            
            # Mostrar ações existentes para editar
            st.subheader("Ações Sugeridas:")
            for j, (index, row) in enumerate(df_dimensao.iterrows()):
                with st.expander(f"Ação {j+1}: {row['Sugestão de Ação'][:50]}...", expanded=False):
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
                                except:
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
                
                # Aplicar formatação condicional baseada no nível de risco
                for row_num, (_, row) in enumerate(df_plano_copia.iterrows(), 1):
                    nivel_risco = row["Nível de Risco"]
                    if nivel_risco in risco_format:
                        worksheet.write(row_num, 1, nivel_risco, risco_format[nivel_risco])
                
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
            
            output.seek(0)
            return output
        
        except Exception as e:
            st.error(f"Erro ao gerar o arquivo Excel: {str(e)}")
            return None
    
    # Botão para exportar plano personalizado
    if st.button("Exportar Plano de Ação Personalizado"):
        # Gerar Excel com o plano personalizado
        output = gerar_excel_plano_personalizado(st.session_state.plano_acao_personalizado)
        if output:
            st.success("Plano de Ação gerado com sucesso!")
            st.download_button(
                label="Baixar Plano de Ação Personalizado",
                data=output,
                file_name="plano_acao_personalizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# Executar a função de plano de ação editável
plano_acao_editavel(df_plano_acao)

# Adicionar resumo do status do plano de ação
if "plano_acao_personalizado" in st.session_state:
    st.subheader("Resumo do Status do Plano")
    
    # Calcular estatísticas
    plano = st.session_state.plano_acao_personalizado
    total_acoes = len(plano)
    acoes_por_status = plano["Status"].value_counts()
    
    # Criar colunas para métricas
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total de Ações", total_acoes)
    
    with col2:
        st.metric("Não Iniciadas", acoes_por_status.get("Não iniciada", 0))
    
    with col3:
        st.metric("Em Andamento", acoes_por_status.get("Em andamento", 0))
    
    with col4:
        st.metric("Concluídas", acoes_por_status.get("Concluída", 0))
    
    # Adicionar gráfico de progresso
    progresso = acoes_por_status.get("Concluída", 0) / total_acoes if total_acoes > 0 else 0
    st.progress(progresso)
    st.write(f"Progresso geral: {progresso*100:.1f}%")
