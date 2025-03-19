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
        pdf.cell(0, 6, "Risco Muito Baixo: Media > 4", ln=True)
        
        # Adicionar descrição das dimensões do HSE-IT
        pdf.ln(5)
        pdf.set_font("Arial", style='B', size=11)
        pdf.cell(0, 8, "Sobre o HSE-IT:", ln=True)
        pdf.set_font("Arial", size=9)
        
        # Corrigido: Transformei o cell/multi_cell incorreto em apenas um multi_cell
        pdf.multi_cell(0, 6, "O questionario HSE-IT avalia 7 dimensoes de fatores psicossociais no trabalho: Demanda, Controle, Apoio da Chefia, Apoio dos Colegas, Relacionamentos, Funcao e Mudanca. Priorize as acoes nas dimensoes com maior risco.", 0)
        
        # Corrigindo o problema de BytesIO
        temp_file = "temp_plano_acao.pdf"
        pdf.output(temp_file)
        
        with open(temp_file, "rb") as file:
            output = io.BytesIO(file.read())
        
        import os
        if os.path.exists(temp_file):
            os.remove(temp_file)
        
        return output
    
    except Exception as e:
        st.error(f"Erro ao gerar o PDF do Plano de Ação: {str(e)}")
        return None

# Função para gerar PDF do relatório de resultados
def gerar_pdf(df_resultados):
    try:
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        
        # Configurando fonte para suporte básico
        pdf.set_font("Arial", style='B', size=14)
        
        # Usando strings sem acentos para evitar problemas de codificação
        pdf.cell(200, 10, "Relatorio de Fatores Psicossociais - HSE-IT", ln=True, align='C')
        pdf.ln(10)
        
        # Adicionar explicação da metodologia
        pdf.set_font("Arial", style='I', size=10)
        pdf.multi_cell(0, 5, "O questionario HSE-IT avalia 7 dimensoes de fatores psicossociais no trabalho. Os resultados sao apresentados em uma escala de 1 a 5, onde valores mais altos indicam melhores resultados.", 0)
        pdf.ln(5)
        
        # Tabela de resultados
        pdf.set_font("Arial", style='B', size=10)
        pdf.cell(80, 7, "Dimensao", 1)
        pdf.cell(25, 7, "Media", 1)
        pdf.cell(60, 7, "Nivel de Risco", 1)
        pdf.ln()
        
        pdf.set_font("Arial", size=10)
        for _, row in df_resultados.iterrows():
            # Remover acentos e caracteres especiais
            dimensao = row['Dimensão'].encode('ascii', 'replace').decode('ascii')
            risco = row['Risco'].split(' ')[0] + ' ' + row['Risco'].split(' ')[1]  # Remove emoji
            
            # Determinar cor de fundo baseada no risco
            if "Muito Alto" in risco:
                pdf.set_fill_color(255, 107, 107)  # Vermelho claro
                text_color = 255  # Branco
            elif "Alto" in risco:
                pdf.set_fill_color(255, 165, 0)  # Laranja
                text_color = 0  # Preto
            elif "Moderado" in risco:
                pdf.set_fill_color(255, 255, 0)  # Amarelo
                text_color = 0  # Preto
            elif "Baixo" in risco:
                pdf.set_fill_color(144, 238, 144)  # Verde claro
                text_color = 0  # Preto
            elif "Muito Baixo" in risco:
                pdf.set_fill_color(187, 143, 206)  # Roxo claro
                text_color = 0  # Preto
            else:
                pdf.set_fill_color(255, 255, 255)  # Branco
                text_color = 0  # Preto
                
            pdf.set_text_color(text_color)
            pdf.cell(80, 7, dimensao, 1, 0, 'L', 1)
            pdf.cell(25, 7, f"{row['Média']:.2f}", 1, 0, 'C', 1)
            pdf.cell(60, 7, risco, 1, 0, 'C', 1)
            pdf.ln()
        
        # Resetar cor do texto
        pdf.set_text_color(0)
        
        # Adicionar descrições das dimensões
        pdf.ln(10)
        pdf.set_font("Arial", style='B', size=12)
        pdf.cell(0, 10, "Descricao das Dimensoes:", ln=True)
        
        for dimensao, descricao in st.session_state.get('DESCRICOES_DIMENSOES', {}).items():
            pdf.set_font("Arial", style='B', size=10)
            pdf.cell(0, 7, dimensao.encode('ascii', 'replace').decode('ascii'), ln=True)
            pdf.set_font("Arial", size=10)
            pdf.multi_cell(0, 5, descricao.encode('ascii', 'replace').decode('ascii'), 0)
            pdf.ln(3)
        
        # Adicionar informações sobre classificação de risco
        pdf.ln(10)
        pdf.set_font("Arial", style='B', size=12)
        pdf.cell(0, 10, "Legenda de Classificacao:", ln=True)
        pdf.set_font("Arial", size=10)
        pdf.cell(0, 8, "Risco Muito Alto: Media <= 1", ln=True)
        pdf.cell(0, 8, "Risco Alto: 1 < Media <= 2", ln=True)
        pdf.cell(0, 8, "Risco Moderado: 2 < Media <= 3", ln=True)
        pdf.cell(0, 8, "Risco Baixo: 3 < Media <= 4", ln=True)
        pdf.cell(0, 8, "Risco Muito Baixo: Media > 4", ln=True)
        
        # Adicionar recomendações gerais
        pdf.ln(10)
        pdf.set_font("Arial", style='B', size=12)
        pdf.cell(0, 10, "Recomendacoes Gerais:", ln=True)
        pdf.set_font("Arial", size=10)
        pdf.multi_cell(0, 5, "Para resultados mais detalhados e um plano de acao personalizado, consulte o arquivo Excel completo ou o documento do Plano de Acao fornecido. Priorize intervencoes nas dimensoes com maior nivel de risco.", 0)
        
        # Corrigindo o problema de BytesIO
        temp_file = "temp_report.pdf"
        pdf.output(temp_file)
        
        with open(temp_file, "rb") as file:
            output = io.BytesIO(file.read())
        
        import os
        if os.path.exists(temp_file):
            os.remove(temp_file)
        
        return output
    
    except Exception as e:
        st.error(f"Erro ao gerar o PDF: {str(e)}")
        return None

# Função para gerar Excel completo com múltiplas abas
def gerar_excel_completo(df, df_perguntas, colunas_filtro, colunas_perguntas):
    output = io.BytesIO()
    
    try:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Aba 1: Empresa Toda
            df_resultados_excel = st.session_state.df_resultados.copy()
            if 'Questões' in df_resultados_excel.columns:
                df_resultados_excel = df_resultados_excel.drop(columns=['Questões'])
                
            df_resultados_excel.to_excel(writer, sheet_name='Empresa Toda', index=False)
            
            # Formatar a planilha
            worksheet = writer.sheets['Empresa Toda']
            
            # Definir formatos
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })

            # Configurar largura das colunas
            worksheet.set_column('A:A', 5)  # Índice
            worksheet.set_column('B:B', 25)  # Dimensão
            worksheet.set_column('C:C', 40)  # Descrição
            worksheet.set_column('D:D', 10)  # Média
            worksheet.set_column('E:E', 20)  # Risco
            worksheet.set_column('F:F', 15)  # Número de Respostas
            
            # Adicionar cabeçalhos formatados
            for col_num, value in enumerate(df_resultados_excel.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Adicionar filtros
            worksheet.autofilter(0, 0, len(df_resultados_excel), len(df_resultados_excel.columns) - 1)
            
            # Congelar painel para manter cabeçalhos visíveis
            worksheet.freeze_panes(1, 0)
            
            # Aba 2: Plano de Ação
            df_plano_acao.to_excel(writer, sheet_name='Plano de Ação', index=False)
            
            # Formatar a aba de plano de ação
            worksheet_plano = writer.sheets['Plano de Ação']
            
            # Definir formatos para plano de ação
            risco_format = {
                'Risco Muito Alto': workbook.add_format({'bg_color': '#FF6B6B', 'font_color': 'white'}),
                'Risco Alto': workbook.add_format({'bg_color': '#FFA500'}),
                'Risco Moderado': workbook.add_format({'bg_color': '#FFFF00'}),
                'Risco Baixo': workbook.add_format({'bg_color': '#90EE90'}),
                'Risco Muito Baixo': workbook.add_format({'bg_color': '#BB8FCE'})
            }

            # Configurar largura das colunas
            worksheet_plano.set_column('A:A', 25)  # Dimensão
            worksheet_plano.set_column('B:B', 15)  # Nível de Risco
            worksheet_plano.set_column('C:C', 10)  # Média
            worksheet_plano.set_column('D:D', 50)  # Sugestão de Ação
            worksheet_plano.set_column('E:E', 15)  # Responsável
            worksheet_plano.set_column('F:F', 15)  # Prazo
            worksheet_plano.set_column('G:G', 15)  # Status

            # Adicionar cabeçalhos formatados
            for col_num, value in enumerate(df_plano_acao.columns.values):
                worksheet_plano.write(0, col_num, value, header_format)

            # Aplicar formatação condicional baseada no nível de risco
            for row_num, (_, row) in enumerate(df_plano_acao.iterrows(), 1):
                nivel_risco = row["Nível de Risco"]
                if nivel_risco in risco_format:
                    worksheet_plano.write(row_num, 1, nivel_risco, risco_format[nivel_risco])

            # Adicionar validação de dados para a coluna Status
            status_options = ['Não iniciada', 'Em andamento', 'Concluída', 'Cancelada']
            worksheet_plano.data_validation('G2:G1000', {'validate': 'list',
                                         'source': status_options,
                                         'input_title': 'Selecione o status:',
                                         'input_message': 'Escolha um status da lista'})

            # Adicionar filtros
            worksheet_plano.autofilter(0, 0, len(df_plano_acao), len(df_plano_acao.columns) - 1)

            # Congelar painel para manter cabeçalhos visíveis durante rolagem
            worksheet_plano.freeze_panes(1, 0)
            
            # Aba 3: Detalhes das Questões (explicação da metodologia)
            worksheet_detalhes = workbook.add_worksheet('Detalhes das Dimensões')
            
            # Cabeçalhos
            headers = ["Dimensão", "Questões", "Descrição"]
            for col, header in enumerate(headers):
                worksheet_detalhes.write(0, col, header, workbook.add_format({'bold': True, 'bg_color': '#D7E4BC'}))
            
            # Configurar largura das colunas
            worksheet_detalhes.set_column('A:A', 20)  # Dimensão
            worksheet_detalhes.set_column('B:B', 30)  # Questões
            worksheet_detalhes.set_column('C:C', 50)  # Descrição
            
            # Preencher dados das dimensões
            DIMENSOES_HSE = st.session_state.get('DIMENSOES_HSE', {})
            DESCRICOES_DIMENSOES = st.session_state.get('DESCRICOES_DIMENSOES', {})
            
            row = 1
            for dimensao, questoes in DIMENSOES_HSE.items():
                worksheet_detalhes.write(row, 0, dimensao)
                worksheet_detalhes.write(row, 1, str(questoes))
                worksheet_detalhes.write(row, 2, DESCRICOES_DIMENSOES.get(dimensao, ""))
                row += 1
            
            # Adicionar abas para cada tipo de filtro
            for filtro in colunas_filtro:
                if filtro != "Carimbo de data/hora":
                    valores_unicos = df[filtro].dropna().unique()
                    
                    # Criar uma aba resumo para este filtro
                    sheet_name = f'Por {filtro}'
                    if len(sheet_name) > 31:  # Excel limita nomes de abas a 31 caracteres
                        sheet_name = sheet_name[:31]
                    
                    # Criar DataFrame para o resumo deste filtro
                    resultados_resumo = []
                    
                    # Para cada valor único do filtro, calcular resultados
                    for valor in valores_unicos:
                        if pd.notna(valor):
                            df_filtrado = df[df[filtro] == valor]
                            indices_filtrados = df.index[df[filtro] == valor].tolist()
                            df_perguntas_filtradas = df_perguntas.loc[indices_filtrados]
                            
                            # Importações dinâmicas para funções necessárias
                            from utils.processamento import calcular_resultados_dimensoes
                            
                            # Calcular resultados para este filtro
                            resultados_filtro = calcular_resultados_dimensoes(df_filtrado, df_perguntas_filtradas, colunas_perguntas)
                            
                            # Adicionar coluna com o valor do filtro
                            for res in resultados_filtro:
                                res[filtro] = valor
                                resultados_resumo.append(res)
                    
                    if resultados_resumo:
                        df_resumo = pd.DataFrame(resultados_resumo)
                        
                        # Remover a coluna de questões para o Excel
                        if 'Questões' in df_resumo.columns:
                            df_resumo = df_resumo.drop(columns=['Questões'])
                        
                        # Pivotear para melhor visualização
                        if len(resultados_resumo) > 0:
                            try:
                                df_pivot = df_resumo.pivot(index='Dimensão', columns=filtro, values='Média')
                                df_pivot.to_excel(writer, sheet_name=sheet_name)
                                
                                # Formatar a planilha pivotada
                                worksheet = writer.sheets[sheet_name]
                                
                                # Formatar a planilha como uma tabela pivot
                                for col in range(len(df_pivot.columns) + 1):
                                    worksheet.set_column(col, col, 15)
                                
                                # Adicionar formatação condicional para as células de dados
                                for col in range(1, len(df_pivot.columns) + 1):
                                    for row in range(1, len(df_pivot) + 1):
                                        worksheet.conditional_format(row, col, row, col, {
                                            'type': 'cell',
                                            'criteria': '<=',
                                            'value': 1,
                                            'format': risco_format['Risco Muito Alto']
                                        })
                                        worksheet.conditional_format(row, col, row, col, {
                                            'type': 'cell',
                                            'criteria': 'between',
                                            'minimum': 1.01,
                                            'maximum': 2,
                                            'format': risco_format['Risco Alto']
                                        })
                                        worksheet.conditional_format(row, col, row, col, {
                                            'type': 'cell',
                                            'criteria': 'between',
                                            'minimum': 2.01,
                                            'maximum': 3,
                                            'format': risco_format['Risco Moderado']
                                        })
                                        worksheet.conditional_format(row, col, row, col, {
                                            'type': 'cell',
                                            'criteria': 'between',
                                            'minimum': 3.01,
                                            'maximum': 4,
                                            'format': risco_format['Risco Baixo']
                                        })
                                        worksheet.conditional_format(row, col, row, col, {
                                            'type': 'cell',
                                            'criteria': '>',
                                            'value': 4,
                                            'format': risco_format['Risco Muito Baixo']
                                        })
                            except Exception as e:
                                # Fallback caso pivot falhe
                                df_resumo.to_excel(writer, sheet_name=sheet_name, index=False)
                                worksheet = writer.sheets[sheet_name]
                                for col_num, value in enumerate(df_resumo.columns.values):
                                    worksheet.write(0, col_num, value, header_format)
            
            # Adicionar aba com gráfico de resultados gerais
            worksheet_grafico = workbook.add_worksheet('Gráfico de Riscos')
            
            # Adicionar os dados para o gráfico
            worksheet_grafico.write_column('A1', ['Dimensão'] + list(df_resultados['Dimensão']))
            worksheet_grafico.write_column('B1', ['Média'] + list(df_resultados['Média']))
            worksheet_grafico.write_column('C1', ['Risco'] + list(df_resultados['Risco']))
            
            # Criar o gráfico
            chart = workbook.add_chart({'type': 'bar'})
            
            # Configurar o gráfico
            chart.add_series({
                'name': 'Média',
                'categories': ['Gráfico de Riscos', 1, 0, len(df_resultados), 0],
                'values': ['Gráfico de Riscos', 1, 1, len(df_resultados), 1],
                'data_labels': {'value': True},
                'fill': {'color': '#5A713D'}  # Cor da Escutaris
            })
            
            # Configurar aparência do gráfico
            chart.set_title({'name': 'HSE-IT: Fatores Psicossociais - Avaliação de Riscos'})
            chart.set_x_axis({'name': 'Dimensão'})
            chart.set_y_axis({'name': 'Média (1-5)', 'min': 0, 'max': 5})
            chart.set_legend({'position': 'bottom'})
            chart.set_size({'width': 720, 'height': 576})
            
            # Adicionar linhas de referência para os níveis de risco
            chart.set_y_axis({
                'name': 'Média (1-5)',
                'min': 0,
                'max': 5,
                'major_gridlines': {'visible': True},
                'minor_gridlines': {'visible': False},
                'major_unit': 1
            })
            
            # Inserir o gráfico na planilha
            worksheet_grafico.insert_chart('E1', chart, {'x_scale': 1.5, 'y_scale': 1.5})
            
            # Adicionar legenda de interpretação dos riscos
            worksheet_grafico.write('A20', 'Interpretação dos Riscos:', workbook.add_format({'bold': True}))
            worksheet_grafico.write('A21', 'Média ≤ 1: Risco Muito Alto')
            worksheet_grafico.write('A22', '1 < Média ≤ 2: Risco Alto')
            worksheet_grafico.write('A23', '2 < Média ≤ 3: Risco Moderado')
            worksheet_grafico.write('A24', '3 < Média ≤ 4: Risco Baixo')
            worksheet_grafico.write('A25', 'Média > 4: Risco Muito Baixo')
        
        output.seek(0)
        return output
    
    except Exception as e:
        st.error(f"Erro ao gerar o arquivo Excel: {str(e)}")
        return None

with col1:
    st.subheader("Relatório Excel Completo")
    st.write("Relatório detalhado com múltiplas abas, incluindo análises por filtros, gráficos e plano de ação.")
    
    # Botão para gerar Excel
    if st.button("Gerar Relatório Excel", use_container_width=True):
        with st.spinner("Gerando relatório Excel..."):
            excel_data = gerar_excel_completo(df, df_perguntas, colunas_filtro, colunas_perguntas)
            if excel_data:
                st.success("Relatório Excel gerado com sucesso!")
                st.session_state.excel_report = excel_data
                st.session_state.excel_ready = True
    
    # Botão para baixar Excel
    if st.session_state.get("excel_ready", False) and st.session_state.get("excel_report", None) is not None:
        st.download_button(
            label="Baixar Relatório Excel Completo",
            data=st.session_state.excel_report,
            file_name=f"relatorio_completo_hse_it_{filtro_opcao}_{filtro_valor}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

with col2:
    st.subheader("Relatórios PDF")
    st.write("Relatórios simplificados para compartilhamento e apresentação.")
    
    # Botões para os relatórios PDF
    col2a, col2b = st.columns(2)
    
    with col2a:
        if st.button("Gerar Relatório de Resultados", use_container_width=True):
            with st.spinner("Gerando PDF de resultados..."):
                pdf_data = gerar_pdf(df_resultados)
                if pdf_data:
                    st.success("Relatório PDF gerado com sucesso!")
                    st.session_state.pdf_report = pdf_data
                    st.session_state.pdf_ready = True
        
        # Botão para baixar PDF de resultados
        if st.session_state.get("pdf_ready", False) and st.session_state.get("pdf_report", None) is not None:
            st.download_button(
                label="Baixar PDF de Resultados",
                data=st.session_state.pdf_report,
                file_name=f"resultados_hse_it_{filtro_opcao}_{filtro_valor}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
    
    with col2b:
        if st.button("Gerar Plano de Ação PDF", use_container_width=True):
            with st.spinner("Gerando PDF do plano de ação..."):
                pdf_plano = gerar_pdf_plano_acao(df_plano_acao)
                if pdf_plano:
                    st.success("Plano de Ação PDF gerado com sucesso!")
                    st.session_state.pdf_plano = pdf_plano
                    st.session_state.plano_ready = True
        
        # Botão para baixar PDF do plano de ação
        if st.session_state.get("plano_ready", False) and st.session_state.get("pdf_plano", None) is not None:
            st.download_button(
                label="Baixar Plano de Ação PDF",
                data=st.session_state.pdf_plano,
                file_name=f"plano_acao_hse_it_{filtro_opcao}_{filtro_valor}.pdf",
                mime="application/pdf",
                use_container_width=True
            )

# Informações sobre como usar os relatórios
with st.expander("Como usar os relatórios"):
    st.write("""
    ### Relatório Excel Completo
    Contém múltiplas abas com análises detalhadas:
    - **Empresa Toda**: Visão geral dos fatores psicossociais nas 7 dimensões do HSE-IT
    - **Plano de Ação**: Sugestões de ações para cada dimensão com maior risco
    - **Detalhes das Dimensões**: Explicação sobre cada dimensão e suas questões
    - **Por Setor/Cargo/etc.**: Análises segmentadas por filtros demográficos
    - **Gráfico de Riscos**: Visualização gráfica das dimensões
    
    ### Relatório PDF de Resultados
    Versão simplificada com os principais resultados, ideal para compartilhamento com a liderança e stakeholders. Contém:
    - Tabela de resultados por dimensão
    - Descrição das dimensões do HSE-IT
    - Interpretação dos níveis de risco
    
    ### Plano de Ação PDF
    Documento específico com as ações sugeridas, contendo espaços para preenchimento manual de responsáveis e prazos. Ideal para:
    - Discussão em reuniões de planejamento
    - Acompanhamento de ações
    - Documentação do programa de saúde psicossocial
    """)
