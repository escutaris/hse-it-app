import streamlit as st
import pandas as pd
import io
import os
import tempfile
from fpdf import FPDF
from datetime import datetime
import plotly.graph_objects as go
from utils.processamento import classificar_risco

# Apply consistent Escutaris styling
def aplicar_estilo_escutaris():
    st.markdown("""
    <style>
    /* Main colors */
    :root {
        --escutaris-verde: #5A713D;
        --escutaris-cinza: #2E2F2F;
        --escutaris-bege: #F5F0EB;
        --escutaris-laranja: #FF5722;
    }
    
    /* Headings */
    h1, h2, h3 {
        color: var(--escutaris-verde) !important;
        font-family: 'Helvetica Neue', Arial, sans-serif !important;
    }
    
    /* Subheadings */
    h4, h5, h6 {
        color: var(--escutaris-cinza) !important;
        font-family: 'Helvetica Neue', Arial, sans-serif !important;
    }
    
    /* Buttons */
    .stButton>button {
        background-color: var(--escutaris-verde) !important;
        color: white !important;
        border-radius: 5px !important;
        border: none !important;
        padding: 0.5rem 1rem !important;
        transition: all 0.3s ease !important;
    }
    
    .stButton>button:hover {
        opacity: 0.9 !important;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1) !important;
    }
    
    /* Card containers */
    .report-card {
        background-color: white;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        margin-bottom: 20px;
        border-top: 4px solid var(--escutaris-verde);
    }
    
    /* Download buttons */
    .download-button {
        margin-top: 10px;
    }
    
    /* Spinner */
    .stSpinner {
        border-top-color: var(--escutaris-verde) !important;
    }
    
    /* Success message */
    .success-msg {
        padding: 10px;
        background-color: #d4edda;
        color: #155724;
        border-radius: 5px;
        margin-bottom: 15px;
    }
    
    /* Error message */
    .error-msg {
        padding: 10px;
        background-color: #f8d7da;
        color: #721c24;
        border-radius: 5px;
        margin-bottom: 15px;
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        background-color: #f5f5f5;
        border-radius: 5px;
    }
    
    /* Columns spacing */
    .report-col {
        padding: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

# Apply the styling
aplicar_estilo_escutaris()

# Page title
st.title("Relatórios - HSE-IT")

# Check if data is available
if "df_resultados" not in st.session_state or st.session_state.df_resultados is None:
    st.warning("Nenhum dado carregado ainda. Por favor, faça upload de um arquivo na página 'Upload de Dados'.")
    st.stop()

# Retrieve data from session state
try:
    df = st.session_state.df
    df_perguntas = st.session_state.df_perguntas
    colunas_filtro = st.session_state.colunas_filtro
    colunas_perguntas = st.session_state.colunas_perguntas
    df_resultados = st.session_state.df_resultados
    df_plano_acao = st.session_state.df_plano_acao
    filtro_opcao = st.session_state.filtro_opcao
    filtro_valor = st.session_state.filtro_valor
except Exception as e:
    st.error(f"Erro ao carregar dados da sessão: {str(e)}")
    st.info("Por favor, retorne à página de upload e carregue seus dados novamente.")
    st.stop()

st.header("Download de Relatórios")
st.write("Escolha abaixo o tipo de relatório que deseja gerar.")

# Function to generate the Action Plan PDF
def gerar_pdf_plano_acao(df_plano_acao):
    try:
        # Criar um PDF com suporte a caracteres UTF-8
        class PDF(FPDF):
            def __init__(self):
                super().__init__()
                # Adicionar fonte com suporte a UTF-8
                try:
                    self.add_font('DejaVu', '', 'DejaVuSans.ttf', uni=True)
                    self.add_font('DejaVu', 'B', 'DejaVuSans-Bold.ttf', uni=True)
                except Exception:
                    pass  # Fallback para fontes padrão
            
            def header(self):
                # Cabeçalho opcional para todas as páginas
                self.set_font('Arial', 'B', 10)
                self.cell(0, 10, 'Escutaris HSE Analytics - Plano de Ação', 0, 1, 'R')
                self.ln(5)
                
            def footer(self):
                # Rodapé com página
                self.set_y(-15)
                self.set_font('Arial', '', 8)
                self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')
        
        # Se não tivermos as fontes DejaVu disponíveis, criar uma alternativa mais simples
        try:
            pdf = PDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            
            # Tentar usar fonte DejaVu para caracteres especiais
            try:
                pdf.set_font("DejaVu", 'B', 16)
            except Exception:
                pdf.set_font("Arial", 'B', 16)
        except Exception:
            # Fallback para versão mais simples sem caracteres especiais
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            
            # Usar fonte padrão ASCII
            pdf.set_font("Arial", style='B', size=16)
            # Aviso para o usuário sobre caracteres especiais
            st.warning("Alguns caracteres especiais podem não aparecer corretamente no PDF. Recomendamos usar o relatório Excel para melhor visualização.")
        
        # Título - remover acentos se não tiver suporte Unicode
        try:
            pdf.cell(200, 10, "Plano de Ação - HSE-IT: Fatores Psicossociais", ln=True, align='C')
        except Exception:
            pdf.cell(200, 10, "Plano de Acao - HSE-IT: Fatores Psicossociais", ln=True, align='C')
        pdf.ln(10)
        
        # Data do relatório
        pdf.set_font("Arial", size=10)
        data_atual = datetime.now().strftime("%d/%m/%Y")
        pdf.cell(0, 10, f"Data do relatorio: {data_atual}", ln=True)
        pdf.ln(5)
        
        # Função para remover acentos
        import unicodedata
        def remover_acentos(texto):
            try:
                # Normaliza para forma NFD e remove diacríticos
                return ''.join(c for c in unicodedata.normalize('NFD', texto)
                            if unicodedata.category(c) != 'Mn')
            except Exception:
                return texto.encode('ascii', 'replace').decode('ascii')
        
        # Group by dimension
        for dimensao in df_plano_acao["Dimensão"].unique():
            df_dimensao = df_plano_acao[df_plano_acao["Dimensão"] == dimensao]
            nivel_risco = df_dimensao["Nível de Risco"].iloc[0]
            media = df_dimensao["Média"].iloc[0]
            
            # Tratar texto para compatibilidade
            dimensao_safe = remover_acentos(dimensao)
            nivel_risco_safe = remover_acentos(nivel_risco)
            
            # Dimension title
            pdf.set_font("Arial", style='B', size=12)
            pdf.cell(0, 10, f"{dimensao_safe}", ln=True)
            
            # Risk information
            pdf.set_font("Arial", style='I', size=10)
            pdf.cell(0, 8, f"Media: {media} - Nivel de Risco: {nivel_risco_safe}", ln=True)
            
            # Suggested actions
            pdf.set_font("Arial", size=10)
            pdf.cell(0, 6, "Acoes Sugeridas:", ln=True)
            
            for _, row in df_dimensao.iterrows():
                pdf.set_x(15)  # Indentation for list
                sugestao = row['Sugestão de Ação']
                
                # Tratar texto para compatibilidade
                sugestao_safe = remover_acentos(sugestao)
                
                pdf.multi_cell(0, 6, f"- {sugestao_safe}", 0)
            
            # Space for implementation table
            pdf.ln(5)
            pdf.cell(0, 6, "Implementacao:", ln=True)
            
            # Create table
            col_width = [100, 40, 40]
            pdf.set_font("Arial", style='B', size=8)
            
            # Table header
            pdf.set_x(15)
            pdf.cell(col_width[0], 7, "Acao", border=1)
            pdf.cell(col_width[1], 7, "Responsavel", border=1)
            pdf.cell(col_width[2], 7, "Prazo", border=1)
            pdf.ln()
            
            # Table rows
            pdf.set_font("Arial", size=8)
            for _, row in df_dimensao.iterrows():
                pdf.set_x(15)
                sugestao = row['Sugestão de Ação']
                
                # Tratar texto para compatibilidade
                sugestao_safe = remover_acentos(sugestao)
                
                # Check if text is too long
                if len(sugestao_safe) > 60:
                    sugestao_safe = sugestao_safe[:57] + "..."
                
                pdf.cell(col_width[0], 7, sugestao_safe, border=1)
                pdf.cell(col_width[1], 7, "", border=1)
                pdf.cell(col_width[2], 7, "", border=1)
                pdf.ln()
            
            pdf.ln(10)
        
        # Add risk interpretation information
        pdf.ln(10)
        pdf.set_font("Arial", style='B', size=11)
        pdf.cell(0, 8, "Legenda de Classificacao de Riscos:", ln=True)
        pdf.set_font("Arial", size=9)
        pdf.cell(0, 6, "Risco Muito Alto: Media <= 1", ln=True)
        pdf.cell(0, 6, "Risco Alto: 1 < Media <= 2", ln=True)
        pdf.cell(0, 6, "Risco Moderado: 2 < Media <= 3", ln=True)
        pdf.cell(0, 6, "Risco Baixo: 3 < Media <= 4", ln=True)
        pdf.cell(0, 6, "Risco Muito Baixo: Media > 4", ln=True)
        
        # Add observations about the action plan
        pdf.ln(10)
        pdf.set_font("Arial", style='B', size=11)
        pdf.cell(0, 8, "Observacoes:", ln=True)
        pdf.set_font("Arial", size=9)
        pdf.multi_cell(0, 6, "Este plano de acao foi gerado automaticamente com base nos resultados da avaliacao HSE-IT. As acoes sugeridas devem ser analisadas e adaptadas ao contexto especifico da organizacao. Recomenda-se definir responsaveis, prazos e indicadores de monitoramento para cada acao implementada.", 0)
        
        # Usar arquivo temporário para evitar problemas com BytesIO
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_file:
            temp_filename = temp_file.name
        
        # Salvar para arquivo temporário
        pdf.output(temp_filename)
        
        # Ler o arquivo temporário para retornar
        with open(temp_filename, "rb") as file:
            output = io.BytesIO(file.read())
        
        # Limpar arquivo temporário
        os.unlink(temp_filename)
        
        return output
    
    except Exception as e:
        st.error(f"Erro ao gerar o PDF do Plano de Ação: {str(e)}")
        
        # Mais informações de erro para debugging
        import traceback
        st.code(traceback.format_exc())
        
        return None

# Function to generate results PDF report
def gerar_pdf(df_resultados):
    try:
        # Criar um PDF com suporte a caracteres UTF-8
        class PDF(FPDF):
            def __init__(self):
                super().__init__()
                # Tentar usar fonte com suporte a UTF-8
                try:
                    self.add_font('DejaVu', '', 'DejaVuSans.ttf', uni=True)
                    self.add_font('DejaVu', 'B', 'DejaVuSans-Bold.ttf', uni=True)
                except Exception:
                    pass  # Fallback para fontes padrão
            
            def header(self):
                # Cabeçalho opcional
                self.set_font('Arial', 'B', 10)
                self.cell(0, 10, 'Escutaris HSE Analytics', 0, 1, 'R')
                self.ln(5)
                
            def footer(self):
                # Rodapé com página
                self.set_y(-15)
                self.set_font('Arial', '', 8)
                self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')
        
        # Tentar criar PDF com fontes avançadas
        try:
            pdf = PDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            
            # Tentar usar a fonte DejaVu para melhor suporte a caracteres especiais
            try:
                pdf.set_font("DejaVu", 'B', 14)
            except Exception:
                pdf.set_font("Arial", style='B', size=14)
        except Exception:
            # Fallback para versão mais simples
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            pdf.set_font("Arial", style='B', size=14)
        
        # Título - sem acentos para compatibilidade
        pdf.cell(200, 10, "Relatorio de Fatores Psicossociais - HSE-IT", ln=True, align='C')
        pdf.ln(10)
        
        # Add methodology explanation
        pdf.set_font("Arial", style='I', size=10)
        pdf.multi_cell(0, 5, "O questionario HSE-IT avalia 7 dimensoes de fatores psicossociais no trabalho. Os resultados sao apresentados em uma escala de 1 a 5, onde valores mais altos indicam melhores resultados.", 0)
        pdf.ln(5)
        
        # Função para remover acentos
        import unicodedata
        def remover_acentos(texto):
            try:
                # Normaliza para forma NFD e remove diacríticos
                return ''.join(c for c in unicodedata.normalize('NFD', texto)
                            if unicodedata.category(c) != 'Mn')
            except Exception:
                return texto.encode('ascii', 'replace').decode('ascii')
        
        # Results table
        pdf.set_font("Arial", style='B', size=10)
        pdf.cell(80, 7, "Dimensao", 1)
        pdf.cell(25, 7, "Media", 1)
        pdf.cell(60, 7, "Nivel de Risco", 1)
        pdf.ln()
        
        pdf.set_font("Arial", size=10)
        for _, row in df_resultados.iterrows():
            # Remover acentos e caracteres especiais para compatibilidade
            dimensao = remover_acentos(row['Dimensão'])
            risco = row['Risco'].split(' ')[0] + ' ' + row['Risco'].split(' ')[1]  # Remove emoji
            
            # Determine background color based on risk
            if "Muito Alto" in risco:
                pdf.set_fill_color(255, 107, 107)  # Light red
                text_color = 255  # White
            elif "Alto" in risco:
                pdf.set_fill_color(255, 165, 0)  # Orange
                text_color = 0  # Black
            elif "Moderado" in risco:
                pdf.set_fill_color(255, 255, 0)  # Yellow
                text_color = 0  # Black
            elif "Baixo" in risco:
                pdf.set_fill_color(144, 238, 144)  # Light green
                text_color = 0  # Black
            elif "Muito Baixo" in risco:
                pdf.set_fill_color(187, 143, 206)  # Light purple
                text_color = 0  # Black
            else:
                pdf.set_fill_color(255, 255, 255)  # White
                text_color = 0  # Black
                
            pdf.set_text_color(text_color)
            pdf.cell(80, 7, dimensao, 1, 0, 'L', 1)
            pdf.cell(25, 7, f"{row['Média']:.2f}", 1, 0, 'C', 1)
            pdf.cell(60, 7, risco, 1, 0, 'C', 1)
            pdf.ln()
        
        # Reset text color
        pdf.set_text_color(0)
        
        # Add dimension descriptions
        pdf.ln(10)
        pdf.set_font("Arial", style='B', size=12)
        pdf.cell(0, 10, "Descricao das Dimensoes:", ln=True)
        
        for dimensao, descricao in st.session_state.get('DESCRICOES_DIMENSOES', {}).items():
            pdf.set_font("Arial", style='B', size=10)
            pdf.cell(0, 7, remover_acentos(dimensao), ln=True)
            pdf.set_font("Arial", size=10)
            pdf.multi_cell(0, 5, remover_acentos(descricao), 0)
            pdf.ln(3)
        
        # Add risk classification information
        pdf.ln(10)
        pdf.set_font("Arial", style='B', size=12)
        pdf.cell(0, 10, "Legenda de Classificacao:", ln=True)
        pdf.set_font("Arial", size=10)
        pdf.cell(0, 8, "Risco Muito Alto: Media <= 1", ln=True)
        pdf.cell(0, 8, "Risco Alto: 1 < Media <= 2", ln=True)
        pdf.cell(0, 8, "Risco Moderado: 2 < Media <= 3", ln=True)
        pdf.cell(0, 8, "Risco Baixo: 3 < Media <= 4", ln=True)
        pdf.cell(0, 8, "Risco Muito Baixo: Media > 4", ln=True)
        
        # Add general recommendations
        pdf.ln(10)
        pdf.set_font("Arial", style='B', size=12)
        pdf.cell(0, 10, "Recomendacoes Gerais:", ln=True)
        pdf.set_font("Arial", size=10)
        pdf.multi_cell(0, 5, "Para resultados mais detalhados e um plano de acao personalizado, consulte o arquivo Excel completo ou o documento do Plano de Acao fornecido. Priorize intervencoes nas dimensoes com maior nivel de risco.", 0)
        
        # Usar arquivo temporário para evitar problemas
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_file:
            temp_filename = temp_file.name
        
        # Salvar para arquivo temporário
        pdf.output(temp_filename)
        
        # Ler o arquivo temporário
        with open(temp_filename, "rb") as file:
            output = io.BytesIO(file.read())
        
        # Limpar arquivo temporário
        os.unlink(temp_filename)
        
        return output
    
    except Exception as e:
        st.error(f"Erro ao gerar o PDF: {str(e)}")
        
        # Mais informações de erro para debugging
        import traceback
        st.code(traceback.format_exc())
        
        return None

# Function to generate Excel report with multiple sheets
def gerar_excel_completo(df, df_perguntas, colunas_filtro, colunas_perguntas):
    output = io.BytesIO()
    
    try:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Sheet 1: Overall company results
            df_resultados_excel = st.session_state.df_resultados.copy()
            if 'Questões' in df_resultados_excel.columns:
                df_resultados_excel = df_resultados_excel.drop(columns=['Questões'])
                
            df_resultados_excel.to_excel(writer, sheet_name='Empresa Toda', index=False)
            
            # Format the worksheet
            worksheet = writer.sheets['Empresa Toda']
            
            # Define formats
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })

            # Configure column widths
            worksheet.set_column('A:A', 5)  # Index
            worksheet.set_column('B:B', 25)  # Dimension
            worksheet.set_column('C:C', 40)  # Description
            worksheet.set_column('D:D', 10)  # Average
            worksheet.set_column('E:E', 20)  # Risk
            worksheet.set_column('F:F', 15)  # Number of Responses
            
            # Add formatted headers
            for col_num, value in enumerate(df_resultados_excel.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Add filters
            worksheet.autofilter(0, 0, len(df_resultados_excel), len(df_resultados_excel.columns) - 1)
            
            # Freeze header row
            worksheet.freeze_panes(1, 0)
            
            # Sheet 2: Action Plan
            df_plano_acao.to_excel(writer, sheet_name='Plano de Ação', index=False)
            
            # Format action plan sheet
            worksheet_plano = writer.sheets['Plano de Ação']
            
            # Define risk formats
            risco_format = {
                'Risco Muito Alto': workbook.add_format({'bg_color': '#FF6B6B', 'font_color': 'white'}),
                'Risco Alto': workbook.add_format({'bg_color': '#FFA500'}),
                'Risco Moderado': workbook.add_format({'bg_color': '#FFFF00'}),
                'Risco Baixo': workbook.add_format({'bg_color': '#90EE90'}),
                'Risco Muito Baixo': workbook.add_format({'bg_color': '#BB8FCE'})
            }

            # Configure column widths
            worksheet_plano.set_column('A:A', 25)  # Dimension
            worksheet_plano.set_column('B:B', 15)  # Risk Level
            worksheet_plano.set_column('C:C', 10)  # Average
            worksheet_plano.set_column('D:D', 50)  # Action Suggestion
            worksheet_plano.set_column('E:E', 15)  # Responsible
            worksheet_plano.set_column('F:F', 15)  # Deadline
            worksheet_plano.set_column('G:G', 15)  # Status

            # Add formatted headers
            for col_num, value in enumerate(df_plano_acao.columns.values):
                worksheet_plano.write(0, col_num, value, header_format)

            # Apply conditional formatting based on risk level
            for row_num, (_, row) in enumerate(df_plano_acao.iterrows(), 1):
                nivel_risco = row["Nível de Risco"]
                if nivel_risco in risco_format:
                    worksheet_plano.write(row_num, 1, nivel_risco, risco_format[nivel_risco])

            # Add data validation for Status column
            status_options = ['Não iniciada', 'Em andamento', 'Concluída', 'Cancelada']
            worksheet_plano.data_validation('G2:G1000', {'validate': 'list',
                                         'source': status_options,
                                         'input_title': 'Selecione o status:',
                                         'input_message': 'Escolha um status da lista'})

            # Add filters
            worksheet_plano.autofilter(0, 0, len(df_plano_acao), len(df_plano_acao.columns) - 1)

            # Freeze header row
            worksheet_plano.freeze_panes(1, 0)
            
            # Sheet 3: Dimension Details (methodology explanation)
            worksheet_detalhes = workbook.add_worksheet('Detalhes das Dimensões')
            
            # Headers
            headers = ["Dimensão", "Questões", "Descrição"]
            for col, header in enumerate(headers):
                worksheet_detalhes.write(0, col, header, workbook.add_format({'bold': True, 'bg_color': '#D7E4BC'}))
            
            # Configure column widths
            worksheet_detalhes.set_column('A:A', 20)  # Dimension
            worksheet_detalhes.set_column('B:B', 30)  # Questions
            worksheet_detalhes.set_column('C:C', 50)  # Description
            
            # Fill dimension data
            DIMENSOES_HSE = st.session_state.get('DIMENSOES_HSE', {})
            DESCRICOES_DIMENSOES = st.session_state.get('DESCRICOES_DIMENSOES', {})
            
            row = 1
            for dimensao, questoes in DIMENSOES_HSE.items():
                worksheet_detalhes.write(row, 0, dimensao)
                worksheet_detalhes.write(row, 1, str(questoes))
                worksheet_detalhes.write(row, 2, DESCRICOES_DIMENSOES.get(dimensao, ""))
                row += 1
            
            # Add sheets for each filter type
            for filtro in colunas_filtro:
                if filtro != "Carimbo de data/hora":
                    valores_unicos = df[filtro].dropna().unique()
                    
                    # Create a summary sheet for this filter
                    sheet_name = f'Por {filtro}'
                    if len(sheet_name) > 31:  # Excel limits sheet names to 31 characters
                        sheet_name = sheet_name[:31]
                    
                    # Create a DataFrame for this filter summary
                    resultados_resumo = []
                    
                    # For each unique value, calculate results
                    for valor in valores_unicos:
                        if pd.notna(valor):
                            df_filtrado = df[df[filtro] == valor]
                            indices_filtrados = df.index[df[filtro] == valor].tolist()
                            df_perguntas_filtradas = df_perguntas.loc[indices_filtrados]
                            
                            # Dynamic imports for required functions
                            from utils.processamento import calcular_resultados_dimensoes
                            
                            # Calculate results for this filter
                            resultados_filtro = calcular_resultados_dimensoes(df_filtrado, df_perguntas_filtradas, colunas_perguntas)
                            
                            # Add a column with filter value
                            for res in resultados_filtro:
                                res[filtro] = valor
                                resultados_resumo.append(res)
                    
                    if resultados_resumo:
                        df_resumo = pd.DataFrame(resultados_resumo)
                        
                        # Remove questions column for Excel output
                        if 'Questões' in df_resumo.columns:
                            df_resumo = df_resumo.drop(columns=['Questões'])
                        
                        # Pivot for better visualization
                        if len(resultados_resumo) > 0:
                            try:
                                df_pivot = df_resumo.pivot(index='Dimensão', columns=filtro, values='Média')
                                df_pivot.to_excel(writer, sheet_name=sheet_name)
                                
                                # Format pivot worksheet
                                worksheet = writer.sheets[sheet_name]
                                
                                # Format as a pivot table
                                for col in range(len(df_pivot.columns) + 1):
                                    worksheet.set_column(col, col, 15)
                                
                                # Add conditional formatting for data cells
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
                                # Fallback if pivot fails
                                df_resumo.to_excel(writer, sheet_name=sheet_name, index=False)
                                worksheet = writer.sheets[sheet_name]
                                for col_num, value in enumerate(df_resumo.columns.values):
                                    worksheet.write(0, col_num, value, header_format)
            
            # Add sheet with chart of overall results
            worksheet_grafico = workbook.add_worksheet('Gráfico de Riscos')
            
            # Add data for the chart
            worksheet_grafico.write_column('A1', ['Dimensão'] + list(df_resultados['Dimensão']))
            worksheet_grafico.write_column('B1', ['Média'] + list(df_resultados['Média']))
            worksheet_grafico.write_column('C1', ['Risco'] + list(df_resultados['Risco']))
            
            # Create chart
            chart = workbook.add_chart({'type': 'bar'})
            
            # Configure chart
            chart.add_series({
                'name': 'Média',
                'categories': ['Gráfico de Riscos', 1, 0, len(df_resultados), 0],
                'values': ['Gráfico de Riscos', 1, 1, len(df_resultados), 1],
                'data_labels': {'value': True},
                'fill': {'color': '#5A713D'}  # Escutaris color
            })
            
            # Configure chart appearance
            chart.set_title({'name': 'HSE-IT: Fatores Psicossociais - Avaliação de Riscos'})
            chart.set_x_axis({'name': 'Dimensão'})
            chart.set_y_axis({'name': 'Média (1-5)', 'min': 0, 'max': 5})
            chart.set_legend({'position': 'bottom'})
            chart.set_size({'width': 720, 'height': 576})
            
            # Add reference lines for risk levels
            chart.set_y_axis({
                'name': 'Média (1-5)',
                'min': 0,
                'max': 5,
                'major_gridlines': {'visible': True},
                'minor_gridlines': {'visible': False},
                'major_unit': 1
            })
            
            # Insert chart into worksheet
            worksheet_grafico.insert_chart('E1', chart, {'x_scale': 1.5, 'y_scale': 1.5})
            
            # Add risk interpretation legend
            worksheet_grafico.write('A20', 'Interpretação dos Riscos:', workbook.add_format({'bold': True}))
            worksheet_grafico.write('A21', 'Média ≤ 1: Risco Muito Alto')
            worksheet_grafico.write('A22', '1 < Média ≤ 2: Risco Alto')
            worksheet_grafico.write('A23', '2 < Média ≤ 3: Risco Moderado')
            worksheet_grafico.write('A24', '3 < Média ≤ 4: Risco Baixo')
            worksheet_grafico.write('A25', 'Média > 4: Risco Muito Baixo')
            
            # Add an executive summary sheet
            worksheet_resumo = workbook.add_worksheet('Resumo Executivo')
            
            # Title and introduction
            worksheet_resumo.merge_range('A1:F1', 'RELATÓRIO EXECUTIVO - AVALIAÇÃO HSE-IT', 
                                        workbook.add_format({
                                            'bold': True, 
                                            'font_size': 16, 
                                            'align': 'center',
                                            'valign': 'vcenter',
                                            'bg_color': '#5A713D',
                                            'font_color': 'white'
                                        }))
            
            # Add date and filter info
            date_format = workbook.add_format({'bold': True, 'align': 'right'})
            worksheet_resumo.write('F3', f'Data: {datetime.now().strftime("%d/%m/%Y")}', date_format)
            worksheet_resumo.write('F4', f'Filtro: {filtro_opcao} - {filtro_valor}', date_format)
            
            # Section header
            section_format = workbook.add_format({
                'bold': True, 
                'bg_color': '#D7E4BC',
                'border': 1,
                'font_size': 12
            })
            
            # Add key findings section
            worksheet_resumo.merge_range('A6:F6', 'PRINCIPAIS DESCOBERTAS', section_format)
            
            # Calculate key metrics
            media_geral = df_resultados["Média"].mean()
            _, cor_geral = classificar_risco(media_geral)
            dimensao_mais_critica = df_resultados.loc[df_resultados["Média"].idxmin()]
            dimensao_melhor = df_resultados.loc[df_resultados["Média"].idxmax()]
            
            # Add key metrics data
            row = 7
            worksheet_resumo.write(row, 0, 'Média Geral:', workbook.add_format({'bold': True}))
            worksheet_resumo.write(row, 1, f"{media_geral:.2f}")
            row += 1
            
            worksheet_resumo.write(row, 0, 'Dimensão Mais Crítica:', workbook.add_format({'bold': True}))
            worksheet_resumo.write(row, 1, f"{dimensao_mais_critica['Dimensão']} ({dimensao_mais_critica['Média']:.2f})")
            row += 1
            
            worksheet_resumo.write(row, 0, 'Dimensão Melhor Avaliada:', workbook.add_format({'bold': True}))
            worksheet_resumo.write(row, 1, f"{dimensao_melhor['Dimensão']} ({dimensao_melhor['Média']:.2f})")
            row += 2
            
            # Add all dimensions results
            worksheet_resumo.merge_range(f'A{row}:F{row}', 'RESULTADOS POR DIMENSÃO', section_format)
            row += 1
            
            # Headers for results table
            worksheet_resumo.write(row, 0, 'Dimensão', header_format)
            worksheet_resumo.write(row, 1, 'Média', header_format)
            worksheet_resumo.write(row, 2, 'Nível de Risco', header_format)
            worksheet_resumo.write(row, 3, 'Prioridade de Ação', header_format)
            row += 1
            
            # Sort dimensions by priority (lowest average first)
            df_sorted = df_resultados.sort_values(by="Média")
            
            # Write dimension data
            for i, (_, dim_row) in enumerate(df_sorted.iterrows()):
                worksheet_resumo.write(row + i, 0, dim_row["Dimensão"])
                worksheet_resumo.write(row + i, 1, dim_row["Média"])
                
                # Style the risk level cell based on risk
                nivel_risco = dim_row["Risco"].split(" ")[0] + " " + dim_row["Risco"].split(" ")[1]
                if nivel_risco in risco_format:
                    worksheet_resumo.write(row + i, 2, nivel_risco, risco_format[nivel_risco])
                else:
                    worksheet_resumo.write(row + i, 2, nivel_risco)
                
                # Set priority based on risk level
                priority_map = {
                    "Risco Muito Alto": "Imediata",
                    "Risco Alto": "Alta",
                    "Risco Moderado": "Média",
                    "Risco Baixo": "Baixa",
                    "Risco Muito Baixo": "Manutenção"
                }
                priority = priority_map.get(nivel_risco, "Não definida")
                worksheet_resumo.write(row + i, 3, priority)
            
            row += len(df_sorted) + 2
            
            # Add recommendations section
            worksheet_resumo.merge_range(f'A{row}:F{row}', 'RECOMENDAÇÕES', section_format)
            row += 1
            
            # Add general recommendations
            worksheet_resumo.merge_range(f'A{row}:F{row}', 'Com base nos resultados da avaliação HSE-IT, recomenda-se:')
            row += 2
            
            # Add specific recommendations for high risk dimensions
            high_risk_dims = df_sorted[df_sorted["Média"] <= 2]
            if not high_risk_dims.empty:
                for i, (_, dim_row) in enumerate(high_risk_dims.iterrows()):
                    dim_name = dim_row["Dimensão"]
                    worksheet_resumo.write(row + i, 0, f"{i+1}. Priorizar ações para a dimensão {dim_name}:")
                    
                    # Add a sample recommendation based on dimension
                    if "df_plano_acao" in st.session_state:
                        sample_actions = st.session_state.df_plano_acao[
                            st.session_state.df_plano_acao["Dimensão"] == dim_name
                        ]["Sugestão de Ação"].head(1).tolist()
                        
                        if sample_actions:
                            worksheet_resumo.merge_range(f'B{row+i}:F{row+i}', sample_actions[0])
            else:
                worksheet_resumo.write(row, 0, "1. Manter as boas práticas atuais, com foco em melhorias contínuas.")
                row += 1
                worksheet_resumo.write(row, 0, "2. Revisar periodicamente os fatores psicossociais para garantir que permaneçam em níveis saudáveis.")
        
        output.seek(0)
        return output
    
    except Exception as e:
        st.error(f"Erro ao gerar o arquivo Excel: {str(e)}")
        
        # More detailed error information for debugging
        import traceback
        st.code(traceback.format_exc())
        
        return None

# Create layout with two report type columns
st.markdown('<div style="display: flex; flex-wrap: wrap; gap: 20px;">', unsafe_allow_html=True)

# Excel Report Card
st.markdown('<div class="report-card" style="flex: 1; min-width: 300px;">', unsafe_allow_html=True)
st.subheader("Relatório Excel Completo")
st.write("""
Este relatório detalhado contém múltiplas abas com análises completas, incluindo:

* Visão geral da empresa
* Análises por filtros demográficos
* Plano de ação detalhado
* Gráficos e visualizações
* Resumo executivo
""")

# Button to generate Excel report
if st.button("Gerar Relatório Excel", key="gen_excel", use_container_width=True):
    with st.spinner("Gerando relatório Excel completo..."):
        excel_data = gerar_excel_completo(df, df_perguntas, colunas_filtro, colunas_perguntas)
        if excel_data:
            st.success("Relatório Excel gerado com sucesso!")
            st.session_state.excel_report = excel_data
            st.session_state.excel_ready = True

# Download Excel button
if st.session_state.get("excel_ready", False) and st.session_state.get("excel_report", None) is not None:
    st.download_button(
        label="Baixar Relatório Excel",
        data=st.session_state.excel_report,
        file_name=f"relatorio_completo_hse_it_{filtro_opcao}_{filtro_valor}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="dl_excel"
    )

st.markdown('</div>', unsafe_allow_html=True)

# PDF Reports Card
st.markdown('<div class="report-card" style="flex: 1; min-width: 300px;">', unsafe_allow_html=True)
st.subheader("Relatórios PDF")
st.write("""
Relatórios resumidos em formato PDF, ideais para apresentações e compartilhamento:

* **Relatório de Resultados**: Resumo dos resultados por dimensão
* **Plano de Ação**: Documento com ações sugeridas e campos para preenchimento
""")

# Create two columns for PDF reports
col1, col2 = st.columns(2)

# Results PDF
with col1:
    if st.button("Gerar Relatório de Resultados", key="gen_results", use_container_width=True):
        with st.spinner("Gerando PDF de resultados..."):
            pdf_data = gerar_pdf(df_resultados)
            if pdf_data:
                st.success("Relatório PDF gerado com sucesso!")
                st.session_state.pdf_report = pdf_data
                st.session_state.pdf_ready = True
    
    # Download PDF button
    if st.session_state.get("pdf_ready", False) and st.session_state.get("pdf_report", None) is not None:
        st.download_button(
            label="Baixar PDF de Resultados",
            data=st.session_state.pdf_report,
            file_name=f"resultados_hse_it_{filtro_opcao}_{filtro_valor}.pdf",
            mime="application/pdf",
            use_container_width=True,
            key="dl_results"
        )

# Action Plan PDF
with col2:
    if st.button("Gerar Plano de Ação PDF", key="gen_plan", use_container_width=True):
        with st.spinner("Gerando PDF do plano de ação..."):
            pdf_plano = gerar_pdf_plano_acao(df_plano_acao)
            if pdf_plano:
                st.success("Plano de Ação PDF gerado com sucesso!")
                st.session_state.pdf_plano = pdf_plano
                st.session_state.plano_ready = True
    
    # Download Action Plan button
    if st.session_state.get("plano_ready", False) and st.session_state.get("pdf_plano", None) is not None:
        st.download_button(
            label="Baixar Plano de Ação PDF",
            data=st.session_state.pdf_plano,
            file_name=f"plano_acao_hse_it_{filtro_opcao}_{filtro_valor}.pdf",
            mime="application/pdf",
            use_container_width=True,
            key="dl_plan"
        )

st.markdown('</div>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)  # Close flex container

# Additional Information
with st.expander("Como usar os relatórios", expanded=False):
    st.write("""
    ### Relatório Excel Completo
    Este relatório abrangente é o mais completo e contém:
    - **Empresa Toda**: Visão geral dos fatores psicossociais nas 7 dimensões do HSE-IT
    - **Plano de Ação**: Sugestões de ações para cada dimensão com maior risco
    - **Detalhes das Dimensões**: Explicação sobre cada dimensão e suas questões
    - **Por Setor/Cargo/etc.**: Análises segmentadas por filtros demográficos
    - **Gráfico de Riscos**: Visualização gráfica das dimensões
    - **Resumo Executivo**: Síntese dos principais resultados e recomendações
    
    Ideal para análises detalhadas e acompanhamento do plano de ação ao longo do tempo.
    
    ### Relatório PDF de Resultados
    Versão simplificada com os principais resultados, ideal para compartilhamento com a liderança e stakeholders. Contém:
    - Tabela de resultados por dimensão
    - Descrição das dimensões do HSE-IT
    - Interpretação dos níveis de risco
    
    Perfeito para reuniões de apresentação dos resultados e sensibilização da liderança.
    
    ### Plano de Ação PDF
    Documento específico com as ações sugeridas, contendo espaços para preenchimento manual de responsáveis e prazos. Ideal para:
    - Discussão em reuniões de planejamento
    - Acompanhamento de ações
    - Documentação do programa de saúde psicossocial
    
    Excelente para workshops de definição de ações e atribuição de responsabilidades.
    """)

# Footer with Escutaris info
st.divider()
st.markdown("""
<div style="text-align: center; color: #666;">
<p>Relatório gerado pela plataforma HSE-IT Analytics da Escutaris</p>
<p><small>Para mais informações, visite <a href="https://escutaris.com.br">escutaris.com.br</a> ou entre em contato pelo email <a href="mailto:contato@escutaris.com.br">contato@escutaris.com.br</a></small></p>
</div>
""", unsafe_allow_html=True)
