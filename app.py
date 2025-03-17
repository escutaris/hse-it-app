import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
from fpdf import FPDF

# Função para classificar os riscos com base na pontuação média
def classificar_risco(media):
    if media <= 1:
        return "Risco Muito Alto 🔴", "red"
    elif media > 1 and media <= 2:
        return "Risco Alto 🟠", "orange"
    elif media > 2 and media <= 3:
        return "Risco Moderado 🟡", "yellow"
    elif media > 3 and media <= 4:
        return "Risco Baixo 🟢", "green"
    else:
        return "Risco Muito Baixo 🟣", "purple"

# Função para gerar o arquivo Excel com os resultados
def gerar_excel(df_resultados):
    output = io.BytesIO()
    try:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_resultados.to_excel(writer, sheet_name='Resultados', index=False)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Erro ao gerar o arquivo Excel: {str(e)}")
        return None

# Função para gerar o arquivo PDF com os resultados
def gerar_pdf(df_resultados):
    try:
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        
        # Configurando fonte para suporte básico
        pdf.set_font("Arial", style='B', size=14)
        
        # Usando strings sem acentos para evitar problemas de codificação
        pdf.cell(200, 10, "Relatorio de Fatores Psicossociais (HSE-IT)", ln=True, align='C')
        pdf.ln(10)
        
        pdf.set_font("Arial", size=12)
        for index, row in df_resultados.iterrows():
            # Remover acentos e caracteres especiais
            fator = row['Fator Psicossocial'].encode('ascii', 'replace').decode('ascii')
            risco = row['Risco'].split(' ')[0] + ' ' + row['Risco'].split(' ')[1]  # Remove emoji
            pdf.cell(0, 10, f"{fator}: {risco}", ln=True)
        
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
        
        output = io.BytesIO()
        pdf.output(output)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Erro ao gerar o PDF: {str(e)}")
        return None

# Interface do aplicativo no Streamlit
st.title("Avaliação de Fatores Psicossociais - HSE-IT")
st.write("Faça upload do arquivo Excel contendo os resultados do questionário.")

# Adiciona uma explicação sobre o formato esperado do arquivo
with st.expander("Informações sobre o formato do arquivo"):
    st.write("""
    O arquivo deve conter:
    - Colunas de filtro (Setor, Cargo, etc.) nas primeiras 7 colunas
    - Colunas com as perguntas numeradas (começando com números)
    - As respostas devem ser valores numéricos de 1 a 5
    """)

uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=["xlsx", "csv"])

if uploaded_file is not None:
    try:
        # Determinar o tipo de arquivo e carregá-lo adequadamente
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, sep=';' if ';' in uploaded_file.getvalue().decode('utf-8', errors='replace')[:1000] else ',')
        else:
            df = pd.read_excel(uploaded_file)
        
        # Verificar se o DataFrame tem dados suficientes
        if df.empty or len(df.columns) < 8:  # 7 colunas de filtro + pelo menos 1 pergunta
            st.error("O arquivo não contém dados suficientes ou está em formato incorreto.")
            st.stop()
        
        # Separar os dados de filtro e as perguntas
        df_filtros = df.iloc[:, :7]  # Colunas com Setor, Cargo, etc.
        
        # Garantir que as colunas das perguntas são corretamente identificadas
        colunas_perguntas = [col for col in df.columns if str(col).strip() and str(col).strip()[0].isdigit()]
        
        if not colunas_perguntas:
            st.error("Erro: Nenhuma pergunta válida foi encontrada no arquivo enviado. As colunas de perguntas devem começar com números.")
            st.stop()
        
        df_perguntas = df[colunas_perguntas]  # Agora contém apenas as perguntas do HSE-IT
        
        # Converter valores para numéricos, tratando erros
        for col in df_perguntas.columns:
            df_perguntas[col] = pd.to_numeric(df_perguntas[col], errors='coerce')
        
        # Verificar se há valores ausentes após a conversão
        missing_values = df_perguntas.isna().sum().sum()
        if missing_values > 0:
            st.warning(f"Atenção: {missing_values} valores não numéricos ou ausentes foram encontrados e serão ignorados nos cálculos.")
        
        fatores = {
            "Gestão organizacional": [15, 19, 25, 26, 28, 32],
            "Violência e assédio moral/sexual no trabalho": [5, 14, 21, 34],
            "Contexto da organização do trabalho": [3, 6, 9, 1, 4, 11, 13, 8, 27, 31, 35],
            "Características das relações sociais no trabalho": [7, 24, 27, 31, 5, 14, 21, 34],
            "Conteúdo das tarefas do trabalho": [6, 9, 12, 18, 20, 22],
            "Discriminação no trabalho": [14, 21],
            "Condições do ambiente de trabalho": [5, 6, 14, 21, 34],
            "Interação pessoa-tarefa": [10, 19, 6, 9, 12, 20],
            "Jornada de trabalho": [16, 18, 30]
        }
        
        # Mostrar informações básicas sobre os dados
        st.write(f"**Dados carregados com sucesso!** Total de {len(df)} respostas.")
        
        # Criar opções de filtro
        col1, col2 = st.columns(2)
        with col1:
            filtro_opcao = st.selectbox("Filtrar por", ["Empresa Toda"] + list(df_filtros.columns))
        
        filtro_valor = "Geral"
        df_filtrado = df
        
        if filtro_opcao != "Empresa Toda":
            with col2:
                valores_unicos = df_filtros[filtro_opcao].dropna().unique()
                if len(valores_unicos) > 0:
                    filtro_valor = st.selectbox(f"Escolha um {filtro_opcao}", valores_unicos)
                    df_filtrado = df[df_filtros[filtro_opcao] == filtro_valor]
                    df_perguntas = df_perguntas[df_filtros[filtro_opcao] == filtro_valor]
                else:
                    st.warning(f"Não há valores únicos para o filtro '{filtro_opcao}'.")
        
        # Verificar se há dados após a filtragem
        if df_filtrado.empty:
            st.error(f"Nenhum dado encontrado para {filtro_opcao}: {filtro_valor}")
            st.stop()
        
        resultados = []
        cores = []
        
        for fator, perguntas in fatores.items():
            # Converter números de pergunta para índices de coluna (se possível)
            indices_validos = []
            for p in perguntas:
                for col in colunas_perguntas:
                    if str(col).strip().startswith(str(p)):
                        indices_validos.append(col)
                        break
            
            if indices_validos:
                # Calcular média diretamente de todas as respostas para este fator
                valores = df_perguntas[indices_validos].values.flatten()
                valores = valores[~pd.isna(valores)]  # Remove NaN
                
                if len(valores) > 0:
                    media = valores.mean()
                    risco, cor = classificar_risco(media)
                    resultados.append({
                        "Fator Psicossocial": fator, 
                        "Média": round(media, 2), 
                        "Risco": risco,
                        "Número de Respostas": len(valores)
                    })
                    cores.append(cor)
        
        if not resultados:
            st.error("Não foi possível calcular resultados com os dados fornecidos.")
            st.stop()
        
        df_resultados = pd.DataFrame(resultados)
        
        # Mostrar estatísticas gerais
        st.write("### Resumo da Avaliação")
        media_geral = df_resultados["Média"].mean()
        st.metric(
            label="Média Geral", 
            value=f"{media_geral:.2f}",
            delta=f"{classificar_risco(media_geral)[0]}"
        )
        
        st.write("### Resultados Detalhados da Avaliação")
        st.dataframe(df_resultados)
        
        # Criar gráfico de barras para visualização dos riscos
        fig, ax = plt.subplots(figsize=(10, 6))
        bars = ax.barh(df_resultados["Fator Psicossocial"], df_resultados["Média"], color=cores)
        
        # Adicionar rótulos de dados nas barras
        for i, bar in enumerate(bars):
            ax.text(
                bar.get_width() + 0.1, 
                bar.get_y() + bar.get_height()/2, 
                f"{df_resultados['Média'].iloc[i]:.2f} ({df_resultados['Risco'].iloc[i].split()[0]} {df_resultados['Risco'].iloc[i].split()[1]})",
                va='center'
            )
        
        # Melhorar o layout e aparência do gráfico
        ax.set_xlabel("Média das Respostas")
        ax.set_title(f"Classificação de Riscos Psicossociais ({filtro_opcao}: {filtro_valor})")
        ax.grid(axis='x', linestyle='--', alpha=0.7)
        ax.set_xlim(0, 5)  # Define o limite dos eixos para escala de 1-5
        
        # Adicionar legenda de cores
        from matplotlib.lines import Line2D
        legend_elements = [
            Line2D([0], [0], color='red', lw=4, label='Risco Muito Alto (≤1)'),
            Line2D([0], [0], color='orange', lw=4, label='Risco Alto (1-2)'),
            Line2D([0], [0], color='yellow', lw=4, label='Risco Moderado (2-3)'),
            Line2D([0], [0], color='green', lw=4, label='Risco Baixo (3-4)'),
            Line2D([0], [0], color='purple', lw=4, label='Risco Muito Baixo (>4)'),
        ]
        ax.legend(handles=legend_elements, loc='lower right')
        
        plt.tight_layout()
        st.pyplot(fig)
        
        # Criar botões para download dos relatórios
        st.write("### Download de Relatórios")
        
        col1, col2 = st.columns(2)
        
        excel_data = gerar_excel(df_resultados)
        if excel_data:
            with col1:
                st.download_button(
                    label="Baixar Relatório Excel", 
                    data=excel_data, 
                    file_name=f"relatorio_hse_{filtro_opcao}_{filtro_valor}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        pdf_data = gerar_pdf(df_resultados)
        if pdf_data:
            with col2:
                st.download_button(
                    label="Baixar Relatório PDF", 
                    data=pdf_data, 
                    file_name=f"relatorio_hse_{filtro_opcao}_{filtro_valor}.pdf", 
                    mime="application/pdf"
                )
                
    except Exception as e:
        st.error(f"Ocorreu um erro inesperado: {str(e)}")
        st.write("Detalhes do erro para debug:")
        st.exception(e)
