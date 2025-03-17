import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
import io

# Função para classificar os riscos com base na pontuação média
def classificar_risco(media):
    if media <= 1:
        return "Risco Muito Alto 🔴"
    elif media > 1 and media <= 2:
        return "Risco Alto 🟠"
    elif media > 2 and media <= 3:
        return "Risco Moderado 🟡"
    elif media > 3 and media <= 4:
        return "Risco Baixo 🟢"
    else:
        return "Risco Muito Baixo 🟣"

# Função para gerar o arquivo Excel com os resultados
def gerar_excel(df_resultados):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_resultados.to_excel(writer, sheet_name='Resultados', index=False)
    output.seek(0)
    return output

# Função para gerar o arquivo PDF com os resultados
def gerar_pdf(df_resultados):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", style='B', size=14)
    pdf.cell(200, 10, "Relatório de Fatores Psicossociais (HSE-IT)", ln=True, align='C')
    pdf.ln(10)
    pdf.set_font("Arial", size=12)
    for index, row in df_resultados.iterrows():
        pdf.cell(0, 10, f"{row['Fator Psicossocial']}: {row['Risco']}", ln=True)
    output = io.BytesIO()
    pdf.output(output)
    output.seek(0)
    return output

# Interface do aplicativo no Streamlit
st.title("Avaliação de Fatores Psicossociais - HSE-IT")
st.write("Faça upload do arquivo Excel contendo os resultados do questionário.")

uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=["xlsx", "csv"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    # Separar os dados de filtro e as perguntas
    df_filtros = df.iloc[:, :7]  # Colunas com Setor, Cargo, etc.
    df_perguntas = df.iloc[:, 7:]  # Colunas com as respostas das perguntas
    
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
    
    # Criar opções de filtro
    filtro_opcao = st.selectbox("Filtrar por", ["Empresa Toda", "Setor", "Cargo", "Escolaridade", "Tempo de Serviço"])
    
    if filtro_opcao != "Empresa Toda":
        filtro_valor = st.selectbox(f"Escolha um {filtro_opcao}", df_filtros[filtro_opcao].unique())
        df_perguntas = df_perguntas[df_filtros[filtro_opcao] == filtro_valor]
    
    resultados = []
    for fator, perguntas in fatores.items():
        media = df_perguntas.iloc[:, perguntas].mean().mean()
        risco = classificar_risco(media)
        resultados.append({"Fator Psicossocial": fator, "Média": round(media, 2), "Risco": risco})
    
    df_resultados = pd.DataFrame(resultados)
    st.write("### Resultados da Avaliação")
    st.dataframe(df_resultados)
    
    # Criar gráfico de barras para visualização dos riscos
    fig, ax = plt.subplots()
    ax.barh(df_resultados["Fator Psicossocial"], df_resultados["Média"], color=['red' if "Muito Alto" in r else 'orange' if "Alto" in r else 'yellow' if "Moderado" in r else 'green' if "Baixo" in r else 'purple' for r in df_resultados["Risco"]])
    ax.set_xlabel("Média das Respostas")
    ax.set_title(f"Classificação de Riscos Psicossociais ({filtro_opcao}: {filtro_valor if filtro_opcao != 'Empresa Toda' else 'Geral'})")
    st.pyplot(fig)
    
    # Criar botões para download dos relatórios
    excel_data = gerar_excel(df_resultados)
    pdf_data = gerar_pdf(df_resultados)
    
    st.download_button(label="Baixar Relatório Excel", data=excel_data, file_name="relatorio_hse.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button(label="Baixar Relatório PDF", data=pdf_data, file_name="relatorio_hse.pdf", mime="application/pdf")
