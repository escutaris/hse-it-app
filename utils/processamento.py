import pandas as pd
import streamlit as st
from utils.constantes import QUESTOES_INVERTIDAS, DIMENSOES_HSE, DESCRICOES_DIMENSOES

# Função para classificar os riscos com base na pontuação média
@st.cache_data
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

# Função para carregar e processar dados
@st.cache_data
def carregar_dados(uploaded_file):
    # Determinar o tipo de arquivo e carregá-lo adequadamente
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file, sep=';' if ';' in uploaded_file.getvalue().decode('utf-8', errors='replace')[:1000] else ',')
    else:
        df = pd.read_excel(uploaded_file)
    
    # Separar os dados de filtro e as perguntas
    colunas_filtro = list(df.columns[:7])
    if "Carimbo de data/hora" in colunas_filtro:
        colunas_filtro.remove("Carimbo de data/hora")
    
    # Garantir que as colunas das perguntas são corretamente identificadas
    colunas_perguntas = [col for col in df.columns if str(col).strip() and str(col).strip()[0].isdigit()]
    
    # Converter valores para numéricos, tratando erros
    df_perguntas = df[colunas_perguntas]
    for col in df_perguntas.columns:
        df_perguntas[col] = pd.to_numeric(df_perguntas[col], errors='coerce')
    
    return df, df_perguntas, colunas_filtro, colunas_perguntas

# Função para processar os dados (incluindo inversão de questões)
def processar_dados_hse(df_perguntas, colunas_perguntas):
    # Copiar o dataframe para não modificar o original
    df_processado = df_perguntas.copy()
    
    # Inverter a pontuação das questões invertidas
    for col in colunas_perguntas:
        # Extrair o número da questão (considerar que o formato é "1. Pergunta")
        try:
            numero_questao = int(str(col).strip().split('.')[0])
            if numero_questao in QUESTOES_INVERTIDAS:
                df_processado[col] = 6 - df_processado[col]  # Inverte a escala (1->5, 2->4, 3->3, 4->2, 5->1)
        except (ValueError, IndexError):
            continue
    
    return df_processado

# Função para calcular resultados por dimensão
@st.cache_data
def calcular_resultados_dimensoes(df, df_perguntas_filtradas, colunas_perguntas):
    # Processar dados (inverter questões necessárias)
    df_processado = processar_dados_hse(df_perguntas_filtradas, colunas_perguntas)
    
    resultados = []
    cores = []
    
    # Número total de respostas após filtragem
    num_total_respostas = len(df_processado)
    
    # Calcular resultados para cada dimensão
    for dimensao, numeros_questoes in DIMENSOES_HSE.items():
        # Converter números de pergunta para índices de coluna
        indices_validos = []
        for q in numeros_questoes:
            for col in colunas_perguntas:
                if str(col).strip().startswith(str(q) + ".") or str(col).strip().startswith(str(q) + " "):
                    indices_validos.append(col)
                    break
        
        if indices_validos:
            # Calcular média para esta dimensão
            valores = df_processado[indices_validos].values.flatten()
            valores = valores[~pd.isna(valores)]  # Remove NaN
            
            if len(valores) > 0:
                media = valores.mean()
                risco, cor = classificar_risco(media)
                
                resultados.append({
                    "Dimensão": dimensao,
                    "Descrição": DESCRICOES_DIMENSOES[dimensao],
                    "Média": round(media, 2),
                    "Risco": risco,
                    "Número de Respostas": num_total_respostas,
                    "Questões": numeros_questoes
                })
                cores.append(cor)
    
    return resultados
