import pandas as pd
import streamlit as st
from datetime import datetime
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
    else:  # media > 4 (implícito no else)
        return "Risco Muito Baixo 🟣", "purple"

# Função para carregar e processar dados
@st.cache_data
def carregar_dados(uploaded_file):
    # Determinar o tipo de arquivo e carregá-lo adequadamente
    if uploaded_file.name.endswith('.csv'):
        # Tentar detectar o separador (vírgula ou ponto-e-vírgula)
        df = pd.read_csv(uploaded_file, sep=';' if ';' in uploaded_file.getvalue().decode('utf-8', errors='replace')[:1000] else ',')
    else:
        # Para arquivos Excel
        try:
            df = pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"Erro ao carregar arquivo Excel: {str(e)}")
            st.info("Certifique-se de que o arquivo está no formato .xlsx ou .xls correto.")
            return None, None, None, None
    
    # Verificar se o DataFrame foi carregado corretamente
    if df is None or len(df) == 0:
        st.error("Não foi possível carregar dados do arquivo ou o arquivo está vazio.")
        return None, None, None, None
    
    # Separar os dados de filtro e as perguntas
    colunas_filtro = list(df.columns[:7])
    if "Carimbo de data/hora" in colunas_filtro:
        colunas_filtro.remove("Carimbo de data/hora")
    
    # Garantir que as colunas das perguntas são corretamente identificadas
    colunas_perguntas = [col for col in df.columns if str(col).strip() and str(col).strip()[0].isdigit()]
    
    # Verificar se foram encontradas colunas de perguntas
    if len(colunas_perguntas) == 0:
        st.error("Não foram encontradas colunas de perguntas no formato correto (começando com números).")
        st.info("O arquivo deve conter perguntas do HSE-IT no formato '1. Pergunta...'")
        return None, None, None, None
    
    # Converter valores para numéricos, tratando erros
    df_perguntas = df[colunas_perguntas]
    for col in df_perguntas.columns:
        df_perguntas[col] = pd.to_numeric(df_perguntas[col], errors='coerce')
    
    # Verificar se há muitos valores ausentes
    valores_ausentes = df_perguntas.isna().sum().sum()
    total_valores = df_perguntas.size
    percentual_ausente = (valores_ausentes / total_valores) * 100
    
    if percentual_ausente > 30:  # Se mais de 30% dos dados estão ausentes
        st.warning(f"Atenção: {percentual_ausente:.1f}% dos valores das respostas estão ausentes. Isso pode afetar a qualidade dos resultados.")
    
    return df, df_perguntas, colunas_filtro, colunas_perguntas

# Função para processar os dados (incluindo inversão de questões)
def processar_dados_hse(df_perguntas, colunas_perguntas):
    # Copiar o dataframe para não modificar o original
    df_processado = df_perguntas.copy()
    
    # Inverter a pontuação das questões invertidas
    for col in colunas_perguntas:
        # Extrair o número da questão (considerar formatos como "1. Pergunta" ou "1 - Pergunta")
        try:
            # Tenta vários formatos possíveis de numeração
            if '.' in str(col):
                numero_questao = int(str(col).strip().split('.')[0])
            elif '-' in str(col):
                numero_questao = int(str(col).strip().split('-')[0])
            elif ' ' in str(col):
                numero_questao = int(str(col).strip().split(' ')[0])
            else:
                numero_questao = int(str(col).strip()[0])
                
            if numero_questao in QUESTOES_INVERTIDAS:
                df_processado[col] = 6 - df_processado[col]  # Inverte a escala (1->5, 2->4, 3->3, 4->2, 5->1)
        except (ValueError, IndexError) as e:
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
                # Melhorar a detecção de questões - considerar vários formatos
                col_texto = str(col).strip()
                if (col_texto.startswith(str(q) + ".") or 
                    col_texto.startswith(str(q) + " ") or 
                    col_texto.startswith(str(q) + "-")):
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
            else:
                # Adicionar um registro mesmo se não houver dados válidos
                resultados.append({
                    "Dimensão": dimensao,
                    "Descrição": DESCRICOES_DIMENSOES[dimensao],
                    "Média": None,
                    "Risco": "Sem dados suficientes",
                    "Número de Respostas": 0,
                    "Questões": numeros_questoes
                })
                cores.append("gray")
    
    return resultados

# Função aprimorada para padronizar o formato de data
def padronizar_formato_data(data_input):
    """
    Converte uma entrada de data para um formato padrão DD/MM/YYYY.
    Aceita strings em vários formatos ou objetos datetime.
    
    Args:
        data_input: String de data ou objeto datetime
        
    Returns:
        String formatada como DD/MM/YYYY ou None se falhar
    """
    # Se for None ou vazio, retorna None
    if data_input is None or (isinstance(data_input, str) and data_input.strip() == ""):
        return None
    
    # Se já for um objeto datetime, formata diretamente
    if isinstance(data_input, datetime):
        return data_input.strftime("%d/%m/%Y")
    
    # Converter para string se não for
    data_str = str(data_input).strip()
    
    # Lista de formatos de data a serem tentados (do mais comum ao menos comum)
    formatos_data = [
        "%d/%m/%Y",    # 31/12/2023
        "%Y-%m-%d",    # 2023-12-31
        "%d-%m-%Y",    # 31-12-2023
        "%m/%d/%Y",    # 12/31/2023
        "%d.%m.%Y",    # 31.12.2023
        "%Y/%m/%d",    # 2023/12/31
        "%d/%m/%y",    # 31/12/23
        "%y-%m-%d",    # 23-12-31
        "%m-%d-%Y",    # 12-31-2023
        "%B %d, %Y",   # December 31, 2023
        "%d %B %Y",    # 31 December 2023
        "%Y%m%d",      # 20231231
    ]
    
    # Tentar cada formato
    for formato in formatos_data:
        try:
            data_obj = datetime.strptime(data_str, formato)
            # Validar se é uma data plausível (não muito no passado ou futuro)
            ano_atual = datetime.now().year
            if data_obj.year < 2000 or data_obj.year > ano_atual + 10:
                continue  # Provavelmente uma interpretação incorreta
            return data_obj.strftime("%d/%m/%Y")
        except ValueError:
            continue
    
    # Se chegou aqui, não conseguiu converter
    return None
