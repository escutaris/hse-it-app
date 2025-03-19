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

# Função para processar os dados (incluindo inversão de questões) - ATUALIZADA
def processar_dados_hse(df_perguntas, colunas_perguntas):
    """
    Processa dados do HSE-IT, incluindo inversão das questões que precisam ser invertidas.
    Lida com diferentes formatos de questões (números, textos com números, etc.)
    """
    # Copiar o dataframe para não modificar o original
    df_processado = df_perguntas.copy()
    
    # Função auxiliar para extrair número da questão de forma robusta
    def extrair_numero_questao(coluna):
        try:
            coluna_str = str(coluna).strip()
            
            # Padrão 1: "Número. Texto" (ex: "1. Pergunta")
            if '.' in coluna_str and coluna_str[0].isdigit():
                return int(coluna_str.split('.')[0])
                
            # Padrão 2: "Número - Texto" (ex: "1 - Pergunta")
            elif '-' in coluna_str and coluna_str[0].isdigit():
                return int(coluna_str.split('-')[0].strip())
                
            # Padrão 3: "Número Texto" (ex: "1 Pergunta")
            elif ' ' in coluna_str and coluna_str[0].isdigit():
                return int(coluna_str.split(' ')[0])
                
            # Padrão 4: Apenas os dígitos iniciais (ex: "1Pergunta")
            elif coluna_str[0].isdigit():
                digits = ''
                for char in coluna_str:
                    if char.isdigit():
                        digits += char
                    else:
                        break
                if digits:
                    return int(digits)
                    
            # Padrão 5: "Questão Número" (ex: "Questão 1")
            elif "questão" in coluna_str.lower() or "questao" in coluna_str.lower():
                # Extrair números após "questão"
                import re
                match = re.search(r'quest[ãa]o\s*(\d+)', coluna_str.lower())
                if match:
                    return int(match.group(1))
                    
            # Se nenhum padrão for encontrado, retornar None
            return None
            
        except Exception as e:
            print(f"Erro ao extrair número da questão '{coluna}': {str(e)}")
            return None
    
    # Inverter a pontuação das questões invertidas
    for col in colunas_perguntas:
        numero_questao = extrair_numero_questao(col)
        
        if numero_questao in QUESTOES_INVERTIDAS:
            # Verificar se a coluna contém valores numéricos
            if pd.api.types.is_numeric_dtype(df_processado[col]):
                # Inverter a escala (1->5, 2->4, 3->3, 4->2, 5->1)
                df_processado[col] = 6 - df_processado[col]
            else:
                # Tentar converter para numérico primeiro
                try:
                    df_processado[col] = pd.to_numeric(df_processado[col], errors='coerce')
                    # Inverter apenas valores não-NA
                    mask = ~pd.isna(df_processado[col])
                    df_processado.loc[mask, col] = 6 - df_processado.loc[mask, col]
                except Exception as e:
                    print(f"Erro ao processar questão invertida {col}: {str(e)}")
    
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

# Adicionar função gerar_sugestoes_acoes
def gerar_sugestoes_acoes(df_resultados):
    # Dicionário com sugestões de ações para cada dimensão por nível de risco
    sugestoes_por_dimensao = {
        "Demanda": {
            "Risco Muito Alto": [
                "Realizar auditoria completa da distribuição de carga de trabalho",
                "Implementar sistema de gestão de tarefas e priorização",
                "Reavaliar prazos e expectativas de produtividade",
                "Contratar pessoal adicional para áreas sobrecarregadas",
                "Implementar pausas obrigatórias durante o dia de trabalho"
            ],
            "Risco Alto": [
                "Mapear atividades e identificar gargalos de processo",
                "Implementar ferramentas para melhor organização e planejamento do trabalho",
                "Revisar e ajustar prazos de entregas e metas",
                "Capacitar gestores em gerenciamento de carga de trabalho das equipes"
            ],
            "Risco Moderado": [
                "Promover treinamentos de gestão do tempo e priorização",
                "Revisar distribuição de tarefas entre membros da equipe",
                "Estabelecer momentos regulares para feedback sobre carga de trabalho"
            ],
            "Risco Baixo": [
                "Manter monitoramento regular das demandas de trabalho",
                "Realizar check-ins periódicos sobre volume de trabalho",
                "Oferecer recursos de apoio para períodos de pico de trabalho"
            ],
            "Risco Muito Baixo": [
                "Documentar boas práticas atuais de gestão de demandas",
                "Compartilhar casos de sucesso no gerenciamento de carga de trabalho",
                "Manter práticas de gestão de demandas e continuar monitorando"
            ]
        },
        # Adicionar os outros dicionários de sugestões das outras dimensões...
        # (copiar o conteúdo do arquivo sugestoes.py)
    }

    plano_acao = []

    # Para cada dimensão no resultado
    for _, row in df_resultados.iterrows():
        dimensao = row['Dimensão']
        risco = row['Risco']
        nivel_risco = risco.split()[0] + " " + risco.split()[1]  # Ex: "Risco Alto"

        # Obter sugestões para esta dimensão e nível de risco
        if dimensao in sugestoes_por_dimensao and nivel_risco in sugestoes_por_dimensao[dimensao]:
            sugestoes = sugestoes_por_dimensao[dimensao][nivel_risco]

            # Adicionar ao plano de ação
            for sugestao in sugestoes:
                plano_acao.append({
                    "Dimensão": dimensao,
                    "Nível de Risco": nivel_risco,
                    "Média": row['Média'],
                    "Sugestão de Ação": sugestao,
                    "Responsável": "",
                    "Prazo": "",
                    "Status": "Não iniciada"
                })

    # Criar DataFrame com o plano de ação
    df_plano_acao = pd.DataFrame(plano_acao)
    return df_plano_acao

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
