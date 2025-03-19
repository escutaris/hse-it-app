import streamlit as st
import pandas as pd  # Adicionado import do pandas que estava faltando
from utils.constantes import DIMENSOES_HSE, DESCRICOES_DIMENSOES, QUESTOES_INVERTIDAS

# Título da página
st.title("Informações HSE-IT")

st.write("""
## Sobre o Questionário HSE-IT

O HSE-IT (Health and Safety Executive - Indicator Tool) é um questionário validado para avaliação de 
fatores psicossociais no ambiente de trabalho, desenvolvido pela instituição britânica de saúde e segurança ocupacional.

O questionário consiste em 35 perguntas que avaliam 7 dimensões de fatores psicossociais, permitindo identificar
áreas de risco que precisam de intervenção.
""")

# Mostrar dimensões do HSE-IT
st.subheader("Dimensões Avaliadas")

for dimensao, questoes in DIMENSOES_HSE.items():
    with st.expander(f"{dimensao} - Questões {questoes}"):
        st.write(f"**Descrição**: {DESCRICOES_DIMENSOES[dimensao]}")
        st.write("**Questões relacionadas**:")
        for q in questoes:
            # Buscar o texto das questões
            texto_questao = f"Questão {q}"
            
            # Alguns exemplos de questões (você pode adicionar mais ou todas as 35)
            questoes_texto = {
                1: "Sei claramente o que é esperado de mim no trabalho",
                2: "Posso decidir quando fazer uma pausa",
                3: "Grupos de trabalho diferentes pedem-me coisas difíceis de conjugar",
                4: "Sei do que necessito para fazer o meu trabalho",
                5: "Sou sujeito a assédio pessoal sob a forma de palavras ou comportamentos incorretos",
                6: "Tenho prazos impossíveis de cumprir",
                7: "Se o trabalho se torna difícil, os colegas ajudam-me",
                8: "Recebo feedback de apoio sobre o trabalho que faço",
                9: "Tenho que trabalhar muito intensivamente",
                10: "Tenho capacidade de decisão sobre a minha rapidez de trabalho",
                11: "Sei claramente os meus deveres e responsabilidades",
                12: "Tenho que negligenciar tarefas porque tenho uma carga elevada para cumprir",
                13: "Sei claramente as metas e objetivos do meu departamento",
                14: "Há fricção ou animosidade entre os colegas",
                15: "Posso decidir como fazer o meu trabalho",
                16: "Não consigo fazer pausas suficientes",
                17: "Compreendo como o meu trabalho se integra no objetivo geral da organização",
                18: "Sou pressionado a trabalhar durante horários longos",
                19: "Tenho poder de escolha para decidir o que faço no trabalho",
                20: "Tenho que trabalhar muito depressa",
                21: "Sou sujeito a intimidação/perseguição no trabalho",
                22: "Tenho pressões de tempo irrealistas",
                23: "Posso estar seguro de que o meu chefe imediato me ajuda num problema de trabalho",
                24: "Tenho ajuda e apoio necessários dos colegas",
                25: "Tenho algum poder de decisão sobre a minha forma de trabalho",
                26: "Tenho oportunidades suficientes para questionar os chefes sobre mudanças no trabalho",
                27: "Sou respeitado como mereço pelos colegas de trabalho",
                28: "O pessoal é sempre consultado sobre mudança no trabalho",
                29: "Posso falar com o meu chefe imediato sobre algo no trabalho que me transtornou ou irritou",
                30: "O meu horário pode ser flexível",
                31: "Os meus colegas estão dispostos a ouvir os meus problemas relacionados com o trabalho",
                32: "Quando são efetuadas mudanças no trabalho, sei claramente como resultarão na prática",
                33: "Recebo apoio durante trabalho que pode ser emocionalmente exigente",
                34: "Os relacionamentos no trabalho estão sob pressão",
                35: "O meu chefe imediato encoraja-me no trabalho"
            }
            
            # Corrigido o problema de indentação aqui
            if q in questoes_texto:
                texto_questao = questoes_texto[q]
                
            if q in QUESTOES_INVERTIDAS:
                st.write(f"- {q}: {texto_questao} *(escala invertida)*")
            else:
                st.write(f"- {q}: {texto_questao}")

# Explicação sobre a interpretação dos resultados
st.subheader("Interpretação dos Resultados")

st.write("""
Os resultados são apresentados em uma escala de 1 a 5, onde valores mais altos geralmente indicam
melhores condições psicossociais (exceto em questões com escala invertida, onde a lógica é oposta).

A classificação de risco é feita com base na média de cada dimensão:
""")

# Tabela de classificação de risco
risco_data = {
   "Classificação": ["Risco Muito Alto 🔴", "Risco Alto 🟠", "Risco Moderado 🟡", "Risco Baixo 🟢", "Risco Muito Baixo 🟣"],
   "Pontuação Média": ["≤ 1", "> 1 e ≤ 2", "> 2 e ≤ 3", "> 3 e ≤ 4", "> 4"],
   "Interpretação": [
       "Situação crítica, requer intervenção imediata", 
       "Condição preocupante, intervenção necessária em curto prazo", 
       "Condição de alerta, melhorias necessárias", 
       "Condição favorável, com oportunidades de melhoria", 
       "Condição excelente, manter as boas práticas"
   ]
}
df_risco = pd.DataFrame(risco_data)
st.table(df_risco)

# Notas sobre questões invertidas
st.subheader("Notas sobre Questões Invertidas")

st.write(f"""
Algumas questões do HSE-IT possuem escala invertida devido à sua formulação negativa. São elas: {QUESTOES_INVERTIDAS}

Nestas questões, uma resposta de valor mais baixo é considerada positiva. Por exemplo, para a questão 
"Tenho prazos impossíveis de cumprir", responder "1 - Nunca" representa uma boa condição de trabalho.

Este aplicativo já realiza automaticamente a inversão destas questões durante o processamento dos dados.
""")

# Template para coleta de dados via Google Forms
st.subheader("Coleta de Dados com Google Forms")

st.write("""
### Opção 1: Template para Google Forms

Você pode utilizar nosso modelo de Google Forms para coletar as respostas do questionário HSE-IT de forma padronizada:
""")

st.markdown("[Acessar Template HSE-IT no Google Forms](https://docs.google.com/forms/d/1BG7LWuuVUXs1CxsUWlELJ33q5QvI4JR8Y9PxJX2xQ-4/copy)", unsafe_allow_html=True)

st.write("""
### Como usar o Google Forms:

1. Clique no link acima e selecione "Fazer uma cópia" para criar sua própria versão
2. Personalize o formulário com o nome da sua empresa
3. Compartilhe o link do formulário com seus colaboradores
4. Após a coleta, vá em "Respostas" > "Ver no Planilhas Google"
5. Baixe a planilha em formato Excel (.xlsx)
6. Faça upload do arquivo na página "Upload de Dados" desta plataforma
""")

# Template para coleta de dados via Excel
st.subheader("Template para Coleta de Dados via Excel")

st.write("""
Alternativamente, você pode baixar um template Excel para aplicação do questionário HSE-IT. Este template contém:
- As 35 questões originais do HSE-IT
- Campos para informações demográficas (Setor, Cargo, etc.)
- Instruções detalhadas para aplicação
- Informações sobre as dimensões e interpretação
""")

def gerar_template_excel():
   import io
   import pandas as pd
   
   output = io.BytesIO()
   
   try:
       with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
           workbook = writer.book
           
           # Criar DataFrame com a estrutura esperada
           colunas = [
               "Setor", "Cargo", "Tempo_Empresa", "Genero", "Faixa_Etaria", "Escolaridade", "Regime_Trabalho",
           ]
           
           # As questões HSE-IT
           questoes_hse = [
               "1. Sei claramente o que é esperado de mim no trabalho",
               "2. Posso decidir quando fazer uma pausa",
               "3. Grupos de trabalho diferentes pedem-me coisas difíceis de conjugar",
               "4. Sei do que necessito para fazer o meu trabalho",
               "5. Sou sujeito a assédio pessoal sob a forma de palavras ou comportamentos incorretos",
               "6. Tenho prazos impossíveis de cumprir",
               "7. Se o trabalho se torna difícil, os colegas ajudam-me",
               "8. Recebo feedback de apoio sobre o trabalho que faço",
               "9. Tenho que trabalhar muito intensivamente",
               "10. Tenho capacidade de decisão sobre a minha rapidez de trabalho",
               "11. Sei claramente os meus deveres e responsabilidades",
               "12. Tenho que negligenciar tarefas porque tenho uma carga elevada para cumprir",
               "13. Sei claramente as metas e objetivos do meu departamento",
               "14. Há fricção ou animosidade entre os colegas",
               "15. Posso decidir como fazer o meu trabalho",
               "16. Não consigo fazer pausas suficientes",
               "17. Compreendo como o meu trabalho se integra no objetivo geral da organização",
               "18. Sou pressionado a trabalhar durante horários longos",
               "19. Tenho poder de escolha para decidir o que faço no trabalho",
               "20. Tenho que trabalhar muito depressa",
               "21. Sou sujeito a intimidação/perseguição no trabalho",
               "22. Tenho pressões de tempo irrealistas",
               "23. Posso estar seguro de que o meu chefe imediato me ajuda num problema de trabalho",
               "24. Tenho ajuda e apoio necessários dos colegas",
               "25. Tenho algum poder de decisão sobre a minha forma de trabalho",
               "26. Tenho oportunidades suficientes para questionar os chefes sobre mudanças no trabalho",
               "27. Sou respeitado como mereço pelos colegas de trabalho",
               "28. O pessoal é sempre consultado sobre mudança no trabalho",
               "29. Posso falar com o meu chefe imediato sobre algo no trabalho que me transtornou ou irritou",
               "30. O meu horário pode ser flexível",
               "31. Os meus colegas estão dispostos a ouvir os meus problemas relacionados com o trabalho",
               "32. Quando são efetuadas mudanças no trabalho, sei claramente como resultarão na prática",
               "33. Recebo apoio durante trabalho que pode ser emocionalmente exigente",
               "34. Os relacionamentos no trabalho estão sob pressão",
               "35. O meu chefe imediato encoraja-me no trabalho"
           ]
           
           # Adicionar questões ao template
           for q in questoes_hse:
               colunas.append(q)
           
           df_template = pd.DataFrame(columns=colunas)
           
           # Adicionar linha de exemplo
           import random
           random.seed(42)  # Para consistência
           valores_exemplo = ["TI", "Analista", "1-3 anos", "Masculino", "25-35", "Superior", "CLT"]
           valores_exemplo.extend([random.randint(1, 5) for _ in range(len(questoes_hse))])
           
           df_template.loc[0] = valores_exemplo
           
           # Exportar para Excel
           df_template.to_excel(writer, sheet_name="Questionário HSE-IT", index=False)
           
           # Formatar a planilha
           worksheet = writer.sheets["Questionário HSE-IT"]
           
           # Configurar largura das colunas
           for i, col in enumerate(df_template.columns):
               # Calcular largura baseada no comprimento da coluna
               col_width = max(len(str(col)), 15)
               worksheet.set_column(i, i, col_width)
           
           # Criar folha de instruções
           worksheet_instrucoes = workbook.add_worksheet("Instruções")
           
           # Formatar cabeçalho das instruções
           header_format = workbook.add_format({'bold': True, 'size': 14})
           worksheet_instrucoes.write('A1', 'Instruções para Aplicação do Questionário HSE-IT', header_format)
           worksheet_instrucoes.set_column('A:A', 100)
           
           # Adicionar instruções
           instrucoes = [
               "",
               "Sobre o HSE-IT:",
               "O HSE-IT (Health and Safety Executive - Indicator Tool) é um questionário validado para avaliação de fatores psicossociais no trabalho, desenvolvido pela instituição britânica de saúde e segurança.",
               "",
               "Instruções de uso:",
               "",
               "1. As primeiras 7 colunas contêm informações demográficas que podem ser adaptadas conforme necessário:",
               "   - Setor: Área/departamento do colaborador",
               "   - Cargo: Função desempenhada",
               "   - Tempo_Empresa: Tempo de permanência na organização",
               "   - Gênero: Identificação de gênero",
               "   - Faixa_Etária: Grupo etário",
               "   - Escolaridade: Nível de formação",
               "   - Regime_Trabalho: Tipo de contrato/regime",
               "",
               "2. As demais colunas (35) contêm as perguntas do questionário HSE-IT. Não altere o conteúdo destas perguntas para preservar a validade do instrumento.",
               "",
               "3. As respostas devem ser preenchidas com valores de 1 a 5, seguindo esta escala:",
               "   - 1: Nunca / Discordo totalmente",
               "   - 2: Raramente / Discordo parcialmente",
               "   - 3: Às vezes / Nem concordo nem discordo",
               "   - 4: Frequentemente / Concordo parcialmente",
               "   - 5: Sempre / Concordo totalmente",
               "",
               f"4. IMPORTANTE: Para algumas questões específicas, a escala é invertida devido à formulação negativa. Estas questões são: {QUESTOES_INVERTIDAS}.",
               "   Exemplo: Para a questão 'Tenho prazos impossíveis de cumprir', responder '1' (Nunca) é positivo.",
               "",
               "5. Mantenha a estrutura exata deste template, incluindo numeração e texto completo das questões.",
               "",
               "6. Aplique o questionário garantindo o anonimato das respostas para obter dados mais honestos e confiáveis.",
               "",
               "7. Recomenda-se aplicar o questionário a pelo menos 60% dos colaboradores para obter um panorama representativo da organização.",
               "",
               "8. Os resultados são agrupados em 7 dimensões: Demanda, Controle, Apoio da Chefia, Apoio dos Colegas, Relacionamentos, Função e Mudança.",
               "",
               "Após a coleta, utilize a plataforma HSE-IT Analytics para processar os dados e gerar relatórios de riscos psicossociais."
           ]
           
           # Escrever instruções
           for i, texto in enumerate(instrucoes, 2):
               worksheet_instrucoes.write(i, 0, texto)
               
           # Adicionar aba com detalhes sobre as dimensões
           worksheet_dimensoes = workbook.add_worksheet("Dimensões HSE-IT")
           
           # Título
           worksheet_dimensoes.write('A1', 'Dimensões do Questionário HSE-IT', header_format)
           worksheet_dimensoes.set_column('A:A', 15)
           worksheet_dimensoes.set_column('B:B', 40)
           worksheet_dimensoes.set_column('C:C', 30)
           
           # Cabeçalhos
           dim_header = workbook.add_format({'bold': True, 'fg_color': '#D7E4BC', 'border': 1})
           worksheet_dimensoes.write('A3', 'Dimensão', dim_header)
           worksheet_dimensoes.write('B3', 'Descrição', dim_header)
           worksheet_dimensoes.write('C3', 'Questões', dim_header)
           
           # Conteúdo
           row = 4
           for dimensao, questoes in DIMENSOES_HSE.items():
               worksheet_dimensoes.write(row, 0, dimensao)
               worksheet_dimensoes.write(row, 1, DESCRICOES_DIMENSOES[dimensao])
               worksheet_dimensoes.write(row, 2, str(questoes))
               row += 1
           
           # Adicionar informações sobre questões invertidas
           row += 2
           worksheet_dimensoes.write(row, 0, 'IMPORTANTE:', workbook.add_format({'bold': True}))
           row += 1
           worksheet_dimensoes.write(row, 0, 'Questões com escala invertida:')
           worksheet_dimensoes.write(row, 1, str(QUESTOES_INVERTIDAS))
           row += 1
           worksheet_dimensoes.write(row, 0, 'Nestas questões, uma pontuação mais baixa é positiva devido à formulação negativa da pergunta.')
       
       output.seek(0)
       return output
   
   except Exception as e:
       st.error(f"Erro ao gerar o template Excel: {str(e)}")
       return None

# Gerar e oferecer o template para download
template_excel = gerar_template_excel()
if template_excel:
   st.download_button(
       label="Baixar Template HSE-IT",
       data=template_excel,
       file_name="template_questionario_hse_it.xlsx",
       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
       help="Template Excel com as 35 questões do HSE-IT e instruções para aplicação."
   )

# Referências
st.subheader("Referências")

st.write("""
1. Health and Safety Executive (HSE). "Managing the causes of work-related stress: A step-by-step approach using the Management Standards." HSE Books, 2007.

2. Cousins, R., et al. "Management Standards and work-related stress in the UK: Practical development." Work & Stress, 18(2), 113-136, 2004.

3. Mackay, C. J., et al. "Management Standards and work-related stress in the UK: Policy background and science." Work & Stress, 18(2), 91-112, 2004.

Para mais informações, visite o site do HSE: [www.hse.gov.uk/stress](https://www.hse.gov.uk/stress/)
""")

# Sobre a Escutaris
st.subheader("Sobre a Escutaris")

# Logo e informações da empresa
col1, col2 = st.columns([1, 3])

with col1:
   # Usar um espaço reservado para o logo (você pode substituir pela URL real da imagem)
   st.image("/api/placeholder/200/200", caption="Escutaris")

with col2:
   st.markdown("""
   ### Transformando ambientes de trabalho
   
   A Escutaris é especializada em diagnósticos psicossociais e saúde mental corporativa. Oferecemos soluções personalizadas para identificar, gerenciar e minimizar riscos psicossociais no ambiente de trabalho.
   
   [Visite nosso site](https://escutaris.com.br) para conhecer mais sobre nossos serviços e soluções.
   """)
