import pandas as pd

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
        "Controle": {
            "Risco Muito Alto": [
                "Redesenhar processos para aumentar a autonomia dos trabalhadores",
                "Implementar esquemas de trabalho flexível",
                "Revisar políticas de microgerenciamento",
                "Treinar gestores em delegação efetiva",
                "Criar espaços para participação em decisões estratégicas"
            ],
            "Risco Alto": [
                "Identificar áreas específicas onde os trabalhadores podem ter mais controle",
                "Envolver colaboradores no planejamento de metas e métodos de trabalho",
                "Implementar sistema de sugestões para melhorias nos processos",
                "Oferecer opções de horários flexíveis"
            ],
            "Risco Moderado": [
                "Aumentar gradualmente a autonomia nas decisões rotineiras",
                "Solicitar feedback regular sobre nível de controle no trabalho",
                "Implementar projetos-piloto para testar maior autonomia"
            ],
            "Risco Baixo": [
                "Manter boas práticas de autonomia",
                "Revisar periodicamente áreas onde o controle pode ser ampliado",
                "Reconhecer e celebrar iniciativas independentes"
            ],
            "Risco Muito Baixo": [
                "Documentar práticas bem-sucedidas de autonomia",
                "Compartilhar casos de sucesso com outras áreas da organização",
                "Manter cultura de confiança e delegação"
            ]
        },
        "Apoio da Chefia": {
            "Risco Muito Alto": [
                "Implementar programa estruturado de desenvolvimento de lideranças",
                "Realizar avaliação 360° para gestores",
                "Estabelecer canais de comunicação direta entre colaboradores e alta direção",
                "Revisar políticas de promoção para garantir que bons líderes sejam reconhecidos",
                "Oferecer coaching individual para gestores com desafios específicos"
            ],
            "Risco Alto": [
                "Treinar gestores em habilidades de suporte e feedback",
                "Implementar reuniões regulares one-on-one entre líderes e liderados",
                "Estabelecer expectativas claras para comportamentos de liderança",
                "Criar fóruns para líderes compartilharem desafios e soluções"
            ],
            "Risco Moderado": [
                "Revisar e melhorar as práticas de feedback das lideranças",
                "Oferecer recursos e ferramentas para líderes apoiarem suas equipes",
                "Promover workshops sobre comunicação efetiva"
            ],
            "Risco Baixo": [
                "Manter programas de desenvolvimento de lideranças",
                "Reconhecer e celebrar boas práticas de liderança",
                "Implementar sistema de mentoria entre líderes"
            ],
            "Risco Muito Baixo": [
                "Documentar e compartilhar práticas exemplares de liderança",
                "Utilizar líderes como mentores para novos gestores",
                "Manter cultura de apoio e desenvolvimento contínuo"
            ]
        },
        "Apoio dos Colegas": {
            "Risco Muito Alto": [
                "Implementar programas estruturados de team building",
                "Revisar a composição e dinâmica das equipes",
                "Estabelecer facilitadores de equipe para melhorar integração",
                "Criar espaços físicos e virtuais para colaboração",
                "Implementar sistema de reconhecimento por comportamentos colaborativos"
            ],
            "Risco Alto": [
                "Promover atividades regulares de integração de equipes",
                "Treinar em habilidades de trabalho em equipe e comunicação",
                "Estabelecer objetivos compartilhados que incentivem a colaboração",
                "Revisar processos que possam estar criando competição indesejada"
            ],
            "Risco Moderado": [
                "Implementar reuniões regulares de equipe para compartilhamento",
                "Criar projetos colaborativos entre diferentes membros",
                "Oferecer oportunidades para pessoas se conhecerem melhor"
            ],
            "Risco Baixo": [
                "Manter momentos regulares de integração",
                "Monitorar dinâmicas de equipe, especialmente com novos membros",
                "Reconhecer comportamentos de apoio entre colegas"
            ],
            "Risco Muito Baixo": [
                "Documentar boas práticas de colaboração",
                "Utilizar a cultura de apoio como exemplo para novos funcionários",
                "Manter ambiente de confiança e colaboração"
            ]
        },
        "Relacionamentos": {
            "Risco Muito Alto": [
                "Implementar política de tolerância zero para assédio e comportamentos inadequados",
                "Criar canais confidenciais para denúncias",
                "Treinar todos os colaboradores em respeito e diversidade",
                "Estabelecer mediação de conflitos com profissionais externos",
                "Auditar clima organizacional e relacionamentos interpessoais"
            ],
            "Risco Alto": [
                "Desenvolver e comunicar política clara sobre comportamentos aceitáveis",
                "Treinar gestores na identificação e gestão de conflitos",
                "Implementar processos estruturados para resolução de conflitos",
                "Promover diálogo aberto sobre relacionamentos saudáveis no trabalho"
            ],
            "Risco Moderado": [
                "Realizar workshops sobre comunicação não-violenta",
                "Estabelecer acordos de equipe sobre comportamentos esperados",
                "Promover atividades que construam confiança entre colegas"
            ],
            "Risco Baixo": [
                "Manter comunicação regular sobre respeito no ambiente de trabalho",
                "Incorporar avaliação de relacionamentos nas pesquisas de clima",
                "Reconhecer exemplos positivos de resolução de conflitos"
            ],
            "Risco Muito Baixo": [
                "Documentar boas práticas de relacionamentos saudáveis",
                "Utilizar a cultura positiva como diferencial da organização",
                "Manter monitoramento contínuo do clima relacional"
            ]
        },
        "Função": {
            "Risco Muito Alto": [
                "Realizar revisão completa de descrições de cargos e responsabilidades",
                "Implementar processo de clarificação de funções para toda a organização",
                "Treinar gestores em delegação clara e definição de expectativas",
                "Estabelecer processos para resolver conflitos de papéis e responsabilidades",
                "Revisar estrutura organizacional para eliminar ambiguidades"
            ],
            "Risco Alto": [
                "Revisar e atualizar descrições de cargo",
                "Implementar reuniões regulares para esclarecer expectativas",
                "Criar matriz RACI (Responsável, Aprovador, Consultado, Informado) para projetos",
                "Treinar equipes em comunicação sobre papéis e responsabilidades"
            ],
            "Risco Moderado": [
                "Revisar processos onde ocorrem conflitos de funções",
                "Promover workshops para clarificar interfaces entre áreas",
                "Estabelecer fóruns para discutir e esclarecer papéis em projetos"
            ],
            "Risco Baixo": [
                "Manter atualizações periódicas de responsabilidades",
                "Incluir clareza de papéis nas avaliações de desempenho",
                "Promover comunicação contínua sobre expectativas"
            ],
            "Risco Muito Baixo": [
                "Documentar boas práticas de clareza de funções",
                "Utilizar como modelo para novos departamentos ou projetos",
                "Manter cultura de transparência sobre papéis e responsabilidades"
            ]
        },
        "Mudança": {
            "Risco Muito Alto": [
                "Implementar metodologia estruturada de gestão de mudanças",
                "Criar comitê representativo para planejamento de mudanças",
                "Estabelecer múltiplos canais de comunicação sobre processos de mudança",
                "Treinar gestores em liderança durante transformações",
                "Avaliar impacto de mudanças anteriores e implementar lições aprendidas"
            ],
            "Risco Alto": [
                "Desenvolver plano de comunicação para mudanças organizacionais",
                "Envolver representantes de diferentes níveis no planejamento",
                "Implementar feedbacks regulares durante processos de mudança",
                "Oferecer suporte adicional para equipes mais afetadas"
            ],
            "Risco Moderado": [
                "Melhorar a transparência sobre razões das mudanças",
                "Criar fóruns para esclarecer dúvidas sobre transformações",
                "Celebrar pequenas vitórias durante processos de mudança"
            ],
            "Risco Baixo": [
                "Manter comunicação proativa sobre possíveis mudanças",
                "Oferecer oportunidades regulares para feedback durante transformações",
                "Reconhecer contribuições positivas em processos de mudança"
            ],
            "Risco Muito Baixo": [
                "Documentar práticas bem-sucedidas de gestão de mudanças",
                "Utilizar abordagem participativa como padrão",
                "Manter cultura de adaptabilidade e melhoria contínua"
            ]
        }
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
