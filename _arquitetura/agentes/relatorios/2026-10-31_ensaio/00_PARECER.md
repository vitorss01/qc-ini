# Parecer consolidado: ciclo 2026-10-31_ensaio (modo completo, Bioquímica)

> ENSAIO DE INTEGRAÇÃO. Nenhum Excel aberto, nenhuma cópia feita, nenhuma suíte rodada. Todos os testes estão
> `NAO_EXECUTADO` (motivo: ensaio). Nada aqui é validação.

No ciclo real eu faria: (1) leria os relatórios Bioquimica_F1..F4 desta pasta; (2) conferiria se cada teste tem estado e
evidência reproduzível (cópia + sha256 curto); (3) montaria a tabela B a partir da F3; (4) juntaria as ocorrências com IDs
provisórios P-xx e deduplicaria; (5) responderia o parecer D só com testes APROVADOS como "validado".

## Relatório A: experiência de uso
- F1, F2, F3, F4: ensaio, sem uso real do sistema. Sem impressões de uso.

## Relatório B: validação matemática
| Analito | Equipamento | Lote | Nível | Período | Dados de entrada | Fórmula encontrada | Resultado independente | Resultado QC_INI | Diferença | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| — | — | — | — | — | — | — | — | — | — | NAO_EXECUTADO (ensaio; F3-01 e F3-02 não rodaram) |

Calculadora independente citada pela F3: `_arquitetura/agentes/ferramentas/calc_independente.py` (não usada neste ciclo).

## Relatório C: ocorrências (IDs provisórios; o definitivo OC-xxxx é do Orquestrador)
| ID | Frente | Módulo | Tipo | Prioridade | Título | Recomendação |
|---|---|---|---|---|---|---|
| P-01 | F3 | Estatística/Sigma | calculo | Crítico | [ENSAIO] Ocorrência fictícia de cálculo | Descartar (fictícia) |
| P-02 | F1 | Entrada CIQ | funcional | Alto | [ENSAIO] Ocorrência fictícia de entrada de CIQ | Descartar (fictícia) |
| P-03 | F2 | CEQ | usabilidade | Médio | [ENSAIO] Ocorrência fictícia de usabilidade na janela do CEQ | Descartar (fictícia) |
| P-04 | F4 | Painel / filtro de período | duvida_metodo | Baixo | [ENSAIO] Trimestre deve ser calculado dos dados brutos? | Descartar (fictícia); no ciclo real iria ao Especialista em CQ |

Todas fictícias, sem evidência; não devem entrar em OCORRENCIAS.md como reais.

## Relatório D: parecer final
| Pergunta | Resposta |
|---|---|
| A rotina funciona? | Não executado (ensaio) |
| Bias validado? | Não executado (ensaio; F3-01) |
| Erro Total validado? | Não executado (ensaio; F3-01) |
| Sigma validado? | Não executado (ensaio; F3-01/F3-02) |
| Integração CIQ+CEQ validada? | Não executado (ensaio; F2-01/F2-02, F3) |
| Filtros de período e lote corretos? | Não executado (ensaio; F4-PER, F4-LOTE, F4-COMB, F4-LJ) |
| Entrada manual e automática validadas? | Não executado (ensaio; F1-MAN, F1-AUTO, F1-EQUIV) |
| Há divergência matemática? | Desconhecido: nenhum cálculo feito |
| Há problema operacional? | Desconhecido: só ocorrências fictícias P-01..P-04 |
| O que fica pendente? | Todo o ciclo real: 13 testes (F1 3, F2 3, F3 2, F4 5), tabela B com 5 analitos por risco, sha256 da produção |
