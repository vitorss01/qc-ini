# Registro de ocorrências da auditoria do QC_INI

Mantido só pelo Orquestrador (`/auditoria-qc`). Uma linha por ocorrência e IDs que nunca se repetem. A evidência fica
no relatório do ciclo indicado.

**Status possíveis:**

1. ABERTA
2. EM ANÁLISE
3. AGUARDA AUTORIZAÇÃO
4. EM CORREÇÃO
5. AGUARDA RETESTE
6. VALIDADA
7. ENCERRADA SEM CORREÇÃO

**Prioridade:** Crítico, Alto, Médio ou Baixo (REGRAS_COMUNS §4).

| ID | Ciclo | Produto | Módulo | Prior. | Descrição | Evidência | Analisa | Corrige | Status | Reteste | Validação final |
|---|---|---|---|---|---|---|---|---|---|---|---|
| OC-0001 | 2026-10-31_ciclo01 | Ambos | ETL manual / motor | Crítico | Equipamento/analito em minúsculas: o manual toma o lugar de um resultado real no Painel/LJ/Westgard e a Estatística conta os dois | relatorios/2026-10-31_ciclo01/parecer_cq G1; Bio F1, Hema F1 | CQ + Tester | Orquestrador | AGUARDA RETESTE (ADR-088, entrega 35, aec2c43) | — | — |
| OC-0002 | 2026-10-31_ciclo01 | Ambos | Painel H7:H9 | Crítico | ET = 1,65·CV quando o bias é SEM EP (regressão do ADR-030 pelo painel.py:705) | relatorios/2026-10-31_ciclo01/parecer_cq G3 | CQ + Tester | Orquestrador | AGUARDA RETESTE (ADR-088, entrega 35, aec2c43) | — | — |
| OC-0003 | 2026-10-31_ciclo01 | Hema | Estatística / troca de lote | Crítico | Trocar o lote deixa a Estatística com o B4 novo e n/média/DP/CV/ET do lote anterior | relatorios/2026-10-31_ciclo01/parecer_cq G13; Hema F2/F4 | CQ + Tester | Orquestrador | AGUARDA RETESTE (ADR-088, entrega 35, aec2c43) | — | — |
| OC-0004 | 2026-10-31_ciclo01 | Ambos | ETL / dados CIQ | Crítico | Resultados ≤ 0 e de nível trocado ficam ATIVO e derrubam Sigma/u(Rw) (56 na Bio). Usuário autorizou em 03/11: ≤ 0 sai da estatística automaticamente | relatorios/2026-10-31_ciclo01/parecer_cq G4 | CQ + Tester | Orquestrador (regra) + laboratório (nível trocado) | AGUARDA RETESTE (ADR-088, entrega 35, aec2c43) | — | — |
| OC-0005 | 2026-10-31_ciclo01 | Ambos | ETL manual | Alto | Data inválida colada (31/02/2026) vira 03/03/2026 e o manual entra ATIVO | relatorios/2026-10-31_ciclo01/parecer_cq G2 | CQ + Tester | Orquestrador | AGUARDA RETESTE (ADR-088, entrega 35, aec2c43) | — | — |
| OC-0006 | 2026-10-31_ciclo01 | Ambos | mCEQ.BiasEQ | Alto | Bias do ET/Sigma inclui amostras com resultado 0 e NAO AVALIADO que o ViesEQ exclui. Usuário autorizou em 03/11: mesmo filtro do ViesEQ | relatorios/2026-10-31_ciclo01/parecer_cq G7 | CQ + Tester | Orquestrador | AGUARDA RETESTE (ADR-088, entrega 35, aec2c43) | — | — |
| OC-0007 | 2026-10-31_ciclo01 | Bio | Lotes / dados | Alto | Lotes de urina/PCR (5659, 6353/6354, 8353, 8177, 8794) ativos em analitos séricos | relatorios/2026-10-31_ciclo01/parecer_cq G5 | CQ + Tester | Laboratório (inativar, SIPEC) + alerta | AGUARDA LABORATÓRIO | — | — |
| OC-0008 | 2026-10-31_ciclo01 | Hema | Analitos / alvo do lote | Alto | Alvo do lote 522611 incompatível com os dados em 12 pares (IG#, RDW-SD N2) | relatorios/2026-10-31_ciclo01/parecer_cq G6 | CQ + Tester | Laboratório (bula) | AGUARDA LABORATÓRIO | — | — |
| OC-0009 | 2026-10-31_ciclo01 | Bio | Estatística resumo | Médio | Resumo "Sem EP" conta 22 níveis, mas são 4 | relatorios/2026-10-31_ciclo01/parecer_cq G8 | CQ + Tester | Orquestrador | AGUARDA RETESTE (ADR-088, entrega 35, aec2c43) | — | — |
| OC-0010 | 2026-10-31_ciclo01 | Ambos | Sigma / Eng_Saida | Médio | Vários Sigmas para o mesmo analito (bloco do lote com CV do cadastro); Sigma legado no Eng_Saida | relatorios/2026-10-31_ciclo01/parecer_cq G9 | CQ + Tester | Orquestrador (rótulos/legado) | ABERTA | — | — |
| OC-0011 | 2026-10-31_ciclo01 | Ambos | Especificações | Médio | Biblioteca de especificações 2026–2029 vazia (Lipase VB, PCR CLIA) | relatorios/2026-10-31_ciclo01/parecer_cq G10 | CQ + Tester | Laboratório | AGUARDA LABORATÓRIO | — | — |
| OC-0012 | 2026-10-31_ciclo01 | Ambos | Painel janela 180 | Alto | "Tudo"/período longo: indicadores do Painel só das 180 corridas mais recentes; o aviso fala só do gráfico | relatorios/2026-10-31_ciclo01/parecer_cq G11 | CQ + Tester | Orquestrador (aviso) | ABERTA | — | — |
| OC-0013 | 2026-10-31_ciclo01 | Ambos | Painel H3 vazio | Médio | Bias e Sigma aparecem no modo TODOS, contra o texto do ADR-086. Usuário decidiu em 03/11: manter, rotulados "do método", e emendar o ADR | relatorios/2026-10-31_ciclo01/parecer_cq G12 | CQ + Tester | Orquestrador | AGUARDA RETESTE (ADR-088, entrega 35, aec2c43) | — | — |
| OC-0014 | 2026-10-31_ciclo01 | Hema | EQA_Base | Baixo | 90 linhas "Controllab" da simulação na tabela do CAP (ADR-034) | relatorios/2026-10-31_ciclo01/parecer_cq G14 | CQ + Tester | Orquestrador (aviso opcional) | ABERTA | — | — |
| OC-0015 | 2026-10-31_ciclo01 | Hema | Especificações RDW | Médio | CEQ no RDW-CV, VB só no RDW-SD; nenhum tem Sigma | relatorios/2026-10-31_ciclo01/parecer_cq G15 | CQ + Tester | Laboratório (fonte) | AGUARDA LABORATÓRIO | — | — |
| OC-0016 | 2026-10-31_ciclo01 | Hema | Especificações CLIA | Médio | ETp "CLIA" de WBC 7 % e PLT 15 % difere do 42 CFR 493.941 | relatorios/2026-10-31_ciclo01/parecer_cq G16 | CQ + Tester | Laboratório (fonte) | AGUARDA LABORATÓRIO | — | — |
| OC-0017 | 2026-10-31_ciclo01 | Hema | Especificações FAB | Baixo | "FAB CVTp" com valores que não são CV % | relatorios/2026-10-31_ciclo01/parecer_cq G17 | CQ + Tester | Laboratório | AGUARDA LABORATÓRIO | — | — |
| OC-0018 | 2026-10-31_ciclo01 | Bio | QA E01 | Baixo | E01 permanente por 4 inativações sem motivo (software correto) | relatorios/2026-10-31_ciclo01/parecer_cq G18 | CQ + Tester | Laboratório | ENCERRADA SEM CORREÇÃO | — | — |
| OC-0019 | 2026-10-31_ciclo01 | Bio | Digitar Resultados | Alto | EQUIPAMENTO vazio recebe o do Painel sem aviso; as instruções não pedem o equipamento | relatorios/2026-10-31_ciclo01/parecer_ux G01 | UX + Tester | Orquestrador (instrução + destaque) | ABERTA | — | — |
| OC-0020 | 2026-10-31_ciclo01 | Ambos | Digitar Resultados | Médio | Manual com analito/lote/equipamento fora do cadastro fica ATIVO e invisível (só aviso: INCOMPLETO não autorizado) | relatorios/2026-10-31_ciclo01/parecer_ux G02 | UX + Tester | Orquestrador (aviso A11) | ABERTA | — | — |
| OC-0021 | 2026-10-31_ciclo01 | Ambos | Digitar Resultados | Baixo | RUN/DETALHE mostram 0; DETALHE sem acento | relatorios/2026-10-31_ciclo01/parecer_ux G03 | UX + Tester | Orquestrador | ABERTA | — | — |
| OC-0022 | 2026-10-31_ciclo01 | Bio | QA I01 | Baixo | I01 lista analitos de URINA sem dizer a matriz | relatorios/2026-10-31_ciclo01/parecer_ux G04 | UX + Tester | Orquestrador | ABERTA | — | — |
| OC-0023 | 2026-10-31_ciclo01 | Ambos | Audit_Log | Baixo | RESULTADO_MANUAL_INCLUIDO gravado para manuais incompletos e em conflito | relatorios/2026-10-31_ciclo01/parecer_ux G05 | UX + Tester | Orquestrador | ABERTA | — | — |
| OC-0024 | 2026-10-31_ciclo01 | Ambos | ResumoQA | Médio | E04 de manual aparece como "ERRO DE GOVERNANÇA" | relatorios/2026-10-31_ciclo01/parecer_ux G06 | UX + Tester | Orquestrador | ABERTA | — | — |
| OC-0025 | 2026-10-31_ciclo01 | Ambos | Painel status | Médio | Status REJEITADO = pior do período, sem dizer isso | relatorios/2026-10-31_ciclo01/parecer_ux G07 | UX + Tester | Orquestrador (rótulo) | ABERTA | — | — |
| OC-0026 | 2026-10-31_ciclo01 | Ambos | Digitar Resultados | Baixo | MOTIVO "Outro" sem complemento | relatorios/2026-10-31_ciclo01/parecer_ux G08 | UX + Tester | Orquestrador | ABERTA | — | — |
| OC-0027 | 2026-10-31_ciclo01 | Ambos | frmCEQ / permissões | Alto | ANALISTA não corrige nem exclui CEQ. Usuário autorizou em 03/11: botões Corrigir/Excluir com trilha | relatorios/2026-10-31_ciclo01/parecer_ux G09 | UX + Tester | Orquestrador | EM CORREÇÃO | — | — |
| OC-0028 | 2026-10-31_ciclo01 | Ambos | CEQ consolidação | Alto | Correção na tabela só chega ao bias depois do Consolidar; Z3 diz "base em dia" | relatorios/2026-10-31_ciclo01/parecer_ux G10 | UX + Tester | Orquestrador | EM CORREÇÃO | — | — |
| OC-0029 | 2026-10-31_ciclo01 | Ambos | frmCEQ mapa | Alto | CEQ com nome fora do mapa/Controllab gravado OK e fora do bias, sem aviso | relatorios/2026-10-31_ciclo01/parecer_ux G11 | UX + Tester | Orquestrador (aviso) | EM CORREÇÃO | — | — |
| OC-0030 | 2026-10-31_ciclo01 | Hema | frmCEQ mapa | Alto | "Lymphs" (LYMPH#) entra no bias do LYMPH% e derruba o Sigma | relatorios/2026-10-31_ciclo01/parecer_ux G12 | UX + Tester | Orquestrador (aviso + confirmação) | EM CORREÇÃO | — | — |
| OC-0031 | 2026-10-31_ciclo01 | Ambos | frmCEQ alvo | Médio | CEQ sem valor-alvo aceito sem aviso (só aviso: obrigatório não autorizado) | relatorios/2026-10-31_ciclo01/parecer_ux G13 | UX + Tester | Orquestrador (aviso) | ABERTA | — | — |
| OC-0032 | 2026-10-31_ciclo01 | Ambos | Audit_Log CEQ | Médio | Edição, exclusão, consolidação e troca de provedor do CEQ sem trilha | relatorios/2026-10-31_ciclo01/parecer_ux G14 | UX + Tester | Orquestrador | EM CORREÇÃO | — | — |
| OC-0033 | 2026-10-31_ciclo01 | Ambos | frmCEQ sessão | Médio | SalvarCEQ grava sem sessão (papel mínimo não autorizado) | relatorios/2026-10-31_ciclo01/parecer_ux G15 | UX + Tester | Orquestrador (exigir sessão) | ABERTA | — | — |
| OC-0034 | 2026-10-31_ciclo01 | Ambos | frmCEQ listas | Baixo | Listas da janela misturam nomes da simulação | relatorios/2026-10-31_ciclo01/parecer_ux G16 | UX + Tester | Orquestrador | ABERTA | — | — |
| OC-0035 | 2026-10-31_ciclo01 | Ambos | GUIA §5–6 | Médio | GUIA manda digitar CEQ nas abas bloqueadas; não cita a janela nem o Consolidar | relatorios/2026-10-31_ciclo01/parecer_ux G17 | UX + Tester | Orquestrador | ABERTA | — | — |
| OC-0036 | 2026-10-31_ciclo01 | Ambos | Estatística AT | Médio | Coluna AT mostra "004 / 000" (TEXT gravado pelo COM em pt-BR) | relatorios/2026-10-31_ciclo01/parecer_ux G18 | UX + Tester | Orquestrador | ABERTA | — | — |
| OC-0037 | 2026-10-31_ciclo01 | Ambos | Estatística H3 vazio | Médio | Estatística no modo TODOS diz "refletido do Painel" e calcula o lote em uso | relatorios/2026-10-31_ciclo01/parecer_ux G19 | UX + Tester | Orquestrador (rótulo) | ABERTA | — | — |
| OC-0038 | 2026-10-31_ciclo01 | Ambos | Painel período | Médio | O período escolhido não volta ao retornar de outro lote | relatorios/2026-10-31_ciclo01/parecer_ux G20 | UX + Tester | Orquestrador | ABERTA | — | — |
| OC-0039 | 2026-10-31_ciclo01 | Ambos | Painel C4/A39 | Médio | C4/A39 mostram o lote em uso (vencido na Hema) em telas de outro lote | relatorios/2026-10-31_ciclo01/parecer_ux G21 | UX + Tester | Orquestrador + laboratório (lote em uso) | ABERTA | — | — |
| OC-0040 | 2026-10-31_ciclo01 | Bio | Estatística MENSAL/TRIM | Médio | Visões MENSAL/TRIMESTRAL inalcançáveis e "Ano 0" | relatorios/2026-10-31_ciclo01/parecer_ux G22 | UX + Tester | Orquestrador | ABERTA | — | — |
| OC-0041 | 2026-10-31_ciclo01 | Ambos | Painel K3/K4 | Médio | K3/K4 com rótulo antigo após De/Até | relatorios/2026-10-31_ciclo01/parecer_ux G23 | UX + Tester | Orquestrador | ABERTA | — | — |
| OC-0042 | 2026-10-31_ciclo01 | Ambos | Levey-Jennings | Médio | LJ suavizado, sem eixo Y e sem ±kDP (mudança de formato não autorizada em 03/11) | relatorios/2026-10-31_ciclo01/parecer_ux G24 | UX + Tester | — | ENCERRADA SEM CORREÇÃO | — | — |
| OC-0043 | 2026-10-31_ciclo01 | Bio | Painel J7:J8 | Baixo | "REJEITADO" cortado | relatorios/2026-10-31_ciclo01/parecer_ux G25 | UX + Tester | Orquestrador | ABERTA | — | — |
| OC-0044 | 2026-10-31_ciclo01 | Bio | Estatística | Baixo | Nome do analito desalinhado na vertical | relatorios/2026-10-31_ciclo01/parecer_ux G26 | UX + Tester | Orquestrador | ABERTA | — | — |
| OC-0045 | 2026-10-31_ciclo01 | Bio | Painel H3 | Médio | Lista do H3 com os lotes de todos os analitos; escolher um troca o analito | relatorios/2026-10-31_ciclo01/parecer_ux G27 | UX + Tester | Orquestrador | ABERTA | — | — |
| OC-0046 | 2026-10-31_ciclo01 | Ambos | Painel violações | Médio | "Últ. violação/Classificação" não é do período | relatorios/2026-10-31_ciclo01/parecer_ux G28 | UX + Tester | Orquestrador (rótulo) | ABERTA | — | — |
| OC-0047 | 2026-10-31_ciclo01 | Bio | Levey-Jennings | Baixo | Ponto de violação cortado no topo do gráfico | relatorios/2026-10-31_ciclo01/parecer_ux G29 | UX + Tester | Orquestrador | ABERTA | — | — |
| OC-0048 | 2026-10-31_ciclo01 | Hema | Estatística Ano De/Até | Médio | Ano De/Até corta lotes que atravessam o ano; Até < De aceito | relatorios/2026-10-31_ciclo01/parecer_ux G30 | UX + Tester | Orquestrador | ABERTA | — | — |
| OC-0049 | 2026-10-31_ciclo01 | Hema | Estatística | Baixo | Linhas vazias mostram 0 e SEM EP | relatorios/2026-10-31_ciclo01/parecer_ux G31 | UX + Tester | Orquestrador | ABERTA | — | — |
| OC-0050 | 2026-10-31_ciclo01 | Ambos | tblCQ_Final cabeçalho | Baixo | "Total: 44392,0" (TEXT pelo COM) | relatorios/2026-10-31_ciclo01/parecer_ux G32 | UX + Tester | Orquestrador | ABERTA | — | — |

## Reprodução independente do ciclo 2026-10-31_ciclo01 (Tester, 03/11/2026)

`relatorios/2026-10-31_ciclo01/tester_reproduzir.md`: 25 das 26 críticas/altas REPRODUZIDAS com números próprios,
1 PARCIAL (u(Rw) 81,2 % da Proteína total não reconferida), 0 NAO_REPRODUZIDA. Por OC: 0001, 0002, 0004 (Bio e Hema),
0003 (só Hema; a Bio não é afetada), 0005, 0006, 0007, 0008, 0012, 0019, 0027, 0028, 0029, 0030 — REPRODUZIDA.
Produção intacta (hashes iguais no fim). Achados fora da rota para o próximo ciclo: validação do EQUIPAMENTO da Bio
recusa "DIMENSION 1" lida por automação (conferir digitando); Amilase N1 do lote 8974 a 7 DP do alvo (n 373).

## Histórico dos ciclos

| Ciclo | Modo | Produtos | Frentes | Ocorrências novas | Resultado |
|---|---|---|---|---|---|
| 2026-10-31_ciclo01 | completo | Bio + Hema | F1–F4 | 80 constatações → 50 OC (deduplicadas por grupo de parecer) | 245 testes (172 A / 62 R / 5 parciais / 6 NE); 174 cálculos independentes, nenhuma divergência de fórmula; decisões do usuário em 03/11: ≤0 fora da estatística, BiasEQ com o filtro do ViesEQ, H3 vazio rotulado, ANALISTA corrige CEQ |
| 2026-10-31_ensaio | ensaio (sem Excel) | Bioquimica | F1–F4 | 0 (4 fictícias [ENSAIO], descartadas) | Ligação conferida: 4 frentes, parecer consolidado, roteamento CQ/UX/Tester correto; agentes fora do registro rodaram pela definição (8/8) |
