# BACKLOG PÓS-RELEASE — QC_INI

Melhorias que **não impedem a entrega**. Não são tarefas da release atual (feature freeze desde 23/11/2026). Quem
classifica e acrescenta: `qc-release-manager`. Prioridade futura: Alta / Média / Baixa. Risco: risco de implementar.

| ID | Melhoria | Componente | Justificativa | Benefício esperado | Prior. futura | Risco |
|---|---|---|---|---|---|---|
| BK-01 | Destacar o manual colado com analito/lote/equipamento fora do cadastro (A11 + realce) | ETL manual / QA (OC-0020) | Hoje fica ATIVO e não aparece no Painel; mitigado por orientação | evita omissão silenciosa de uma corrida | Alta | Baixo |
| BK-02 | Lote FORA_DO_CQ/SO_ANALITOS não pedir Média/DP ("SEM CORRIDA" explicado) | Painel/Analitos (OC-0070) | número certo; mensagem confunde | menos dúvida do usuário | Média | Baixo |
| BK-03 | Rótulos: "pior corrida do período" no status; "Últ. violação"/"Classificação" = lote inteiro | Painel (OC-0025, 0046) | interpretação | leitura correta da tela | Média | Baixo |
| BK-04 | Rotular o Sigma dos blocos da aba Analitos como "do cadastro" (não oficial) | Analitos (OC-0010) | 3 Sigmas na tela; o oficial é Estatística L = Painel I | evita confusão | Média | Baixo |
| BK-05 | Trocar a senha das abas (texto puro no repositório público) | Segurança (OC-0067) | barreira real é o login PBKDF2 + permissão da pasta | reduz exposição | Alta (até ~2 semanas) | Médio (entrega própria) |
| BK-06 | Aviso quando o provedor digitado no CEQ ≠ provedor do bias do analito | mCEQForm (OC-0069) | latente (mapa Controllab vazio) | evita "entra no bias" falso | Média | Baixo |
| BK-07 | Checagem de papel (SUPERVISOR/ADMIN) para editar a coluna Material | mMaterialLote / Configuração (ADR-099) | hoje vale a proteção da aba | governança | Média | Baixo |
| BK-08 | LÍQUOR e SANGUE na lista de material da Bio | de/para + ADR-099 | exige linhas novas no de/para (decisão do laboratório) | cobre outros controles | Baixa | Médio |
| BK-09 | Desempenho do Inativar (20–30 s) e do "aguarde N s" fixo | mIntegracao / UX (OC-0060, 0063) | usabilidade, sem efeito no resultado | rotina mais rápida | Média | Médio |
| BK-10 | Botão CIQ (20 primeiros) passar a 40 como o ADR-100 | mLotes.PreencherVaziosCIQ | coerência com o alvo do laboratório | padronização | Baixa | Baixo |
| BK-11 | Exclusão de outliers (|z|>3) e DP ≥ 20 dias no alvo do laboratório | aplicar_alvos_ciq (ADR-100) | CLSI C24; hoje sem exclusão (revisão do RT pelo CSV) | alvo mais robusto | Média | Médio |
| BK-12 | Alvo por equipamento na Bio (DIMENSION 1 × 2) | LotesStore (ADR-073 limitação) | alvo único mascara viés entre aparelhos | CQ por aparelho | Média | Alto (estrutura) |
| BK-13 | Power BI: BiasEQ com o mesmo filtro de provedor/rodada da tela | mBI (Bio) | BI diverge da tela | coerência | Baixa | Baixo |
| BK-14 | `NumeroCEQ`: "1.234" sem vírgula lido como 1,234 | mCEQForm | CEQ ≥ 1000 com ponto de milhar | entrada robusta | Baixa | Baixo |
| BK-15 | Testes: R12c (qa_adr088), qa_corrida_lote e qa_blocos_lote BN escolherem dados que cumpram a pré-condição | testes/ | falham em cópia com dados antigos (ambiente) | gate menos sensível a dados | Baixa | Baixo |
| BK-16 | Teste automático do retrato na seleção (R10) | testes/qa_adr097 (OC-0087) | comportamento validado à mão | cobertura | Baixa | Baixo |
| BK-17 | Itens de rótulo/visual: I01 sem matriz (0022), "Outro" sem complemento (0026), H3 (0037, 0045), período ao voltar (0038), ponto cortado (0047), lote de urina sem gráfico (0061), GUIA "etapa k de 7" (0064), bloco por ano no Eng_Saida (0077) | Painel/QA/GUIA | todos com orientação (PRONTIDAO §2, classe B) | acabamento | Baixa | Baixo |
| BK-18 | Lentidão pela tarefa DB_SEAC a cada 2 min | automação do laboratório (OC-0065) | ambiente | rotina mais fluida | Baixa | Médio |
| BK-19 | Triagem do `--simular` (ADR-100): critério de DP inflado (DP novo > 2× bula), teste binomial no % fora de 3 DP, ≥ 20 dias (CLSI C24), régua D1×D2 pelo DP da bula | aplicar_alvos_ciq / revisão do RT | parecer CQ 24/11: os critérios (a)–(e) deixam passar o N1 dos lotes 8974/8977 | revisão do RT mais segura | Alta | Baixo (só relatório) |
| BK-20 | F04 do qa_fluxo_manual (R_4s com manual) com dado sintético em vez de depender dos dados reais | testes/qa_fluxo_manual | NAO_EXECUTADO nos dois produtos em 24/11 (sem caso válido nos dados) | cobertura da regra | Média | Baixo |
