# Tester/QA — ciclo 2026-11-14_reteste04 — modo VALIDAR (entrega 41)

Data: 14/11/2026. Correções: ADR-090, ADR-092, ADR-093, ADR-094 e OC-0066, em produção desde a entrega 41 (commit
1a72892; o commit 22c27f3 só corrige um teste). Quem corrigiu foi o agente programador, por meio do Orquestrador. Eu não
corrigi nada e não editei `src/`, instaladores, fórmulas nem suítes.

## 0. Resumo

| Produto | Item | Veredito | Evidência (cópia nova da produção; números próprios) |
|---|---|---|---|
| Bio | OC-0012 | **VALIDADA** | Glicose · DIM 1 · 8974 · Tudo: M1 = 39 = 219 corridas elegíveis − 180 (oráculo por RUN); AF1 "N1 174 · N2 161" = minha contagem; janela N1 140 / N2 131, média 89,15 / 281,98473, DP 16,96799 / 4,92245 = minha conta sobre as 180 corridas mais recentes (dif. ≤ 1,6e-11). A39 "…fora do gráfico **e dos indicadores do Painel** (máx. 180). Período inteiro: N1 174 · N2 161…". Ampliação no **DIMENSION 2** (PCR · 8353): M1 83 = 263 − 180, AF1 "N1 244 · N2 246" = meu, n 170/166 = meu. Vizinho Nov/2025: M1 0, sem aviso, n 14/12 = meu |
| Bio | OC-0019 | **VALIDADA** | Painel em **DIMENSION 1**, manual sem EQUIPAMENTO: "DIMENSION 1", fundo 10284031 (RGB 255,235,156), fonte 22428, nota "Preenchido pelo Painel com DIMENSION 1: confira antes de ATUALIZAR DADOS."; DIMENSION 2 escrito por cima → fundo 16777215, sem nota; A3 cita EQUIPAMENTO e o âmbar. (A validação de dados da coluna, R4-01, fica fora, como pedido) |
| Bio | OC-0023 | **VALIDADA** | 1º ATUALIZAR: sem NÍVEL (MAN_0001), sem HORA (MAN_0002) e RESULTADO "xyz" (MAN_0005) → só RESULTADO_MANUAL_RECUSADO; conflito +3 min (MAN_0003, real BIO-368740) → só RESULTADO_MANUAL_EM_CONFLITO. Completar o NÍVEL + ATUALIZAR só entradas → ATIVO, RUN 26102301, exatamente 1 RESULTADO_MANUAL_INCLUIDO e nenhuma trilha repetida dos outros 9 manuais |
| Bio | OC-0024 | **VALIDADA** | E04 = 3, E01 = 4 (os 4 E01 reais de governança já descritos no reteste 03: inativações sem motivo). ResumoQA: "• 3 resultado(s) manual(is) incompleto(s)…" + "• 4 erro(s) de governança…"; nenhum "ERRO DE GOVERNANÇA" |
| Bio | OC-0031 | **VALIDADA** (função; clique na caixa Sim/Não não executado) | Fosfatase alcalina (CAP "Alkaline Phosphatase"): sem alvo → "OK\|547\|Gravado fora do bias (sem valor-alvo)."; G **bit a bit igual** (10,106578562965522); só muda a coluna S (verificação dos limites, "15 → 16 dentro"), como a confirmação avisa. Com alvo 100 (+10 %): G = 10,079933922224143; oráculo das duas etapas com k = 3 rodadas (C-A/C-B/C-C 2025 na EQA_Base): (3·G0 + 10)/4 = 10,079933922224143 |
| Bio | OC-0033 | **VALIDADA** | Arquivo recém-aberto: SalvarCEQ "ERRO\|0\|Sessão encerrada…", 545 → 545 linhas, 0 trilha; AbrirCEQ → caixa "Controle externo (CEQ)" / "Faça o login para registrar resultados do CEQ." (texto da janela do meu PID), sem frmCEQ |
| Bio | OC-0034 | **VALIDADA** | DistintosCEQ (CAP) = oráculo das linhas reais: rodada 3/3, analito 31/31, avaliação 2/2; 0 itens só da simulação; Controllab vazio |
| Bio | OC-0035 | **VALIDADA** (texto) | GUIA §3 cita EQUIPAMENTO e âmbar; §5–6 têm os 9 termos (janela, login, valor-alvo, Corrigir, Excluir, Consolidar, Z3, ATUALIZAR, FÓRMULA SOBRESCRITA) e "não digite"; nenhuma frase manda digitar nas abas EQA; botões existem ("Registrar resultado do CEQ" nos 2 Resumos, "Consolidar" nas 2 abas de dados, ATUALIZAR) |
| Bio | OC-0036 | **VALIDADA** | 62 textos na AT: 58 "x,xx / y,yy" (29 valores distintos, ex.: AT14 "3,69 / 0,26") e 4 "SEM EP / SEM EP"; nenhum "004 / 000" |
| Bio | OC-0040 | **VALIDADA** (ressalva: R4-02/R4-F2-03, já registradas) | D4 `=IF(Painel!K3="","",Painel!K3)`, B4 travada (escrita recusada), A4 "Visão (segue o Painel)". Ago/2026, T2/2026, Ano 2026, De/Até 20/12/2025–10/01/2026 e Tudo: nenhum "Ano 0"; De/Até atravessando o ano → K3 e D4 vazios. Glicose 8977 DIM 1: n/média/DP/CV da Estatística = meus (Ago 20/19; Ano 90/88; dif. ≤ 3e-13). No estado Tudo continuam D4 "2026" e B5/D5 "00/01/1900" (resíduo de apresentação já registrado pelo Usuário) |
| Bio | OC-0041 | **VALIDADA** | Painel protegido: K3 2026 + T1 → 01/01–31/03; De/Até 15/08–20/10/2026 → K4 "Personalizado", K3 2026; K4 Set → 01/09–30/09; De/Até 10/11/2025–05/02/2026 → K4 "Personalizado", K3 vazio, nenhum "Ano 0" |
| Bio | OC-0043 | **VALIDADA** | J7/J8: 48 pt úteis, ShrinkToFit, sem quebra, Segoe UI 10 (medido em negrito, conservador). Os 7 textos curtos tirados da **própria fórmula** do J7 + REJEITADO/ALERTA/OK: pior fator 0,736 (SEM CORRIDA) ≥ 0,70; REJEITADO 0,928. Q4 tem a forma longa dos 7. Estados reais: 8974 Tudo "REJEITADO"; Fev/2024 "PER. VAZIO" (Q4 "SEM CORRIDA NO PERÍODO"); 6353 "SEM MÉD/DP" |
| Bio | OC-0044 | **VALIDADA** | 96 linhas de analito, A:B = −4108 (centro) em todas, inclusive as de 33 pt (Lactato, Cálcio, Colesterol total, HDL, Creatinina) |
| Bio | OC-0050 | **VALIDADA** | A5 "Total: 35.906 · participam: 35.763 · X vermelho: 4 · inativados: 11 · manuais: 0 · substituídos por manual: 0" = minha contagem; depois do teste "35.916 · 35.765 · 4 · 12 · 10 · 0" = minha contagem |
| Bio | OC-0052 | **VALIDADA** | Mapa real (37 nomes): só o par de palavra "Lactate" → cita “Lactate Dehydrogenase”. Par **sem espaço** sintético na cópia ("TESTE-AUDITORIA R4V" → ALT, "…R4VX" → AST): o menor cita “TESTE-AUDITORIA R4VX”, o maior não cita par; linhas apagadas |
| Bio | OC-0054 | **VALIDADA** | EQA.CAP_Dados linha 7 (AST, 149), AB1 selecionada: mesmo valor → 0 trilha, Z3 "base em dia"; 151 → 1 CEQ_ALTERADO "Seu Resultado: 149 → 151", sem "?", Z3 "HÁ EDIÇÃO NÃO CONSOLIDADA" |
| Bio | OC-0055 | **VALIDADA** | "CAP TESTE-AUDITORIA R4V 2025 · Alkaline Phosphatase · TV-02 (resultado 110)"; lista da rodada, trilha e duplicidade ("ERRO\|5\|Já existe este registro (…) na linha 548…") sem "2025 2025" |
| Bio | OC-0056 / ADR-093 | **VALIDADA** (caso ambíguo NAO_EXECUTADO) | Oráculo próprio (de/para × tblDB_Recebimento): na URINA aceitam Creatinina→CRE2, Sódio→NA, Glicose→GLUC, Potássio→K, Cloro→CL. Lote 6354: "Creatinina" (MAN_0008) e **"Sódio"** (MAN_0009) → ATIVO, PARTICIPA SIM, ANALITO = ANALITO_ORIGEM = CRE2 / NA, ANALITO_QCINI e BLOCO vazios, MATRIZ URINA, DETALHE "…lancado como 'CRE2'/'NA'…", QA **I05** (e A06, manual no dia de resultados do equipamento), trilha INCLUIDO com a nota. **"Ácido úrico"** (URCA nunca recebido na urina, MAN_0010) → LOTE_OUTRO_MATERIAL, fora da estatística, **A13**, trilha FORA_DA_ESTATISTICA |
| Bio | OC-0057 | **VALIDADA** | Depois do Logout (SessaoAtiva False): Corrigir e Excluir "ERRO\|0\|Sessão encerrada…", F548 110 → 110, linha não apagada, 547 → 547 linhas, nenhuma trilha CEQ_*/EQA_*, G igual. Login de novo: Corrigir 105 → G = 8,829933922224143 = (3·G0 + 5)/4 (dif. 2e-16); excluir tudo → Estatística A14:AU133 **bit a bit** igual ao início |
| Bio | OC-0058 | **VALIDADA** | 1º ATUALIZAR: IMPLAUSIVEL (−1, MAN_0004), LOTE_FORA_DO_CQ (8794, MAN_0006), LOTE_OUTRO_MATERIAL por SO_ANALITOS (8353 + Glicose, MAN_0007) e por MATRIZ (6354 + Ácido úrico, MAN_0010) → só RESULTADO_MANUAL_FORA_DA_ESTATISTICA; ATIVOS → INCLUIDO; 2º ATUALIZAR não repete |
| Bio | OC-0059 | **VALIDADA** (documentação) | ADR-090.md:20 e :36 chamam a correção do G23 de OC-0041 ("antes rotulado OC-0040"); nenhuma linha atribui o G23 só a OC-0040; OCORRENCIAS: OC-0040 = "Estatística MENSAL/TRIM", OC-0041 = "Painel K3/K4" |
| Bio | OC-0066 | **VALIDADA** | Inativei BIO-343077 e rodei o ATUALIZAR completo (34,9 s): tblCQ_Final[DATA_INATIVACAO] "14/11/2026 20:32:11" (largura 17,71, sem ####); Sync_Log 18 → 19 linhas, nova INICIO "14/11/2026 20:32:53", FIM "…20:33:13" |
| Bio | ADR-091 | **VALIDADA** | 17 regras (10 MATRIZ, 6 FORA, 1 SO). Oráculo próprio antes e depois do ATUALIZAR: **0 violações**; 8177 (29) e 8794 (30 → 31) todos LOTE_FORA_DO_CQ; 8353: PCR 977 ATIVO + 1 IMPLAUSIVEL, outros 30 → 31 LOTE_OUTRO_MATERIAL, nenhum participando; 5657/5658/5659/6353/6354: 1.154 → 1.157 linhas, todas URINA, 0 analito de soro participando; A13 = LOTE_* = 89 → 92. Painel no lote 8353 (PCR, DIM 2): n = meu oráculo dos que participam |
| Bio | ADR-093 | **VALIDADA** | Ver OC-0056 acima (mesma execução) |
| Bio | ADR-094 | **VALIDADA** | Varredura própria de **todas** as tabelas (26 colunas com formato de data, 638 células: 1ª, última, 25 espalhadas e as linhas tocadas): 0 "####" e 0 data-hora sem segundos no escopo; texto = minha formatação do Value2 (dia antes do mês). DATA só data; DB A5 "último resultado: 13/11/2026 08:39:01 · última chegada: 14/11/2026 14:34:33"; manuais REGISTRADO_EM "14/11/2026 20:31:48". Fora do escopo do ADR (decisão do ADR): Audit_Log DataCorrida só data, BI_Data aaaa-mm-dd |
| Bio | (efeito colateral) | **sem efeito** | Mesmos dados, antes × depois da entrega 41 (`_entrega_2026-11-14_0715`): Estatística A14:AU133 + Painel B7:I8 iguais **bit a bit** em 3 estados do Painel (≈1.700–1.780 células numéricas cada) |
| Hema | OC-0023 | **VALIDADA** | Sem NÍVEL, sem HORA, RESULTADO "xyz" → só RECUSADO; conflito +3 min (real HEM-367846, "coincide… (2 min)") → só EM_CONFLITO; completar o NÍVEL → ATIVO, RUN 26102201, 1 INCLUIDO, nada repetido |
| Hema | OC-0024 | **VALIDADA** | E04 = 3, nenhum outro ERRO; ResumoQA "QA da integração: 3 erro(s)…" + "• 3 resultado(s) manual(is) incompleto(s)…", sem "governança" |
| Hema | OC-0031 | **VALIDADA** (função; clique não executado) | MCHC: sem alvo → "OK\|467\|Gravado fora do bias…", G **bit a bit igual** (1,606529458857288; só a coluna R, \|SDI\| dos limites, muda, como avisado); com alvo 100 → G 3,704897094142966 = (3·G0 + 10)/4 exato (k = 3 = FH9-A/B/C 2025) |
| Hema | OC-0033 | **VALIDADA** | SalvarCEQ sem sessão "ERRO\|0\|…", 465 → 465, 0 trilha; AbrirCEQ: caixa "Faça o login para registrar resultados do CEQ.", sem frmCEQ |
| Hema | OC-0034 | **VALIDADA** | DistintosCEQ = oráculo: rodada 3/3, analito 25/25, avaliação 3/3; 0 da simulação |
| Hema | OC-0035 | **VALIDADA** (texto) | Mesmo GUIA (comum): mesmos resultados da Bio; botões existem na Hema |
| Hema | OC-0036 | **VALIDADA** | 84 textos na AT: todos "x,xx / y,yy" (20 valores distintos, AT14 "2,49 / 0,04") ou "SEM EP / SEM EP" (analitos sem CEQ, a partir da linha 68); nenhum "004 / 000" |
| Hema | OC-0041 | **VALIDADA** | Mesma sequência da Bio, Painel protegido: T1 → 01/01–31/03/2026; De/Até → "Personalizado"; Set → 01/09–30/09; atravessando o ano → K3 vazio, sem "Ano 0" |
| Hema | OC-0049 | **VALIDADA** | Linhas 98–133 (36) sem nenhum texto de A a AU, no arquivo aberto e depois do ATUALIZAR; 84 linhas com analito e níveis 1/2/3; B98 `=IF($A98="","",1)`. Antes × depois da entrega 41: a única diferença na Estatística é essa (B98:B133 1/2/3 → vazio, 36 células) |
| Hema | OC-0050 | **VALIDADA** | A5 "Total: 48.391 · participam: 47.709 · X vermelho: 0 · inativados: 0 · manuais: 0 · substituídos: 0" = meu; depois "48.399 · 39.399 · 0 · 1 · 8 · 0" = meu (com as regras de teste) |
| Hema | OC-0052 | **VALIDADA** | Mapa real (29 nomes): 7 pares por prefixo, todos citados; "MCH" → “MCHC” (MCHC); "Lymphs" → “Lymphs Absolute” |
| Hema | OC-0054 | **VALIDADA** | Linha 7 (Basophils Absolute, 0,3): mesmo valor → 0 trilha; 2,3 → 1 CEQ_ALTERADO "0,3 → 2,3", sem "?", Z3 "HÁ EDIÇÃO NÃO CONSOLIDADA" |
| Hema | OC-0055 | **VALIDADA** | "CAP TESTE-AUDITORIA R4V 2025 · MCHC · TV-02 (resultado 110)"; lista, trilha e duplicidade sem "2025 2025" |
| Hema | OC-0057 | **VALIDADA** | Sem sessão: Corrigir/Excluir "ERRO\|0\|…", F468 110 → 110, linha mantida, 0 trilha CEQ; com sessão: Corrigir 105 → G 2,454897094142966 = (3·G0 + 5)/4 exato; excluir tudo → Estatística bit a bit igual |
| Hema | OC-0058 | **VALIDADA** | Regras de teste na cópia (602911 FORA_DO_CQ; 522611 SO_ANALITOS WBC): −1 (IMPLAUSIVEL), 602911 WBC (LOTE_FORA_DO_CQ), 522611 RBC (LOTE_OUTRO_MATERIAL) → só FORA_DA_ESTATISTICA; 522611 WBC → ATIVO + INCLUIDO |
| Hema | OC-0066 | **VALIDADA** | Cenário exato da OC: DATA_INATIVACAO **vazia** na produção. Inativei HEM-348142 + ATUALIZAR completo (88,5 s): "14/11/2026 20:37:10" (largura 17,71); Sync_Log nova linha INICIO "14/11/2026 20:38:50", FIM "…20:39:25" |
| Hema | ADR-091 | **VALIDADA** | Produção: 0 regras, 0 LOTE_*, A13 = 0. Com as 2 regras de teste + ATUALIZAR: 0 violações; 602911: 3.752 linhas, todas LOTE_FORA_DO_CQ; 522611: WBC 153 ATIVO + 3 DUPLICIDADE_ORIGEM, outros 4.561 LOTE_OUTRO_MATERIAL + 90 DUPLICIDADE_ORIGEM, nenhum participando; A13 = LOTE_* = 8.313; Principal "participam: 39.399" = meu. Painel dos lotes de teste: não conferido por mim |
| Hema | ADR-094 | **VALIDADA** | 26 colunas de data, 538 células: 0 "####", 0 data-hora sem segundos no escopo; DB A5 com segundos; fora do escopo: Audit_Log DataCorrida, BI_Data |
| Hema | (efeito colateral) | **sem efeito** (fora o OC-0049) | Antes × depois da entrega 41, mesmos dados: ver §5 |

**Totais:** Bio 25 itens: 25 VALIDADA (1 subcaso NAO_EXECUTADO: ADR-093 com código ambíguo). Hema 18 itens: 18 VALIDADA.
0 NAO_VALIDADA, 0 NAO_EXECUTADO no nível do item. Regressão (versão da entrega 41): Bio 52 PASS / 0 FAIL, Hema 54 PASS / 0 FAIL, ver §6.

Não validei (como pedido): OC-0021, OC-0012 da Hema e R4-01.

## 1. Ambiente e método

| Item | Valor |
|---|---|
| Produção copiada (nunca aberta) | `QC_Bioquimica.xlsm` sha256 `8ec815ab653db405`, `QC_Hematologia.xlsm` `aab0a54f464e0232`; o sha de cada cópia é igual ao da produção; a produção ficou igual depois de cada execução (`producao_intacta.igual = true`) e no fim da cadeia |
| Cópias | `%TEMP%\qcw\auditoria\2026-11-14_reteste04\tester\<produto>_valida_<hhmmss>.xlsm` (nova a cada execução, nunca salva, apagada no fim); regressão em `tester\regressao\<produto>\` |
| Excel | um por vez, aberto e fechado pelo `qa_final.QA`: validação Bio 26080, Hema 5620, reexecução estática da Bio 16620; comparações 7576, 14300, 12792, 20940, 28028, 23732; suítes 2200, 20340, 30804, 12292, 20888, 12060, 12324, 20068, 9776, 4432, 8304, 1648, 25844, 7552. Os EXCEL.EXE `/automation` da tarefa agendada (17396, 30160, 17068) não foram tocados; o último saiu sozinho. No fim não ficou nenhum Excel meu, e as 19 cópias `.xlsm` do `%TEMP%` foram apagadas. Nenhuma captura de tela: as caixas do AbrirCEQ foram lidas pelo texto das janelas do **meu** PID |
| Scripts (meus) | `_arquitetura/agentes/ferramentas/valida_reteste04.py` (validação) e `compara_entrega41.py` (antes × depois da entrega). Não reaproveitam os scripts do Usuário Simulado (`reteste04_usuario_*.py`) nem as suítes `qa_adr09x`. Reaproveitam só a infraestrutura: `qa_final.QA`, `repro_comum`, as classes `Audit`/`Vigia` do meu `valida_adr090_091.py`, `auditoria_abas.largura_pt`, `qa_inativar_ux.criar_usuario/fazer_login` |
| Cenários variados | em relação ao reteste 03 e ao Usuário: CEQ em Fosfatase alcalina (Bio) e MCHC (Hema), escolhidos pelo mapa; âmbar com o Painel em DIMENSION 1; janela de 180 corridas no DIMENSION 2; conflito +3 min; ADR-093 com Sódio (aceito) e Ácido úrico (recusado) escolhidos pelo meu oráculo; lotes de teste da Hema escolhidos pela varredura (602911, 522611); edição direta na linha 7 |
| Dados de teste (só nas cópias) | usuário `TESTE_AUDITORIA_R4V` (ANALISTA); manuais MAN_0001..0010 (Bio) e MAN_0001..0008 (Hema) com MOTIVO "Outro: TESTE-AUDITORIA", valores 222,222 / 211,111 / −1 / "xyz" / 50 / 51 / 12,345 / 3,21 / 4,32 (Bio), 7,11 / 4,22 / 7,33 (Hema); CEQ rodada "TESTE-AUDITORIA R4V 2025", amostras TV-00..TV-02 (110, alvo vazio/100; correção 105); 2 linhas de mapa sintéticas na EQA_Base da Bio (apagadas); regras tblLotesMaterial NOTA "TESTE-AUDITORIA R4V" (Hema); inativação "TESTE-AUDITORIA R4V inativação" e comentário de teste |
| Caminho | filtros pelas células do Painel com eventos ligados (Painel protegido como abre); manuais célula a célula na Digitar Resultados com eventos ligados; ATUALIZAR pelo núcleo do botão (`AtualizarDadosAutomatico`, DB_SEAC só lido; 0 resultados novos) e `AtualizarSoEntradas`; CEQ pelo núcleo dos botões (`SalvarCEQ`, `CorrigirCEQ`, `ExcluirCEQ`, `ConfirmacoesCEQ`, `AvisoAnalitoCEQ`, `DescricaoCEQ`, `RegistrosDaRodada`, `DistintosCEQ`) e o `AbrirCEQ` real; inativar/comentar pelo núcleo da janela do ponto |
| Logs/JSON (fora do repositório) | `tester\valida_<produto>.json/.log`, `valida_Bioquimica_estaticas.json/.log`, `compara_entrega41_<produto>.json/.log`, `regressao\<produto>\rodar_suites.log` e `saida\*.log/json`, `cadeia.log`. Não copiei os JSON para o repositório porque a trilha do Audit_Log da cópia leva nome, login e máquina de quem roda (o script mascara, mas a máscara da 1ª execução da Bio também cobriu o trecho "ini" de palavras, então os números deste relatório foram lidos com a máscara desfeita) |

**Não executado (sem interação com a tela):** o clique Sim/Não da caixa "Gravar mesmo assim?" do frmCEQ (validei a
função `ConfirmacoesCEQ` e o retorno do `SalvarCEQ`) e a digitação física em célula com validação de dados.

**Rodadas e ajustes do meu script (transparência).** Bio rodou inteira uma vez (346 s); depois de ler o JSON corrigi
três oráculos **do meu script** e repeti só a seção estática (48 s):
- O36: a AT vem de `mCEQ.ViesEQ` (não de células), então o meu oráculo passou a ser o formato "x,xx / y,yy" de todos os
  textos (não recalculo o viés);
- O43: o regex dos textos da fórmula do J7 quebrava nas aspas duplas (`""&engLote`); passei a ler os literais com o
  escape e a medir em negrito pela `DisplayFormat`;
- O50: o A5 separa os itens com dois espaços; normalizei os espaços.
Três vereditos internos do 1º JSON eram falso-negativos do oráculo e foram conferidos à mão no JSON: OC-0031 (o meu
critério exigia a Estatística toda igual, mas a amostra sem alvo **deve** contar na verificação dos limites — coluna S
na Bio, R na Hema — e o G ficou igual bit a bit; a contagem de rodadas da base tinha sido feita depois de gravar a
rodada de teste: 4 = 3 + a de teste), OC-0057 (a trilha LOGOUT entrava no "nenhuma trilha"; não há CEQ_*), OC-0012
ampliação (dif. de DP 1,06e-9 vem do arredondamento a 9 casas do `qa.painel_stats`, não do sistema; tolerância
justificada: 5e-10 absoluto sobre o valor arredondado). E, no ADR-093, o meu 1º oráculo dava a Creatinina como recusada
porque a tblDeParaAnalitos tem a linha URINA·CRE2 **sem** ANALITO_QCINI ("sem cadastro no QC_INI"); pelo ADR-093 essa
linha não é um de/para (é assim que a urina chega), e com essa leitura o oráculo dá CRE2 aceito, como o sistema fez.
A Hema rodou uma vez (372 s), já com os ajustes de O36, O50, O57 e da contagem de rodadas; o O31 da Hema teve o mesmo falso-negativo da coluna de limites (R) e foi conferido da mesma forma. O script versionado já tem todos os oráculos corrigidos (G no lugar da Estatística inteira no O31, linha URINA sem ANALITO_QCINI no ADR-093), para quem for repetir.

## 2. Bioquímica — detalhes

### OC-0012 (janela de 180 corridas)

| Cenário | M1 | AF1 | Meu (corridas elegíveis − 180; contagem do período) | n da janela (Painel / meu) | média e DP da janela |
|---|---|---|---|---|---|
| Glicose · DIM 1 · 8974 · Tudo | 39 | N1 174 · N2 161 | 219 − 180 = 39; N1 174 · N2 161 | 140/131 = 140/131 | 89,15 / 16,967987 e 281,984733 / 4,922452 = meus (≤ 1,6e-11) |
| PCR · DIM 2 · 8353 · Tudo (ampliação) | 83 | N1 244 · N2 246 | 263 − 180 = 83; igual | 170/166 = 170/166 | 0,992059 / 0,344545 e 3,638916 / 0,364938 = meus (≤ 1,1e-9, arredondamento do helper) |
| Glicose · DIM 1 · 8974 · Nov/2025 (vizinho) | 0 | vazio | 16 corridas, sem corte | 14/12 = 14/12 | = meus (≤ 9e-12) |

Oráculo: resultados com PARTICIPA = SIM e RESULTADO numérico do analito, lote, equipamento e período, na tblCQ_Final da
cópia; corridas = RUN distintos; janela = os 180 RUN mais recentes. Nenhuma corrida só com X vermelho nesses cenários.

### OC-0040 (Estatística segue o Painel)

| Estado do Painel | K3 / K4 | D4 | B5–D5 / H5–J5 | Glicose 8977 DIM 1, n N1/N2 (Estatística = meu) |
|---|---|---|---|---|
| Ago/2026 | 2026 / Ago | 2026 | 01/08–31/08/2026 | 20/19 = 20/19 |
| T2/2026 | 2026 / T2 | 2026 | 01/04–30/06/2026 | 0/0 = 0/0 |
| Ano 2026 | 2026 / Ano inteiro | 2026 | 01/01–31/12/2026 | 90/88 = 90/88 (média 77,3 / 272,4659; DP ≤ 3e-13) |
| De/Até 20/12/2025–10/01/2026 | vazio / Personalizado | vazio | 20/12/2025–10/01/2026 | 0/0 = 0/0 |
| Tudo | 2026 / Tudo | 2026 | "00/01/1900" / "histórico inteiro do lote" | 90/88 = 90/88 |

O defeito ("Ano 0", visões inalcançáveis) não ocorre. O estado Tudo continua com o cabeçalho confuso (D4 "2026",
B5/D5 "00/01/1900"). Isso já foi registrado pelo Usuário Simulado como R4-02 e R4-F2-03 (Baixo). Não é efeito novo da
correção: antes ela mostrava "Ano 0" também nesse estado.

### Manuais (OC-0019, 0023, 0024, 0056, 0058; ADR-091/093)

| Manual | Entrada | STATUS (tblCQ_Final) | Trilha do 1º ATUALIZAR | QA |
|---|---|---|---|---|
| MAN_0001 | Glicose 8977, sem EQUIPAMENTO e sem NÍVEL | MANUAL_INCOMPLETO ("…: NIVEL") → depois ATIVO, RUN 26102301 | RECUSADO → (2º) INCLUIDO | E04 |
| MAN_0002 | sem HORA | MANUAL_INCOMPLETO ("…: HORA") | RECUSADO | E04 |
| MAN_0003 | +3 min do real BIO-368740 | CONFLITO_MANUAL ("coincide… (3 min)") | EM_CONFLITO | A01 |
| MAN_0004 | −1 | IMPLAUSIVEL | FORA_DA_ESTATISTICA | A11 |
| MAN_0005 | RESULTADO "xyz" | MANUAL_INCOMPLETO ("…: RESULTADO") | RECUSADO | E04 |
| MAN_0006 | 8794 Glicose 50 | LOTE_FORA_DO_CQ | FORA_DA_ESTATISTICA | A13 |
| MAN_0007 | 8353 Glicose 51 | LOTE_OUTRO_MATERIAL (SO_ANALITOS) | FORA_DA_ESTATISTICA | A13 |
| MAN_0008 | 6354 "Creatinina" 12,345 | ATIVO, CRE2, URINA, nota | INCLUIDO + nota | I05, A06 |
| MAN_0009 | 6354 "Sódio" 3,21 | ATIVO, NA, URINA, nota | INCLUIDO + nota | I05, A06 |
| MAN_0010 | 6354 "Ácido úrico" 4,32 | LOTE_OUTRO_MATERIAL (MATRIZ) | FORA_DA_ESTATISTICA | A13 |

O A06 dos dois aceitos é "Resultado manual e do interfaceamento no mesmo dia (fora da tolerância)": o lote 6354 tem
urina do equipamento em 23/10, então é esperado. O caminho do código ambíguo (mais de um ANALITO_ORIGEM) ficou
`NAO_EXECUTADO`: nenhum analito do catálogo da Bio tem mais de um código no de/para.

## 3. Hematologia — detalhes

- Os mesmos fluxos da Bio, com WBC N1 do lote 625311 (referência HEM-367846, 22/10/2026), sem âmbar (XN1000, fundo
  branco, sem nota: a OC-0019 é só da Bio).
- **Regras de teste** (a Hema não tem regras na produção, como diz o ADR-091 §5): escolhi os dois lotes menores
  diferentes do da referência, 602911 (3.751 linhas) como FORA_DO_CQ e 522611 (4.805 linhas) como SO_ANALITOS WBC.
  Consequência só na cópia: participam 47.709 → 39.399.
- ATUALIZAR completo: 88,5 s. Só entradas: 35,1 s. A Bio levou 34,9 s e 20,2 s. Os dois estão dentro do que o Usuário
  mediu hoje (Hema 27–32 s sem regras; com 8.313 linhas reclassificadas o tempo maior é esperado).

## 4. ADR-094 / OC-0066 — varredura

O meu oráculo **não** usa a lista `COLUNAS` do instalador. Ele anda por todas as tabelas de todas as abas e pega cada
coluna cujo formato tem dia/mês/ano. Em cada uma confere a 1ª e a última linha, 25 linhas espalhadas e as linhas tocadas
pelo teste:
- `#` no texto é falha;
- valor com hora tem de dar texto `dd/mm/aaaa hh:mm:ss`, igual ao Value2 formatado por mim, arredondado ao segundo;
- coluna DATA tem de ficar só com a data.

| | Colunas de data | Células | No escopo | Fora do escopo (decisão do ADR-094) |
|---|---|---|---|---|
| Bio | 26 | 638 | 0 achados | Audit_Log DataCorrida só data; BI_Data aaaa-mm-dd |
| Hema | 26 | 538 | 0 achados | idem |

## 5. Efeito colateral nas estatísticas (antes × depois da entrega 41, mesmos dados)

Abri sozinhos, cada um numa cópia, `_entrega_2026-11-14_0715/<produto>/producao_original` (sha Bio `8b1cda1345d2`, Hema
`f1907656b663`: a produção de antes da entrega) e o instalado da entrega 41 (Bio `c9517a151ee7`, Hema `4c6a626280de`).
Pus os mesmos filtros no Painel e li Estatística A14:AU133 e Painel B7:J{6+n} (Value2).

| Produto | Estado | Células numéricas | Diferenças |
|---|---|---|---|
| Bio | salvo (Proteína total 8977, 2026, Out) | 1.701 | **0** |
| Bio | Glicose 8977 DIM 1, Ano 2026 | 1.776 | **0** |
| Bio | Cloro 8974 DIM 2, Tudo | 1.760 | **0** |
| Hema | salvo (RBC 522611 Tudo) | 1.951 | 36 = B98:B133 1/2/3 → vazio (OC-0049) |
| Hema | WBC 522611, Ano 2026 | 1.951 | 39 = as mesmas 36 + J7:J9 "SEM CORRIDA NO PERÍODO" → "PER. VAZIO" (forma curta do OC-0043, ADR-092; o Q4 mantém a longa) |
| Hema | PLT Tudo | 1.951 | 36 = idem |

## 6. Regressão (suítes, cópia da produção atual, uma por vez)

Cada suíte rodou numa cópia nova da produção atual (sha igual ao da produção), pelo `rodar_suites.py --suites`, uma por
vez. A cadeia toda foi sequencial, com um Excel por vez.

**Atenção à versão das suítes.** O repositório já está à frente da entrega 41. Os commits 69a1f65, 39084f1 e e7490ba
trazem o ADR-095, da próxima entrega, e mudaram `qa_adr090.py` e `qa_adr092.py`. Por isso rodei as cinco suítes na
árvore atual e repeti as duas alteradas na árvore do commit **22c27f3**, que é o teste da entrega 41 (`git archive`
para `%TEMP%`, fora do repositório).

| Produto | Suíte | Árvore | PASS / FAIL | Tempo | Histórico (entrega 41, 14/11) |
|---|---|---|---|---|---|
| Bio | qa_adr089 | atual (= 22c27f3) | 10 / 0 | 234 s | 223 s |
| Bio | qa_adr090 | atual | 21 / 0 | 256 s | 248 s |
| Bio | qa_adr090 | 22c27f3 | 21 / 0 | 243 s | 248 s |
| Bio | qa_adr091 | atual (= 22c27f3) | 7 / 0 | 259 s | — |
| Bio | qa_adr092 | atual | 7 / 0 | 53 s | ~53 s |
| Bio | qa_adr092 | 22c27f3 | 7 / 0 | 53 s | ~53 s |
| Bio | qa_adr094 | atual (= 22c27f3) | 7 / 0 | 45 s | ~47 s |
| Hema | qa_adr089 | atual (= 22c27f3) | 10 / 0 | 207 s | 201 s |
| Hema | qa_adr090 | atual | **22 / 1** (O12) | 250 s | 258 s |
| Hema | qa_adr090 | 22c27f3 | **23 / 0** | 252 s | 258 s |
| Hema | qa_adr091 | atual (= 22c27f3) | 7 / 0 | 226 s | — |
| Hema | qa_adr092 | atual | 7 / 0 | 55 s | ~55 s |
| Hema | qa_adr092 | 22c27f3 | 7 / 0 | 55 s | ~55 s |
| Hema | qa_adr094 | atual (= 22c27f3) | 7 / 0 | 57 s | ~53 s |

- **Totais na versão da entrega 41:**
  - Bio: 52 PASS / 0 FAIL nas 5 suítes;
  - Hema: 54 PASS / 0 FAIL.
- **Totais na árvore atual:**
  - Bio: 52 / 0, em 849 s;
  - Hema: 53 / 1, em 796 s.
- **Tempo:** sem lentidão. Os tempos ficaram dentro de ±5 % dos da entrega 41 (logs em `_entrega_2026-11-14_0715/<produto>/`).
- **O FAIL do O12 na Hema não é regressão.**
  - O `qa_adr090.py` atual (ADR-095, linhas 182–186) já espera o texto novo do aviso, sem "veja a aba Estatística" e
    com "Período inteiro (todos os lotes): …".
  - Esse texto ainda não está em produção.
  - Os números conferem: 278 corridas cortadas e AF1 "N1 374 · N2 400 · N3 340" = contagem independente.
  - É o mesmo ponto da OC-0012 da Hema (R4-H01), que ficou fora desta validação.
  - Na versão do teste da entrega 41, o O12 passa.
- **Ressalva:** as suítes não alteradas rodaram com o `camada_dados.py` e os instaladores atuais. Elas só os leem, e
  essas mudanças também são do ADR-095. Não repeti o qa_adr089, o qa_adr091 nem o qa_adr094 na árvore 22c27f3.
- A produção ficou igual no fim das duas cadeias: Bio `8ec815ab653db405…`, Hema `aab0a54f464e0232…`.
- Logs:
  - `%TEMP%\qcw\auditoria\2026-11-14_reteste04\tester\regressao\<produto>\rodar_suites.log` e `saida\*.log/json`;
  - `…\tester\regressao22c\<produto>\…`;
  - `cadeia.log` e `cadeia2.log`.

### Testes automatizados propostos

Nenhum caso novo é exigido pelos itens validados. Proponho só tornar permanente o que este VALIDAR fez à mão:
- **qa_adr092 K07** — antes × depois da entrega, com os mesmos dados: a Estatística A14:AU133 e o Painel ficam iguais
  bit a bit em 3 estados do Painel, salvo as células listadas pela própria entrega. É o `compara_entrega41.py`, que
  usa a `producao_original` que o `entregar.py` já guarda.
- **qa_adr093 (nova) A01** — o oráculo do ADR-093 pela tblDeParaAnalitos × tblDB_Recebimento. Uma linha URINA sem
  ANALITO_QCINI **não** conta como de/para. O teste cobre um aceito diferente da Creatinina (ex.: Sódio → NA) e um
  recusado por "nunca recebido" (ex.: Ácido úrico → URCA).

## 7. Ocorrências novas

Nenhuma ocorrência nova nos itens validados. Resíduos já registrados neste ciclo e que confirmei de passagem:
- R4-02 / R4-F2-03 (Estatística no estado "Tudo");
- OC-0021: o DETALHE do IMPLAUSIVEL, LOTE_* e da nota do ADR-093 continua sem acento ("implausivel", "estatistica",
  "lancado"; ex.: MAN_0004, MAN_0008).
