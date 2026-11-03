# Bioquímica — F3: auditoria matemática independente (ciclo 2026-10-31_ciclo01, completo)

**Agente:** Usuário Simulado · **Data:** 31/10/2026, reexecutada por inteiro em 01/11/2026 · **Excel:** não aberto
nesta frente (instrução do Orquestrador; outra frente usava o Excel).

**Reexecução de 01/11/2026 (terceira passada, pedida pelo Orquestrador):** a produção continua com o sha256 curto
`fe10dba1f1fd` (última gravação 30/10/2026 16:33) e `src/` não mudou desde o commit `0a25f5a` (`git status` limpo fora
de `_arquitetura/agentes/`). Fiz uma cópia nova (`shutil.copy2`, hash conferido igual) e repeti tudo:

- a calculadora deu de novo 62 linhas × 19 colunas e 0 divergências;
- escrevi uma segunda opinião do zero para **outros 6 analitos** × 2 níveis (F3-T24, 96/96 valores);
- reconferi o Painel contra o independente (F3-T10);
- reconferi, na cópia nova, a evidência de cada ocorrência: fórmulas de H7/H8 e B103, AT14/AT20, G76:G93,
  biblioteca 2026–2029, os IDs da OC-3 e os 253 suspeitos.

Nada mudou. As ocorrências OC-1 a OC-7 continuam válidas, com a mesma evidência.

**Cópia auditada:** `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\Bioquimica_F3.xlsm` (`shutil.copy2` da raiz).
**sha256 curto da produção copiada:** `fe10dba1f1fd` (cópia idêntica, conferida).
**Reexecução de conferência (31/10/2026, fim da tarde):** a produção continua com o mesmo sha256 (`fe10dba1f1fd`,
última gravação 30/10/2026 16:33). Copiei de novo para a pasta do ciclo, rodei a calculadora (versão atual, que já
inclui a extensão da Hematologia) e o resultado se repetiu: 62 linhas × 19 colunas, 0 divergências. Além disso,
escrevi uma **segunda opinião do zero** (script curto, sem importar a calculadora) para 4 analitos × 2 níveis
(F3-T23), conferi de novo no código as definições da seção 1 e reconferi a evidência de cada ocorrência.
**Leitura:** openpyxl, `read_only=True`, `data_only=True`. Os valores do QC_INI são os gravados na última
recalculação do Excel (Value2, precisão total). Os dados brutos vêm das tabelas.
**Calculadora:** `_arquitetura/agentes/ferramentas/calc_independente.py`. É Python puro (`statistics`/`math`) e não lê
nenhuma fórmula do QC_INI para calcular. O cabeçalho documenta o uso e a definição. Para reproduzir:

```
set PYTHONIOENCODING=utf-8
python -I _arquitetura\agentes\ferramentas\calc_independente.py "%TEMP%\qcw\auditoria\2026-10-31_ciclo01\Bioquimica_F3.xlsm" --json f3.json
```

**Filtros em vigor na cópia (entradas do usuário):**

| Filtro | Valor |
|---|---|
| Painel | Proteína total, lote em análise H3 = 8977, equipamento L4 = DIMENSION 2, De/Até 01/10/2026–31/10/2026 |
| Estatística | Visão PERSONALIZADO (01/10–31/10/2026); lote H4 = 8977; Provedor por analito (Especificações AR = CAP nos 31); Ano EP N4 = 2025; Rodada P4 = TODAS; Sigma L11 = 6 meses (29/04/2026–28/10/2026); sem exclusões |
| Anos de referência do lote 8977 | CLIA 2025 · VB 2025 |

Esses filtros também estão no JSON.

## 1. Definição confirmada no código (antes de calcular)

| Grandeza | Onde | Definição que vale (conferida) |
|---|---|---|
| n, Média, DP, CV (Estatística C:F) | `src/bio/mEstatPeriodo.bas` `EstatPeriodo`/`Agregar`; `mDados.CarregarDB` | tblCQ_Final, `PARTICIPA_ESTATISTICA = SIM`, equipamento = Painel!L4, `DATA_HORA` truncada ao dia dentro da janela e fora das exclusões, RESULTADO numérico, LOTE = Estatística!H4. DP amostral (n−1), CV = DP/média × 100 |
| Bias EQC (G abs, T sinal, AC n, AD rodadas) | `src/comum/mCEQ.bas` `BiasEQ`, `Casa`, `AnoVigente`; ADR-035/071/072 | EQA_Base, Analito_Canonico, `Uso_Analitico <> NAO`, provedor de Especificações!AR. Com provedor e ano explícitos, TODAS usa **exatamente** o ano. Duas etapas: média dos \|bias\| de cada rodada (ano\|rodada), depois média das médias |
| ETp, CVTp (K, J) | Especificações S/T/U; ADR-069 | CLIA → K (biblioteca do ano CLIA do lote); VB → 1,65·f·CVi + f/2·√(CVi²+CVg²), f = 0,25/0,50/0,75 (ÓTI/DES/MIN); FAB → M. Vazio ou ≤ 0 = "-" (ADR-068) |
| ET % (H) | fórmula da Estatística; ADR-071 nota Q11 | 1,65 · CV do lote no período + \|bias\| |
| CV do Sigma | `src/comum/mIncerteza.bas` `IncertezaCIQ "CV"`; ADR-064/071 | janela [EDATE(MU_Fim, −6)+1, MU_Fim]. MU_Fim = maior DATA com PARTICIPA = SIM (tabela inteira). Equipamento do Painel; lotes com n ≥ 20 e média ≠ 0; CVp = √(Σ(nᵢ−1)CVᵢ² / Σ(nᵢ−1)) |
| Sigma (L) | Estatística L; ADR-035/068/071 | (ETp − \|bias\|) / CVp |
| Margem, classe, DPM, rendimento (N, O, M, V, W) | `mPlanoQC`, `Cfg_PlanoQC`; ADR-035 | N = ETp − ET; O = N/ETp × 100; DPM = [1 − Φ(σ − 1,5)] × 10⁶; rendimento = 100 − DPM/10⁴; faixas ≥6/5/4/3 |
| Viés com sinal e RMS (AR, AT) e u(Rw) 12 m (AF) | `mCEQ.ViesEQ`, `mIncerteza`; ADR-064/067 | só amostras avaliadas, resultado e alvo ≠ 0, cada chave uma vez; u(Rw) = o mesmo CV pooled numa janela de 12 meses |

## 2. Amostra de risco (5 analitos, todos os níveis)

| Analito | Risco que representa | Por quê |
|---|---|---|
| Proteína total | **2 lotes** | Na janela de 6 m do Sigma entram os lotes 8974 (n = 44) e 8977 (n = 90). É o analito do Painel, o que permite comparar Painel × Estatística × independente. A u(Rw) de 12 m estava anômala (81 %) |
| Bilirrubina direta | **ETp de fonte diferente (VB, desempenho MIN)** | Usa ETp = 39,915 % pela variação biológica, tem o maior bias do CEQ (30,8 %) e entra em "Margem crítica" |
| Hemoglobina glicada | **sem ETp utilizável no Sigma / sem CEQ** | Nenhum analito está sem ETp no ano de referência 2025 (ver T13). A HbA1c é o caso de "sem bias, sem Sigma": não tem CEQ ("SEM EP"), e no lote 8977 também não tem CIQ (n = 0) |
| Ureia | **CEQ em mais de uma rodada + ETp excedido** | 3 rodadas CAP 2025 (C-A, C-B, C-C), com bias por rodada bem diferente. ET de 18 % e 26 % contra ETp de 9 % |
| Glicose | **rotina estável** | n de 19 e 20 no mês, CV de 1,3 % e 1,4 %. Mostrou o efeito de um único resultado implausível no Sigma (OC-3 abaixo) |

Além da amostra, a calculadora recalculou **todas as 62 linhas analito × nível** da Estatística em 19 colunas.

## 3. Testes executados

| ID | Teste | Estado | Evidência |
|---|---|---|---|
| F3-T01 | Achar os dados brutos | APROVADO | tblCQ_Final (aba Principal - Resultados, 34.254 linhas); EQA_Base (545 linhas: CAP 2025, rodadas C-A/C-B/C-C e 90 linhas de simulação A/B/C com Uso = NAO); Especificações A4:AR43 e biblioteca 46–89; ano de referência do lote no bloco da aba Analitos |
| F3-T02 | Confirmar a definição no código e nos ADR | APROVADO | Seção 1 (mCEQ.BiasEQ duas etapas; mEstatPeriodo; mIncerteza.IncertezaCIQ; ADR-035/064/068/069/071/072) |
| F3-T03 | n, Média, DP, CV por analito × nível (lote 8977, DIM 2, out/2026) | APROVADO | 62/62 linhas em C, D, E, F; maior dif. relativa 1,5e-12 (Fósforo N2, DP: soma de quadrados × duas passagens) |
| F3-T04 | Bias do CEQ (abs, com sinal, n amostras, n rodadas) | APROVADO | 62/62 em G, T, AC, AD; maior dif. 9e-15. Bias recalculado de Resultado e Valor_Alvo: as 545 amostras batem com as colunas Bias/Bias_Abs da EQA_Base (0 divergências) |
| F3-T05 | ET % | APROVADO | 58/58 com valor e 4/4 vazios (sem CEQ) em H; maior dif. 6,7e-13 |
| F3-T06 | ETp e CVTp contra o cadastro (CLIA/VB) | APROVADO | 62/62 em K e J; VB recalculado da biblioteca 2025: Lactato 15,2097 (ÓTI), Bilirrubina direta 39,9151 (MIN), Lipase 18,8455 (MIN); CLIA = biblioteca 2025; CVTp CLIA = ETp/3 |
| F3-T07 | CV pooled de 6 m e Sigma | APROVADO | Janela igual (29/04–28/10/2026; MU_Fim = 28/10). L: 58/58 com valor e 4/4 vazios; maior dif. 4,2e-13 |
| F3-T08 | Margem, classe Sigma, DPM, rendimento | APROVADO | N, O, M, V, W: 62/62; maior dif. 1,9e-11 (margem de Potássio N1 = −0,157 p.p., cancelamento perto de zero; absoluto 3e-12) |
| F3-T09 | u(Rw) 12 m e viés do CEQ com sinal (incerteza) | APROVADO | AF 62/62 e AR 62/62; maior dif. 5,7e-13 |
| F3-T10 | Painel = Estatística = independente (Proteína total, N1 e N2) | APROVADO | Painel B7:I8 (n, média, DP, CV do motor; ETp; bias; ET; Sigma) e C46:C47 (CVp), J46:J47, M46:M47: todos dentro de 4,3e-13 do independente. Reconferido em 01/11 na cópia nova: C46 = 3,3257310225463 contra 3,3257310225449 (4,2e-13), I7 = 1,99073935777791 contra 1,99073935777875 |
| F3-T11 | Sinal e escala percentual do bias | APROVADO | bias = (X − alvo)/alvo × 100. Positivo = laboratório acima do grupo (Proteína total C-A/CHM-05: 9,0 vs 8,39 → +7,2706 %). O \|bias\| entra como magnitude no ET e no Sigma |
| F3-T12 | Provedor, ano e rodada certos | APROVADO | Só CAP 2025 entra. As 90 linhas de simulação (Uso = NAO, rodadas A/B/C) ficaram fora: AD = 3 rodadas (seriam 6). K5 diz "3 rodada(s), 435 linha(s)", que é igual a 29 analitos × 15 amostras. A separação CAP × Controllab **não pôde ser exercitada com dado real**: o Controllab não tem nenhuma linha. Pelo código, `Casa()` filtra o provedor |
| F3-T13 | ETp ausente aparece como "-" (ADR-068) | NAO_EXECUTADO | Os 31 analitos têm ETp positivo no ano de referência do lote 8977 (2025), então não há caso real. Para provocar o caso seria preciso apagar o ETp no Excel, que fica com a F1/F2 ou o próximo ciclo. Pela leitura da fórmula, K = "-" quando N(ETp) ≤ 0 e L a AA seguem K. O risco real está na OC-7 |
| F3-T14 | O valor muda ao incluir ou corrigir um CEQ/CIQ | NAO_EXECUTADO | Fica com a F1/F2 (instrução do Orquestrador) |
| F3-T15 | 2 lotes no cálculo (Proteína total) | APROVADO | Estatística C:F usa só o lote 8977 (n = 14 e 15 no mês). O CVp do Sigma junta os lotes 8974 (n = 44/47) e 8977 (n = 90/83) e deixa fora os lotes com n < 20 (8177, 8794 com n = 1), como manda o ADR-071 |
| F3-T16 | Resumo Six Sigma (Estatística B98:B103, F98:F100) | REPROVADO | Classes (0/0/3/2/53) e orçamento (46/2/10) batem com o independente. "Sem EP" = **22**, mas só há **4** níveis sem CEQ (HbA1c e PCR). Ver OC-5 |
| F3-T17 | Coluna AT (RMS do viés / u(Cref)) | REPROVADO | AT14 mostra "004 / 000" para RMS = 3,6944 %. Ver OC-4 |
| F3-T18 | ET % do Painel (H7:H8) quando não há CEQ | REPROVADO | Pela leitura da fórmula e pelos dados (sem Excel): `H7 = E7*1,65 + ABS(IF(ISNUMBER(G7),G7,0))` trata "SEM EP" como bias 0. Ver OC-2. Confirmação visual fica para a F4 |
| F3-T19 | Plausibilidade dos resultados que entram no CVp e na u(Rw) | REPROVADO | Há resultados de outro nível e valores ≤ 0 ativos na estatística, e eles derrubam o Sigma. Ver OC-3 |
| F3-T20 | Seleção das amostras do bias do ET/Sigma | APROVADO (segue a definição) | O BiasEQ inclui amostras com resultado 0 (−100 %), que o ViesEQ exclui. É diferença de método, encaminhada ao Especialista em CQ (OC-1) |
| F3-T21 | Bloco do lote na aba Analitos (AU:BD) | APROVADO | O recálculo de CV = DP/Média do cadastro, ET e Sigma do bloco bate (lotes 8974 e 8977, 7 analitos). A diferença de definição em relação à Estatística está na OC-6 |
| F3-T22 | Biblioteca anual coerente entre os anos | REPROVADO | Em 2026–2029 faltam o CVi/CVg da Lipase e o ETp CLIA da PCR, que existem em 2025. Ver OC-7 |
| F3-T23 | Segunda opinião independente da calculadora (reexecução) | APROVADO | Script novo, só `statistics`/`math`, lendo direto tblCQ_Final e EQA_Base: Proteína total, Bilirrubina direta, Ureia e Glicose, N1 e N2 (n, média, DP, CV, bias, ET, Sigma, com CVp de 6 m por lote). 56/56 valores iguais à Estatística, maior dif. relativa 5,8e-13. Reconferidos no código: `mCEQ.BiasEQ` (duas etapas, `Casa` com Uso ≠ NAO e provedor, `AnoVigente` com ano exato do ADR-072), `mIncerteza.AgregarMU`/`IncertezaCIQ` (equipamento, PARTICIPA = SIM via `CarregarDB`, n ≥ 20, média ≠ 0, CVp por gl), `MU_N_LOTE_MIN = 20` |
| F3-T24 | Segunda opinião do zero com **outros 6 analitos** (reexecução de 01/11) | APROVADO | Script novo, sem importar a calculadora, lendo direto tblCQ_Final, EQA_Base e a biblioteca 2025: Lactato (VB ÓTI), Lipase (VB MIN), Creatinina, Potássio, Bicarbonato e Colesterol total, N1 e N2. Recalculei n, média, DP, CV, bias, ET, ETp e Sigma (CVp de 6 m, lotes com n ≥ 20, sem lotes de n = 1 como 8177 e 6354). **96/96** valores iguais à Estatística; maior dif. relativa 8,9e-13 (DP do Potássio N2, soma de quadrados × duas passagens). Tabela 4.2b |
| F3-T25 | Reconferir as evidências das ocorrências na cópia nova (01/11) | APROVADO | Painel!H7 = `IF($E7="","",$E7*1.65+ABS(IF(ISNUMBER($G7),$G7,0)))` (OC-2); B103 = `COUNTIF($G$14:$G$93,"SEM EP")` = 22, com A76:A93 = 0 e G = "SEM EP" (OC-5); AT14 = "004 / 000" e AT20 = "024 / 001" (OC-4); G20 = 30,77093324845204 e AR20 = −20,12030759436774 (OC-1); biblioteca 2026–2029 sem Lipase VB e sem PCR CLIA (OC-7); BIO-319726, 215815, 215820, 268082, 232288 e 232289 ativos com PARTICIPA = SIM, e 253/253 suspeitos de 13/10 ativos (OC-3) |

## 4. Relatório B — validação matemática

### 4.1 Resultado global

**62 linhas × 19 colunas = 1.178 comparações. 0 divergências** na tolerância de 1e-9 relativo (REGRAS §3.5). A maior
diferença foi 1,9e-11, numa margem perto de zero (cancelamento ETp − ET, absoluto 3e-12). Esse nível vem da soma de
quadrados do VBA, comparada com o desvio em duas passagens do `statistics.stdev`, e não de método diferente.

| Coluna | O quê | OK / DIF | Maior dif. relativa |
|---|---|---|---|
| C, D, E, F | n, Média, DP, CV | 62/0 cada | 1,5e-12 |
| G, T, AC, AD | Bias abs, com sinal, amostras, rodadas | 62/0 cada | 9,0e-15 |
| H | ET % | 62/0 | 6,7e-13 |
| J, K | CVTp, ETp | 62/0 cada | 0 |
| L, M | Sigma, classe | 62/0 cada | 4,2e-13 |
| N, O | Margem p.p., % | 62/0 cada | 1,9e-11 |
| V, W | DPM, rendimento | 62/0 cada | 9,6e-13 |
| AF, AR | u(Rw) 12 m, viés com sinal | 62/0 cada | 5,7e-13 |

### 4.2 Tabela por analito da amostra (Estatística, lote 8977, DIMENSION 2)

| Analito | Equipamento | Lote | Nível | Período | Dados de entrada | Fórmula encontrada | Resultado independente | Resultado QC_INI | Diferença (rel.) | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| Glicose | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=19; CEQ 15 amostras/3 rodadas; CVp gl=127 | CV % (periodo, lote): DP/media*100; DP amostral n-1 | 1.299696056 | 1.299696056 | 5.5e-13 | APROVADO |
| Glicose | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=19; CEQ 15 amostras/3 rodadas; CVp gl=127 | Bias EQC abs %: media das medias por rodada de abs((X-alvo)/alvo*100) | 2.733294443 | 2.733294443 | 1.6e-16 | APROVADO |
| Glicose | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=19; CEQ 15 amostras/3 rodadas; CVp gl=127 | Bias com sinal %: 2 etapas, com sinal | -1.326437105 | -1.326437105 | 1.7e-16 | APROVADO |
| Glicose | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=19; CEQ 15 amostras/3 rodadas; CVp gl=127 | ET %: 1,65*CV + abs(bias) | 4.877792936 | 4.877792936 | 2.4e-13 | APROVADO |
| Glicose | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=19; CEQ 15 amostras/3 rodadas; CVp gl=127 | ETp %: fonte S: CLIA K / VB 1,65fCVi+f/2*raiz(CVi2+CVg2) | 10 | 10 | 0 | APROVADO |
| Glicose | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=19; CEQ 15 amostras/3 rodadas; CVp gl=127 | Sigma: (ETp - abs(bias))/CVp 6 m | 0.3297940308 | 0.3297940308 | 1.2e-15 | APROVADO |
| Glicose | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=19; CEQ 15 amostras/3 rodadas; CVp gl=127 | DPM: (1-Phi(S-1,5))*1e6 | 879040.9544 | 879040.9544 | 2.6e-16 | APROVADO |
| Glicose | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=20; CEQ 15 amostras/3 rodadas; CVp gl=124 | CV % (periodo, lote): DP/media*100; DP amostral n-1 | 1.446218239 | 1.446218239 | 2.1e-13 | APROVADO |
| Glicose | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=20; CEQ 15 amostras/3 rodadas; CVp gl=124 | Bias EQC abs %: media das medias por rodada de abs((X-alvo)/alvo*100) | 2.733294443 | 2.733294443 | 1.6e-16 | APROVADO |
| Glicose | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=20; CEQ 15 amostras/3 rodadas; CVp gl=124 | Bias com sinal %: 2 etapas, com sinal | -1.326437105 | -1.326437105 | 1.7e-16 | APROVADO |
| Glicose | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=20; CEQ 15 amostras/3 rodadas; CVp gl=124 | ET %: 1,65*CV + abs(bias) | 5.119554538 | 5.119554538 | 9.9e-14 | APROVADO |
| Glicose | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=20; CEQ 15 amostras/3 rodadas; CVp gl=124 | ETp %: fonte S: CLIA K / VB 1,65fCVi+f/2*raiz(CVi2+CVg2) | 10 | 10 | 0 | APROVADO |
| Glicose | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=20; CEQ 15 amostras/3 rodadas; CVp gl=124 | Sigma: (ETp - abs(bias))/CVp 6 m | 1.134183687 | 1.134183687 | 1.6e-14 | APROVADO |
| Glicose | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=20; CEQ 15 amostras/3 rodadas; CVp gl=124 | DPM: (1-Phi(S-1,5))*1e6 | 642748.9267 | 642748.9267 | 1.0e-14 | APROVADO |
| Proteína total | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=14; CEQ 15 amostras/3 rodadas; CVp gl=132 | CV % (periodo, lote): DP/media*100; DP amostral n-1 | 1.86014752 | 1.86014752 | 5.8e-13 | APROVADO |
| Proteína total | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=14; CEQ 15 amostras/3 rodadas; CVp gl=132 | Bias EQC abs %: media das medias por rodada de abs((X-alvo)/alvo*100) | 3.37933636 | 3.37933636 | 0 | APROVADO |
| Proteína total | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=14; CEQ 15 amostras/3 rodadas; CVp gl=132 | Bias com sinal %: 2 etapas, com sinal | 2.280078802 | 2.280078802 | 0 | APROVADO |
| Proteína total | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=14; CEQ 15 amostras/3 rodadas; CVp gl=132 | ET %: 1,65*CV + abs(bias) | 6.448579769 | 6.448579769 | 2.8e-13 | APROVADO |
| Proteína total | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=14; CEQ 15 amostras/3 rodadas; CVp gl=132 | ETp %: fonte S: CLIA K / VB 1,65fCVi+f/2*raiz(CVi2+CVg2) | 10 | 10 | 0 | APROVADO |
| Proteína total | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=14; CEQ 15 amostras/3 rodadas; CVp gl=132 | Sigma: (ETp - abs(bias))/CVp 6 m | 1.990739358 | 1.990739358 | 4.2e-13 | APROVADO |
| Proteína total | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=14; CEQ 15 amostras/3 rodadas; CVp gl=132 | DPM: (1-Phi(S-1,5))*1e6 | 311805.4029 | 311805.4029 | 9.6e-13 | APROVADO |
| Proteína total | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=15; CEQ 15 amostras/3 rodadas; CVp gl=128 | CV % (periodo, lote): DP/media*100; DP amostral n-1 | 2.20127836 | 2.20127836 | 5.3e-14 | APROVADO |
| Proteína total | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=15; CEQ 15 amostras/3 rodadas; CVp gl=128 | Bias EQC abs %: media das medias por rodada de abs((X-alvo)/alvo*100) | 3.37933636 | 3.37933636 | 0 | APROVADO |
| Proteína total | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=15; CEQ 15 amostras/3 rodadas; CVp gl=128 | Bias com sinal %: 2 etapas, com sinal | 2.280078802 | 2.280078802 | 0 | APROVADO |
| Proteína total | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=15; CEQ 15 amostras/3 rodadas; CVp gl=128 | ET %: 1,65*CV + abs(bias) | 7.011445654 | 7.011445654 | 2.8e-14 | APROVADO |
| Proteína total | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=15; CEQ 15 amostras/3 rodadas; CVp gl=128 | ETp %: fonte S: CLIA K / VB 1,65fCVi+f/2*raiz(CVi2+CVg2) | 10 | 10 | 0 | APROVADO |
| Proteína total | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=15; CEQ 15 amostras/3 rodadas; CVp gl=128 | Sigma: (ETp - abs(bias))/CVp 6 m | 1.117602291 | 1.117602291 | 1.4e-14 | APROVADO |
| Proteína total | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=15; CEQ 15 amostras/3 rodadas; CVp gl=128 | DPM: (1-Phi(S-1,5))*1e6 | 648916.8052 | 648916.8052 | 9.0e-15 | APROVADO |
| Bilirrubina direta | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=13; CEQ 15 amostras/3 rodadas; CVp gl=126 | CV % (periodo, lote): DP/media*100; DP amostral n-1 | 3.844300435 | 3.844300435 | 3.9e-14 | APROVADO |
| Bilirrubina direta | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=13; CEQ 15 amostras/3 rodadas; CVp gl=126 | Bias EQC abs %: media das medias por rodada de abs((X-alvo)/alvo*100) | 30.77093325 | 30.77093325 | 1.2e-16 | APROVADO |
| Bilirrubina direta | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=13; CEQ 15 amostras/3 rodadas; CVp gl=126 | Bias com sinal %: 2 etapas, com sinal | -30.77093325 | -30.77093325 | 1.2e-16 | APROVADO |
| Bilirrubina direta | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=13; CEQ 15 amostras/3 rodadas; CVp gl=126 | ET %: 1,65*CV + abs(bias) | 37.11402897 | 37.11402897 | 6.7e-15 | APROVADO |
| Bilirrubina direta | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=13; CEQ 15 amostras/3 rodadas; CVp gl=126 | ETp %: fonte S: CLIA K / VB 1,65fCVi+f/2*raiz(CVi2+CVg2) | 39.91509558 | 39.91509558 | 0 | APROVADO |
| Bilirrubina direta | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=13; CEQ 15 amostras/3 rodadas; CVp gl=126 | Sigma: (ETp - abs(bias))/CVp 6 m | 0.4348467994 | 0.4348467994 | 7.0e-15 | APROVADO |
| Bilirrubina direta | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=13; CEQ 15 amostras/3 rodadas; CVp gl=126 | DPM: (1-Phi(S-1,5))*1e6 | 856596.6961 | 856596.6961 | 6.8e-16 | APROVADO |
| Bilirrubina direta | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=14; CEQ 15 amostras/3 rodadas; CVp gl=125 | CV % (periodo, lote): DP/media*100; DP amostral n-1 | 1.745726395 | 1.745726395 | 2.5e-13 | APROVADO |
| Bilirrubina direta | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=14; CEQ 15 amostras/3 rodadas; CVp gl=125 | Bias EQC abs %: media das medias por rodada de abs((X-alvo)/alvo*100) | 30.77093325 | 30.77093325 | 1.2e-16 | APROVADO |
| Bilirrubina direta | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=14; CEQ 15 amostras/3 rodadas; CVp gl=125 | Bias com sinal %: 2 etapas, com sinal | -30.77093325 | -30.77093325 | 1.2e-16 | APROVADO |
| Bilirrubina direta | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=14; CEQ 15 amostras/3 rodadas; CVp gl=125 | ET %: 1,65*CV + abs(bias) | 33.6513818 | 33.6513818 | 2.2e-14 | APROVADO |
| Bilirrubina direta | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=14; CEQ 15 amostras/3 rodadas; CVp gl=125 | ETp %: fonte S: CLIA K / VB 1,65fCVi+f/2*raiz(CVi2+CVg2) | 39.91509558 | 39.91509558 | 0 | APROVADO |
| Bilirrubina direta | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=14; CEQ 15 amostras/3 rodadas; CVp gl=125 | Sigma: (ETp - abs(bias))/CVp 6 m | 1.238400231 | 1.238400231 | 6.3e-14 | APROVADO |
| Bilirrubina direta | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=14; CEQ 15 amostras/3 rodadas; CVp gl=125 | DPM: (1-Phi(S-1,5))*1e6 | 603184.9889 | 603184.9889 | 5.0e-14 | APROVADO |
| Ureia | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=18; CEQ 15 amostras/3 rodadas; CVp gl=141 | CV % (periodo, lote): DP/media*100; DP amostral n-1 | 7.56360155 | 7.56360155 | 1.2e-16 | APROVADO |
| Ureia | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=18; CEQ 15 amostras/3 rodadas; CVp gl=141 | Bias EQC abs %: media das medias por rodada de abs((X-alvo)/alvo*100) | 5.910450581 | 5.910450581 | 0 | APROVADO |
| Ureia | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=18; CEQ 15 amostras/3 rodadas; CVp gl=141 | Bias com sinal %: 2 etapas, com sinal | 1.368810052 | 1.368810052 | 0 | APROVADO |
| Ureia | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=18; CEQ 15 amostras/3 rodadas; CVp gl=141 | ET %: 1,65*CV + abs(bias) | 18.39039314 | 18.39039314 | 0 | APROVADO |
| Ureia | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=18; CEQ 15 amostras/3 rodadas; CVp gl=141 | ETp %: fonte S: CLIA K / VB 1,65fCVi+f/2*raiz(CVi2+CVg2) | 9 | 9 | 0 | APROVADO |
| Ureia | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=18; CEQ 15 amostras/3 rodadas; CVp gl=141 | Sigma: (ETp - abs(bias))/CVp 6 m | 0.1850999117 | 0.1850999117 | 1.3e-15 | APROVADO |
| Ureia | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=18; CEQ 15 amostras/3 rodadas; CVp gl=141 | DPM: (1-Phi(S-1,5))*1e6 | 905728.2585 | 905728.2585 | 0 | APROVADO |
| Ureia | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=21; CEQ 15 amostras/3 rodadas; CVp gl=140 | CV % (periodo, lote): DP/media*100; DP amostral n-1 | 12.46570043 | 12.46570043 | 1.4e-16 | APROVADO |
| Ureia | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=21; CEQ 15 amostras/3 rodadas; CVp gl=140 | Bias EQC abs %: media das medias por rodada de abs((X-alvo)/alvo*100) | 5.910450581 | 5.910450581 | 0 | APROVADO |
| Ureia | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=21; CEQ 15 amostras/3 rodadas; CVp gl=140 | Bias com sinal %: 2 etapas, com sinal | 1.368810052 | 1.368810052 | 0 | APROVADO |
| Ureia | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=21; CEQ 15 amostras/3 rodadas; CVp gl=140 | ET %: 1,65*CV + abs(bias) | 26.47885629 | 26.47885629 | 0 | APROVADO |
| Ureia | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=21; CEQ 15 amostras/3 rodadas; CVp gl=140 | ETp %: fonte S: CLIA K / VB 1,65fCVi+f/2*raiz(CVi2+CVg2) | 9 | 9 | 0 | APROVADO |
| Ureia | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=21; CEQ 15 amostras/3 rodadas; CVp gl=140 | Sigma: (ETp - abs(bias))/CVp 6 m | 0.4004936761 | 0.4004936761 | 3.5e-15 | APROVADO |
| Ureia | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=21; CEQ 15 amostras/3 rodadas; CVp gl=140 | DPM: (1-Phi(S-1,5))*1e6 | 864226.3614 | 864226.3614 | 4.0e-16 | APROVADO |
| Hemoglobina glicada | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=0; CEQ SEM EP amostras/SEM EP rodadas; CVp gl=138 | CV % (periodo, lote): DP/media*100; DP amostral n-1 | - | - | 0 | APROVADO |
| Hemoglobina glicada | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=0; CEQ SEM EP amostras/SEM EP rodadas; CVp gl=138 | Bias EQC abs %: media das medias por rodada de abs((X-alvo)/alvo*100) | SEM EP | SEM EP | igual | APROVADO |
| Hemoglobina glicada | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=0; CEQ SEM EP amostras/SEM EP rodadas; CVp gl=138 | Bias com sinal %: 2 etapas, com sinal | SEM EP | SEM EP | igual | APROVADO |
| Hemoglobina glicada | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=0; CEQ SEM EP amostras/SEM EP rodadas; CVp gl=138 | ET %: 1,65*CV + abs(bias) | - | - | 0 | APROVADO |
| Hemoglobina glicada | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=0; CEQ SEM EP amostras/SEM EP rodadas; CVp gl=138 | ETp %: fonte S: CLIA K / VB 1,65fCVi+f/2*raiz(CVi2+CVg2) | 8 | 8 | 0 | APROVADO |
| Hemoglobina glicada | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=0; CEQ SEM EP amostras/SEM EP rodadas; CVp gl=138 | Sigma: (ETp - abs(bias))/CVp 6 m | - | - | 0 | APROVADO |
| Hemoglobina glicada | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=0; CEQ SEM EP amostras/SEM EP rodadas; CVp gl=138 | DPM: (1-Phi(S-1,5))*1e6 | - | - | 0 | APROVADO |
| Hemoglobina glicada | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=0; CEQ SEM EP amostras/SEM EP rodadas; CVp gl=151 | CV % (periodo, lote): DP/media*100; DP amostral n-1 | - | - | 0 | APROVADO |
| Hemoglobina glicada | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=0; CEQ SEM EP amostras/SEM EP rodadas; CVp gl=151 | Bias EQC abs %: media das medias por rodada de abs((X-alvo)/alvo*100) | SEM EP | SEM EP | igual | APROVADO |
| Hemoglobina glicada | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=0; CEQ SEM EP amostras/SEM EP rodadas; CVp gl=151 | Bias com sinal %: 2 etapas, com sinal | SEM EP | SEM EP | igual | APROVADO |
| Hemoglobina glicada | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=0; CEQ SEM EP amostras/SEM EP rodadas; CVp gl=151 | ET %: 1,65*CV + abs(bias) | - | - | 0 | APROVADO |
| Hemoglobina glicada | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=0; CEQ SEM EP amostras/SEM EP rodadas; CVp gl=151 | ETp %: fonte S: CLIA K / VB 1,65fCVi+f/2*raiz(CVi2+CVg2) | 8 | 8 | 0 | APROVADO |
| Hemoglobina glicada | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=0; CEQ SEM EP amostras/SEM EP rodadas; CVp gl=151 | Sigma: (ETp - abs(bias))/CVp 6 m | - | - | 0 | APROVADO |
| Hemoglobina glicada | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 (CVp: 29/04-28/10/2026) | n=0; CEQ SEM EP amostras/SEM EP rodadas; CVp gl=151 | DPM: (1-Phi(S-1,5))*1e6 | - | - | 0 | APROVADO |

### 4.2b Segunda opinião de 01/11 (F3-T24): outros 6 analitos, script do zero

Estatística, lote 8977, DIMENSION 2, 01–31/10/2026. CEQ: CAP 2025, 3 rodadas, Uso ≠ NAO. CVp: 29/04–28/10/2026.
As colunas n, média, DP, CV, bias e ETp também foram comparadas (todas dentro de 1e-9; ver T24). Abaixo, o ET e o
Sigma, que juntam tudo.

| Analito | Equipamento | Lote | Nível | Período | Dados de entrada | Fórmula encontrada | Resultado independente | Resultado QC_INI | Diferença (rel.) | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| Lactato | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 | n=20; CV 2,689142; bias 2,700944 | ET = 1,65·CV + \|bias\| | 7.138027732 | 7.138027732 | 8,6e-14 | APROVADO |
| Lactato | DIMENSION 2 | 8977 | 1 | CVp 29/04-28/10 | ETp VB ÓTI 15,209694; CVp 7,314608 (8974 n=49, 8977 n=93) | Sigma = (ETp − \|bias\|)/CVp | 1.710105318 | 1.710105318 | 8,3e-15 | APROVADO |
| Lactato | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 | n=16; CV 7,608851 | ET | 15.25554779 | 15.25554779 | 1,0e-14 | APROVADO |
| Lactato | DIMENSION 2 | 8977 | 2 | CVp 29/04-28/10 | CVp 24,055539 (8177 n=1 fora) | Sigma | 0.5199945587 | 0.5199945587 | 1,5e-15 | APROVADO |
| Lipase | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 | n=13; CV 3,795409; bias 7,624035 | ET | 13.8864592 | 13.8864592 | 6,4e-15 | APROVADO |
| Lipase | DIMENSION 2 | 8977 | 1 | CVp 29/04-28/10 | ETp VB MIN 18,845471; CVp 4,533547 | Sigma | 2.475200344 | 2.475200344 | 1,4e-14 | APROVADO |
| Lipase | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 | n=14; CV 4,980323 | ET | 15.84156881 | 15.84156881 | 5,6e-15 | APROVADO |
| Lipase | DIMENSION 2 | 8977 | 2 | CVp 29/04-28/10 | CVp 4,020948 | Sigma | 2.790743801 | 2.790743801 | 1,6e-14 | APROVADO |
| Creatinina | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 | n=16; CV 3,588171; bias 2,929870 | ET | 8.850351965 | 8.850351965 | 6,8e-14 | APROVADO |
| Creatinina | DIMENSION 2 | 8977 | 1 | CVp 29/04-28/10 | ETp CLIA 15; CVp 12,887422 (6354 n=1 fora) | Sigma | 0.9365821873 | 0.9365821873 | 7,7e-15 | APROVADO |
| Creatinina | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 | n=17; CV 3,101723 | ET | 8.047712543 | 8.047712543 | 4,0e-14 | APROVADO |
| Creatinina | DIMENSION 2 | 8977 | 2 | CVp 29/04-28/10 | CVp 5,732755 (6354 n=5 e 8177 n=1 fora) | Sigma | 2.105467424 | 2.105467424 | 3,3e-14 | APROVADO |
| Potássio | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 | n=15; CV 2,167690; bias 2,579837 | ET | 6.156525243 | 6.156525243 | 4,8e-13 | APROVADO |
| Potássio | DIMENSION 2 | 8977 | 1 | CVp 29/04-28/10 | ETp CLIA 6; CVp 6,139874 | Sigma | 0.5570412735 | 0.5570412735 | 2,1e-14 | APROVADO |
| Potássio | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 | n=15; CV 1,632448 | ET | 5.273376205 | 5.273376205 | 4,6e-13 | APROVADO |
| Potássio | DIMENSION 2 | 8977 | 2 | CVp 29/04-28/10 | CVp 3,825415 | Sigma | 0.894063397 | 0.894063397 | 3,9e-13 | APROVADO |
| Bicarbonato | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 | n=14; CV 6,335948; bias 16,838844 | ET | 27.29315787 | 27.29315787 | 2,1e-14 | APROVADO |
| Bicarbonato | DIMENSION 2 | 8977 | 1 | CVp 29/04-28/10 | ETp CLIA 20; CVp 14,474702 | Sigma | 0.2183917513 | 0.2183917513 | 2,5e-16 | APROVADO |
| Bicarbonato | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 | n=14; CV 8,963256 | ET | 31.62821767 | 31.62821767 | 8,9e-15 | APROVADO |
| Bicarbonato | DIMENSION 2 | 8977 | 2 | CVp 29/04-28/10 | CVp 16,033242 | Sigma | 0.19716259 | 0.19716259 | 1,3e-14 | APROVADO |
| Colesterol total | DIMENSION 2 | 8977 | 1 | 01-31/10/2026 | n=14; CV 0,857163; bias 2,242593 | ET | 3.65691126 | 3.65691126 | 9,7e-14 | APROVADO |
| Colesterol total | DIMENSION 2 | 8977 | 1 | CVp 29/04-28/10 | ETp CLIA 10; CVp 5,657736 | Sigma | 1.371114979 | 1.371114979 | 1,9e-14 | APROVADO |
| Colesterol total | DIMENSION 2 | 8977 | 2 | 01-31/10/2026 | n=15; CV 1,475795 | ET | 4.677654901 | 4.677654901 | 1,8e-13 | APROVADO |
| Colesterol total | DIMENSION 2 | 8977 | 2 | CVp 29/04-28/10 | CVp 15,138955 | Sigma | 0.5124136291 | 0.5124136291 | 2,4e-15 | APROVADO |

O CVp de 6 m do Colesterol N2 (15,1 %), do Lactato N2 (24,1 %) e da Creatinina N1 (12,9 %) é muito maior que o CV do
mês (1,5 %, 7,6 % e 3,6 %). É o mesmo padrão da OC-3 (resultados implausíveis ativos na janela). A conta está certa; o
que entra nela é que é suspeito.

As linhas da HbA1c com "-" nos dois lados são vazio = vazio ("SEM EP" no bias, sem n no lote 8977). Como não há bias,
não há ET nem Sigma, e o QC_INI não inventa número (ADR-068/030). O ETp de 8 % (CLIA 2025) bate.

### 4.3 Etapas do cálculo (para outro especialista reproduzir)

**Proteína total, N1, lote 8977, DIMENSION 2, 01–31/10/2026**

1. **CIQ (14 resultados ativos):** BIO-359368 = 6,5; 360019 = 6,5; 360701 = 6,5; 361344 = 6,8; 362138 = 6,5;
   362550 = 6,4; 363221 = 6,3; 365293 = 6,6; 366575 = 6,6; 368239 = 6,7; 368754 = 6,5; 369273 = 6,6; 369967 = 6,5;
   370822 = 6,5.
   - Resultado: média 6,535714; DP (n−1) 0,1215739; CV 1,860148 %.
   - QC_INI: Estatística D44/E44/F44 e Painel C7/D7/E7, iguais.
2. **CEQ CAP 2025 (15 amostras, Uso = SIM):**
   - C-A: média de \|bias\| = 4,615032 (CHM-01..05: +2,6087; +2,6253; −6,4039; +4,1667; +7,2706).
   - C-B: 2,355033 (+0,7958; +2,4390; +2,7254; +1,6484; +4,1667).
   - C-C: 3,167944 (+5,2632; −1,8405; +1,8519; +0,7557; +6,1286).
   - Bias = (4,615032 + 2,355033 + 3,167944)/3 = **3,379336 %** (G44 = 3,3793363600341007).
   - Com sinal, pelas mesmas duas etapas: 2,280079 % (T44).
3. **ET** = 1,65 × 1,860148 + 3,379336 = **6,448580 %** (H44).
4. **ETp:** fonte CLIA, biblioteca 2025 = **10 %** (K44); CVTp = 10/3.
5. **CVp de 6 m** (29/04–28/10/2026, DIMENSION 2):
   - lote 8974: n = 44, CV 1,825 %;
   - lote 8977: n = 90, CV 3,846 %;
   - CVp = √[(43·1,825² + 89·3,846²)/132] = **3,325731 %** (Painel C46).
6. **Sigma** = (10 − 3,379336)/3,325731 = **1,990739** (L44 e Painel I7). Classe: Desempenho inadequado.
7. **Margem** = 10 − 6,448580 = 3,551420 p.p. = 35,51 % do ETp ("Dentro do orcamento").
8. **DPM** = [1 − Φ(0,490739)] × 10⁶ = 311.805,4 (V44); rendimento 68,8195 %.

**Bilirrubina direta (VB):**

- ETp = 1,65 × 0,75 × 20,9 + 0,375 × √(20,9² + 31,1²) = 25,8638 + 14,0513 = **39,9151 %**; CVTp = 0,75 × 20,9 = 15,675.
- Bias = média das rodadas C-A 44,8571 / C-B 21,6335 / C-C 25,8222 = **30,7709 %**. A C-A inclui CHM-04 (0,0 contra
  alvo 0,05 = −100 %) e a C-C inclui CHM-13 (0,0 contra 0,03 = −100 %).

**Ureia (3 rodadas):**

- bias 5,9105 %;
- ET N1 = 1,65 × 7,5636 + 5,9105 = 18,3904 %, contra ETp de 9 % → margem −104,3 % ("ETp excedido");
- CVp N1 = 16,69 %, puxado pelo lote 8977 (CV 20,9 %).

**Glicose (rotina):**

- CV do mês 1,30 %, ET 4,88 % (margem de 51 %);
- CVp N1 = 22,03 %, porque o lote 8977 tem CV de 27,2 % na janela;
- Sigma 0,33. Sem o resultado BIO-319726, o Sigma seria 4,3 (OC-3).

### 4.4 Caminho do dado (conferido ponto a ponto)

```
EQA.CAP_Dados (digitação) ─► EQA_Base (consolidação mEQA; Bias = (X−alvo)/alvo·100, Uso_Analitico)
   └─► mCEQ.BiasEQ "ABS" (provedor Especificações!AR, Ano N4, Rodada P4; duas etapas) ─► Estatística G
         ├─► ET %  H = 1,65·F + |G| ─────────────► Estatística N/O/P (orçamento) ─► Painel C63:K64
         └─► Sigma L = (K − |G|)/CVp ─► M classe, V DPM, W rendimento, X/Y/Z plano (Cfg_PlanoQC)
                                         └─► Painel I7:I8, J46:O47, plano C52:P53, O4 "Sigma do plano"
DB_SEAC/manual/inativação ─► tblCQ_Final (PARTICIPA_ESTATISTICA)
   ├─► mEstatPeriodo (lote H4, janela H5:J5, equipamento Painel L4) ─► Estatística C:F ─► H (ET)
   ├─► motor mEstatistica (Eng_Saida 185:186) ─► Painel B7:F8 ─► Painel H7:H8 (ET do Painel, ver OC-2)
   └─► mIncerteza.IncertezaCIQ (6 m, lotes n≥20, equipamento) ─► CVp ─► Estatística L e Painel C46:C47
Especificações (S fonte, biblioteca do ano do lote) ─► K (ETp), J (CVTp)
```

## 5. Ocorrências

### OC-1 — O bias do ET/Sigma inclui amostras do CEQ com resultado 0 (bias de −100 %) que o próprio projeto exclui da incerteza

- **Tipo:** duvida_metodo · **Prioridade:** Alto.
- **Tela/módulo:** Estatística G/H/L (mCEQ.BiasEQ) × AR/AJ (mCEQ.ViesEQ).
- **Ação:** recalcular o bias da Bilirrubina direta pelas duas regras.
- **Esperado:** a mesma amostra entra ou sai nos dois indicadores do CEQ, pelo mesmo critério. O ADR-064 (achado A-01)
  diz que, perto do limite de detecção, o bias relativo vira ±100 % e não descreve o método.
- **Observado:** CHM-04 (C-A) e CHM-13 (C-C) têm resultado do laboratório = 0,0 contra alvos de 0,05 e 0,03 (bias de
  −100 % cada).
  - O BiasEQ (ET/Sigma) as inclui: bias = **30,77 %**.
  - O ViesEQ (incerteza) as exclui: média com sinal = −20,12 % com n = 13.
  - Sem elas, o bias de duas etapas seria **19,99 %**: ET N1 iria de 37,1 % para 26,3 %, e o Sigma N1 de 0,43 para 0,95
    (CVp de 21,03 %).
  - Há ainda uma questão de resolução: o resultado é laudado com 0,1 mg/dL (0,2 contra 0,34), o que infla o % em todas
    as amostras baixas.
- **Dificuldade:** o usuário vê "Margem crítica" na Bilirrubina direta sem saber que duas amostras de −100 % pesam
  cerca de 11 p.p. no bias.
- **Consequência:** o bias, o ET e o Sigma de um analito ficam distorcidos por um critério que o projeto já reconheceu
  como inadequado em outra coluna.
- **Sugestão (ao Especialista em CQ, sem alterar nada):** decidir se o BiasEQ "ABS" deve aplicar o mesmo filtro do
  ViesEQ (resultado/alvo ≠ 0 e NAO AVALIADO fora), e se deve haver regra de resolução ou de limite inferior para o
  bias %.
- **Evidência:** EQA_Base linhas da Bilirrubina direta (chaves `CAP|2025|C-A 2025|…|CHM-04` e `…|C-C 2025|…|CHM-13`);
  Estatística G20 = 30,77093324845204, AR20 = −20,12030759436774, AU20 = "13 / 3"; `calc_independente.py`
  (`bias_ceq`, `vies_nordtest`).

### OC-2 — O ET % do Painel conta bias 0 quando o analito não tem CEQ

- **Tipo:** calculo · **Prioridade:** Alto.
- **Tela/módulo:** Painel, INDICADORES POR NÍVEL, H7:H8.
- **Ação:** ler a fórmula e simular com os dados de um analito sem CEQ (HbA1c, lote 8590, DIMENSION 2, out/2026).
- **Esperado:** sem bias não há ET (ADR-030/068, "ausência não é zero"), como faz a Estatística H ("") e o bloco
  ERRO TOTAL do próprio Painel (C63, que lê a Estatística).
- **Observado:** `H7 = IF($E7="","",$E7*1.65+ABS(IF(ISNUMBER($G7),$G7,0)))`. Com G7 = "SEM EP", o ET vira 1,65·CV.
  Na HbA1c, lote 8590, DIM 2, out/2026, os CVs são 1,5383 % e 1,6874 %, então o Painel mostraria ET de **2,54 %** e
  **2,78 %**, com o bias tratado como zero. O bloco ERRO TOTAL do mesmo Painel fica vazio. O mesmo acontece com a PCR
  (lote 8353), que dá 2,37 % e 2,01 %.
- **Dificuldade:** dois ETs diferentes na mesma tela, e o da linha principal parece ótimo.
- **Consequência:** estatística incorreta exibida (ET subestimado) para os analitos sem CEQ.
- **Sugestão:** o ET do Painel seguir a Estatística (vazio ou "SEM EP" quando G não for número). A decisão é do
  Orquestrador.
- **Evidência:** fórmula Painel!H7/H8 (leitura com openpyxl `data_only=False`); Estatística G40:G41 e G66:G67 =
  "SEM EP"; estatística do lote 8590 recalculada (`Calculadora.ciq_stats`). **Confirmação visual no Excel fica para a
  F4.**

### OC-3 — Resultados de outro nível e valores ≤ 0 continuam ativos e derrubam o Sigma e a u(Rw)

- **Tipo:** dados · **Prioridade:** Crítico. Classifico como associação errada entre níveis/material que tira a
  confiabilidade do Sigma (REGRAS §4).
- **Tela/módulo:** tblCQ_Final → CVp (Estatística L, Painel C46:C47 e I7:I8), u(Rw) (AF), plano de CQ.
- **Ação:** varrer o que entra no CVp com o critério do estudo de 13/10 (`estudos/cv_investigacao.py`: valor ≤ 0 ou
  mais de 5 DP robustos da mediana do lote).
- **Esperado:** o que entra na estatística é resultado do controle certo. O GUIA §6 pede que o laboratório revise
  `estudos/cv_suspeitos_bioquimica.csv` e inative os não conformes.
- **Observado:**
  - Os **253** suspeitos listados em 13/10 continuam **todos ativos** (PARTICIPA = SIM, ATIVO).
  - A varredura de hoje acha **287** na janela de 6 m (106 fora da lista, porque a janela andou) e **672** na de 12 m.
  - Há indícios claros de nível ou material trocado:
    - Corrida 25110602 (DIMENSION 1, 06/11/2025): N1 com valores de N2. Amilase BIO-215815 = 496 e Glicose
      BIO-215820 = 296, quando as medianas do N1 são 75 e 89.
    - Corrida 26032701 (DIMENSION 2, 27/03/2026): valores ≈ 0 ou negativos. Albumina −0,16; Amilase −2; Proteína total
      N2 BIO-268082 = −0,5 (lote 8974).
    - Glicose N1 BIO-319726 = 277 mg/dL (lote 8977, 16/07/2026).
    - Proteína total BIO-232288 = 101,5 (N1) e BIO-232289 = 65,3 (N2), em 05/01/2026.
  - Impacto medido:
    - **Glicose N1:** Sigma **0,33** com o resultado e **4,32** sem ele (CVp de 22,03 % para 1,68 %).
    - **Proteína total N1:** Sigma 1,99 → 3,40.
    - **Lactato N1:** Sigma 1,71 → 4,01.
    - **Proteína total:** u(Rw) de 12 m = **81,2 %** por causa de 2 resultados.
  - Resultado: 53 dos 58 níveis aparecem como "Desempenho inadequado", e 41 deles estão "Dentro do orçamento" no ET.
- **Dificuldade:** o usuário vê Sigma < 1 em analitos de rotina estável e não tem, na tela, nada que aponte para o
  resultado causador.
- **Consequência:** o plano de CQ (regras, N, run size) é recomendado sobre um Sigma artificial em quase todos os
  analitos, e a incerteza é superestimada.
- **Sugestão:**
  - **Laboratório:** revisar e inativar com motivo, pela aba Inativar, os IDs da lista e os novos (JSON da F3).
  - **Desenvolvimento (só com autorização):** um alerta no QA_INTEGRACAO para corrida cujo nível casa com o alvo de
    outro nível, ou para valor ≤ 0.
- **Evidência:** IDs acima; `cv_suspeitos_bioquimica.csv` comparado com o PARTICIPA atual (253/253 ativos);
  `Calculadora.cv_pooled` por lote; CVp da Glicose N1 por lote (8974 = 1,206 %, 8977 = 27,242 %).

### OC-4 — A coluna AT da Estatística mostra "004 / 000" em vez de "3,69 / 0,xx"

- **Tipo:** funcional · **Prioridade:** Médio.
- **Tela/módulo:** Estatística AT14:AT75 (RMS do viés / u(Cref), incerteza).
- **Ação:** comparar o texto exibido com o RMS recalculado.
- **Esperado:** "3,69 / 0,4x" ou similar.
- **Observado:** a fórmula usa `TEXT(…,"0.00")`. No Excel pt-BR o ponto é o separador de milhar, então o texto vira
  "004 / 000" (Lactato, RMS = 3,6944), "024 / 001" (Bilirrubina direta, RMS = 24,1859) etc. O próprio código
  (`mCEQ.ResumoFiltroEQ`, comentário) registra o "defeito pt-BR de AT14", que foi corrigido no K5 mas continua no AT.
- **Dificuldade:** o número fica ilegível e engana: 004 parece 4 mil ou 4.
- **Consequência:** a memória de cálculo da incerteza não pode ser conferida pela tela.
- **Sugestão:** montar o texto no VBA (CStr/Format), como já se faz no K5.
- **Evidência:** Estatística AT14 = '004 / 000', AT20 = '024 / 001'; RMS independente 3,6944 e 24,1859
  (`vies_nordtest`).

### OC-5 — O resumo "Sem EP (sem bias → sem Sigma)" conta 22 níveis, mas são 4

- **Tipo:** calculo · **Prioridade:** Médio.
- **Tela/módulo:** Estatística B103 (`=COUNTIF($G$14:$G$93,"SEM EP")`).
- **Ação:** recontar.
- **Esperado:** 4 (HbA1c e PCR, 2 níveis cada).
- **Observado:** **22**. As 18 linhas sem analito (A76:A93 = 0, catálogo vazio) chamam `BiasEQ("0")` e devolvem
  "SEM EP", que entra na contagem.
- **Dificuldade:** o resumo sugere que um terço da rotina está sem controle externo.
- **Consequência:** indicador de gestão errado.
- **Sugestão:** contar só as linhas com analito (por exemplo, COUNTIFS com A ≠ 0).
- **Evidência:** Estatística G76:G93 = "SEM EP" com A = 0; B103 = 22.

### OC-6 — Três Sigmas diferentes para o mesmo analito, e ET × Sigma com CVs diferentes

- **Tipo:** duvida_metodo · **Prioridade:** Médio.
- **Tela/módulo:** Estatística H/L, Analitos (bloco do lote AU:BD), Eng_Saida 190:269 (`engEstat`).
- **Ação:** comparar os Sigmas de Proteína total N1.
- **Esperado:** um Sigma oficial, ou rótulos que deixem claro por que são diferentes.
- **Observado:**
  - Estatística/Painel: **1,99** (CVp de 6 m).
  - Bloco do lote 8974: **4,00** (CV = DP/Média do cadastro).
  - Bloco do lote 8977: **2,50** (DP de TESTE = 1,5 × o do lote em uso).
  - Eng_Saida I185:I186 (indicadores do motor): um quarto Sigma, **5,03** (N1) e 3,83 (N2), calculado com o bias
    contra o alvo do lote. O Painel lê só as colunas B:F dessa faixa. Além disso, o bloco `engEstat` (Eng_Saida
    190:269), regravado a cada atualização pelo `AtualizarEstatisticaAba`, traz mais um Sigma por linha, e nenhuma
    fórmula lê esse bloco.
  - Na Estatística, o ET usa o CV do mês e o Sigma usa o CVp de 6 m. Com isso, **41 de 58** níveis estão "Dentro do
    orçamento" e, ao mesmo tempo, "Desempenho inadequado". O Painel mostra os dois lado a lado. A nota da Estatística
    Q11 explica, mas o usuário comum não lê.
- **Dificuldade:** o usuário não sabe qual Sigma vale para o plano.
- **Consequência:** risco de decisão contraditória (liberar pelo ET e intensificar pelo Sigma).
- **Sugestão:** levar ao Especialista em CQ (qual CV deve entrar no ET do orçamento) e ao UX (rotular "Sigma do
  cadastro" no bloco). Retirar ou documentar o bloco legado do Eng_Saida.
- **Evidência:** Estatística L44 = 1,9907; Analitos AX21 = 3,9991 (lote 8974) e AX885 = 2,4984 (lote 8977);
  Eng_Saida I185 = 5,0257 (G185 = −0,651, bias contra o alvo); `engEstat` sem referência em nenhuma aba (busca no XML das planilhas). Contagem de ET dentro com Sigma < 3 pela calculadora.

### OC-7 — Biblioteca de especificações 2026–2029 sem a VB da Lipase e sem o ETp CLIA da PCR

- **Tipo:** dados · **Prioridade:** Médio.
- **Tela/módulo:** Especificações, biblioteca anual (linhas 55–89).
- **Ação:** comparar os anos.
- **Esperado:** os anos copiados de 2025 trazerem as mesmas especificações, ou o laboratório decidir conscientemente
  deixá-los vazios.
- **Observado:** 2025 tem Lipase CVi = 7,6 / CVg = 24 e PCR CLIA = 30. Em 2026, 2027, 2028 e 2029 essas células estão
  vazias (os demais analitos foram copiados).
- **Dificuldade:** nenhuma hoje, porque todos os lotes estão em 2025. Quando um lote novo nascer com ano de referência
  2026 (S2), a Lipase e a PCR ficam sem ETp ("-") e sem Sigma, sem aviso.
- **Consequência:** perda silenciosa de Sigma/plano de dois analitos na troca de ano.
- **Sugestão:** o laboratório completar a biblioteca 2026+, ou o QA avisar quando o ano escolhido não tem a
  especificação da fonte do analito.
- **Evidência:** Especificações Y50/Y51 e AB48 (2025) preenchidas; Y59/Y60, AB57 e equivalentes de 2027–2029
  vazias; LotesStore ESPEC_ANO = 2025 em todos os lotes.

## 6. Dúvidas de regra de negócio (para o Especialista em CQ, via Orquestrador)

1. O bias do ET/Sigma (BiasEQ "ABS") deve excluir as amostras com resultado ou alvo = 0 e as NAO AVALIADO, como o
   ViesEQ? Deve haver um limite de resolução para o bias % em concentrações baixas? (OC-1)
2. O "orçamento de erro" (ET %) deve usar o CV do lote no período (como hoje) ou o mesmo CVp do Sigma? Hoje 41 níveis
   ficam "dentro do orçamento" e "inadequados" ao mesmo tempo. (OC-6)
3. Qual Sigma é o oficial para o plano de CQ? O do bloco do lote (CV do cadastro) é só informativo? (OC-6)
4. O CVp do Sigma junta todos os lotes do equipamento na janela e ignora o lote em análise (H3/H4). A janela termina no
   último resultado de qualquer equipamento (MU_Fim global). É a intenção? Se um equipamento parar, a janela dele
   continua ancorada no outro.
5. HbA1c e PCR não têm CEQ cadastrado. O laboratório participa de outro programa (Controllab?) que deveria ser
   digitado? Hoje esses analitos não têm Sigma nem plano.

## 7. Relatório A — impressões de uso (o que o usuário vê na F3)

- **Os números batem.** Em 1.178 comparações não houve nenhuma divergência de cálculo. A matemática do QC_INI segue
  exatamente a definição dos ADR, inclusive as duas etapas do bias, a regra do ano exato (ADR-072), a exclusão das
  simulações e o CV pooled só com lotes de n ≥ 20. O Painel repete a Estatística sem recalcular.
- **O que preocupa não é a conta, é o que entra nela.** Um único resultado de N2 gravado como N1 transforma uma glicose
  de Sigma 4,3 em 0,33. A tela não dá nenhuma pista do resultado causador. Para achar, precisei recalcular por lote.
  Como profissional de CQ, eu desconfiaria do sistema ("tudo inadequado?") antes de desconfiar dos dados.
- **Leitura confusa do Sigma.** Na mesma tela convivem ET "Dentro do orçamento" e Sigma "Desempenho inadequado"; o bloco
  do lote mostra outro Sigma. Falta dizer, perto do número, de onde vem cada CV.
- **A Estatística acompanha o lote do Painel.** Com o lote 8977, a HbA1c e a PCR aparecem com n = 0 (o lote não tem
  esses analitos), sem explicação na linha. O usuário pode achar que não houve corrida.
- **Rastreabilidade boa:** o K5 diz o que entrou (3 rodadas, 435 linhas), e os IDs e as chaves do CEQ permitem chegar
  à amostra. A coluna AT ilegível (OC-4) é a exceção.

## 8. Arquivos

- Calculadora: `_arquitetura/agentes/ferramentas/calc_independente.py`.
- JSON completo da comparação (62 linhas × 19 colunas, com o resultado independente): gerado no scratchpad da sessão
  pelo comando da introdução. Não é versionado: tem valores de resultado de controle, mas nenhum dado de paciente.
- Segunda opinião de 01/11 (F3-T24): script avulso, só `statistics`/`math`, sem importar a calculadora. Ficou no
  scratchpad da sessão; a saída ficou em `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\f3r3\segunda_opiniao.txt`, e o
  JSON da calculadora em `f3r3\bio_f3.json`. Nenhum dos dois é versionado.
- Cópia auditada: `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\Bioquimica_F3.xlsm` descartada ao fim da frente (REGRAS §2.3), e a produção não foi tocada. Para reproduzir, copie de novo a produção e confira se o sha256 curto ainda é `fe10dba1f1fd`.
