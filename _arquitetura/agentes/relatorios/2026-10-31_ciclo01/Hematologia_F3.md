# Hematologia — F3: auditoria matemática independente (ciclo 2026-10-31_ciclo01, completo)

**Agente:** Usuário Simulado · **Data:** 31/10/2026, reexecutada por inteiro em 01/11/2026 · **Excel:** não aberto
nesta frente (instrução do Orquestrador: outra frente usava o Excel).

**Terceira passada (01/11/2026, pedida pelo Orquestrador):** a produção continua com o sha256 curto `0b45fa270d2a`
(última gravação 30/10/2026 22:11) e `src/` não mudou desde o commit `0a25f5a` (`git status` limpo fora de
`_arquitetura/agentes/`). Fiz uma cópia nova (`shutil.copy2`, hash conferido igual) e repeti a frente inteira:

- a calculadora deu de novo 84 linhas × 27 colunas e 45 células do Painel, 0 divergências, e o JSON ficou
  **idêntico** ao da segunda passada (`hema_f3r3/hema_f3_r3.json` × `hema_f3_r2.json`);
- escrevi uma **terceira implementação, do zero**, para **outros 9 analitos** × 3 níveis (HGB, MCH, MCHC, NEUT%,
  MONO%, LYMPH#, EO#, MPV e RDW-CV). Ela recalcula também o ETp a partir da biblioteca anual e o DPM por `erfc`.
  Resultado: **459/459** valores iguais à Estatística (F3-T24, tabela 4.3b). Somando as passadas, 14 dos 28
  analitos foram conferidos por uma implementação que não importa a calculadora;
- reli no código as definições da seção 1 (`EstatBasica` com o filtro `Year(DATA)` entre B3 e D3, `BiasEQ` com
  `Casa`/`AnoVigente`, `IncertezaCIQ`) e reconferi, na cópia nova, a evidência de cada ocorrência (F3-T25);
- conferi três coisas que as passadas anteriores não tinham coberto: o contador de alertas da coluna U
  (F3-T26), o ET e o Sigma do **bloco do lote** na aba Analitos (AU:BD, 9 lotes, F3-T27) e a **plausibilidade do
  alvo cadastrado** (Média/DP) contra os dados do próprio lote (F3-T28).

As ocorrências OC-H1 a OC-H12 continuam válidas, com a mesma evidência. A terceira passada acrescenta a **OC-H13**
(alvo do lote 522611 incompatível com os dados em 12 pares, entre eles IG# com valores de IG% e RDW-SD N2 = 14,78 fL)
e a **OC-H14** (o bloco do lote mostra outro ET e outro Sigma para o mesmo lote, em 21 pares calculados com DP de
TESTE).

**Cópia auditada:** `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\Hematologia_F3.xlsm` (`shutil.copy2` da raiz).
**sha256 curto da produção copiada:** `0b45fa270d2a` (cópia idêntica, conferida). Arquivo de produção de 30/10/2026 22:11.
**Leitura:** openpyxl, `read_only=True`, `data_only=True`. Os valores do QC_INI são os gravados na última recalculação
do Excel (Value2, precisão total). Os dados brutos vêm das tabelas.
**Calculadora:** `_arquitetura/agentes/ferramentas/calc_independente.py`, estendida nesta frente com
`DadosHema`/`CalculadoraHema`/`auditar_hema` e detecção automática do produto (`--produto auto|bio|hema`). É Python puro
(`statistics`/`math`) e não lê nenhuma fórmula do QC_INI para calcular. As fórmulas foram lidas (`data_only=False`) só
para **confirmar a definição**. Para reproduzir:

```
set PYTHONIOENCODING=utf-8
python _arquitetura\agentes\ferramentas\calc_independente.py "%TEMP%\qcw\auditoria\2026-10-31_ciclo01\Hematologia_F3.xlsm" --json hema_f3.json
```

**Reexecução (31/10/2026, 20:48):** a frente foi refeita do zero numa cópia nova (`Hematologia_F3.xlsm`, sha256 curto
da produção `0b45fa270d2a`, o mesmo da primeira execução: a produção não mudou). Resultado idêntico: 84 × 27 e 45
células do Painel, 0 divergências (`hema_f3_r2.txt`/`.json`). Para não depender só da calculadora, uma **segunda
implementação, escrita do zero e sem importar a calculadora** (`_arquitetura/agentes/ferramentas/conferencia_cruzada_hema.py`),
recalculou n, CV, bias, ET, CVp e Sigma de WBC, RBC, HCT, MCV e PLT (15 níveis) direto das tabelas brutas: maior
diferença relativa contra a Estatística de **3,8e-12** (Sigma do RBC N2), dentro de 1e-9 (F3-T22). As definições da
seção 1 foram relidas no código nesta reexecução (`mCEQ.BiasEQ`/`Casa`/`AnoVigente`, `src/hema/mEstatistica.EstatBasica`)
e as evidências das OC-H1, H2, H4 e H5 foram conferidas de novo na cópia nova (fórmula Painel!H7:H9, AT14/AT23/AT41/AT44,
Especificações linhas 48–53, contagens por lote e ano).

**Regressão da calculadora na Bioquímica:** a mesma ferramenta rodou antes e depois da extensão sobre uma cópia nova
de `QC_Bioquimica.xlsm` (sha `fe10dba1f1fd`). A saída e o JSON ficaram **idênticos** (62 × 19, 0 divergências).

**Filtros em vigor na cópia (entradas do usuário):**

| Filtro | Valor |
|---|---|
| Painel | analito C3 = RBC; lote em análise H3 = 522611; Ano K3 = 2025; Período K4 = "Tudo (sem filtro)"; De/Até G3:G4 vazios |
| Estatística | Ano De B3 = 2025, Ano Até D3 = 2025; Lote B4 = 522611 (= loteAnalise); Provedor R4 = CAP; Ano EP N4 = 2025; Rodada P4 = TODAS; Sigma L11 = 6 meses (02/04/2026–01/10/2026, MU_Fim = 01/10/2026); u(Rw) 12 m (02/10/2025–01/10/2026) |
| Equipamento | Cfg_Integracao!B11 (selEquipamento = EQUIPAMENTO_PADRAO) = XN1000 |
| Anos de referência do lote 522611 | CLIA 2025 · VB 2025 (bloco do lote na aba Analitos; os 9 lotes estão em 2025/2025) |
| Configuração | "Deseja utilizar o lote" C20 = 522611, validade 31/03/2026 |

**Contexto da base:** tblCQ_Final com 44.392 linhas (43.710 com PARTICIPA = SIM; 589 DUPLICIDADE_ORIGEM e 93 SEM_VALOR
fora), todas do XN1000 e do interfaceamento, em 9 lotes de 01/10/2025 a 01/10/2026. Todas foram recebidas numa carga
só, em 02/10/2026 15:04, e a Sync_Log está vazia. A estatística reflete, portanto, uma base congelada em 01/10/2026
(a entrada de dados é assunto da F1). EQA_Base: 465 linhas, todas do CAP 2025. São 3 rodadas (FH9-A/B/C 2025), com
125 linhas cada. Há ainda 90 linhas de simulação (Uso_Analitico = NAO, rodadas 1/2/3) e 15 linhas de IPF sem analito
canônico.

## 1. Definição confirmada no código (antes de calcular)

| Grandeza | Onde | Definição que vale (conferida) |
|---|---|---|
| n, Média, DP, CV (Estatística C:F) | `src/hema/mEstatistica.bas` `AtualizarEstatisticaAba` → `EstatBasica`, publicada em Eng_Saida 190:309 (`engEstat`) e lida por `INDEX`; `mDados.CarregarDB` (COL_DATA = DATA_HORA) | Linhas com PARTICIPA = SIM, do equipamento `EquipFiltro()` (XN1000), do analito e nível, com LOTE = B4 e RESULTADO numérico. **ANO de DATA_HORA entre B3 e D3**: não é uma janela de dias e não há exclusões. Se Ano Até < Ano De, vale Ano De. DP amostral (n−1). n < 2 deixa o DP vazio; média 0 deixa o CV vazio |
| Bias EQC (G abs, T sinal, AC amostras, AD rodadas) | `src/comum/mCEQ.bas` `BiasEQ`/`Casa`/`AnoVigente`; ADR-035/071/072 | EQA_Base pelo Analito_Canonico, sem as linhas Uso = NAO. **Provedor único em R4** (a Bioquímica usa Especificações!AR). Com o ano explícito, TODAS usa exatamente aquele ano. O bias sai em duas etapas: média dos \|bias\| de cada rodada (ano\|rodada), depois a média das médias. **Não filtra Status_Padronizado** |
| ETp, CVTp, fonte (K, J, I) | Especificações (Hema): K/M/N/O/P (biblioteca do ano CLIA/VB do lote; célula vazia = 0), R (VB), S (fonte), T (ETp em uso), U (CVTp = `cvtpOficial`), V = K/3, W = fi(Q)·CVi (tabela X4:Y6); ADR-069/075 | CLIA → K; VB → 1,65·f·CVi + f/2·√(CVi²+CVg²), com f = 0,25/0,50/0,75 (ÓTI/DES/MIN); FAB → M. Estatística K = T se N(T) > 0, senão "-" (ADR-068) |
| ET % (H) | fórmula da Estatística | 1,65 · F + \|G\|. Fica vazio se F ou G não for número |
| CV do Sigma | `src/comum/mIncerteza.bas` `IncertezaCIQ "CV"` (`AgregarMU`); ADR-064/071 | Janela [EDATE(MU_Fim, −L11)+1, MU_Fim], com MU_Fim = INT(MAXIFS(DATA; PARTICIPA = SIM)). Mesmo equipamento; lotes com n ≥ 20 e média ≠ 0; CVp = √(Σ(nᵢ−1)CVᵢ² / Σ(nᵢ−1)). **Sem exclusões** (a fórmula L não passa o 6º argumento) |
| Sigma (L) | Estatística L; ADR-035/068/071 | "-" se K = "-"; senão (K − \|G\|) / CVp |
| Classe (M), margem (N, O, P), DPM (V), rendimento (W) | `mQualidade`, `Cfg_PlanoQC`, `mPlanoQC.DPMdoSigma` | Faixas ≥ 6 / 5 / 4 / 3 (Cfg_PlanoQC). N = K − H; O = N/K × 100. P: < 0 = "ETp excedido", ≤ 10 = "Margem critica", senão "Dentro do orcamento". DPM = (1 − Φ(σ − 1,5)) × 10⁶; rendimento = 100 − DPM/10⁴ |
| Status CV (Q) e SDI (R) | fórmula Q; `mCEQ.StatusSDIeq` | Q: "OK" se F ≤ J. R: pior \|SDI\| (coluna SDI do provedor) do ano vigente ≤ 2 |
| u(Rw) 12 m (AF..AI), viés (AR), amostras/rodadas (AU) | `mIncerteza.IncertezaCIQ` CV/GL/LOTES/DIAS/VALIDADE; `mCEQ.ViesEQ`; ADR-064/067 | AF vazio se a validade for INSUFICIENTE (< 90 dias ou gl < 30). O ViesEQ usa só amostras avaliadas, com resultado e alvo ≠ 0 e cada chave uma vez |
| Painel B7:F9 | `AtualizarPainelEng` (motor por corrida, filtro De/Até do Painel) | Mesmo lote e equipamento, valor por corrida e nível no filtro `filtroDe/filtroAte`. G, I, C45:O47 e C61:K63 **leem a Estatística**. H7:H9 é calculado no próprio Painel (ver OC-H2) |

## 2. Amostra de risco (6 analitos, todos os níveis)

| Analito | Risco que representa | Por quê |
|---|---|---|
| RBC | **2 lotes** (e analito do Painel) | O CVp de 6 m junta 4 lotes: 602911 (n = 20, exatamente o mínimo), 608511, 614011 e 619611. O 625311 fica fora (n = 1–2). C:F usa só o lote 522611. Permite comparar Painel × Estatística × independente |
| MCV | **ETp de fonte diferente (VB, MIN)** | ETp de 2,4096 % pela variação biológica (CVi 0,8; CVg 3,7). Está "ETp excedido" nos 3 níveis |
| IG% | **sem ETp** (fonte CLIA sem valor → "-") **e CEQ "NAO AVALIADO"** | Testa o ADR-068 com caso real. As 15 amostras do CAP são NAO AVALIADO: entram no bias do ET (G = 3,01 %) e saem do viés da incerteza (AR = SEM EP) |
| HCT | **CEQ em mais de uma rodada, com sinal oscilante** | FH9-A tem vieses negativos e FH9-B positivos: \|bias\| = 1,138 % contra bias com sinal de 0,324 %. Isso mostra que o \|bias\| não cancela o sinal |
| WBC | **rotina estável** | CLIA 7 %; CVp de 2,26 / 1,43 / 0,90 %; Sigma de 2,27 / 3,58 / 5,68 |
| RET% (extra) | **ETp sem CEQ** | ETp VB de 22,80 %, sem nenhuma linha de CEQ (SEM EP). Serve para a OC-H2 (ET do Painel) |

Além da amostra, a calculadora recalculou **todas as 84 linhas analito × nível** da Estatística, em 27 colunas, e
**45 células do Painel**. Na terceira passada, a segunda opinião cobriu outros 9 analitos escolhidos por risco: HGB
(o outlier da OC-H7 e alvo N2 a −14,6 DP), MCH/MCHC (VB, Sigma < 1,6; MCHC com Sigma 0,09), NEUT%/LYMPH# (Sigma
alto, DPM perto de zero, onde 1 − Φ perde dígitos), MONO% (outliers da OC-H7, CVp de 12 %), EO# (VB com CVg de 65 %),
MPV (VB, DP de TESTE no cadastro) e RDW-CV (CEQ sem ETp, ADR-068).

## 3. Testes executados

| ID | Teste | Estado | Evidência |
|---|---|---|---|
| F3-T01 | Achar os dados brutos | APROVADO | tblCQ_Final (aba Principal - Resultados, 44.392 linhas); EQA_Base A:U (465 linhas); Especificações A4:W43 + X4:Y6 + biblioteca 46–89 (só 2025 preenchida); ano de referência do lote no bloco da aba Analitos |
| F3-T02 | Confirmar a definição no código e nos ADR | APROVADO | Seção 1 (EstatBasica/AtualizarEstatisticaAba, mCEQ.BiasEQ, mIncerteza.IncertezaCIQ/AgregarMU, mPlanoQC.DPMdoSigma, ClassificarMargem; ADR-035/064/068/069/071/072/075) |
| F3-T03 | n, Média, DP, CV por analito × nível (lote 522611, XN1000, ano 2025) | APROVADO | 84/84 em C, D, E, F; maior dif. relativa 7,9e-16 |
| F3-T04 | Bias do CEQ (abs, com sinal, amostras, rodadas) | APROVADO | 84/84 em G, T, AC, AD; maior dif. 1,0e-15. Bias recalculado de Resultado e Valor_Alvo: as 465 linhas batem com a coluna Bias da EQA_Base (0 divergências) |
| F3-T05 | ET % | APROVADO | 84/84 (72 com valor; 12 vazios = RDW-SD, RET%, RET#, IRF sem CEQ); maior dif. 5,4e-16 |
| F3-T06 | ETp, CVTp e fonte contra o cadastro (CLIA/VB) | APROVADO | 84/84 em I, J, K. VB recalculado da biblioteca 2025, por exemplo MCV 2,409562 (MIN), MCHC 1,737874, NEUT% 27,582810, RDW-SD 3,665415. CLIA = biblioteca 2025 (WBC 7, RBC/HGB/HCT 4, PLT 15). CVTp CLIA = ETp/3; CVTp VB = fi·CVi |
| F3-T07 | CV pooled de 6 m e Sigma | APROVADO | Janela igual à gravada (N11:P11 = 02/04/2026–01/10/2026). L: 84/84 (57 com valor, 21 "-", 6 vazios sem CEQ); maior dif. 9,4e-12 |
| F3-T08 | Classe, margem, DPM, rendimento | APROVADO | M, N, O, P, W: 84/84 (maior dif. 6,7e-15 em N/O, 2,4e-12 em W). V: 84/84 com a tolerância **absoluta** de 1e-9 DPM em 3 linhas de Sigma > 6,8 (BASO% N2, LYMPH# N2, BASO# N2: diferença de 1,1e-10 DPM, cancelamento em 1 − Φ dos dois lados; justificativa na seção 4.1) |
| F3-T09 | Status CV (Q) e Status SDI (R) | APROVADO | 84/84 nas duas colunas (por exemplo, RBC "FORA (\|SDI\| max 3,40 > 2)") |
| F3-T10 | u(Rw) 12 m, gl, lotes, dias, viés com sinal, amostras/rodadas | APROVADO | AF, AG, AH, AI, AR, AU: 84/84; maior dif. 2,8e-12 (janela 02/10/2025–01/10/2026 igual à gravada) |
| F3-T11 | Painel = Estatística = independente (RBC, N1–N3) | APROVADO | B7:I9, C45:C47, J45:J47, M45:M47, O45:O47, C61:C63, H61:H63 e E74:E76 (11 / 4 / 42): **45/45**, maior dif. 2,4e-11 |
| F3-T12 | Sinal e escala percentual do bias | APROVADO | bias = (X − alvo)/alvo × 100. RBC FH9-01: 4,30 vs 4,257 → +1,0101 %; positivo = laboratório acima do grupo. No WBC, nenhuma amostra fica acima do grupo (13 negativas, 2 iguais): T = −1,8648 = −G. O \|bias\| entra como magnitude no ET e no Sigma |
| F3-T13 | Provedor, ano e rodada certos | APROVADO | Só CAP 2025 entra. As 90 linhas de simulação ficaram fora: WBC/RBC/HGB/HCT/MCV/MCH com AC = 15 e AD = 3 (seriam 30 e 6). K5 = "3 rodada(s), 360 linha(s)" = 24 analitos canônicos × 15. A separação CAP × Controllab **não pôde ser exercitada com dado real** (o Controllab tem 0 linhas). Pelo código, `Casa()` filtra pelo provedor de R4 |
| F3-T14 | ETp ausente aparece como "-" (ADR-068) | APROVADO | Caso real nesta pasta: IG%, IG#, NRBC%, NRBC#, RDW-CV, RET# e IRF (fonte CLIA sem ETp na biblioteca 2025), 21 linhas. Nelas K = "-" e J, L, M, N, O, P, V, W = "-", tudo conferido pelo independente |
| F3-T15 | O valor muda ao incluir ou corrigir um CEQ/CIQ | NAO_EXECUTADO | Fica com a F1/F2 (instrução do Orquestrador: sem Excel nesta frente) |
| F3-T16 | 2 lotes no cálculo | APROVADO | C:F usa só o lote 522611 (n = 50/51/51). O CVp junta 4 lotes com n ≥ 20 e deixa fora o 625311 (n = 1–2). Em RBC N1: 602911 (20; 1,0256 %), 608511 (57; 0,9596 %), 614011 (60; 0,8055 %), 619611 (54; 1,1344 %) → gl 187, CVp 0,975760 % (Painel C45 = 0,9757600160893) |
| F3-T17 | Coluna AT (RMS do viés / u(Cref)) | REPROVADO | AT14 = "002 / 000" para RMS 2,4915 %; AT41 = "005 / 000". Mesma raiz da OC-4 da Bioquímica. Ver OC-H1 |
| F3-T18 | ET % do Painel (H7:H9) quando não há CEQ | REPROVADO | Pela leitura da fórmula e pelos dados (sem Excel): RET% e RDW-SD têm ETp e não têm CEQ. Ver OC-H2. A confirmação visual fica com a F4 |
| F3-T19 | Plausibilidade do que entra no CVp e na u(Rw) | APROVADO | Nenhum resultado ativo ≤ 0 e nenhum RUN/nível/lote repetido entre os ativos. Dois casos isolados de 3–4 DP (observação OC-H7): efeito pequeno no Sigma e +20 % na u(Rw) da HGB N3 |
| F3-T20 | Seleção das amostras do bias do ET/Sigma (NAO AVALIADO) | APROVADO (segue a definição) | O BiasEQ inclui as 15 amostras NAO AVALIADO de IG%, IG#, NRBC% e NRBC#, que o ViesEQ exclui (AR = SEM EP). É a mesma diferença de método da OC-1 da Bioquímica. Aqui o impacto é só no ET (H), porque esses analitos não têm ETp. Ver OC-H6 |
| F3-T21 | Conferência cruzada catálogo × CEQ | REPROVADO | O CAP avalia o RDW como RDW-CV, mas o ETp (VB) está cadastrado só no RDW-SD. O resultado é que nenhum dos dois tem Sigma. Ver OC-H3 |
| F3-T22 | A calculadora conferida por uma 2ª implementação independente | APROVADO | `conferencia_cruzada_hema.py` (não importa a calculadora): WBC, RBC, HCT, MCV, PLT × N1–N3 = 15 níveis, colunas C, F, G, H, L; maior dif. relativa 3,8e-12. Armadilha registrada no cabeçalho do script: a EQA_Base tem uma tabela de/para em W:Y com outro cabeçalho "Provedor"; quem indexar a linha 1 inteira lê o provedor errado (a calculadora lê só A:U) |
| F3-T23 | Reprodutibilidade (reexecução numa cópia nova) | APROVADO | Mesma produção (sha `0b45fa270d2a`); `hema_f3_r2.txt` = 84/0 em todas as 27 colunas e Painel 45/0, igual à 1ª execução. Terceira passada (01/11): `hema_f3r3/hema_f3_r3.txt` = 84/0 e 45/0; JSON idêntico ao da 2ª |
| F3-T24 | Terceira implementação do zero, **outros 9 analitos** (01/11) | APROVADO | `segunda_opiniao_hema.py` (não importa nenhuma das outras): HGB, MCH, MCHC, NEUT%, MONO%, LYMPH#, EO#, MPV, RDW-CV × N1–N3 em C, D, E, F, G, T, AC, AD, H, I, J, K, L, N, O, V, AF. **459/459** iguais; maior dif. relativa 1,6e-10 (DPM da HGB N3: `erfc` exato × 1 − Φ do Excel). ETp recalculado da biblioteca 2025: HGB 4 (CLIA), MCH 2,630170, MCHC 1,737874, NEUT% 27,582810, MONO% 26,163428, LYMPH# 22,757983, EO# 43,555633, MPV 5,474090 (VB MIN); RDW-CV "-". Tabela 4.3b |
| F3-T25 | Reconferir as evidências das ocorrências na cópia nova (01/11) | APROVADO | Painel!H7:H9 = `IF($E7="","",$E7*1.65+ABS(IF(ISNUMBER($G7),$G7,0)))` (OC-H2); AT14 = "002 / 000", AT23 = "001 / 000", AT41 = "005 / 000", AT44 = AT47 = "008 / 000", fórmula com `TEXT(…,"0.00")` (OC-H1); G68 = 3,0093784575914833, AR68 = AU68 = "SEM EP" (OC-H6); G80 = G89 = "SEM EP" (OC-H2/H3); biblioteca 2026–2029 (linhas 55–89) sem nenhum valor (OC-H9); X53:AA53 = 48,1 / 15,1 / 9,4 / 1,2 (OC-H10); A98 = 0 com G/R/T = "SEM EP" (OC-H11); Principal!A5 = "Total: 44392,0 · participam: 43710,0" (OC-H12); Eng_Saida I185 = −0,1148, L190 = 3,1271 "Marginal" (OC-H8); Configuração C20 = 522611 (OC-H5) |
| F3-T26 | Contador de alertas de orçamento (Estatística U14:U133) | APROVADO | U = posição acumulada das linhas com P = "ETp excedido" ou "Margem critica". 120/120 iguais ao recálculo; total 15 (11 excedidos + 4 críticos), coerente com a distribuição de P recalculada (42 dentro, 21 "-", 6 vazios). A Hematologia não tem o resumo "Sem EP" que falhou na Bioquímica (OC-5 de lá) |
| F3-T27 | ET e Sigma do bloco do lote (aba Analitos AU:BD, 9 lotes) | APROVADO no cálculo · ver OC-H14 | Recálculo de AV/AY/BB = 100·DP/Média do cadastro, AW/AZ/BC = 1,65·CV + \|bias\| (AU = \|Estatística G\|) e AX/BA/BD = (ETp − \|bias\|)/CV: **2.268/2.268** iguais, maior dif. 3,8e-16. A definição, porém, difere da Estatística (CV do cadastro em vez do CV medido e do CVp), e 21 pares do lote 522611 têm DP de TESTE. Ver OC-H14 |
| F3-T28 | Plausibilidade do alvo cadastrado (Média/DP) contra os dados do próprio lote | REPROVADO | Lote 522611: 12 pares com a média medida a mais de 5 DP do alvo (IG# N1–N3 a −29 DP, alvo ≈ IG%; RDW-SD N2 a +398 DP, alvo 14,78 fL; RDW-CV N2 +25,7; HGB N2 −14,6; MCH N2 −12,4; MCV N2 −7,3; RDW-CV N3 −6,5; HCT N2 −5,7; RET# N1 +5,6; RET% N1 +5,2). Os outros 8 lotes: 0 pares. Ver OC-H13 |

## 4. Relatório B — validação matemática

### 4.1 Resultado global

**84 linhas × 27 colunas = 2.268 comparações na Estatística, mais 45 células do Painel. 0 divergências** na
tolerância de 1e-9 relativo (REGRAS §3.5). A exceção justificada é a coluna V (DPM), em 3 linhas, que usa tolerância
absoluta. Com Sigma acima de cerca de 6,5, Φ(σ − 1,5) fica a menos de 1e-11 de 1, e o "1 − Φ" perde os dígitos
relativos dos dois lados (o QC_INI com o Φ do Excel, a calculadora com `math.erf`). As diferenças de 1,1e-10 DPM
equivalem a cerca de 1e-16 em probabilidade, a precisão da máquina. O valor exato (`erfc`) do BASO% N2 é 3,388248e-06
DPM: o QC_INI dá 3,388290e-06 e a calculadora 3,388179e-06. O método é o mesmo nos dois lados, e não há ocorrência.

| Coluna | O quê | OK / DIF | Maior dif. relativa |
|---|---|---|---|
| C, D, E, F | n, Média, DP, CV | 84/0 cada | 7,9e-16 |
| G, T, AC, AD | Bias abs, com sinal, amostras, rodadas | 84/0 cada | 1,0e-15 |
| H | ET % | 84/0 | 5,4e-16 |
| I, J, K | Fonte, CVTp, ETp | 84/0 cada | 0 |
| L, M | Sigma, classe | 84/0 cada | 9,4e-12 |
| N, O, P | Margem p.p., %, situação | 84/0 cada | 6,7e-15 |
| Q, R | Status CV, status SDI | 84/0 cada | — |
| V, W | DPM (tol. abs. 1e-9), rendimento | 84/0 cada | 3,3e-05 (V, abs. 1,1e-10) · 2,4e-12 (W) |
| AF, AG, AH, AI | u(Rw) 12 m, gl, lotes, dias | 84/0 cada | 2,8e-12 |
| AR, AU | Viés com sinal, amostras/rodadas | 84/0 cada | 5,5e-16 |
| Painel (45 células) | B7:I9, C/J/M/O45:47, C/H61:63, E74:E76 | 45/0 | 2,4e-11 |

Distribuição (independente = QC_INI): 57 níveis com Sigma. São 16 de classe mundial, 4 excelentes, 8 bons, 7 marginais
e 22 com desempenho inadequado. Outros 21 ficam "-" (sem ETp) e 6 vazios (com ETp e sem CEQ: RDW-SD e RET%).
Orçamento: 11 com ETp excedido, 4 em margem crítica e 42 dentro do orçamento. Em 7 níveis, o ET está "Dentro do
orçamento" e o Sigma é < 3 (OC-6 da Bioquímica: ET e Sigma usam CVs diferentes).

### 4.2 Etapas do cálculo (exemplo reproduzível: RBC N1, lote 522611)

1. **CIQ do lote (C:F):** tblCQ_Final, ANALITO = RBC, NIVEL = 1, LOTE = 522611, EQUIPAMENTO = XN1000, PARTICIPA = SIM,
   ano(DATA_HORA) = 2025 → n = 50, média = 2,4022, DP (n−1) = 0,030256724669982, CV = 1,259542280825165 %.
2. **Bias (G):** EQA_Base CAP 2025, RBC (Red Blood Cell Count), Uso = SIM. As médias do \|bias\| por rodada são
   FH9-A = 0,805702, FH9-B = 1,481856 e FH9-C = 2,313302 (5 amostras cada), e a média das médias é **1,533620125276759 %**.
   Todas as amostras estão acima do grupo, então T = +1,5336.
3. **ET (H):** 1,65 × 1,259542280825165 + 1,533620125276759 = **3,611864888638281 %**.
4. **ETp (K):** fonte CLIA, biblioteca 2025 (ano CLIA do lote) → **4 %**; CVTp = 4/3 = 1,3333 %.
5. **CVp de 6 m:** janela 02/04/2026–01/10/2026. Os lotes com n ≥ 20 são 602911 (20; CV 1,0256 %), 608511 (57; 0,9596 %),
   614011 (60; 0,8055 %) e 619611 (54; 1,1344 %); o 625311 (n = 1) fica fora. Com Σ(n−1) = 187, o CVp é
   **0,975760016091096 %**.
6. **Sigma (L):** (4 − 1,533620125276759) / 0,975760016091096 = **2,52765007179079** → "Desempenho inadequado".
7. **Margem:** N = 4 − 3,611865 = 0,388135 p.p.; O = 9,7034 % → "Margem critica".
8. **DPM:** (1 − Φ(2,52765 − 1,5)) × 10⁶ = **152.057,23**; rendimento = 84,7943 %.
9. **Painel:** B7:E7 (motor por corrida, "Tudo", lote 522611) = os mesmos n/média/DP/CV. G7, I7, C45, J45, M45, O45,
   C61 e H61 batem com os passos 2–8.

### 4.3 Tabela por analito da amostra (Estatística, lote 522611, XN1000)

| Analito | Equipamento | Lote | Nível | Período | Dados de entrada | Fórmula encontrada | Resultado independente | Resultado QC_INI | Diferença (rel.) | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| RBC | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | CV % = DP/média×100 (DP n−1), lote 522611, ano 2025 | 1.259542281 | 1.259542281 | 3.5e-16 | APROVADO |
| RBC | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | Bias abs % = média das médias por rodada de abs((X−alvo)/alvo×100) | 1.533620125 | 1.533620125 | 0 | APROVADO |
| RBC | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | ET % = 1,65·CV(F) + abs(bias) | 3.611864889 | 3.611864889 | 2.5e-16 | APROVADO |
| RBC | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | ETp: S=CLIA → biblioteca CLIA do ano do lote; S=VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 4 | 4 | 0 | APROVADO |
| RBC | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | Sigma = (ETp − abs(bias))/CVp 6 m | 2.527650072 | 2.527650072 | 1.8e-12 | APROVADO |
| RBC | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | DPM = (1 − Φ(σ − 1,5))·10⁶ | 152057.2308 | 152057.2308 | 7.2e-12 | APROVADO |
| RBC | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | CV % = DP/média×100 (DP n−1), lote 522611, ano 2025 | 0.9550870524 | 0.9550870524 | 1.2e-16 | APROVADO |
| RBC | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | Bias abs % = média das médias por rodada de abs((X−alvo)/alvo×100) | 1.533620125 | 1.533620125 | 0 | APROVADO |
| RBC | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | ET % = 1,65·CV(F) + abs(bias) | 3.109513762 | 3.109513762 | 0 | APROVADO |
| RBC | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | ETp: S=CLIA → biblioteca CLIA do ano do lote; S=VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 4 | 4 | 0 | APROVADO |
| RBC | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | Sigma = (ETp − abs(bias))/CVp 6 m | 3.092197681 | 3.092197681 | 3.8e-12 | APROVADO |
| RBC | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | DPM = (1 − Φ(σ − 1,5))·10⁶ | 55670.14702 | 55670.14703 | 2.4e-11 | APROVADO |
| RBC | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | CV % = DP/média×100 (DP n−1), lote 522611, ano 2025 | 0.9204085696 | 0.9204085696 | 2.4e-16 | APROVADO |
| RBC | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | Bias abs % = média das médias por rodada de abs((X−alvo)/alvo×100) | 1.533620125 | 1.533620125 | 0 | APROVADO |
| RBC | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | ET % = 1,65·CV(F) + abs(bias) | 3.052294265 | 3.052294265 | 0 | APROVADO |
| RBC | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | ETp: S=CLIA → biblioteca CLIA do ano do lote; S=VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 4 | 4 | 0 | APROVADO |
| RBC | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | Sigma = (ETp − abs(bias))/CVp 6 m | 3.055344083 | 3.055344083 | 2.4e-12 | APROVADO |
| RBC | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | DPM = (1 − Φ(σ − 1,5))·10⁶ | 59932.07186 | 59932.07186 | 1.5e-11 | APROVADO |
| WBC | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | CV % = DP/média×100 (DP n−1), lote 522611, ano 2025 | 1.804141645 | 1.804141645 | 2.5e-16 | APROVADO |
| WBC | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | Bias abs % = média das médias por rodada de abs((X−alvo)/alvo×100) | 1.864814601 | 1.864814601 | 0 | APROVADO |
| WBC | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | ET % = 1,65·CV(F) + abs(bias) | 4.841648316 | 4.841648316 | 1.8e-16 | APROVADO |
| WBC | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | ETp: S=CLIA → biblioteca CLIA do ano do lote; S=VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 7 | 7 | 0 | APROVADO |
| WBC | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | Sigma = (ETp − abs(bias))/CVp 6 m | 2.272109227 | 2.272109227 | 6.3e-14 | APROVADO |
| WBC | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | DPM = (1 − Φ(σ − 1,5))·10⁶ | 220024.8689 | 220024.8689 | 1.9e-13 | APROVADO |
| WBC | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | CV % = DP/média×100 (DP n−1), lote 522611, ano 2025 | 1.503999378 | 1.503999378 | 4.4e-16 | APROVADO |
| WBC | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | Bias abs % = média das médias por rodada de abs((X−alvo)/alvo×100) | 1.864814601 | 1.864814601 | 0 | APROVADO |
| WBC | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | ET % = 1,65·CV(F) + abs(bias) | 4.346413575 | 4.346413575 | 2.0e-16 | APROVADO |
| WBC | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | ETp: S=CLIA → biblioteca CLIA do ano do lote; S=VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 7 | 7 | 0 | APROVADO |
| WBC | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | Sigma = (ETp − abs(bias))/CVp 6 m | 3.579684881 | 3.579684881 | 1.4e-12 | APROVADO |
| WBC | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | DPM = (1 − Φ(σ − 1,5))·10⁶ | 18777.22289 | 18777.22289 | 1.3e-11 | APROVADO |
| WBC | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | CV % = DP/média×100 (DP n−1), lote 522611, ano 2025 | 0.9063066804 | 0.9063066804 | 1.2e-16 | APROVADO |
| WBC | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | Bias abs % = média das médias por rodada de abs((X−alvo)/alvo×100) | 1.864814601 | 1.864814601 | 0 | APROVADO |
| WBC | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | ET % = 1,65·CV(F) + abs(bias) | 3.360220624 | 3.360220624 | 1.3e-16 | APROVADO |
| WBC | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | ETp: S=CLIA → biblioteca CLIA do ano do lote; S=VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 7 | 7 | 0 | APROVADO |
| WBC | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | Sigma = (ETp − abs(bias))/CVp 6 m | 5.68330323 | 5.68330323 | 2.4e-13 | APROVADO |
| WBC | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | DPM = (1 − Φ(σ − 1,5))·10⁶ | 14.36518841 | 14.36518841 | 0 | APROVADO |
| HCT | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | CV % = DP/média×100 (DP n−1), lote 522611, ano 2025 | 2.085314934 | 2.085314934 | 0 | APROVADO |
| HCT | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | Bias abs % = média das médias por rodada de abs((X−alvo)/alvo×100) | 1.138359618 | 1.138359618 | 2.0e-16 | APROVADO |
| HCT | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | ET % = 1,65·CV(F) + abs(bias) | 4.57912926 | 4.57912926 | 0 | APROVADO |
| HCT | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | ETp: S=CLIA → biblioteca CLIA do ano do lote; S=VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 4 | 4 | 0 | APROVADO |
| HCT | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | Sigma = (ETp − abs(bias))/CVp 6 m | 2.100217682 | 2.100217682 | 1.3e-12 | APROVADO |
| HCT | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | DPM = (1 − Φ(σ − 1,5))·10⁶ | 274180.5856 | 274180.5856 | 3.3e-12 | APROVADO |
| HCT | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | CV % = DP/média×100 (DP n−1), lote 522611, ano 2025 | 1.579803922 | 1.579803922 | 1.4e-16 | APROVADO |
| HCT | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | Bias abs % = média das médias por rodada de abs((X−alvo)/alvo×100) | 1.138359618 | 1.138359618 | 2.0e-16 | APROVADO |
| HCT | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | ET % = 1,65·CV(F) + abs(bias) | 3.745036088 | 3.745036088 | 0 | APROVADO |
| HCT | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | ETp: S=CLIA → biblioteca CLIA do ano do lote; S=VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 4 | 4 | 0 | APROVADO |
| HCT | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | Sigma = (ETp − abs(bias))/CVp 6 m | 2.407025386 | 2.407025386 | 2.0e-12 | APROVADO |
| HCT | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | DPM = (1 − Φ(σ − 1,5))·10⁶ | 182196.6865 | 182196.6865 | 7.0e-12 | APROVADO |
| HCT | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | CV % = DP/média×100 (DP n−1), lote 522611, ano 2025 | 1.548211586 | 1.548211586 | 0 | APROVADO |
| HCT | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | Bias abs % = média das médias por rodada de abs((X−alvo)/alvo×100) | 1.138359618 | 1.138359618 | 2.0e-16 | APROVADO |
| HCT | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | ET % = 1,65·CV(F) + abs(bias) | 3.692908734 | 3.692908734 | 1.2e-16 | APROVADO |
| HCT | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | ETp: S=CLIA → biblioteca CLIA do ano do lote; S=VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 4 | 4 | 0 | APROVADO |
| HCT | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | Sigma = (ETp − abs(bias))/CVp 6 m | 2.579560879 | 2.579560879 | 2.9e-13 | APROVADO |
| HCT | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | DPM = (1 − Φ(σ − 1,5))·10⁶ | 140168.885 | 140168.885 | 1.2e-12 | APROVADO |
| MCV | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | CV % = DP/média×100 (DP n−1), lote 522611, ano 2025 | 1.193358142 | 1.193358142 | 1.9e-16 | APROVADO |
| MCV | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | Bias abs % = média das médias por rodada de abs((X−alvo)/alvo×100) | 1.318118232 | 1.318118232 | 0 | APROVADO |
| MCV | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | ET % = 1,65·CV(F) + abs(bias) | 3.287159166 | 3.287159166 | 2.7e-16 | APROVADO |
| MCV | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | ETp: S=CLIA → biblioteca CLIA do ano do lote; S=VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 2.409561992 | 2.409561992 | 0 | APROVADO |
| MCV | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | Sigma = (ETp − abs(bias))/CVp 6 m | 1.060576843 | 1.060576843 | 8.0e-13 | APROVADO |
| MCV | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | DPM = (1 − Φ(σ − 1,5))·10⁶ | 669822.5247 | 669822.5247 | 4.6e-13 | APROVADO |
| MCV | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | CV % = DP/média×100 (DP n−1), lote 522611, ano 2025 | 0.9297982451 | 0.9297982451 | 1.2e-16 | APROVADO |
| MCV | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | Bias abs % = média das médias por rodada de abs((X−alvo)/alvo×100) | 1.318118232 | 1.318118232 | 0 | APROVADO |
| MCV | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | ET % = 1,65·CV(F) + abs(bias) | 2.852285336 | 2.852285336 | 1.6e-16 | APROVADO |
| MCV | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | ETp: S=CLIA → biblioteca CLIA do ano do lote; S=VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 2.409561992 | 2.409561992 | 0 | APROVADO |
| MCV | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | Sigma = (ETp − abs(bias))/CVp 6 m | 1.126890502 | 1.126890502 | 6.8e-13 | APROVADO |
| MCV | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | DPM = (1 − Φ(σ − 1,5))·10⁶ | 645466.5252 | 645466.5252 | 4.4e-13 | APROVADO |
| MCV | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | CV % = DP/média×100 (DP n−1), lote 522611, ano 2025 | 0.8842915959 | 0.8842915959 | 0 | APROVADO |
| MCV | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | Bias abs % = média das médias por rodada de abs((X−alvo)/alvo×100) | 1.318118232 | 1.318118232 | 0 | APROVADO |
| MCV | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | ET % = 1,65·CV(F) + abs(bias) | 2.777199365 | 2.777199365 | 0 | APROVADO |
| MCV | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | ETp: S=CLIA → biblioteca CLIA do ano do lote; S=VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 2.409561992 | 2.409561992 | 0 | APROVADO |
| MCV | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | Sigma = (ETp − abs(bias))/CVp 6 m | 1.282860619 | 1.282860619 | 5.1e-13 | APROVADO |
| MCV | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | DPM = (1 − Φ(σ − 1,5))·10⁶ | 585950.1378 | 585950.1378 | 4.3e-13 | APROVADO |
| IG% | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | CV % = DP/média×100 (DP n−1), lote 522611, ano 2025 | 3.113790198 | 3.113790198 | 1.4e-16 | APROVADO |
| IG% | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | Bias abs % = média das médias por rodada de abs((X−alvo)/alvo×100) | 3.009378458 | 3.009378458 | 0 | APROVADO |
| IG% | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | ET % = 1,65·CV(F) + abs(bias) | 8.147132284 | 8.147132284 | 0 | APROVADO |
| IG% | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | ETp: S=CLIA → biblioteca CLIA do ano do lote; S=VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | - | - | 0 | APROVADO |
| IG% | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | Sigma = (ETp − abs(bias))/CVp 6 m | - | - | 0 | APROVADO |
| IG% | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ 15 amostras/3 rod.; CVp gl=187 | DPM = (1 − Φ(σ − 1,5))·10⁶ | - | - | 0 | APROVADO |
| IG% | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | CV % = DP/média×100 (DP n−1), lote 522611, ano 2025 | 2.847808883 | 2.847808883 | 4.7e-16 | APROVADO |
| IG% | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | Bias abs % = média das médias por rodada de abs((X−alvo)/alvo×100) | 3.009378458 | 3.009378458 | 0 | APROVADO |
| IG% | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | ET % = 1,65·CV(F) + abs(bias) | 7.708263114 | 7.708263114 | 2.3e-16 | APROVADO |
| IG% | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | ETp: S=CLIA → biblioteca CLIA do ano do lote; S=VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | - | - | 0 | APROVADO |
| IG% | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | Sigma = (ETp − abs(bias))/CVp 6 m | - | - | 0 | APROVADO |
| IG% | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=194 | DPM = (1 − Φ(σ − 1,5))·10⁶ | - | - | 0 | APROVADO |
| IG% | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | CV % = DP/média×100 (DP n−1), lote 522611, ano 2025 | 2.744679338 | 2.744679338 | 0 | APROVADO |
| IG% | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | Bias abs % = média das médias por rodada de abs((X−alvo)/alvo×100) | 3.009378458 | 3.009378458 | 0 | APROVADO |
| IG% | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | ET % = 1,65·CV(F) + abs(bias) | 7.538099365 | 7.538099365 | 0 | APROVADO |
| IG% | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | ETp: S=CLIA → biblioteca CLIA do ano do lote; S=VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | - | - | 0 | APROVADO |
| IG% | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | Sigma = (ETp − abs(bias))/CVp 6 m | - | - | 0 | APROVADO |
| IG% | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ 15 amostras/3 rod.; CVp gl=171 | DPM = (1 − Φ(σ − 1,5))·10⁶ | - | - | 0 | APROVADO |
| RET% | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ SEM EP amostras/SEM EP rod.; CVp gl=187 | CV % = DP/média×100 (DP n−1), lote 522611, ano 2025 | 3.393624152 | 3.393624152 | 3.9e-16 | APROVADO |
| RET% | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ SEM EP amostras/SEM EP rod.; CVp gl=187 | Bias abs % = média das médias por rodada de abs((X−alvo)/alvo×100) | SEM EP | SEM EP | 0 | APROVADO |
| RET% | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ SEM EP amostras/SEM EP rod.; CVp gl=187 | ET % = 1,65·CV(F) + abs(bias) | (vazio) | (vazio) | 0 | APROVADO |
| RET% | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ SEM EP amostras/SEM EP rod.; CVp gl=187 | ETp: S=CLIA → biblioteca CLIA do ano do lote; S=VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 22.79762847 | 22.79762847 | 0 | APROVADO |
| RET% | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ SEM EP amostras/SEM EP rod.; CVp gl=187 | Sigma = (ETp − abs(bias))/CVp 6 m | (vazio) | (vazio) | 0 | APROVADO |
| RET% | XN1000 | 522611 | 1 | CIQ 2025 (CVp 02/04–01/10/2026) | n=50; CEQ SEM EP amostras/SEM EP rod.; CVp gl=187 | DPM = (1 − Φ(σ − 1,5))·10⁶ | (vazio) | (vazio) | 0 | APROVADO |
| RET% | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ SEM EP amostras/SEM EP rod.; CVp gl=194 | CV % = DP/média×100 (DP n−1), lote 522611, ano 2025 | 3.938084502 | 3.938084502 | 2.3e-16 | APROVADO |
| RET% | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ SEM EP amostras/SEM EP rod.; CVp gl=194 | Bias abs % = média das médias por rodada de abs((X−alvo)/alvo×100) | SEM EP | SEM EP | 0 | APROVADO |
| RET% | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ SEM EP amostras/SEM EP rod.; CVp gl=194 | ET % = 1,65·CV(F) + abs(bias) | (vazio) | (vazio) | 0 | APROVADO |
| RET% | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ SEM EP amostras/SEM EP rod.; CVp gl=194 | ETp: S=CLIA → biblioteca CLIA do ano do lote; S=VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 22.79762847 | 22.79762847 | 0 | APROVADO |
| RET% | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ SEM EP amostras/SEM EP rod.; CVp gl=194 | Sigma = (ETp − abs(bias))/CVp 6 m | (vazio) | (vazio) | 0 | APROVADO |
| RET% | XN1000 | 522611 | 2 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ SEM EP amostras/SEM EP rod.; CVp gl=194 | DPM = (1 − Φ(σ − 1,5))·10⁶ | (vazio) | (vazio) | 0 | APROVADO |
| RET% | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ SEM EP amostras/SEM EP rod.; CVp gl=171 | CV % = DP/média×100 (DP n−1), lote 522611, ano 2025 | 4.651902821 | 4.651902821 | 0 | APROVADO |
| RET% | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ SEM EP amostras/SEM EP rod.; CVp gl=171 | Bias abs % = média das médias por rodada de abs((X−alvo)/alvo×100) | SEM EP | SEM EP | 0 | APROVADO |
| RET% | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ SEM EP amostras/SEM EP rod.; CVp gl=171 | ET % = 1,65·CV(F) + abs(bias) | (vazio) | (vazio) | 0 | APROVADO |
| RET% | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ SEM EP amostras/SEM EP rod.; CVp gl=171 | ETp: S=CLIA → biblioteca CLIA do ano do lote; S=VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 22.79762847 | 22.79762847 | 0 | APROVADO |
| RET% | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ SEM EP amostras/SEM EP rod.; CVp gl=171 | Sigma = (ETp − abs(bias))/CVp 6 m | (vazio) | (vazio) | 0 | APROVADO |
| RET% | XN1000 | 522611 | 3 | CIQ 2025 (CVp 02/04–01/10/2026) | n=51; CEQ SEM EP amostras/SEM EP rod.; CVp gl=171 | DPM = (1 − Φ(σ − 1,5))·10⁶ | (vazio) | (vazio) | 0 | APROVADO |

### 4.3b Terceira implementação de 01/11 (F3-T24): outros 9 analitos, script do zero

Script: `_arquitetura/agentes/ferramentas/segunda_opiniao_hema.py`, sem importar a calculadora nem a conferência
cruzada. Lê só tblCQ_Final, EQA_Base (A:U), o cadastro de fonte/desempenho e a biblioteca anual da aba
Especificações. Saída completa (459 comparações, 17 colunas por nível) em `hema_f3r3/segunda_opiniao.txt`. Abaixo,
CV, ET e Sigma de todos os níveis, e o bias e o ETp (que são do analito) só no N1. Todas as linhas são do lote
522611, XN1000. O CVp de 6 m usa a janela de 02/04/2026 a 01/10/2026.

| Analito | Equipamento | Lote | Nível | Período | Dados de entrada | Fórmula encontrada | Resultado independente | Resultado QC_INI | Diferença (rel.) | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| HGB | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte CLIA | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 1.091388641 | 1.091388641 | 6.1e-16 | APROVADO |
| HGB | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte CLIA | Bias abs %: média das médias por rodada de abs((X−alvo)/alvo×100) | 1.443112971 | 1.443112971 | 0.0e+00 | APROVADO |
| HGB | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte CLIA | ET %: 1,65·CV + abs(bias) | 3.243904228 | 3.243904228 | 2.7e-16 | APROVADO |
| HGB | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte CLIA | ETp %: CLIA → biblioteca 2025; VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 4 | 4 | 0.0e+00 | APROVADO |
| HGB | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte CLIA | Sigma: (ETp − abs(bias))/CVp 6 m | 2.859826795 | 2.859826795 | 1.1e-12 | APROVADO |
| HGB | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte CLIA | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 0.6959305974 | 0.6959305974 | 1.6e-16 | APROVADO |
| HGB | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte CLIA | ET %: 1,65·CV + abs(bias) | 2.591398457 | 2.591398457 | 1.7e-16 | APROVADO |
| HGB | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte CLIA | Sigma: (ETp − abs(bias))/CVp 6 m | 3.787588569 | 3.787588569 | 7.0e-12 | APROVADO |
| HGB | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte CLIA | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 0.6648898448 | 0.6648898448 | 3.3e-16 | APROVADO |
| HGB | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte CLIA | ET %: 1,65·CV + abs(bias) | 2.540181215 | 2.540181215 | 0.0e+00 | APROVADO |
| HGB | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte CLIA | Sigma: (ETp − abs(bias))/CVp 6 m | 4.828934134 | 4.828934134 | 9.4e-12 | APROVADO |
| MCH | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 1.177272898 | 1.177272898 | 0.0e+00 | APROVADO |
| MCH | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | Bias abs %: média das médias por rodada de abs((X−alvo)/alvo×100) | 1.330224645 | 1.330224645 | 0.0e+00 | APROVADO |
| MCH | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | ET %: 1,65·CV + abs(bias) | 3.272724927 | 3.272724927 | 0.0e+00 | APROVADO |
| MCH | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | ETp %: CLIA → biblioteca 2025; VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 2.63016958 | 2.63016958 | 0.0e+00 | APROVADO |
| MCH | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 1.106172794 | 1.106172794 | 3.9e-13 | APROVADO |
| MCH | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 0.8256376245 | 0.8256376245 | 1.3e-16 | APROVADO |
| MCH | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | ET %: 1,65·CV + abs(bias) | 2.692526726 | 2.692526726 | 0.0e+00 | APROVADO |
| MCH | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 1.540179087 | 1.540179087 | 2.0e-12 | APROVADO |
| MCH | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 0.7479118004 | 0.7479118004 | 1.5e-16 | APROVADO |
| MCH | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | ET %: 1,65·CV + abs(bias) | 2.564279116 | 2.564279116 | 1.7e-16 | APROVADO |
| MCH | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 1.568145221 | 1.568145221 | 1.6e-13 | APROVADO |
| MCHC | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 1.840459381 | 1.840459381 | 1.2e-16 | APROVADO |
| MCHC | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | Bias abs %: média das médias por rodada de abs((X−alvo)/alvo×100) | 1.606529459 | 1.606529459 | 0.0e+00 | APROVADO |
| MCHC | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | ET %: 1,65·CV + abs(bias) | 4.643287438 | 4.643287438 | 0.0e+00 | APROVADO |
| MCHC | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | ETp %: CLIA → biblioteca 2025; VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 1.737874387 | 1.737874387 | 0.0e+00 | APROVADO |
| MCHC | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 0.08904652649 | 0.08904652649 | 1.5e-12 | APROVADO |
| MCHC | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 1.535112708 | 1.535112708 | 0.0e+00 | APROVADO |
| MCHC | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | ET %: 1,65·CV + abs(bias) | 4.139465427 | 4.139465427 | 0.0e+00 | APROVADO |
| MCHC | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 0.105486745 | 0.105486745 | 1.3e-13 | APROVADO |
| MCHC | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 1.27007716 | 1.27007716 | 3.5e-16 | APROVADO |
| MCHC | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | ET %: 1,65·CV + abs(bias) | 3.702156772 | 3.702156772 | 0.0e+00 | APROVADO |
| MCHC | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 0.1111573371 | 0.1111573371 | 1.1e-12 | APROVADO |
| NEUT% | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 2.473972925 | 2.473972925 | 3.6e-16 | APROVADO |
| NEUT% | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | Bias abs %: média das médias por rodada de abs((X−alvo)/alvo×100) | 2.46517352 | 2.46517352 | 0.0e+00 | APROVADO |
| NEUT% | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | ET %: 1,65·CV + abs(bias) | 6.547228845 | 6.547228845 | 2.7e-16 | APROVADO |
| NEUT% | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | ETp %: CLIA → biblioteca 2025; VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 27.58280952 | 27.58280952 | 1.3e-16 | APROVADO |
| NEUT% | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 9.542097069 | 9.542097069 | 2.6e-13 | APROVADO |
| NEUT% | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 1.896084704 | 1.896084704 | 1.2e-16 | APROVADO |
| NEUT% | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | ET %: 1,65·CV + abs(bias) | 5.593713281 | 5.593713281 | 1.6e-16 | APROVADO |
| NEUT% | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 12.97402096 | 12.97402096 | 4.4e-13 | APROVADO |
| NEUT% | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 1.780205607 | 1.780205607 | 0.0e+00 | APROVADO |
| NEUT% | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | ET %: 1,65·CV + abs(bias) | 5.402512771 | 5.402512771 | 0.0e+00 | APROVADO |
| NEUT% | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 14.89658478 | 14.89658478 | 2.1e-13 | APROVADO |
| MONO% | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 9.174015161 | 9.174015161 | 0.0e+00 | APROVADO |
| MONO% | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | Bias abs %: média das médias por rodada de abs((X−alvo)/alvo×100) | 6.491155467 | 6.491155467 | 0.0e+00 | APROVADO |
| MONO% | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | ET %: 1,65·CV + abs(bias) | 21.62828048 | 21.62828048 | 0.0e+00 | APROVADO |
| MONO% | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | ETp %: CLIA → biblioteca 2025; VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 26.16342832 | 26.16342832 | 0.0e+00 | APROVADO |
| MONO% | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 1.615771685 | 1.615771685 | 1.7e-14 | APROVADO |
| MONO% | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 10.1914208 | 10.1914208 | 1.7e-16 | APROVADO |
| MONO% | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | ET %: 1,65·CV + abs(bias) | 23.30699979 | 23.30699979 | 1.5e-16 | APROVADO |
| MONO% | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 3.066149594 | 3.066149594 | 1.9e-15 | APROVADO |
| MONO% | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 7.324166329 | 7.324166329 | 0.0e+00 | APROVADO |
| MONO% | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | ET %: 1,65·CV + abs(bias) | 18.57602991 | 18.57602991 | 0.0e+00 | APROVADO |
| MONO% | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 4.81733868 | 4.81733868 | 3.7e-15 | APROVADO |
| LYMPH# | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 4.366816642 | 4.366816642 | 2.0e-16 | APROVADO |
| LYMPH# | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | Bias abs %: média das médias por rodada de abs((X−alvo)/alvo×100) | 3.270526679 | 3.270526679 | 0.0e+00 | APROVADO |
| LYMPH# | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | ET %: 1,65·CV + abs(bias) | 10.47577414 | 10.47577414 | 1.7e-16 | APROVADO |
| LYMPH# | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | ETp %: CLIA → biblioteca 2025; VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 22.75798275 | 22.75798275 | 1.6e-16 | APROVADO |
| LYMPH# | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 3.906499102 | 3.906499102 | 1.0e-14 | APROVADO |
| LYMPH# | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 4.236274252 | 4.236274252 | 6.3e-16 | APROVADO |
| LYMPH# | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | ET %: 1,65·CV + abs(bias) | 10.2603792 | 10.2603792 | 3.5e-16 | APROVADO |
| LYMPH# | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 6.805828507 | 6.805828507 | 3.6e-13 | APROVADO |
| LYMPH# | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 3.416892622 | 3.416892622 | 3.9e-16 | APROVADO |
| LYMPH# | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | ET %: 1,65·CV + abs(bias) | 8.908399506 | 8.908399506 | 2.0e-16 | APROVADO |
| LYMPH# | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 9.97255507 | 9.97255507 | 2.5e-13 | APROVADO |
| EO# | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 6.392262708 | 6.392262708 | 2.8e-16 | APROVADO |
| EO# | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | Bias abs %: média das médias por rodada de abs((X−alvo)/alvo×100) | 8.181630966 | 8.181630966 | 0.0e+00 | APROVADO |
| EO# | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | ET %: 1,65·CV + abs(bias) | 18.72886443 | 18.72886443 | 1.9e-16 | APROVADO |
| EO# | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | ETp %: CLIA → biblioteca 2025; VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 43.55563302 | 43.55563302 | 0.0e+00 | APROVADO |
| EO# | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 4.995538163 | 4.995538163 | 3.1e-14 | APROVADO |
| EO# | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 7.720001536 | 7.720001536 | 2.3e-16 | APROVADO |
| EO# | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | ET %: 1,65·CV + abs(bias) | 20.9196335 | 20.9196335 | 1.7e-16 | APROVADO |
| EO# | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 4.81577567 | 4.81577567 | 1.8e-14 | APROVADO |
| EO# | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 6.956651447 | 6.956651447 | 3.8e-16 | APROVADO |
| EO# | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | ET %: 1,65·CV + abs(bias) | 19.66010585 | 19.66010585 | 1.8e-16 | APROVADO |
| EO# | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 5.724466535 | 5.724466535 | 3.1e-14 | APROVADO |
| MPV | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 3.513697672 | 3.513697672 | 0.0e+00 | APROVADO |
| MPV | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | Bias abs %: média das médias por rodada de abs((X−alvo)/alvo×100) | 2.883454665 | 2.883454665 | 0.0e+00 | APROVADO |
| MPV | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | ET %: 1,65·CV + abs(bias) | 8.681055823 | 8.681055823 | 0.0e+00 | APROVADO |
| MPV | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | ETp %: CLIA → biblioteca 2025; VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | 5.474090449 | 5.474090449 | 0.0e+00 | APROVADO |
| MPV | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 0.848511571 | 0.848511571 | 1.0e-13 | APROVADO |
| MPV | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 1.779384062 | 1.779384062 | 3.7e-16 | APROVADO |
| MPV | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | ET %: 1,65·CV + abs(bias) | 5.819438366 | 5.819438366 | 3.1e-16 | APROVADO |
| MPV | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 1.349972605 | 1.349972605 | 3.0e-15 | APROVADO |
| MPV | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 0.9885983803 | 0.9885983803 | 3.4e-16 | APROVADO |
| MPV | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | ET %: 1,65·CV + abs(bias) | 4.514641992 | 4.514641992 | 2.0e-16 | APROVADO |
| MPV | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte VB | Sigma: (ETp − abs(bias))/CVp 6 m | 1.717867813 | 1.717867813 | 6.0e-13 | APROVADO |
| RDW-CV | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte CLIA | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 1.275009657 | 1.275009657 | 5.2e-16 | APROVADO |
| RDW-CV | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte CLIA | Bias abs %: média das médias por rodada de abs((X−alvo)/alvo×100) | 1.023082561 | 1.023082561 | 0.0e+00 | APROVADO |
| RDW-CV | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte CLIA | ET %: 1,65·CV + abs(bias) | 3.126848494 | 3.126848494 | 2.8e-16 | APROVADO |
| RDW-CV | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte CLIA | ETp %: CLIA → biblioteca 2025; VB → 1,65·f·CVi + f/2·√(CVi²+CVg²) | - | - | 0.0e+00 | APROVADO |
| RDW-CV | XN1000 | 522611 | 1 | CIQ 2025; CVp 02/04–01/10/2026 | n=50; CEQ CAP 2025 15 am./3 rod.; CVp gl=187; fonte CLIA | Sigma: (ETp − abs(bias))/CVp 6 m | - | - | 0.0e+00 | APROVADO |
| RDW-CV | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte CLIA | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 0.7209171929 | 0.7209171929 | 6.2e-16 | APROVADO |
| RDW-CV | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte CLIA | ET %: 1,65·CV + abs(bias) | 2.212595929 | 2.212595929 | 4.0e-16 | APROVADO |
| RDW-CV | XN1000 | 522611 | 2 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=194; fonte CLIA | Sigma: (ETp − abs(bias))/CVp 6 m | - | - | 0.0e+00 | APROVADO |
| RDW-CV | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte CLIA | CV %: DP(n−1)/média×100, lote 522611, ano 2025 | 0.6134810481 | 0.6134810481 | 7.2e-16 | APROVADO |
| RDW-CV | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte CLIA | ET %: 1,65·CV + abs(bias) | 2.03532629 | 2.03532629 | 2.2e-16 | APROVADO |
| RDW-CV | XN1000 | 522611 | 3 | CIQ 2025; CVp 02/04–01/10/2026 | n=51; CEQ CAP 2025 15 am./3 rod.; CVp gl=171; fonte CLIA | Sigma: (ETp − abs(bias))/CVp 6 m | - | - | 0.0e+00 | APROVADO |

Etapas, para reproduzir um caso (MONO% N3): n = 51 (lote 522611, ano 2025); CV = 7,324166 %; bias de duas etapas do
CAP 2025 = 6,491155 %; ET = 1,65 × 7,324166 + 6,491155 = 18,576030 %; ETp VB MIN = 1,65 × 0,75 × 13,3 + 0,375 ×
√(13,3² + 22,2²) = 26,163428 %; CVp de 6 m (lotes com n ≥ 20, gl 171) = 4,083639 %; Sigma = (26,163428 − 6,491155) /
4,083639 = **4,817339** (Estatística L46 = 4,817338679814015); DPM = 0,5·erfc((4,817339 − 1,5)/√2)·10⁶ = 454,397
(V46 = 454,39690017212706).

### 4.4 Caminho do dado (Hematologia)

```
CEQ (aba EQA.CAP_Dados) -> EQA_Base (Analito_Canonico, Uso_Analitico, Provedor, Ano, Rodada, Resultado, Valor_Alvo, Bias)
   -> mCEQ.BiasEQ(analito, N4, "ABS", R4, P4)  [2 etapas por rodada]  -> Estatística G (|bias|) / T (sinal) / AC / AD
   -> mCEQ.ViesEQ (só avaliadas)                                     -> AR, AJ (u(bias)), AT, AU  (incerteza)

CIQ (DB_SEAC -> tblCQ_Final, PARTICIPA_ESTATISTICA)
   -> mEstatistica.EstatBasica(lote B4, ano B3..D3, XN1000) -> Eng_Saida 190:309 -> Estatística C:F (n, média, DP, CV)
   -> mIncerteza.IncertezaCIQ("CV", Sigma_Ini..MU_Fim)      -> CVp 6 m (lotes n>=20)  -> Estatística L, Painel C45:C47
   -> mIncerteza.IncertezaCIQ("CV", 12 m)                   -> u(Rw) AF

Especificações (S fonte; biblioteca 2025 pelo ano do lote) -> T (ETp) / U (CVTp) -> Estatística K / J

ET  (H) = 1,65·F + |G|                    -> margem N/O/P -> Painel C61:K63, E74:E76
Sigma (L) = (K − |G|) / CVp 6 m           -> M (classe), V (DPM), W, X:AA (plano) -> Painel I7:I9, O4 (Sigma do plano), C45:O54
Painel H7:H9 = 1,65·E(Painel) + |G| (SEM EP vira 0)   <- OC-H2

Analitos, bloco do lote (AU:BD): CV = DP/Média do CADASTRO; ET = 1,65·CV + |G|; Sigma = (ETp − |G|)/CV   <- OC-H14
Cadastro Média/DP do lote (Analitos E:J -> LotesStore) -> limites do LJ e Westgard (ADR-078/083)        <- OC-H13
```

## 5. Ocorrências

### OC-H1 — A coluna AT da Estatística mostra "002 / 000" em vez de "2,49 / 0,xx"

- **Tipo:** funcional · **Prioridade:** Médio. Mesma raiz da OC-4 da Bioquímica, agora confirmada na Hematologia.
- **Tela/módulo:** Estatística AT14:AT97 (RMS do viés / u(Cref), incerteza).
- **Ação:** comparar o texto exibido com o RMS recalculado.
- **Esperado:** "2,49 / 0,xx" para o WBC.
- **Observado:** `TEXT(…,"0.00")` no Excel pt-BR (o ponto é o separador de milhar) produz "002 / 000" (WBC, RMS =
  2,4915), "001 / 000" (HCT, RMS = 1,2958), "005 / 000" (AT41) e "008 / 000" (AT44, AT47).
- **Dificuldade:** o número fica ilegível e engana.
- **Consequência:** a memória de cálculo da incerteza não pode ser conferida pela tela.
- **Sugestão:** a mesma correção da Bioquímica: montar o texto no VBA, como já se faz no K5.
- **Evidência:** Estatística AT14 = '002 / 000', AT23 = '001 / 000'; RMS independente 2,491510594848434 (WBC) e
  1,2958189889137146 (HCT) (`vies_nordtest`).

### OC-H2 — O ET % do Painel conta bias 0 quando o analito não tem CEQ (RET%, RDW-SD)

- **Tipo:** calculo · **Prioridade:** Alto. Mesma raiz da OC-2 da Bioquímica.
- **Tela/módulo:** Painel, INDICADORES POR NÍVEL, H7:H9.
- **Ação:** ler a fórmula e simular com os dados dos analitos que têm ETp e não têm CEQ.
- **Esperado:** sem bias não há ET (ADR-030/068, "ausência não é zero"), como fazem a Estatística H ("") e o bloco
  ERRO TOTAL do próprio Painel (C61:C63, que lê a Estatística).
- **Observado:** `H7 = IF($E7="","",$E7*1.65+ABS(IF(ISNUMBER($G7),$G7,0)))`. Com G = "SEM EP", o ET vira 1,65·CV.
  - **RET%**, lote 522611 (ETp VB 22,80 %): o Painel mostraria **5,60 / 6,50 / 7,68 %**.
  - **RDW-SD** (ETp VB 3,67 %): o Painel mostraria **1,69 / 1,43 / 2,91 %**, com aparência de folga.
  - Nos dois casos, a Estatística H e o Painel C61:C63 ficam vazios.
- **Dificuldade:** dois ETs na mesma tela, e o da linha principal parece ótimo.
- **Consequência:** estatística incorreta exibida (ET subestimado) para 6 níveis da Hematologia.
- **Sugestão:** o ET do Painel seguir a Estatística (vazio ou "SEM EP" quando G não for número).
- **Evidência:** fórmula Painel!H7:H9 (`data_only=False`); Estatística G89:G91 (RET%) e G80:G82 (RDW-SD) = "SEM EP";
  CV recalculado (`CalculadoraHema.ciq_ano`): RET% 3,393624 / 3,938085 / 4,651903 e RDW-SD 1,025866 / 0,868839 /
  1,761059. **A confirmação visual no Excel (trocar C3 para RET%) fica com a F4.**

### OC-H3 — RDW: o CAP avalia o RDW-CV, mas a especificação (VB) está só no RDW-SD, e nenhum dos dois tem Sigma

- **Tipo:** duvida_metodo · **Prioridade:** Médio.
- **Tela/módulo:** Especificações (biblioteca 2025, colunas X/Y) × EQA_Base (de/para "RDW" → RDW-CV) × Estatística.
- **Ação:** cruzar, por analito, quem tem ETp e quem tem CEQ.
- **Esperado:** o analito avaliado no CEQ ter a especificação correspondente, ou uma decisão registrada de que não terá
  Sigma.
- **Observado:**
  - O de/para do CAP leva "RDW" (%) para **RDW-CV**: 15 amostras, bias de duas etapas calculado, mas a fonte é CLIA sem
    valor, então ETp = "-" e Sigma = "-".
  - O **RDW-SD** tem CVi 1,6 / CVg 4,2 (VB, MIN → ETp 3,665 %), mas nenhuma linha de CEQ: SEM EP, Sigma vazio.
  - O mesmo vale para IG%, IG#, NRBC%, NRBC#, RET# e IRF, que ficam sem ETp. Os quatro primeiros têm CEQ "NAO AVALIADO".
- **Dificuldade:** o usuário vê RDW com CEQ e RDW com ETp, e não entende por que não há Sigma em nenhum dos dois.
- **Consequência:** o plano de CQ do RDW não é calculado.
- **Sugestão (ao Especialista em CQ e ao laboratório):** confirmar a que medida (CV ou SD) se referem os CVi/CVg
  cadastrados, e se a especificação deveria estar no RDW-CV.
- **Evidência:** Especificações X50/X51 (RDW-SD 1,6/4,2) e Y50/Y51 vazios (RDW-CV); EQA_Base W:Y ("CAP | RDW |
  RDW-CV"); Estatística K/L de RDW-CV = "-" e G de RDW-SD = "SEM EP".

### OC-H4 — ETp "CLIA" 2025 da Hematologia diferente dos limites CLIA publicados para WBC e PLT

- **Tipo:** duvida_metodo · **Prioridade:** Médio.
- **Tela/módulo:** Especificações, biblioteca 2025, linha 48 (CLIA ETp%).
- **Ação:** conferir o ETp cadastrado com a fonte que o rótulo declara.
- **Esperado:** "CLIA" = limite de aceitação do CLIA vigente, ou o rótulo diz que é uma meta própria.
- **Observado:** WBC = 7 %, RBC = 4 %, HGB = 4 %, HCT = 4 % e PLT = 15 %. Pelo conhecimento do agente, a regra CLIA em
  vigor desde 2024 é WBC ±10 %, RBC/HGB/HCT ±4 % e PLT ±25 %. Os valores de RBC, HGB e HCT coincidem; os de WBC e PLT
  são mais rígidos. **Não é erro de cálculo:** o QC_INI usa corretamente o que está cadastrado. A dúvida é de cadastro
  e de rotulagem, e precisa ser confirmada pelo Especialista em CQ na fonte oficial.
- **Dificuldade:** o usuário lê "CLIA" e assume o limite regulatório.
- **Consequência:** se for meta interna, o Sigma do WBC (2,27 / 3,58 / 5,68) e o do PLT ficam conservadores em relação
  ao CLIA. Se for engano de digitação, o plano de CQ fica mais pesado do que o necessário.
- **Sugestão:** o Especialista confirmar a fonte e o laboratório corrigir o valor, ou trocar a fonte para FAB/meta
  própria, com registro.
- **Evidência:** Especificações B48 = 7, C48:E48 = 4, I48 = 15; Estatística K14:K16 = 7 e K35:K37 = 15.

### OC-H5 — O filtro "Ano De/Até" da Estatística não acompanha o lote e corta lotes que atravessam o ano

- **Tipo:** usabilidade · **Prioridade:** Médio.
- **Tela/módulo:** Estatística B3/D3 (Ano De/Até) × B4 (lote do Painel) → C:F e, por consequência, H, N, O, P e Q.
- **Ação:** recalcular C:F para outros lotes com o filtro de ano gravado (2025–2025).
- **Esperado:** ao escolher um lote no Painel, a estatística do lote aparecer inteira, ou o usuário ser avisado de que o
  filtro de ano exclui o lote.
- **Observado:**
  - Lote **619611** (ago–set/2026) com Ano 2025: **n = 0** em todos os analitos (com 2026, RBC N1 tem n = 54). O
    mesmo acontece com o lote em uso **625311**.
  - O lote **533811** (RBC N1: 22/12/2025–06/02/2026; reconferido na reexecução) com Ano 2025: RBC N1 **n = 8, CV 0,65 %**. Com o lote inteiro, n = 76 e
    CV de 1,37 %. O ET (H) passaria a usar o CV de 8 pontos.
  - Estado atual da produção: o lote em análise é o **522611**, de out/2025, com validade de 31/03/2026
    (Configuração C20). O ET (H) usa o CV de out/2025, enquanto o Sigma usa o CVp de abr–out/2026 e o bias usa o CAP
    2025. Nas células lidas do Painel não há aviso de lote vencido.
- **Dificuldade:** o usuário troca o lote no Painel e vê a Estatística vazia ou com um n pequeno, sem saber que o ano
  da B3:D3 é a causa.
- **Consequência:** o orçamento de erro (ET/margem) é calculado sobre um período errado ou fica em branco.
- **Sugestão:** levar ao UX (avisar quando o ano não cobre o lote, ou propor o ano do lote) e ao Especialista (se o ET
  deve usar o lote inteiro). A confirmação na tela fica com a F4.
- **Evidência:** `CalculadoraHema.ciq_ano('RBC', 1, lote, anoDe, anoAte)`: 619611/2025 → n = 0; 619611/2026 → n = 54;
  533811/2025 → n = 8 (CV 0,6517); 533811/sem filtro → n = 76 (CV 1,3716). Configuração C20/D21 = 522611 / 31/03/2026.

### OC-H6 — O bias do ET inclui amostras "NAO AVALIADO" que a incerteza exclui (IG%, IG#, NRBC%, NRBC#)

- **Tipo:** duvida_metodo · **Prioridade:** Baixo. Mesma questão da OC-1 da Bioquímica, onde o impacto é maior.
- **Tela/módulo:** Estatística G/H (mCEQ.BiasEQ) × AR/AU (mCEQ.ViesEQ).
- **Ação:** recalcular o bias de IG% pelas duas regras.
- **Esperado:** o mesmo critério de seleção de amostras nos dois indicadores do CEQ.
- **Observado:** as 15 amostras de IG% (e de IG#, NRBC% e NRBC#) estão NAO AVALIADO no CAP.
  - O BiasEQ as usa: G = **3,0094 %**, e o ET de IG% fica entre 7,5 % e 8,1 %.
  - O ViesEQ as exclui: AR = "SEM EP", AU = "SEM EP / SEM EP".
  - Como esses analitos não têm ETp, não há Sigma afetado. Só o ET é exibido.
- **Consequência:** ET exibido com base em amostras que o provedor não avaliou.
- **Sugestão:** a mesma decisão pedida na OC-1 (Especialista em CQ).
- **Evidência:** EQA_Base, chaves `CAP|2025|FH9-A 2025|IG|FH9-01` a `…|FH9-C 2025|IG|FH9-15` com Status_Padronizado =
  NAO AVALIADO; Estatística G68 = 3,0093784575914833, AR68 = "SEM EP".

### OC-H7 — Resultados isolados de 3–4 DP continuam ativos e pesam na u(Rw) da HGB N3 (+20 %)

- **Tipo:** dados · **Prioridade:** Baixo. Não há erro de cálculo: o sistema depende da inativação pelo laboratório.
- **Tela/módulo:** tblCQ_Final → u(Rw) (AF) e CVp (L).
- **Ação:** varrer os ativos procurando valores ≤ 0, RUN/nível repetido e valores mais próximos de outro nível.
- **Observado:**
  - Nenhum valor ≤ 0 e nenhuma repetição.
  - Casos isolados:
    - **HEM-241016**, HGB N3 = 13,6 (lote 533811, 02/02/2026; a mediana do N3 é 15,35). No mesmo RUN, MCH 25,4 e
      MCHC 29,1 também estão fora. A u(Rw) de 12 m da HGB N3 fica em **1,0025 %**; sem ele, seria **0,8341 %**.
    - LYMPH% e MONO% N1 do lote 614011, corridas de 29/06 e 06/07/2026: HEM-309465/309466 e HEM-314337/314338, a
      −3,96/−3,03 DP e +3,58/+3,51 DP. O Sigma do MONO% N1 iria de 1,616 para 1,728 sem eles, e o do LYMPH% N1 de
      4,326 para 4,642.
- **Consequência:** pequena. A aba já mostra o alerta "CV heterogêneo" na HGB.
- **Sugestão:** o laboratório revisar esses IDs e inativar os que forem não conformes, com motivo.
- **Evidência:** IDs acima; `cv_pooled_ext` com e sem os IDs (script da F3).

### OC-H8 — Um Sigma "legado" com bias contra o alvo do lote continua sendo gravado no Eng_Saida

- **Tipo:** duvida_metodo · **Prioridade:** Baixo. Confirma, na Hematologia, parte da OC-6 da Bioquímica.
- **Tela/módulo:** Eng_Saida I185:I187 (`engPainel`) e L190:L309 (`engEstat`).
- **Observado:**
  - Eng_Saida I185 dá RBC N1 = **−0,1148**, com G185 = 4,1446 % (bias do CIQ contra o alvo do lote).
  - `engEstat` L190 dá WBC N1 = **3,127 "Marginal"**, contra o Sigma oficial de 2,272 (Estatística L14).
  - Nenhuma célula visível lê essas colunas: o Painel lê B:F, e a Estatística lê C:F.
- **Consequência:** nenhuma na tela hoje. O risco é alguém (ou o BI) ler a coluna errada.
- **Sugestão:** retirar ou documentar as colunas legadas (Especialista e Orquestrador).
- **Evidência:** Eng_Saida linhas 184–193 (lidas com openpyxl).

### OC-H9 — Biblioteca de especificações 2026–2029 inteiramente vazia na Hematologia

- **Tipo:** dados · **Prioridade:** Médio. O mesmo risco da OC-7 da Bioquímica, aqui com todos os analitos.
- **Tela/módulo:** Especificações, linhas 55–89.
- **Observado:** só 2025 está preenchida. O GUIA §4 e o ADR-075 dizem que 2026–2029 ficam para o laboratório. Hoje os
  9 lotes têm ano de referência 2025 e S2 = 2025. Quando um lote novo nascer com o ano 2026, todos os ETp viram "-" e
  todos os Sigmas somem, **sem aviso**.
- **Sugestão:** o laboratório preencher 2026, ou o QA avisar quando o ano do lote não tiver a especificação da fonte.
- **Evidência:** Especificações B57:AC62 (2026) a B84:AC89 (2029) vazias; Analitos, blocos dos 9 lotes = 2025/2025.

### OC-H10 — "FAB CVTp" 2025 com valores que não parecem CV % (RDW-SD 48,1; RDW-CV 15,1; MPV 9,4; RET% 1,2)

- **Tipo:** dados · **Prioridade:** Baixo. Não afeta nenhum cálculo hoje, porque nenhum analito usa a fonte FAB.
- **Observado:** pelo ADR-075, a linha 53 recebeu o antigo "CV Fab". Para os demais analitos os valores fazem sentido
  (3, 1,5, 1, 4…), mas RDW-SD 48,1, RDW-CV 15,1, MPV 9,4 e RET% 1,2 têm cara de valor de controle, e não de CV %. A
  linha "FAB ETp%" (52) está vazia: se o laboratório trocar a fonte para FAB, o ETp vira "-", e a coluna U mostraria um
  CVTp de 48,1 % no RDW-SD.
- **Sugestão:** o laboratório revisar a linha 53 antes de usar a fonte FAB.
- **Evidência:** Especificações X53 = 48,1, Y53 = 15,1, Z53 = 9,4, AA53 = 1,2; B52:AC52 vazios.

### OC-H11 — Linhas vazias da Estatística mostram "0" e "SEM EP"

- **Tipo:** usabilidade · **Prioridade:** Baixo. A raiz é a mesma da OC-5 da Bioquímica, mas na Hematologia não há
  contador afetado.
- **Observado:** A98:A133 = 0 (catálogo vazio a partir de Especificações A32) e G, R, S, T = "SEM EP". A fórmula testa
  `$A="" ` e o A vale 0.
- **Sugestão:** testar também A = 0 (`OR($A="",$A=0)`, como já fazem as colunas K e AF).
- **Evidência:** Estatística A98:T133.

### OC-H12 — O cabeçalho da tblCQ_Final mostra contagens com decimal ("Total: 44392,0 · participam: 43710,0")

- **Tipo:** usabilidade · **Prioridade:** Baixo.
- **Observado:** a linha 5 da aba Principal - Resultados formata contagens inteiras com ",0".
- **Sugestão:** formatar sem decimais.
- **Evidência:** Principal - Resultados!A5.

### OC-H13 — Alvo (Média/DP) do lote 522611 incompatível com os dados do lote em 12 pares: IG# com valores de IG% e RDW-SD N2 = 14,78 fL

- **Tipo:** dados · **Prioridade:** Alto. Encontrada na terceira passada, ao ler o cadastro do lote para conferir o
  bloco da aba Analitos. Não é erro de cálculo: o QC_INI usa o alvo que está cadastrado.
- **Tela/módulo:** Analitos, bloco do lote 522611 (E6:J33) → LotesStore → limites do Levey-Jennings e Westgard
  (ADR-078/083) e bloco AU:BD (OC-H14). O mesmo bloco aparece em Especificações E:J ("lote em edição").
- **Ação:** comparar a média medida de cada analito × nível (PARTICIPA = SIM, lote inteiro) com o alvo cadastrado,
  em DP do alvo, nos 9 lotes.
- **Esperado:** alvo coerente com a grandeza (unidade e ordem de grandeza) e com o material daquele nível. Uma
  diferença de alguns DP entre o alvo do fabricante e a média do laboratório é comum. Dezenas ou centenas de DP
  indicam um valor do analito ou do nível errado.
- **Observado (lote 522611; os outros 8 lotes têm 0 pares acima de 5 DP):**
  - **IG# N1/N2/N3:** alvo 10,29 / 11,05 / 11,88 contra médias de 0,304 / 0,780 / 2,009 (z = −28,6 / −31,6 / −29,5).
    O alvo tem a ordem de grandeza do **IG%** (cadastro do IG%: 9,98 / 11,51 / 11,99).
  - **RDW-SD N2:** alvo 14,7793 ± 0,0819 fL contra a média de 47,38 fL (z = +398). Um RDW-SD de 14,8 fL não é
    plausível; o valor tem a ordem de grandeza de um RDW-CV.
  - **RDW-CV N2:** alvo 9,8345 contra 17,54 (z = +25,7); RDW-CV N3 a −6,5 DP.
  - **Nível 2 da série vermelha:** HGB N2 −14,6 DP, MCH N2 −12,4, MCV N2 −7,3 e HCT N2 −5,7. Os N1 e N3 dos mesmos
    analitos ficam a menos de 2,2 DP.
  - **RET% N1 +5,2 DP e RET# N1 +5,6 DP.**
  - Outros 21 pares do mesmo lote são valores de TESTE (média do lote e DP = 1,5 × DP do lote), e o próprio sistema
    os sinaliza (Analitos B2: "contém valores de TESTE (ADR-076) em 21 par(es)"). **Os 12 pares acima não estão
    nesse aviso.**
- **Consequência:** os limites do LJ e o Westgard desses 12 pares ficam deslocados. A Eventos_Westgard do lote 522611
  tem 6.155 eventos, e os analitos do topo são exatamente os afetados: IG# 714 (152 de 1_3s), RDW-SD 661 (117),
  RDW-CV 609 (111), RET# 606, RET% 595, MCV 428, HGB 413, MCH 408 e HCT 350. Quem abre o Painel no lote em análise
  (522611, o lote do Painel e de Configuração C20) vê rejeições que vêm do cadastro, e não do analisador. O ET e o
  Sigma do bloco do lote (OC-H14) também usam esse DP. A Estatística (C:W) não usa o alvo e não é afetada.
- **Dificuldade:** o laboratório não tem como perceber pelo aviso de TESTE, que não cobre esses pares.
- **Sugestão:** o laboratório conferir na bula do lote 522611 o IG#, o RDW-SD/RDW-CV N2, a série vermelha do N2 e o
  RET N1, e corrigir pelo caminho com trilha da aba Analitos. Ao Orquestrador/UX, avaliar um aviso de plausibilidade
  (por exemplo, a média do lote a mais de k DP do alvo depois de n ≥ 20 resultados). A confirmação na tela do LJ fica
  com a F4.
- **Evidência:** Analitos E25:J25 (IG# 10,2897 ± 0,3488 / 11,0517 ± 0,3247 / 11,8759 ± 0,3345), G28:H28 (RDW-SD N2
  14,7793 ± 0,0819), G29:H29 (RDW-CV N2 9,8345 ± 0,3003), G8:H8 (HGB N2 12,2759 ± 0,083); médias do lote recalculadas
  de tblCQ_Final (PARTICIPA = SIM, LOTE = 522611): IG# 0,3036 / 0,7798 / 2,0094, RDW-SD N2 47,3765 (n = 51), HGB N2
  11,0647; contagem por analito da Eventos_Westgard (cabeçalho "Total de eventos: 6155 · Lote: 522611"). Scripts
  `hema_f3r3/alvos.py` e `teste_pairs.py` (no diretório do ciclo).

### OC-H14 — O bloco do lote (Analitos AU:BD) mostra outro ET e outro Sigma para o mesmo lote, e em 21 pares eles saem de um DP de TESTE

- **Tipo:** duvida_metodo · **Prioridade:** Médio. Mesma raiz da OC-6 da Bioquímica (vários Sigmas para o mesmo
  analito), aqui agravada pelos valores de TESTE e pela OC-H13.
- **Tela/módulo:** Analitos, colunas AU:BD de cada bloco de lote (Bias CEQ %, CV % N1, ET % N1, Sigma N1… N3).
- **Ação:** recalcular o bloco e comparar com a Estatística do mesmo lote (522611).
- **Esperado:** um só ET e um só Sigma por analito, nível e lote, ou um rótulo que diga que o bloco usa o CV do
  cadastro (DP/Média do alvo), e não o CV medido.
- **Observado:** o cálculo do bloco está certo pela própria definição (2.268/2.268, F3-T27). A definição é outra:
  - CV = 100·DP/Média **do cadastro**;
  - ET = 1,65·CV + |G|;
  - Sigma = (ETp − |G|)/CV.

  A Estatística usa o CV medido (ET) e o CVp de 6 m (Sigma). No lote 522611, por exemplo:

  | Analito × nível | CV do cadastro | CV medido | Sigma do bloco | Sigma da Estatística |
  |---|---|---|---|---|
  | MONO% N3 | 10,99 % | 7,32 % | **1,79** | **4,82** |
  | MONO% N2 | 15,29 % | 10,19 % | 1,29 | 3,07 |
  | HCT N1 | 1,86 % | 2,09 % | 1,54 | 2,10 |
  | PLT N1 | 6,64 % | 4,76 % | 1,51 | 2,26 |
  | WBC N1 | 1,76 % | 1,80 % | 2,92 | 2,27 |

  No MONO%, o CV do cadastro é exatamente 1,5 × o CV medido, porque o DP cadastrado é um valor de TESTE
  (DP = 1,5 × DP do lote). O mesmo vale para os 21 pares de TESTE (NEUT%, LYMPH%, MONO%, EO%, BASO%, IG% e MPV): ali o
  Sigma do bloco fica 1,5 vez menor do que ficaria com o DP real do lote.
- **Dificuldade:** o usuário vê dois Sigmas para o mesmo lote (Analitos e Estatística/Painel), com classes
  diferentes (MONO% N3 "Desempenho inadequado" × "Bom").
- **Consequência:** a decisão de plano de CQ pode ser tomada sobre o número errado. Hoje o Painel e o plano leem a
  Estatística, e não o bloco.
- **Sugestão:** o Especialista em CQ decidir se o bloco do lote deve mostrar ET/Sigma e, se sim, com qual CV (e o UX
  rotular). Enquanto houver pares de TESTE, ocultar o ET/Sigma do bloco nesses pares.
- **Evidência:** fórmulas Analitos AV6:AX6 (`100*$F6/ABS($E6)`, `1.65*$AV6+$AU6`, `($T6-$AU6)/$AV6`); valores AV/AX
  do bloco 522611 contra Estatística F/L (`hema_f3r3/blocos.py`, `blocos2.py`); lista dos 21 pares de TESTE
  (`teste_pairs.py`).

## 6. Dúvidas de regra de negócio (para o Especialista em CQ, via Orquestrador)

1. O BiasEQ "ABS" deve excluir as amostras NAO AVALIADO (e as de resultado ou alvo 0), como o ViesEQ? (OC-H6 / OC-1)
2. Na Hematologia, o ET (H) usa o CV do lote no ano B3:D3, e o Sigma usa o CVp de 6 m. Com o lote 522611 (out/2025) e
   a janela abr–out/2026, ET e Sigma descrevem períodos distantes um ano. Qual CV deve entrar no orçamento de erro?
   O filtro por **ano civil** é o desejado, se ele corta lotes que atravessam o ano (lote 533811: 8 de 76 pontos)?
   (OC-H5 / OC-6)
3. Os ETp "CLIA" de WBC (7 %) e PLT (15 %) são metas internas ou deveriam ser 10 % e 25 %? (OC-H4)
4. A especificação de variação biológica do RDW é do RDW-CV ou do RDW-SD? (OC-H3)
5. IG%, IG#, NRBC%, NRBC#, RET# e IRF ficam sem ETp. É decisão consciente (não terão Sigma nem plano)?
6. A base da Hematologia na produção termina em 01/10/2026 (uma carga única em 02/10, Sync_Log vazia). A janela do Sigma
   está ancorada nessa data. É o estado esperado até a virada do MODO_FONTE no laboratório? (assunto da F1)
7. O bloco do lote na aba Analitos deve mostrar ET e Sigma? Se sim, com o CV do cadastro, com o CV medido ou com o CVp
   do Sigma? (OC-H14 / OC-6)
8. O alvo do lote 522611 (IG#, RDW-SD/RDW-CV N2, série vermelha N2, RET N1) é o da bula? O sistema deve avisar
   quando a média do lote se afasta muitos DP do alvo cadastrado? (OC-H13)

## 7. Relatório A — impressões de uso (o que o usuário vê na F3)

- **Os números batem.** Em 2.268 comparações na Estatística e 45 no Painel, não houve nenhuma divergência de cálculo,
  e a matemática segue a definição dos ADR:
  - as duas etapas do bias;
  - o ano exato do ADR-072;
  - a exclusão das 90 linhas de simulação;
  - o CVp só com lotes de n ≥ 20 (o 602911 entra com n = 20 exato);
  - o "-" do ADR-068 em 21 linhas reais.
  
  O Painel repete a Estatística sem recalcular, exceto no ET da linha principal (OC-H2).
- **O que confunde é o período.** A Estatística da Hematologia filtra por ano (B3:D3) e pelo lote do Painel, enquanto o
  Sigma olha os últimos 6 meses de todos os lotes. Hoje a tela mostra o lote 522611, de outubro de 2025 e já vencido,
  ao lado de um Sigma de abril a outubro de 2026. Se eu trocar para o lote atual sem mexer no ano, a Estatística
  zera. Como usuário, eu acharia que o sistema perdeu os dados.
- **Faltam especificações e CEQ em 9 dos 28 analitos.** Sete não têm ETp, e RDW-SD e RET% não têm CEQ. O RDW tem CEQ
  num nome e ETp no outro (OC-H3). Nada na tela explica o motivo da ausência de Sigma, além do "-" e do "SEM EP".
- **O alvo do lote engana mais do que a estatística (terceira passada).** O lote que o Painel abre (522611) tem o IG#
  cadastrado com números de IG% e um RDW-SD de 14,8 fL no N2. O aviso de TESTE da aba Analitos cobre outros 21 pares,
  mas não esses. Como usuário, eu veria centenas de rejeições de Westgard no IG# e no RDW e acharia que o analisador
  está com problema (OC-H13). No mesmo lote, a aba Analitos mostra um Sigma diferente do da Estatística, e eu não
  saberia qual vale (OC-H14).
- **Rastreabilidade boa.** O K5 diz o que entrou (CAP, 2025, 3 rodadas, 360 linhas). As chaves do CEQ chegam à
  amostra, e os IDs do CIQ chegam ao resultado. A coluna AT ilegível (OC-H1) é a exceção.

## 8. Arquivos

- Calculadora: `_arquitetura/agentes/ferramentas/calc_independente.py` (estendida para a Hematologia; regressão da
  Bioquímica idêntica).
- Saídas da execução, no diretório do ciclo em `%TEMP%\qcw\auditoria\2026-10-31_ciclo01\`, não versionadas (têm
  resultados de controle, mas nenhum dado de paciente):
  - `hema_f3.json`: JSON completo (84 linhas × 27 colunas + 45 células do Painel);
  - `hema_f3.txt`: resumo;
  - `bio_antes.*` e `bio_depois.*`: regressão;
  - `hema_f3_r2.json` / `hema_f3_r2.txt`: reexecução (idêntica);
  - `hema_f3_cruzado.txt`: saída da 2ª implementação (F3-T22);
  - `hema_f3r3/` (terceira passada, 01/11): `hema_f3_r3.json`/`.txt` (calculadora), `segunda_opiniao.txt` (F3-T24),
    `tabela_4_3b*.md`, e os scripts de conferência pontual `evid.py` (F3-T25), `extra.py` (F3-T26), `blocos.py` e
    `blocos2.py` (F3-T27/OC-H14), `alvos.py` e `teste_pairs.py` (F3-T28/OC-H13), `rdw.py` (médias por lote).
- Conferência cruzada: `_arquitetura/agentes/ferramentas/conferencia_cruzada_hema.py` (documentada no cabeçalho).
- Terceira implementação: `_arquitetura/agentes/ferramentas/segunda_opiniao_hema.py` (documentada no cabeçalho; os
  filtros da cópia são parâmetros no topo). O cabeçalho da calculadora agora aponta para as duas conferências.
- Cópias `Hematologia_F3.xlsm` e `Hematologia_F3_regressao_bio.xlsm`: descartadas ao fim da frente (REGRAS §2.3); a
  cópia da terceira passada também. A
  produção não foi tocada. Para reproduzir, copie de novo a produção e confira se o sha256 curto ainda é
  `0b45fa270d2a` (Hematologia) e `fe10dba1f1fd` (Bioquímica).
