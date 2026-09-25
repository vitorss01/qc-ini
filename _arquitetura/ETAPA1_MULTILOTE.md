# Etapa 1 — Múltiplos lotes de controle (Bioquímica)

**Data:** 19/09/2026 · **Arquivo:** `QC_Bioquimica.xlsm` · **Decisão:** ADR-049
**Fontes e testes:** `_arquitetura/etapa1_multilote/` (`src/`, `aplicar_etapa1.py`,
`testes/`, `resultados/`)

---

## 1. Auditoria da implementação anterior (as 9 perguntas)

**1. Como os lotes estavam armazenados?** O cadastro ficava em `Configuração!C26:C50`
(25 posições) e o lote em uso em `Configuração!C20`. Cada resultado grava o código
`QC-<lote><nível>` em `DB_Resultados!D`, e a coluna BA guarda o núcleo do lote.

**2. Onde ficavam média, DP e limites?** Em dois lugares:
- `Analitos!E:J`: a tela do lote carregado.
- `LotesStore`: um bloco de 40 linhas por lote, posicionado pela ordem do lote no cadastro.

Os limites não eram configurados em nenhum dos dois; eram derivados como média ± 1/2/3 DP.
No arquivo real só o lote **8974** tinha parâmetros. O **8973** (jan–fev/2025, 39 corridas) não tinha nenhum.

**3. Como o gráfico obtinha esses valores?** Pela tela: `Calc!AX1:BA1` apontava para
`Analitos!E:H`, e os eixos para `Analitos!AC:AH`. Nunca pela `LotesStore`.

**4. As fórmulas dependiam de um lote específico?** Dependiam do **lote ativo**
(`loteAtivo`) em `Calc!B/C/F/AB`. O motor, porém, escolhia o lote **pelo ano**
(`LoteDoAnoOuAtivo`), e a Estatística somava **todos** os lotes. Eram três lotes
diferentes na mesma tela.

**5. Algum ponto impedia múltiplos lotes?** Sim:
- A troca de lote gravava 12 colunas guardadas sobre as 14 de `aInput`. Isso transformava
  as fórmulas de especificação `Analitos!K:P` em valores velhos e punha `#N/D` em Q:R.
- O cadastro tinha teto de 25 lotes.
- A descoberta de corridas parava em 180, na ordem do banco. Num lote longo, as corridas
  **mais novas** sumiriam do gráfico.

**6. Havia risco de mistura?** Sim, e ela acontecia:
- **Resultados de um lote interpretados com média/DP de outro:** o z usava a tela
  carregada, fosse qual fosse o lote dos resultados.
- **Lote novo herdava média/DP do anterior.**
- **Lote sem parâmetro saía todo verde:** DP 0 dava z = 0, lido como "no alvo".
- **Estatística juntava lotes:** somava CV de lotes diferentes e calculava o bias contra
  o alvo do lote ativo.

**7. Era escalável para 5 anos?** Não. Toda escrita de célula custava cerca de 12 s,
porque um nome com `OFFSET` (volátil) arrastava cerca de 400 funções da Estatística para
cada recálculo. Além disso, o cadastro tinha teto de 25 lotes e o gráfico teto silencioso
de 180 corridas.

**8. Ponto mais frágil:** média/DP lidos da **tela** (uma cópia que depende de qual lote
estava carregado) em vez de lidos pela **chave do lote**. Todos os outros defeitos de
mistura decorrem disso.

**9. O que era necessário:** uma única fonte de parâmetros por chave (lote + analito);
um seletor de lote de análise separado do lote de lançamento; recusar a interpretação
quando não há parâmetro; série cronológica; e o gráfico sempre coerente com o motor.

## 2. O que mudou (A–E)

**A/B. Alterações e onde:**

| Onde | O quê |
|---|---|
| `mLotes` (reescrito) | `ParametrosLote(lote, analito, nível)` passa a ser a **única** leitura de média/DP. Também: troca do lote de análise, edição auditada, conferência da store e código de lote canônico ("010" não vira 10). |
| `mEstatistica` | `AlvoAnalito` lê do lote pedido; corridas em ordem cronológica (`OrdenarCorridas`); janela das 180 mais recentes até o fim do período; nível sem parâmetro não entra no Westgard; filtro de período por dia; um recálculo por operação; spinner refaz o motor. |
| `mEstatPeriodo`, `mBI` | Alvo pela mesma chave; o BI usa a mesma ordem cronológica. |
| Módulos das abas Painel, Analitos e EstaPastaDeTrabalho | Painel: E3 troca o lote e De/Até refazem o motor. Analitos: a digitação grava na hora. Ao abrir o arquivo, os três lotes são sincronizados. |
| `LotesStore` | Colunas O = Analito e P = Chave (`LOTE\|ANALITO`); capacidade para 100 lotes × 40 analitos. |
| `Painel` | **E3 = seletor de lote**; faixa O3 com lote e alvo, ou aviso de falta de parâmetro; M7/M8 dizem "SEM MÉDIA/DP DO LOTE — não avaliado" em vez de "Sem violação"; C4 mostra o lote de lançamento. |
| `Calc` | Corridas lidas do motor (vazio se o motor for de outro analito ou lote); média/DP lidos da store pela chave; eixos derivados da média/DP; filtro de período por dia; ponto sem parâmetro não vira "violação". |
| `Configuração` | Cadastro vai a **100 lotes**; lista compacta `lstLotes`. |
| `Estatística` | H4 reflete o lote do Painel. |
| Nomes | `loteSel`, `loteAnalise`, `loteParam`, `ls*`, `lstLotes`, `eng*`; `lstAnosCIQ` e `lstAnosCEQ` sem `OFFSET`. |

**C. Seleção do lote:** o usuário escolhe o lote em **Painel!E3**, numa lista com os
lotes cadastrados. O `Worksheet_Change` do Painel chama `mLotes.TrocarLoteAnalise`, que:
1. grava a tela de parâmetros do lote anterior;
2. carrega a do novo;
3. refaz o motor.

O lote **de lançamento** (`Configuração!C20`) é outra coisa. Escolher um lote antigo
para análise **não** muda o lote em que o próximo resultado é gravado. Trocar o lote de
lançamento leva o Painel junto.

**D. Recuperação de média/DP/limites:** a `LotesStore` tem uma linha por
(lote, analito), com a chave na coluna P.
- **VBA:** `ParametrosLote` devolve média/DP ou **False**. Ausente, texto ou DP ≤ 0
  contam como "sem parâmetro".
- **Fórmulas:** `Calc!BH1` localiza a linha pela chave e `AX1:BA1` leem média/DP N1/N2.
- **Limites:** ±1/2/3 DP e o eixo são derivados desses valores. Não há como o gráfico
  ter a média de um lote e os limites de outro.

**E. Gráfico:** as 14 séries leem `Calc`. As linhas ±1/2/3 DP e a média vêm de
`AX1:BA1`, os pontos vêm das corridas publicadas pelo motor **para aquele lote e
analito**, e as cores vêm do veredicto do motor, que calculou z com a mesma média/DP.
Se o motor não for do lote e analito em tela, o gráfico fica vazio: nunca mostra dado
de um lote com parâmetro de outro.

## 3. Testes (F–G)

Bateria automatizada (`testes/suite_etapa1.py`, `testes/t10_historico.py`). Tudo pelo
caminho do usuário: cadastro na Configuração, seletor do Painel, digitação na aba
Analitos e **importação pela aba Importar**. As conferências leem o Calc, **as séries do
próprio objeto gráfico**, a escala do eixo, o Painel, o Eng_Saida, os Eventos_Westgard,
a Estatística, o BI_Data e o sentinela Excel × BI.

| Teste | Cenário | Resultado |
|---|---|---|
| T0 | Seletor ligado de verdade (eventos): E3 digitado troca o lote; digitar média na aba Analitos grava na store e não toca o outro lote | PASS |
| T1 | Um lote: média/DP, limites, séries do gráfico, eixo, z, 1_3s plantado, M7, store | PASS |
| T2 | Lote A (100/5) × B (200/10): tudo troca; o 235 (z=+3,5 em B) é 1_3s só em B e não aparece em A | PASS |
| T3 | A→B→C→A→C→B, mais troca de analito no meio: sem memória do anterior; store intacta | PASS |
| T4 | Limites 5±0,1 × 5000±250: séries e escala do eixo acompanham, nas duas direções | PASS |
| T5 | Dois lotes nas **mesmas datas** (corrida paralela), dois analitos: RUNs distintos por lote; cada linha do BI com o alvo do SEU lote; sentinela Excel × BI = 0 divergências | PASS |
| T6 | Lote 2601 → entra 2602 (vira lote em uso, sem herdar parâmetro): o Calc, os gráficos, o Painel, a store, o banco e o BI do 2601 ficam **idênticos** aos de antes | PASS |
| T7 | 10 lotes (com "001" e alfanumérico "L-007A"): todos corretos; nenhuma aba ou gráfico novo | PASS |
| T8 | Sem parâmetro; só N1; DP 0; DP negativo; média em texto; lote sem dado: nada é interpretado, avisos corretos, o lote configurado fica intacto | PASS |
| T9 | Alterar média/DP: efeito imediato no histórico **só daquele lote**; Audit_Log com antes/depois; desfazer reproduz o gráfico original | PASS |
| **T0–T9** | | **760/760 verificações** |
| T10 | 5 anos: 10 lotes, 1.625 corridas, 100.750 resultados | ver §4 |

**Regressão no arquivo real:** comparei 4 analitos, 129 corridas cada, antes e depois
da Etapa 1.
- **Iguais:** as mesmas corridas, **os mesmos valores, z e limites**, e a mesma Estatística.
- **Diferentes:** a ordem do eixo (agora cronológica) e 7 veredictos, todos explicados.
  - O RUN 93 é de 03/04/2026, mas era avaliado como se fosse do fim de maio.
  - Na Creatinina, as corridas de dezembro/2025 passam a preceder as de janeiro/2026.

## 4. Longevidade (T10)

A longevidade foi medida duas vezes.
- **T10** (antes da camada de desempenho): 10 lotes e 100.750 resultados. Trocar de
  analito levava 25,8 s, importar uma corrida levava 25–33 s e abrir o arquivo levava 14 s.
- **Cinco anos com o número real de lotes:** 20 na Bioquímica e 40 na Hematologia,
  depois do ADR-050. O clique no spinner ficou em 0,06–0,15 s, a troca de lote em
  0,2–1,1 s, a corrida do dia em 4–5 s e a abertura em 2,6 s. Foram **34/34 verificações
  em cada produto**. As tabelas completas estão em `ENTREGA_5ANOS.md` §2.

## 5. Problemas encontrados e corrigidos (H)

1. Troca de lote destruía as fórmulas `Analitos!K:R` (colunas desalinhadas).
2. Lote novo herdava média/DP do anterior.
3. Lote sem média/DP saía todo verde (z = 0).
4. Gráfico, motor e Estatística usavam três lotes diferentes.
5. RUN não cronológico: regras de sequência avaliadas fora do tempo.
6. Teto silencioso de 180 corridas cortava as **mais novas**.
7. Spinner de analito não refazia o motor: o Painel mostrava "Sem violação" para a
   Glicose, que reprova com 3 violações (provado no arquivo real).
8. Toda escrita de célula custava cerca de 12 s (`OFFSET` volátil), e com a correção
   passou a 0,00 s.
9. Filtro de período por data e hora: resultados com hora sumiam do último dia do período.
10. Cadastro limitado a 25 lotes; ampliado para 100.
11. ETp do motor lido da coluna errada ("ETp VB" em vez de "ETp em uso").
12. Códigos de lote com zero à esquerda ("010") viravam número no seletor.

## 6. Riscos que permanecem (I)

Ver `ENTREGA_5ANOS.md` §5. Os principais riscos são dois:
- **Versionamento de parâmetro:** mudar média/DP reinterpreta o histórico do lote. A
  mudança fica auditada, mas não é congelada.
- **Teto do banco:** 150 mil linhas na Bioquímica e 200 mil na Hematologia. Acima disso,
  a gravação é recusada com mensagem.

## 7. Conclusão (J)

A arquitetura agora suporta vários lotes sem mistura. Cada lote tem média, DP e
limites próprios, lidos pela chave `LOTE|ANALITO`. Um lote sem parâmetro não é
interpretado. O gráfico só mostra o que o motor avaliou. A arquitetura também suporta
cinco anos de uso com resposta de aplicativo (ADR-050). Isso foi comprovado pela suíte
T0–T9 (760/760) e pelos testes de cinco anos com 20 lotes (Bioquímica) e 40 lotes
(Hematologia). A Hematologia, que estava congelada, foi levada ao mesmo motor
(ADR-052).
