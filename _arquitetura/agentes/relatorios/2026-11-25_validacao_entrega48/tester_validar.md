# Tester/QA: VALIDAR, entrega 48 (ADR-102), FASE 1 sem Excel

Data: 25/11/2026. Commit 8e38d99. Escopo: `src/comum/mAlvosRef.bas`, `src/{bio,hema}/Painel.cls`, `mIntegracao.bas`
(gancho no fim do ATUALIZAR), `mSeguranca.bas` (lista do UnlockApp), `src/comum/EstaPastaDeTrabalho.cls`,
`instalar_adr102.py`, `testes/qa_alvos_manual.py`, `agentes/ferramentas/aplicar_alvos_ciq.py`.

- **Independência:** o Tester não implementou nada desta entrega.
- **Restrição desta fase:** uma entrega está rodando. Não abri Excel/COM e não rodei nenhuma suíte. Os resultados de
  execução citados abaixo são do implementador e **não** foram reproduzidos.

## Veredito preliminar: VALIDADA COM RESSALVAS (só revisão estática), a confirmar na fase 2

- Não encontrei defeito crítico nem alto na lógica do ADR-102.
- Os desvios achados são de rastreabilidade e de baixo impacto (D-01 a D-03).
- A validação final depende de duas execuções na cópia da produção nova, depois da entrega:
  - o `repro_e48_alvos.py`, nos dois produtos;
  - a regressão do gate, que já é rodada pelo `entregar.py`.

## 1. Revisão da lógica contra o ADR (`mAlvosRef.bas`, salvo outra indicação)

| Item do ADR | Onde | Resultado |
|---|---|---|
| Só ativos (PARTICIPA = SIM) | `:486`; o `mDados.bas:97-101` mapeia SIM para "Ativo" | conforme |
| Todos os equipamentos | `Apurar` `:474-513`, sem filtro de EQUIPAMENTO e sem ler L4 | conforme |
| Lote, analito e nível exatos | `:487-490` (UCase+Trim; lote canônico pela `LoteCanonico`) | conforme |
| Datas inclusivas; vazio = sem limite | `:492-500` (comparação por `Int()` do dia; registro sem data só entra sem limites) | conforme |
| DP n−1 em duas passadas | `mEstatistica.bas:286-304` (Bio) e `:303-311` (Hema); DP = 0 vira `DP_ZERO` em `:446-452` | conforme |
| N < 2 / 2–19 / 20–89 / ≥ 90, configurável | `CondicaoN` `:430-442`; `CfgNum`/`LimitesN` `:406-420` | conforme |
| Trava: vazio/TODOS, inexistente, duplicada | `:682-708`; `TravaDoLote` `:1008-1020`. `LinhaStore` (`:528`) e `mLotes.IndiceStore` (`mLotes.bas:232`) usam a mesma coluna P e a 1ª ocorrência, então o `GravarParametroStore` grava na mesma linha `rLs` | conforme |
| Grava só o campo escolhido (1/2/3); cancelar não grava | `PerguntarModo` `:170-188`; gravação `:829-848`; bit TESTE só do campo `:856-861`; retornos CANCELADO em `:736-741`, `:759-764`, `:776-781` | conforme |
| Exceção: só depois de liberar | `:1142-1153` (memória `mLiberado` **e** marca Z105 da mesma chave) | conforme |
| Exceção: justificativa ≥ 10 caracteres | `:1157-1162` | conforme |
| Exceção: valores > 0 | `:1169-1183` | conforme |
| Exceção: DP do lote anterior | puxado da LotesStore em `:1095-1101` e conferido na gravação em `:1196-1204`; o próprio lote é recusado em `:1190` | conforme |
| Exceção: marca por campo | `:1316-1321`; o N e o período só são limpos com os dois campos (`:1320`) | conforme |
| Exceção: 2ª confirmação sobre valor PRÓPRIO | `:1246-1268` | conforme |
| Exceção: rebloqueio | depois de gravar (`:1337`); em falha (`:1349`); troca de chave (`:998-1005`, Painel B3:C3 e B98/E98/G98); analito trocado sem evento é pego no gravar (`:1147`); Open e BeforeClose (`EstaPastaDeTrabalho.cls:6` e `:82`) | conforme |
| Aplicar modo 3 limpa a marca do DP | `:806-812`, por análise: `EXCECAO_TEMPORARIA_DP` + modo 3 próprio dá `ACEITAVEL`/`IDEAL`; se o N for baixo, dá `…TEMPORARIO_DP` | conforme (falta prova dinâmica: X06) |
| LI/LS: igual ao limite = dentro | `:1606-1611` (`<` e `>` estritos) | conforme |
| Observações: só nova alerta; PENDENTE some e vira RETIRADA; RETIRADA volta e vira PENDENTE; CIENTE nunca muda | `:1641-1662` e `:1659-1690` | conforme |
| Parecer | `:1919-2012`: texto ≥ 10, confirmação, RefBulaParecer, PENDENTE vira CIENTE só no lote, Audit_Log | conforme |
| Referência nunca grava na LotesStore | depois de `:1353`, nenhum `ABA_LS`, `LS_C_` ou `GravarParametroStore`. `GradeEditada` só escreve na RefBulaStore, `VerificarNucleo` só na RefBulaObs, o parecer só na RefBulaParecer. O `GravarParametroStore` só aparece em `:831-832` e `:1289` | conforme |
| Falha da vigilância não derruba o ATUALIZAR | `mIntegracao.bas:191-196` (`On Error Resume Next` local); `VerificarAoAtualizar` `:1756-1780` | conforme |
| Abas RefBula* VeryHidden para quem não é ADM | `mSeguranca.bas:692-694` | conforme |

### Defeitos e desvios (nenhum crítico ou alto)

- **D-01 (Baixo, rastreabilidade): AA:AO descreve só o campo transferido.**
  - Onde: `mAlvosRef.bas:862-866`.
  - O que acontece: na transferência parcial (modo 2 ou 3), `N_USADO_Nk`, `DATA_INI_Nk` e `DATA_FIM_Nk` são trocados pelo
    N e pelo período da apuração nova. O campo que não foi transferido vinha de outro N e de outro período.
  - Consequência: a auditoria AA:AO passa a descrever só o campo transferido. O Audit_Log guarda o histórico completo.
  - Proposta: documentar no ADR ou no backlog. Não precisa corrigir agora.
- **D-02 (Baixo): condição própria herdada por um campo que não veio do CIQ.**
  - Onde: `mAlvosRef.bas:806-812`, com efeito em `:1246-1268` e `:1296`.
  - O que acontece: na transferência parcial sobre uma linha sem AA:AO (valor antigo do bloco ou da bula), a condição
    gravada recebe `ACEITAVEL`/`IDEAL` e o campo não transferido passa a contar como "próprio".
  - Consequência:
    - mais tarde, a exceção sobre esse campo pede a 2ª confirmação "substituir valor PRÓPRIO";
    - o Audit_Log registra "substituiu valor PROPRIO" para um valor que não veio do CIQ.
  - Efeito: conservador (pede confirmação a mais), mas a trilha fica com um texto inexato. Vai para o backlog.
- **D-03 (Baixo): a origem W de um nível mantido perde o texto anterior.**
  - Onde: `mAlvosRef.bas:855`.
  - O que acontece: o W é reescrito para a linha inteira. Num nível mantido (N < 2 ou DP = 0), a origem textual
    anterior (ex.: "digitado…", "TESTE…") se perde. Só sobrevive o que tinha marca em AA:AO.
  - Atenuantes: o bit TESTE em X:Z fica e o Audit_Log guarda o histórico. Vai para o backlog.
- **Observação de UX (Baixo):** o analito trocado pelo spinner (B3 sem Change) não refaz as linhas 100–10x, e o Z98
  continua com o analito antigo. O `AplicarAlvos` recalcula para o analito atual e a confirmação mostra o analito
  certo, então nada é gravado errado.
- **Pré-existente, não é desta entrega:** `testes/qa_final.py:64` tem a senha literal (BK-05, aceitação do risco ainda
  não registrada).

## 2. O `qa_alvos_manual` prova cada item com oráculo independente?

O oráculo de N, Média e DP é independente:
- é o `calcular_alvos` do `aplicar_alvos_ciq.py`, em Python, sobre a `tblCQ_Final` lida pela COM;
- o filtro de PARTICIPA e de datas é feito no próprio Python.

A marca por campo (`cond_apos_aplicar`) é uma reimplementação da especificação. Vale como oráculo do texto do ADR, não
de uma regra externa.

| Lacuna | Situação na suíte | Coberto pelo repro |
|---|---|---|
| L1 Resultado inativado no período não entra | só implícito (se houver inativo na janela); nada é afirmado | X01 |
| L2 Dois ou mais equipamentos somados (Bio); L4 não interfere | não afirmado | X02 |
| L3 Data final inclusiva com resultado de hora ≠ 00:00 | implícito nas janelas A06–A08 | X03 (explícito: D, D−1, De = D, vazio) |
| L4 Valores estáticos depois de mudar lote, datas, analito, H3 e ATUALIZAR | não testado | X04 |
| L5 LJ e Westgard usam o alvo gravado | só `mLotes.MediaDoLote` (leitura da store) | X05b (Calc, séries) e X05c (1_3s × oráculo) |
| L6 Exceção só do DP seguida de Aplicar modo 3 | não testado (E08 é modo 1; E09 termina na exceção) | X06 |
| L7 Rebloqueio pela troca de analito (C3) e pelo spinner | só B98 (E07) | X07 |
| L8 Igual ao LI/LS = dentro | implícito (LI e LS escolhidos entre os valores) | X08 (IDs no limite sem observação) |
| L9 PENDENTE inativada vira RETIRADA | não testado (R03 só cobre a troca de limite) | X09 |
| L10 Inativado fora da faixa não gera observação | não testado | X10 |
| L11 DP = 0 (DP_ZERO) | não testado | não (precisa de dado sintético; fica para o backlog de testes) |
| L12 Bit TESTE do campo não transferido (modo 2/3) preservado | não testado | não |
| L13 Rebloqueio no Workbook_Open/BeforeClose | não testado | não (só estático: `EstaPastaDeTrabalho.cls:6`, `:82`) |
| L14 Hema "IDEAL" com 90 | só com `ALVO_N_IDEAL = 78` na cópia (dados sem N1 ≥ 90) | não (limitação de dados) |
| L15 Sessão com login (`Quem()` pela sessão) | a suíte roda sem login, então grava "(Windows)" | não |

Proposta de casos na suíte (pós-release): `qa_alvos_manual` A11 (DP_ZERO sintético), M06 (bit TESTE no modo 2/3) e
E11 (troca de analito tranca). Os cenários X01–X10 podem virar os casos A12–A21 se o Orquestrador quiser.

## 3. Script de reprodução

- **Script:** `_arquitetura/agentes/ferramentas/repro_e48_alvos.py`.
- **Validações já feitas sem Excel:** `py_compile` OK, e o oráculo foi testado em Python com dados sintéticos (apuração
  com inativo e datas; marcas e regra do N).
- **Como roda:**
  - copia a produção (`shutil.copy2`) para `%TEMP%\qcw\auditoria\e48\` e trabalha só nessa cópia;
  - abre um Excel (`qa_final.QA`, que imprime `EXCEL_PID`) e fecha só esse Excel;
  - nunca salva e não captura tela;
  - a senha vem só por import do `instalar_integracao`;
  - marca os dados sintéticos com `TESTE-AUDITORIA E48`;
  - confere o SHA-256 da origem no fim (Y00).

### Comando para a fase 2 (PowerShell; um produto de cada vez, só depois do fim do `entregar.py`)

```
New-Item -ItemType Directory -Force "$env:TEMP\qcw\auditoria\e48" | Out-Null
$env:PYTHONIOENCODING = 'utf-8'
python -u "<repo>\_arquitetura\agentes\ferramentas\repro_e48_alvos.py" Bioquimica "$env:TEMP\qcw\auditoria\e48\repro_bio.json" > "$env:TEMP\qcw\auditoria\e48\repro_bio.log"
python -u "<repo>\_arquitetura\agentes\ferramentas\repro_e48_alvos.py" Hematologia "$env:TEMP\qcw\auditoria\e48\repro_hema.json" > "$env:TEMP\qcw\auditoria\e48\repro_hema.log"
```

- O script usa por padrão `QC_RAIZ\QC_<produto>.xlsm` ou a raiz do repositório. Para outro arquivo: `--origem <arquivo>`.
- Saída: código 0 se não houver REPROVADO.
- Duração estimada: 15 a 30 min por produto (três ATUALIZAR só de entradas: X01, X04, X09).

## Estado dos testes nesta fase

| Teste | Estado |
|---|---|
| Revisão estática da lógica (tarefa 1) | APROVADO com ressalvas D-01 a D-03 (Baixo) |
| Avaliação da suíte (tarefa 2) | concluída; lacunas L1–L15 |
| `repro_e48_alvos.py` (Bio e Hema) | NAO_EXECUTADO: entrega em andamento usando o Excel (fase 2) |
| Regressão das suítes afetadas | NAO_EXECUTADO: idem; a lista é a do ADR-102 §Testes, a validar no gate do `entregar.py` |
| `qa_alvos_manual` Bio 35/0 e Hema 35/0 | resultado do implementador; não reproduzido pelo Tester |

## Fase 2 — execução em cópias da produção nova (25/11/2026, depois da troca)

`repro_e48_alvos.py`, uma cópia por produto, nunca salva; produção conferida pelo SHA-256 antes e depois (Y00).

| Produto | Produção (sha) | Resultado |
|---|---|---|
| Bioquímica | `0eca3fac…` | 14/14 APROVADO |
| Hematologia | `846b2bf4…` | 14/14 APROVADO |

Cenários: X00 pré-condições · X01 inativado no período não entra (= oráculo) · X02 todos os equipamentos · X03 datas
inclusivas · X04 valores estáticos (filtros e ATUALIZAR não mudam o alvo) · X05a "Aplicar" modo 1 = oráculo, condição pela
regra do N · X05b LJ usa o alvo gravado · X05c Westgard 1-3s usa o alvo gravado · X06 exceção só do DP e depois modo 3 ·
X07 trocar o analito tranca de novo; spinner sem evento recusa gravar · X08 resultado igual a LI/LS fica dentro ·
X09 pendente de resultado inativado vira RETIRADA · X10 inativado abaixo do LI não gera observação · Y00 produção intacta.

**Veredito final: VALIDADA** (as ressalvas D-01..D-03 da fase 1 são baixas e estão no BK-23).
