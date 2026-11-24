# Tester/QA — Modo VALIDAR — Entrega 46 (24/11/2026)

Validador independente: não implementei nenhuma das mudanças. Código validado: commit 9c4021a, mais os testes cfbb9c7.
A produção já foi entregue. Todas as execuções rodaram em cópias novas da produção, em `%TEMP%\qcw\tester46\`, uma por
suíte, com um Excel por vez. Matei só os PIDs que os scripts criaram e não capturei a tela. As cópias foram descartadas
no fim.

| Produção | sha256 (início) | sha256 (fim) |
|---|---|---|
| QC_Bioquimica.xlsm | 9693db90d03f15e0… | 9693db90d03f15e0… (inalterado) |
| QC_Hematologia.xlsm | a8d6fcb20cbcb1a0… | a8d6fcb20cbcb1a0… (inalterado) |

Scripts próprios, em `_arquitetura/agentes/ferramentas/`:
- `repro_e46_liberacao.py`: cenário próprio da Liberação e integridade (T00–T06);
- `repro_e46_sigma.py`: oráculo próprio do ADR-101 (G01–G04);
- `diag_e46_liberstore.py`: diagnóstico do achado A-01.

## Vereditos

| # | Item | Veredito | Bio PASS/FAIL | Hema PASS/FAIL |
|---|---|---|---|---|
| 1 | ADR-099 (Material + avisos do CEQ) | **VALIDADA** | qa_adr099 8/0; qa_fluxo_ceq 13/0 + 1 ALERTA | qa_adr099 8/0; qa_fluxo_ceq 13/0 + 1 ALERTA |
| 2 | ADR-101 (Sigma em 6 faixas, piso 0, legenda V5:X8) | **VALIDADA** | qa_incerteza 15/0; qa_adr090 21/0; repro_e46_sigma 5/0 | qa_incerteza 16/0; qa_adr090 22/0 + 1 INFO; repro_e46_sigma 5/0 |
| 3 | P1 Liberação: a assinatura acompanha o RUN | **VALIDADA** (com o achado A-01, P2, anterior à entrega) | qa_liberacao_run 4/0; qa_fluxo_manual 19/0 + 1 NAO_EXECUTADO (F04) + 1 INFO; repro_e46_liberacao 8/0 | qa_liberacao_run 4/0; qa_fluxo_manual 19/0 + 1 NAO_EXECUTADO (F04) + 1 INFO; repro_e46_liberacao 8/0 no estado de sessão com login (3/5 em automação pura, ver A-01) |
| 4 | Integridade (status, proteção, sha) | **VALIDADA** | qa_seguranca 20/0; T00 PASS; sha inalterado | qa_seguranca 20/0; T00 PASS; sha inalterado |

### 1. ADR-099

- `qa_adr099` tem estes casos: M01, M02, M04, M05, M06, M07, M09 e Y00.
  - A suíte e o ADR não têm casos M03 nem M08: a numeração pula esses dois.
  - Todos passaram nos dois produtos.
- Os avisos do CEQ estão no `qa_fluxo_ceq`:
  - F10 (avaliação "Adequado" não reconhecida): PASS nos dois produtos. O texto da confirmação é o do ADR.
  - F11 (rodada de 2026 com o Ano EP em 2025): PASS nos dois. O G não muda e o aviso do Ano EP aparece.
- O ALERTA é o F12 (OC-0069, provedor Controllab num analito cujo bias vem do CAP). Ele já está no backlog como BK-06 e
  não é regressão.

### 2. ADR-101

O oráculo próprio é o `repro_e46_sigma`.

- **G01:** percorre todas as linhas de Estatística L14:L133 (Bio: 58 linhas com Sigma; Hema: 57).
  - Nenhum Sigma é menor que 0.
  - A classe M bate com a faixa calculada de forma independente em todas as linhas.
  - Faixas presentes na Bio: inaceitável 41, ruim 12, marginal 2, bom 3. Na Hema, as 6 faixas aparecem.
- **G02:** a `tblPlanoQC_Sigma` tem as 6 faixas. "Ruim" é 2–3 e "Inaceitável" é menor que 2.
- **G03:** testei um analito por faixa no Painel.
  - O valor de I é igual ao de Estatística L, com tolerância de 1e-9.
  - K mostra a classe da faixa.
  - A cor exibida (`DisplayFormat`) é a mesma em I e K, única por faixa e diferente entre faixas.
- **G04:** a legenda em V5:X8 está preenchida (7 células) e nenhum gráfico a cruza (Bio tem 2 gráficos, Hema tem 3).
  Há 6 regras de cor em Painel I7 e em Estatística L14.
- O piso max(0, …) contra o oráculo de Sigma é coberto pelo `qa_incerteza`, que passou em tudo.

### 3. Liberação pelo RUN

- **Suítes:**
  - `qa_liberacao_run` passou com 4/0 nos dois produtos (L01–L03 e Y00).
  - No `qa_fluxo_manual`, F05e e F07b passaram nos dois produtos.
  - O F04 (R_4s) ficou **NAO_EXECUTADO** nos dois. Os dados atuais não têm a pré-condição: nenhuma corrida C2 tem o N1
    com 1,2 ≤ |z| ≤ 2,0. Isso é dos dados, não do produto. Não declaro PASS do F04.
- **Cenário próprio (`repro_e46_liberacao`)**, que as suítes não cobrem:
  - A assinatura é real, pelo `mSeguranca.AssinarCom`, com o usuário local de teste QA_TESTER46 (ANALISTA). A senha é
    sintética e só existe na cópia.
    - Assinei INICIADO e FINALIZADO num RUN do meio da lista e INICIADO no 1º RUN.
    - Bio: RUN 25112502 na linha 104 e RUN 25100101 na linha 4.
    - Hema: RUN 25101601 na linha 32 e RUN 25100101 na linha 4.
  - **T02:** lancei **dois** manuais (MOTIVO TESTE-AUDITORIA, valor igual à mediana dos N1 ativos) em dois dias sem
    resultado do lote, antes do RUN assinado. Isso criou 2 RUNs novos.
    - A assinatura desceu 2 linhas (Bio de 104 para 106; Hema de 32 para 34) e continuou ao lado do mesmo RUN.
    - A assinatura do 1º RUN não se mexeu.
    - A trilha registrou 2 RESULTADO_MANUAL_INCLUIDO e nenhum LIBERACAO_ASSINATURA_SEM_CORRIDA.
  - **T03:** troquei o lote em uso para outro e voltei (`mApp.UsarLote`). As assinaturas voltaram ao lado dos mesmos
    RUNs, e nada do teste apareceu no outro lote.
  - **T04:** depois do `FlushLoteAtual` (o caminho do BeforeSave), cada assinatura tem o seu RUN na coluna E do bloco da
    LiberStore.
  - **T05:** retirei os manuais com Delete e rodei o ATUALIZAR.
    - A lista voltou a ser idêntica à inicial, e as assinaturas voltaram às linhas iniciais.
    - A trilha registrou 2 RESULTADO_MANUAL_RETIRADO.
  - **T06:** as linhas ASSINATURA_* do Audit_Log apontam exatamente os RUNs que estão ao lado das assinaturas na tela
    (Bio: {25112502: 2, 25100101: 1}; Hema: {25101601: 2, 25100101: 1}).
  - Resultado: Bio 8/0 em automação pura. Hema 8/0 no estado de uma sessão com login (`ReprotectAll`). Sem esse estado,
    a Hema dá 3/5: T03–T06 falham. Ver A-01.

### 4. Integridade

- **T00** foi lido na cópia como ela foi entregue, antes de qualquer ATUALIZAR.
  - Bio: ATIVO 36826, INATIVADO 11, IMPLAUSIVEL 44, LOTE_FORA_DO_CQ 59, LOTE_OUTRO_MATERIAL 30.
  - Hema: ATIVO 48546, DUPLICIDADE_ORIGEM 589, SEM_VALOR 93.
  - Todos esses status são documentados no `mEtlCQ`, no ADR-091 e no ADR-092.
  - As 11 linhas da aba Inativar da Bio estão INATIVADO na tblCQ_Final. A Hema tem 0 linhas na aba Inativar.
  - Nenhum resultado não-ATIVO participa da estatística.
- **qa_seguranca** passou com 20/0 nos dois produtos. No XML: 39 abas, só a Login visível, a lista de abas sem proteção
  está vazia e a estrutura está travada.
- **sha** da produção está inalterado no fim (tabela acima).

## Achado novo

**A-01 — P2 (moderado), latente, anterior à entrega 46**

**O que acontece.** Na troca do lote em uso, `mLotes.SalvarViewNoBloco` e `CarregarBlocoNaView` gravam na LiberStore sem
liberar a aba:
- arquivos `src/hema/mLotes.bas` e `src/bio/mLotes.bas`, cerca das linhas 636–675;
- não chamam `LiberarEscrita`;
- usam `On Error Resume Next`.

Se a LiberStore está protegida e sem `UserInterfaceOnly` na sessão, a gravação falha calada. Nesse caso, as assinaturas
do lote que sai somem: ao voltar para o lote, a Liberação aparece sem assinatura.

**Reprodução** (`diag_e46_liberstore.py`, sha das cópias intacto):

| Estado da sessão | Bio | Hema |
|---|---|---|
| Automação pura (como `qa_final` abre a cópia) | o marcador vai para a LiberStore (A:E) e volta | bloco vazio; marcador perdido ao voltar |
| Com `mSeguranca.ReprotectAll` (o que o login faz) | preserva | preserva |

A LiberStore está protegida nos dois arquivos, com as células travadas. Não identifiquei por que a Bio grava mesmo sem
`UserInterfaceOnly`.

**Por que P2 e não P1.**
- No uso real, o login aplica `UserInterfaceOnly` (LockApp/UnlockApp e `ReprotectAll`), e o ADM trabalha desprotegido.
  Nesses estados o caminho funciona nos dois produtos.
- A produção não tinha nenhuma assinatura na Liberação (ETAPA 6 da entrega).

**O risco** é qualquer caminho sem login: macro rodada por automação, ou proteção reaplicada sem `UserInterfaceOnly`
(por exemplo, se o BeforeSave ou o blindar reprotege e a sessão continua). Nesses casos a troca de lote apagaria
assinaturas sem aviso, o que é um problema de integridade da Liberação.

**Correção sugerida** (não editei `src/`): em `SalvarViewNoBloco` e `CarregarBlocoNaView`, envolver a escrita na
LiberStore (e na RegistrosStore) com `prot = LiberarEscrita(stL)` … `RestaurarProtecao stL, prot`, como o ADR-046 manda
("cada rotina que escreve trata a própria janela").

**Teste proposto:** um caso novo L04 no `qa_liberacao_run`: assinar, trocar o lote ida e volta com a LiberStore protegida
sem `UserInterfaceOnly` e conferir a assinatura. É o T03 do `repro_e46_liberacao` sem `REPRO_LOGIN`.

**Efeito colateral da entrega 46:** nenhum observado nas suítes rodadas. Os tempos estão dentro do normal: o
qa_fluxo_manual levou 275 s na Bio e 318 s na Hema; o qa_adr090, 238 s e 247 s.

## Evidência (logs e JSON fora do repositório, em `%TEMP%\qcw\tester46\`)

- Suítes: pastas `suites_bio/`, `suites_hema/`, `ceq_bio/` e `ceq_hema/`, um `.log` e um `.json` por suíte.
- Scripts próprios: pasta `own/`, com os logs `lib2_bio.log`, `lib_hema.log` (automação pura), `lib2_hema_login.log`,
  `sig_bio.log`, `sig_hema.log` e `diag_*_login{0,1}.log`.
- Os dados de teste ficaram só nas cópias: o usuário QA_TESTER46, os manuais TESTE-AUDITORIA e os marcadores. As cópias
  foram apagadas no fim.
