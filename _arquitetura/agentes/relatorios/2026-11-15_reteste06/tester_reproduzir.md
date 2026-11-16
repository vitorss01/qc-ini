# Tester/QA: modo REPRODUZIR, ciclo 2026-11-15_reteste06

Ocorrência reproduzida: **P-01** (Bio R6-01 e Hema R6-01). Apagar ou mover uma linha manual deixa USUARIO e
REGISTRADO_EM na linha. O próximo manual digitado nela herda esse carimbo, é recusado ("DATA/HORA posterior ao
lançamento") e fica atribuído ao usuário antigo.

## Ambiente

| Item | Valor |
|---|---|
| Produção copiada | `QC_Bioquimica.xlsm` sha256 `bd5314426ab8`; `QC_Hematologia.xlsm` sha256 `49f68fba99d5`. O sha foi conferido antes e depois e não mudou |
| Cópias | Uma cópia nova por produto, em `%TEMP%\qcw\auditoria\2026-11-15_reteste06\tester\` (`Bioquimica_P01_220600.xlsm`, `Hematologia_P01_*.xlsm`). Nenhuma foi salva, e as duas foram apagadas no fim |
| Excel | Um por vez, aberto e encerrado pelo `qa_final.QA`: PID 28380 (Bio) e PID 10840 (Hema). Não toquei nos PIDs 13084 e 8640, que são de outros processos e da tarefa agendada. Nenhuma captura de tela |
| Script | `_arquitetura/agentes/ferramentas/repro_p01_carimbo_residual.py <Bioquimica\|Hematologia>`. Escrevi o roteiro do zero e **não** reaproveitei `reteste06_carimbo_residual.py` nem `reteste06_usuario_f1*.py`. Da infraestrutura, usei só `qa_final.QA`, `repro_comum`, `Vigia` e `criar_usuario`/`fazer_login` |
| Saídas | `repro_P01_Bioquimica.json`/`.log` e `repro_P01_Hematologia.json`/`.log`, na pasta acima (fora do repositório) |
| Usuários de teste | `TESTE_AUDITORIA_R6T1` e `TESTE_AUDITORIA_R6T2`, os dois ANALISTA, criados só na cópia |
| Dados de teste | OBS "TESTE-AUDITORIA P01", MOTIVO "Outro", DATAs 14/11/2026 e 15/11/2026. Cada valor sintético tem 3 decimais: a média das 20 últimas corridas ativas do lote ± 0,001 a 0,004. **Bio**, lote 8977 no DIMENSION 1: Glicose N1 (77,251; 77,254; 77,249; 77,246), ALT (TGP) N2 (100,402; 100,398) e Glicose N2 (271,253; 271,247). **Hema**, lote 625311 no XN1000: WBC N1 (3,136; 3,139; 3,135; 3,131), BASO# N3 (0,842; 0,838) e WBC N2 (7,314; 7,308) |

## Roteiro próprio, mais amplo que o do Usuário Simulado

1. Login como T1. O botão Digitar Resultados (`mIntegracao.IrDigitarResultados`) leva a B7, a linha 1 da tabela.
2. T1 digita quatro manuais em L0..L3 (planilha 7..10), com a aba protegida e os eventos ligados.
3. Com a aba protegida:
   - L0, L2 e L3: seleciono ID…MOTIVO (A:K) e uso `ClearContents`, o equivalente da tecla Delete;
   - L1: recorto A:K e colo em A16, numa linha nunca usada.
4. Espero 75 s depois do carimbo de A.
5. Login como **T2, outro usuário**. Clico de novo no botão.
6. T2 digita:
   - **B** em L0: o mesmo analito e nível de A, com DATA/HORA = carimbo de A + 40 s;
   - **B2** em L1, a origem do recorte: **outro analito e outro nível**;
   - **D** em L2: DATA/HORA **anterior** ao carimbo, 15/11 06:00;
   - **C**, o controle, numa linha nunca usada (planilha 13): igual ao B, com a mesma DATA/HORA.

   L3 fica vazia.
7. Rodo o ATUALIZAR DADOS (`AtualizarDadosAutomatico`, o mesmo núcleo do botão).
8. Leio a tabela manual, a `tblCQ_Final`, a `tblQA_Integracao` e a `tblAuditoria`.

O oráculo é independente. Apliquei em Python a regra do motor (`mEtlCQ.bas:1104`, DATA+HORA > REGISTRADO_EM → recusa)
aos valores lidos da tela antes do ATUALIZAR, e comparei com o que o motor gravou. Nenhuma caixa de mensagem apareceu
na digitação nem no ATUALIZAR.

## Resultado: Bioquímica, **REPRODUZIDA**

| Linha (planilha) | Situação | USUARIO / REGISTRADO_EM depois de digitar | Oráculo | Motor (tblCQ_Final) | QA |
|---|---|---|---|---|---|
| L0 (7): A `MAN_0001` apagado | Depois do Delete | T1 / 15/11 22:06:31 (**resíduo**) | | | |
| L0 (7): B `MAN_0005`, Glicose N1, 15/11 22:07:11, digitado por **T2** | O botão levou o cursor a B7 | **T1 / 22:06:31** (herdado) | recusa | **MANUAL_INCOMPLETO**, PARTICIPA NÃO, NAO_PLOTAR, "DATA/HORA posterior ao lancamento", USUARIO_MANUAL **T1** | E04 |
| L1 (8): B2 `MAN_0006`, **ALT (TGP) N2**, 15/11 22:07:17, T2 | Origem do recorte | **T1 / 22:06:37** (herdado) | recusa | **MANUAL_INCOMPLETO**, igual ao B | E04 |
| L2 (9): D `MAN_0007`, Glicose N2, 15/11 06:00, T2 | Hora anterior ao carimbo | **T1 / 22:06:43** (herdado) | aceita | ATIVO, RUN 26111501, mas USUARIO_MANUAL **T1** | — |
| L3 (10): apagada e deixada vazia | Depois do ATUALIZAR | T1 / 22:06:48 **continua** | | | |
| Planilha 13: C `MAN_0008`, controle, igual ao B, T2 | Linha nunca usada | T2 / 22:08:07 | aceita | **ATIVO**, RUN 26111502, USUARIO_MANUAL T2 | — |
| Planilha 16: A2 `MAN_0002` movido | Destino do recorte | T1 / **22:06:58** (carimbo novo, no momento do recorte; o original era 22:06:37) | aceita | ATIVO, RUN 26111401 | — |

A mensagem do ATUALIZAR diz: "2 resultado(s) manual(is) incompleto(s): veja o DETALHE na aba Digitar Resultados". O
DETALHE na tela é "campos obrigatórios ausentes/inválidos: DATA/HORA posterior ao lançamento", e a DATA/HORA digitada
está correta.

Trilha (`tblAuditoria`, UsuarioSistema = T2, que rodou o ATUALIZAR):

- `RESULTADO_MANUAL_RECUSADO` ID MAN_0005: "… DATA/HORA posterior ao lancamento | manual lancado por
  **TESTE_AUDITORIA_R6T1 em 15/11/2026 22:06:31**". Quem digitou foi T2, às 22:07:5x;
- `RESULTADO_MANUAL_RECUSADO` ID MAN_0006: "… manual lancado por TESTE_AUDITORIA_R6T1 em 15/11/2026 22:06:37";
- `RESULTADO_MANUAL_INCLUIDO` ID MAN_0007: "manual lancado por **TESTE_AUDITORIA_R6T1** em 15/11/2026 22:06:43". Ele
  foi aceito, mas a autoria está errada;
- `RESULTADO_MANUAL_INCLUIDO` ID MAN_0008 (controle): "manual lancado por TESTE_AUDITORIA_R6T2 em 15/11/2026 22:08:07",
  que está correto.

## Resultado: Hematologia, **REPRODUZIDA**

| Linha (planilha) | USUARIO / REGISTRADO_EM | Motor | QA |
|---|---|---|---|
| L0 (7): B `MAN_0005`, WBC N1, 15/11 22:10:14, T2 (o botão levou a B7) | **T1 / 22:09:34** (herdado) | **MANUAL_INCOMPLETO** "DATA/HORA posterior ao lancamento", USUARIO_MANUAL T1 | E04 |
| L1 (8): B2 `MAN_0006`, **BASO# N3**, 15/11 22:10:18, T2 (origem do recorte) | **T1 / 22:09:38** (herdado) | **MANUAL_INCOMPLETO**, igual ao B | E04 |
| L2 (9): D `MAN_0007`, WBC N2, 15/11 06:00, T2 | **T1 / 22:09:42** (herdado) | ATIVO, RUN 26111501, USUARIO_MANUAL **T1** | — |
| L3 (10): vazia | T1 / 22:09:46 **continua depois do ATUALIZAR** | | |
| Planilha 13: C `MAN_0008`, controle, T2 | T2 / 22:11:08 | **ATIVO**, RUN 26111502 | — |
| Planilha 16: A2 `MAN_0002` movido | T1 / **22:09:54** (carimbo do recorte; o original era 22:09:38) | ATIVO, RUN 26111401 | — |

A trilha repete o padrão da Bio. MAN_0005 e MAN_0006 aparecem como RECUSADO "lancado por TESTE_AUDITORIA_R6T1 em
22:09:34/22:09:38". MAN_0007 aparece como INCLUIDO "lancado por TESTE_AUDITORIA_R6T1 em 22:09:42", embora T2 o tenha
digitado. O controle MAN_0008 sai certo, com T2.

## Ampliação

| Variação | Resultado |
|---|---|
| Outro produto | Acontece igual na Bio e na Hema, com o mesmo código (`src/bio` e `src/hema/mIntegracao.bas`, mesmas linhas) |
| Outro analito e nível | Glicose N1, ALT (TGP) N2 (Bio), WBC N1 e BASO# N3 (Hema): todos recusados. O defeito não depende do analito, do nível nem do lote. Ele depende só da linha |
| Mover (recortar A:K) | A **origem** fica com o resíduo, e o manual digitado nela é recusado (B2). O **destino** ganha um carimbo novo com a hora do recorte e o usuário logado nesse momento. Quando outra pessoa move, a autoria do manual movido passa a ser dela. Isso é um efeito colateral, menor que a P-01 |
| Outro usuário digita na linha apagada | O manual fica com o USUARIO de quem apagou (T1), tanto na tela quanto em `tblCQ_Final.USUARIO_MANUAL` e no Audit_Log. **Isso vale também quando o manual é aceito** (D, com DATA/HORA anterior ao carimbo). A falha de trilha atinge todo manual digitado numa linha apagada, recusado ou não |
| O ATUALIZAR limpa o resíduo? | Não. A L3 vazia continua com USUARIO e REGISTRADO_EM depois do ATUALIZAR: o `PrepararEntradas` não limpa |
| Colunas travadas | USUARIO, REGISTRADO_EM, RUN, STATUS e DETALHE têm `Locked=True`, e a aba fica protegida. ID…MOTIVO e SUBSTITUI_ID ficam livres. O usuário não consegue limpar o carimbo, e o botão leva exatamente à linha com resíduo |
| Não testado (NAO_EXECUTADO) | (a) Colar um manual movido **sobre** uma linha com resíduo. Pelo código, ele herdaria o carimbo da linha de destino e poderia ser recusado. (b) A base da entrega 42: o Usuário Simulado já mostrou que lá acontece o mesmo, e não repeti |

## Causa (leitura do código; não editei)

- `src/bio/mIntegracao.bas:800` (o mesmo vale para `src/hema` e `src/comum`): no `PrepararLinha` da TB_MANUAL, a linha
  sem dado e sem ID sai com `Exit Sub`, sem limpar USUARIO e REGISTRADO_EM.
- `src/bio/mIntegracao.bas:1709-1712` (`Carimbar`): só preenche o vazio. Por isso o manual novo herda o carimbo.
- `src/bio/mIntegracao.bas:1920` (`EntradaMudou`, ramo `ABA_MANUAL`): não há equivalente ao que o ramo `ABA_INAT` faz
  na linha 1895 (ADR-070), que limpa o carimbo quando o ID é apagado.
- `src/comum/mEtlCQ.bas:1096` e `1104`: o motor compara DATA/HORA com REGISTRADO_EM e recusa. A regra está correta,
  mas recebe o carimbo errado.

## Teste automatizado proposto

Proponho uma suíte nova, a da correção (`testes/qa_adr097.py`, ou o número do ADR que corrigir), entrando na lista
`SUITES` e em `SUITES_DO_LOTE` do `entregar.py`, nos dois produtos. Os casos:

| ID | Caso | Critério |
|---|---|---|
| CARIMBO-01 | Digitar A, apagar A:K | USUARIO e REGISTRADO_EM vazios na linha |
| CARIMBO-02 | Digitar B na linha apagada com DATA/HORA > carimbo de A, outro usuário, e ATUALIZAR | B tem carimbo próprio (usuário atual, hora ≥ início da digitação) e fica ATIVO, PARTICIPA SIM, sem E04 |
| CARIMBO-03 | Recortar A:K para outra linha | A origem fica sem carimbo, e o manual digitado nela fica ATIVO. Decisão pendente para o destino: preservar o carimbo original ou recarimbar |
| CARIMBO-04 | D com DATA/HORA anterior ao carimbo residual, digitado por outro usuário | USUARIO_MANUAL e "manual lancado por" mostram o usuário que digitou D |
| CARIMBO-05 | Linha com resíduo já existente (estado legado), depois PrepararEntradas/ATUALIZAR | Resíduo limpo |
| CARIMBO-06 | Regressão: editar o RESULTADO de um manual já carimbado | O carimbo **não** muda (a edição não pode recarimbar), e a trilha registra a edição como hoje |

Regressão a rodar na validação: `qa_final.py`, `qa_manual_substitui.py`, `qa_adr088.py`, `qa_adr090.py`,
`qa_adr095.py` e `qa_etl.py`. O `qa_etl.py` só entra se o `PrepararEntradas` mudar.
