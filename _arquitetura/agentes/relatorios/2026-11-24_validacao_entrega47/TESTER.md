# Tester/QA — modo VALIDAR — entrega 47 (commit c362922) — 24–25/11/2026

Validação independente. A correção A-01 e o ajuste do `instalar_adr101.py` não foram feitos pelo Tester; quem corrigiu
foi o Orquestrador (commit c362922). Não alterei `src/` nem os instaladores.

## Ambiente e integridade

- Produção conferida no início e no fim, sem abrir: `QC_Bioquimica.xlsm` sha256 `2752646953b22099…` e
  `QC_Hematologia.xlsm` `53fe8cc8ea48efad…`, **iguais nos dois momentos**. A trava `~$QC_Bioquimica.xlsm` de 17/11 continua
  lá, intocada.
- Todas as execuções rodaram em cópias novas em `%TEMP%\qcw\tester47\`, uma por vez. Cada script matou só o Excel que
  ele mesmo abriu (os EXCEL_PID estão nos logs). Os Excel 17088 e 18976 do usuário e o da tarefa agendada não foram tocados.
- No fim, apaguei as cópias `.xlsm` (29, inclusive o backup que a ferramenta do ADR-100 criou na cópia e as cópias de
  trabalho dela em `%TEMP%\qcw\alvos\`). Os logs e JSON ficaram em `%TEMP%\qcw\tester47\`. Os CSV/JSON da ferramenta
  ficaram em `%TEMP%\qcw\alvos\`: são alvos calculados, sem dados de paciente, e estão fora do repositório.
- Script novo do Tester: `_arquitetura/agentes/ferramentas/repro_e47_troca_lote.py`.

## A. A-01: a troca de lote não apaga as assinaturas — **VALIDADA**

| Execução | Hematologia | Bioquímica | Log |
|---|---|---|---|
| `repro_e46_liberacao.py` sem REPRO_LOGIN (o estado que falhava) | **8/8 PASS** (T03: assinaturas iguais, linhas iguais) | **8/8 PASS** | `A_<prod>_repro_semlogin.log/.json` |
| `repro_e46_liberacao.py` com REPRO_LOGIN=1 | **8/8 PASS** | **8/8 PASS** | `A_<prod>_repro_login.log/.json` |
| `diag_e46_liberstore.py` sem DIAG_LOGIN | LiberStore protegida e **sem** UserInterfaceOnly. O marcador foi para o bloco (linha 2, col. E = RUN 25100101) e voltou na Liberação A4 | idem | `A_<prod>_diag_semlogin.log` |

Na entrega 46, a Hema sem login fazia 4/8. Agora são 8/8 nos dois estados e nos dois produtos. O critério foi atendido.

**Cenário próprio** (`repro_e47_troca_lote.py`, sem login, abas protegidas sem UserInterfaceOnly). O roteiro foi:
marcadores sintéticos `QA47`/`TESTE-AUDITORIA` em C/D; **duas trocas seguidas para lotes diferentes** L0→L1→L2→L0 e
volta ao L1; Save da cópia pelo caminho do `Workbook_BeforeSave` (FlushLoteAtual → TrancarParaGravar, que reprotege com
ReprotectAll → Save → RestaurarAposGravar); nova troca na mesma sessão; fechar; **reabrir a cópia salva** e trocar de novo.

| Caso | Hema (L0 522611, L1 528211, L2 533811) | Bio (L0 8974, L1 8588, L2 8352) |
|---|---|---|
| P01 L0→L1→L2→L0: marcadores do L0 ao lado dos mesmos RUNs; nada no L2 | PASS | PASS |
| P02 volta ao L1: o marcador do L1 no mesmo RUN | PASS | PASS |
| P03 ProtectContents=True em LiberStore/Liberação/RegistrosStore depois de **cada** troca | PASS (5 trocas) | PASS (6 trocas) |
| P04 Save (sequência do BeforeSave) + L0→L1→L0 na mesma sessão | PASS | PASS |
| P05 cópia reaberta: L0 com os marcadores ao abrir; L0→L1→L0 intactos; proteção ativa | PASS | PASS |
| P06 bloco do L0 na LiberStore com os marcadores e o RUN na coluna E | PASS | PASS |

Logs: `P2_<prod>.log/.json`. As proteções ficaram ativas o tempo todo: LiberStore e Liberação com ProtectContents=True
(com UserInterfaceOnly depois da troca) e RegistrosStore com ProtectContents=True. Na Hema e na Bio a RegistrosStore
não é mais gravada (ADR-087).

**Correção do oráculo depois de FAIL (declarada).** A 1ª rodada (`P_<prod>.log`) deu 4 PASS / 2 FAIL (P05, P06) nos dois
produtos. A causa estava no harness, não no A-01: o `wb.Save()` pela COM **com eventos ligados** não chegou ao disco. O
sha da cópia ficou igual ao da produção, e por isso a reabertura via o arquivo original. Na 2ª rodada
(`REPRO_SAVE=manual`), o script chama a mesma sequência do BeforeSave (`EstaPastaDeTrabalho.cls` l.23-47) passo a passo.
Ela gravou no disco (`gravou no disco: True` duas vezes) e deu 6/6 nos dois produtos. Ver a observação O-1.

## B. `instalar_adr101.py` reinstalável — **VALIDADA**

Fiz uma cópia nova da produção para cada produto e tirei uma foto antes, depois da 1ª instalação e depois da 2ª
(`snapB.py`, só leitura). A foto cobre: Painel V5:X12 (fórmula/valor, cor do fundo e da fonte, negrito, tamanho,
ShrinkToFit, Wrap, padrão), as larguras de V:X, K6:K9, a formatação condicional de Painel I7/K7/I9 e de Estatística
L14/M14/L133 (tipo, fórmula, AppliesTo, cores, negrito, StopIfTrue), as fórmulas de Estatística L14:L133, a
tblPlanoQC_Sigma (cabeçalho, linhas, endereço) e a proteção e visibilidade das 39 abas.

| | Hema | Bio |
|---|---|---|
| 1ª instalação | rc 0; fumaça `ok: Painel I7 = '2,60' 'Ruim'`; salvo | rc 0; fumaça `ok: '1,85' 'Inaceitável'`; salvo |
| 2ª instalação na mesma cópia | rc 0; fumaça OK; salvo | rc 0; fumaça OK; salvo |
| Foto antes × depois da 1ª e × depois da 2ª | **0 diferenças** | **0 diferenças** |
| Sanidade da foto | 6 regras em I7, K7, L14 e M14; plano com as 6 faixas; L com MAX(0;…); 39/39 abas protegidas; estrutura travada | idem |

As duas instalações registram `MUDANCAS 1: Painel I/K cores e legenda $V$5:$X$8`, porque a legenda é apagada e
reescrita. O resultado é idêntico ao da produção. O sha muda a cada instalação porque o VBA é reinstalado e o arquivo é
salvo de novo: isso é esperado. Logs: `B_<prod>_inst{1,2}.log`, `B_<prod>_foto{0,1,2}.json`, `B_<prod>_resumo.json`.

## C. Ferramenta do ADR-100 (`aplicar_alvos_ciq.py`) — **VALIDADA** (em cópias; produção não tocada)

A ferramenta aceita a cópia explicitamente (`--arquivo <cópia>`, `--so-copia <saída>`). Por isso executei a aplicação
real **só em cópias**. Nenhuma execução sem `--simular` apontou para a produção.

| Verificação | Resultado | Evidência |
|---|---|---|
| Autoteste do cálculo (sem Excel) | 19/19 PASS | `C1_autoteste.log` |
| Recusa com trava `~$` ao lado da entrada (sem `--simular`) | rc 1, Y00 FAIL, recusou **antes de abrir o Excel**, sha igual | `C2_trava.log` |
| `--simular` Hema / Bio | rc 0; nada gravado (sha igual); Hema: 756 pares "aplicado"; Bio: 126 "aplicado" e 131 "insuficiente" | `C3_simular_Hema.log`, `C7_simular_Bio.log` |
| `--so-copia` (Hema) | rc 0, 14 PASS (A00–A11, Y00, F00); A02 com 252 linhas e 0 divergência; A05 com 42 conferidos e 0 divergência; a entrada **não** foi trocada (sha igual); saída com sha próprio | `C4_socopia_Hema.log` |
| Modo troca sobre uma **cópia** (sem flag, `--arquivo` cópia) | rc 0; backup `_backup_pre_alvos_<stamp>/` com sha = original (`53fe8cc8…`); entrada trocada (`→ 223da2d3…`) e sha conferido (A11) | `C5_troca_Hema.log` |
| Reaplicar sem `--forcar` | rc 1, A00 FAIL "ja aplicado neste arquivo (marca blAlvosCIQ)"; nada gravado | `C6_reaplicar_sem_forcar.log` |

Leitura do código:
- Y00/F00 conferem o sha.
- O Excel nunca abre a entrada: ele trabalha numa cópia em `%TEMP%\qcw\alvos`.
- É tudo ou nada: qualquer FAIL antes do Save leva a `Abortar` e nada é salvo nem trocado. A08b exige o mesmo alvo
  depois do ATUALIZAR.
- A trava `~$` é conferida no início e de novo imediatamente antes da troca. O backup é feito antes da troca.

Observações P3 (backlog, não bloqueiam):
- O sha do backup não é conferido contra h0.
- `os.makedirs(destino_rel)` sem `exist_ok` falha se houver duas aplicações no mesmo minuto. Isso é seguro: falha antes
  da troca.
- **Operacional:** a trava antiga `~$QC_Bioquimica.xlsm` de 17/11 fará a ferramenta **recusar** a aplicação real na Bio
  de produção (Y00). O usuário precisará removê-la, com o Excel da Bio fechado, antes de aplicar o ADR-100 na Bio.

## D. Regressão mínima — **VALIDADA**

`testes/rodar_suites.py --suites qa_liberacao_run.py,qa_seguranca.py`, cada suíte numa cópia nova:

| Suíte | Hema | Bio |
|---|---|---|
| qa_liberacao_run | 4/4 PASS (158 s) | 4/4 PASS (97 s) |
| qa_seguranca | 20/20 PASS (117 s); XML travado | 20/20 PASS (120 s); XML travado |

Logs: `D_<prod>/qa_liberacao_run.log`, `D_<prod>/qa_seguranca.log` (+ .json).

## Observação nova

**O-1 — P2, a confirmar.** Pela COM, `Workbook.Save()` com eventos ligados **não grava no disco**, e isso ocorre com e
sem sessão.
- `diag_save.py`, cópia Hema sem sessão: sha igual depois do Save com eventos e `Saved=False`. Sem eventos, o sha mudou.
- `diag_save_login.py`, com login ANALISTA de teste: sha igual depois do Save com eventos, mas **`Saved=True`**.
- A causa provável está em `src/comum/EstaPastaDeTrabalho.cls` l.32-47. O BeforeSave cancela (`Cancel = True`) e chama
  `Me.Save` aninhado. Sob automação, esse Save interno não grava. Mesmo assim, `RestaurarAposGravar`
  (`src/bio/mSeguranca.bas` l.819) marca `ThisWorkbook.Saved = True` ("o disco já tem tudo").
- Efeito no QA: o S19 do `qa_seguranca` passa com `saved: true` sem que o arquivo tenha mudado. O S20 confere o XML de
  um arquivo que já vinha travado. O oráculo não prova a gravação.
- Não sei se o Ctrl+S manual do usuário tem o mesmo comportamento. O ADR-065 diz que isso foi medido, e eu não posso
  testar a interface sem capturar a tela. Se tiver, o usuário perde o que fez sem aviso (seria P0).
- O defeito não foi introduzido pelo c362922 (não toca no BeforeSave) e não afeta os instaladores nem as ferramentas,
  que salvam com eventos desligados.
- Recomendação: o Orquestrador confirma num Ctrl+S manual numa cópia (comparar a data de gravação ou o sha antes e
  depois). Em seguida, o S19 passa a exigir que o sha ou a data de gravação mude. Suíte `qa_seguranca`, caso novo
  **S19b**.

## Efeitos colaterais da entrega 47

Nenhum encontrado:
- A Liberação, a LiberStore e a RegistrosStore continuam protegidas depois das trocas.
- A integridade da cópia (T00) passa nos dois produtos.
- A segurança passa 20/20.
