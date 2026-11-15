# Reteste 04 — Usuário Simulado — Hematologia, frente F1

Data: 14/11/2026 · Agente: AGENTE_USUARIO_SIMULADO · Modo: RETESTE (14 ocorrências) · Produção: entrega 41.

## Ambiente e método

| Item | Valor |
|---|---|
| Produção copiada (nunca aberta) | `QC_Hematologia.xlsm` sha256 curto **aab0a54f464e**, igual antes e depois das 2 execuções (`producao_intacta.igual = true`) |
| Cópias (novas, nenhuma salva) | `%TEMP%\qcw\auditoria\2026-11-14_reteste04\Hematologia_F1.xlsm` (fluxo principal) e `Hematologia_F1_comp.xlsm` (só leitura de tela: Painel no modo TODOS, Estatística, botões Consolidar) |
| Excel | um por vez, aberto e encerrado pelo `qa_final.QA`: PIDs 5424 e 1740. Antes de cada execução conferi que não havia outro Excel aberto. Nenhuma captura de tela; a caixa do `AbrirCEQ` foi lida pelo texto das janelas do meu PID |
| Script | `_arquitetura/agentes/ferramentas/reteste04_usuario_f1_hema.py` (modos: padrão e `--complemento`), documentado no cabeçalho. Reaproveita `qa_final.QA`, `repro_comum` e as funções do `valida_adr090_091.py` com produto Hematologia (sem o âmbar e os DIMENSION da Bio). Os oráculos novos são meus: janela de 180 corridas, vizinhos de período e lote, listas de validação |
| Saídas | `Hematologia_F1.json/.log`, `Hematologia_F1_comp.json/.log` na mesma pasta do `%TEMP%` |
| Usuário | `TESTE_AUDITORIA_R4`, papel ANALISTA, criado só na cópia; login pelo `mSeguranca.DoLogin` |
| Dados de teste | manuais MAN_0001..0005 (WBC N1, lote 625311, dia 09/10/2026) com MOTIVO "Outro: TESTE-AUDITORIA", valores 123,456 / 111,111 / −1 / "abc"; CEQ rodada "TESTE-AUDITORIA C-A 2025", amostras TA-31/55/71/81 (RBC, "Red Blood Cell Count"), arquivo TESTE-AUDITORIA.pdf |
| tblCQ_Final da cópia | 48.391 linhas, 47.709 participam; o ATUALIZAR não trouxe resultado novo (último recebido 13/11/2026 13:10:20) |

**Caminho usado.** Filtros do Painel pelas células (C3, H3, K3, K4, G3/G4) com os eventos ligados; na OC-0041 o Painel
foi protegido, como na tela. Manuais: navegação pelo botão (`mIntegracao.IrDigitarResultados`, que levou à
Digitar Resultados, célula B7), digitação célula a célula com os eventos ligados e **ATUALIZAR DADOS pelo núcleo do
botão** (`mIntegracao.AtualizarDadosAutomatico`, que lê o DB_SEAC só para leitura; 31,9 s e 27,3 s). CEQ: o núcleo que
os botões da janela chamam (`mCEQForm.SalvarCEQ`, `CorrigirCEQ`, `ExcluirCEQ`, `ConfirmacoesCEQ`, `AvisoAnalitoCEQ`,
`DescricaoCEQ`, `RegistrosDaRodada`, `DistintosCEQ`) e o `mCEQForm.AbrirCEQ` real. Edição direta da EQA.CAP_Dados com a
aba liberada, como o administrador.

**Não executado por automação (sem captura de tela e sem clique em caixa):** o clique "Não"/"Sim" na caixa "Gravar mesmo
assim?" do frmCEQ e a digitação física numa célula com validação de dados. Na validação usei a avaliação do próprio
Excel (`Range.Validation.Value`), que é o critério que a caixa "Parar" aplica.

## Resultado por ocorrência

| OC | Estado | Evidência (cópia `Hematologia_F1.xlsm`, salvo indicação) |
|---|---|---|
| OC-0012 | **EFEITO_COLATERAL** (o defeito original está resolvido; ver R4-H01) | WBC, H3 vazio (todos os lotes), K3 2026 + "Ano inteiro": `Eng_Saida!M1` = 278, `AF1` = "N1 374 · N2 400 · N3 340" = contagem independente do período; A39 "⚠ 278 corrida(s) antigas do período fora do gráfico **e dos indicadores do Painel** (máx. 180). Período inteiro: N1 374 · N2 400 · N3 340 — veja a aba Estatística." Janela do Painel (engPainel) n = 120/158/112. Oráculo próprio, conferido no fim: no modo TODOS a corrida do motor é o par (RUN, LOTE) com resultado que participa (ADR-080). Com essa definição o oráculo dá 278 cortadas e uma janela de 120/158/112, **igual**. Contando só por RUN dá 272 e 121/159/113: um RUN com dois lotes conta como duas corridas. Vizinhos: lote 528211 em "Tudo" (89 corridas): M1 0, sem aviso, n 83/86/81 = independente = Estatística; mês Nov/2025 do mesmo lote: M1 0, n 45/49/45 = independente. **Efeito colateral:** o texto novo manda olhar a aba Estatística, mas na Hema ela mostra outro recorte (lote 522611, anos 2025–2025, n 50/51/51). Os números 374/400/340 não aparecem em lugar nenhum da Estatística (R4-H01) |
| OC-0021 | **PERSISTE** (parcial) | RUN vazio (Text "") nos MANUAL_INCOMPLETO (MAN_0001, 0003, 0005) e no CONFLITO_MANUAL (MAN_0002); DETALHE "" no ATIVO (MAN_0001 depois de completar o NÍVEL, RUN 26100901); nenhum "0" em RUN/DETALHE nas 300 linhas; DETALHE acentuado em "campos obrigatórios ausentes/inválidos: NIVEL/HORA/RESULTADO". **Continua sem acento** o DETALHE do IMPLAUSIVEL (MAN_0004, −1): "resultado <= 0: implausivel para material de controle (fora da estatistica e do grafico)". Mesmo resultado da Bio |
| OC-0023 | **RESOLVIDA** | 1º ATUALIZAR: sem NÍVEL (MAN_0001), sem HORA (MAN_0003) e RESULTADO "abc" (MAN_0005) → só RESULTADO_MANUAL_RECUSADO, com o campo no detalhe; conflito +5 min (MAN_0002, "coincide com resultado do interfaceamento (4 min)", real HEM-364445 às 10:32:44) → RESULTADO_MANUAL_EM_CONFLITO. Completar o NÍVEL e atualizar → ATIVO, PARTICIPA SIM, RUN 26100901, **uma** RESULTADO_MANUAL_INCLUIDO. Vizinhos: −1 (MAN_0004) → RESULTADO_MANUAL_FORA_DA_ESTATISTICA (ADR-092); o 2º ATUALIZAR não repetiu a trilha dos outros manuais |
| OC-0024 | **RESOLVIDA** | Mensagem do ATUALIZAR e ResumoQA: "QA da integração: 3 erro(s), 591 alerta(s), 34 informativo(s)." + "• 3 resultado(s) manual(is) incompleto(s): veja o DETALHE na aba Digitar Resultados — não entraram na estatística." Os 3 ERRO são E04 (MAN_0001/0003/0005); nenhuma linha de governança, nenhum "ERRO DE GOVERNANÇA" (178 caracteres). Depois de completar o MAN_0001: "2 erro(s)" e "• 2 resultado(s) manual(is) incompleto(s)…" |
| OC-0031 | **RESOLVIDA** (função; clique na caixa não executado) | `ConfirmacoesCEQ` sem alvo: "Sem o valor-alvo, esta amostra NÃO entra no bias (nem no ET e no Sigma). Ela conta só na verificação dos limites."; com alvo: "". `SalvarCEQ` sem alvo: "OK\|467\|Gravado fora do bias (sem valor-alvo)."; G (bias) de RBC N1/N2/N3 1,533620125276759 → **igual bit a bit**; com alvo 100 → 3,650215093957569 (o mesmo do reteste 03). Rótulos "Valor-alvo ¹" e nota no `lblSub` |
| OC-0033 | **RESOLVIDA** | Arquivo recém-aberto: `SalvarCEQ` "ERRO\|0\|Sessão encerrada: faça o login de novo. Nada foi gravado.", 465 → 465 linhas usadas, nenhuma trilha; `AbrirCEQ` → caixa "Controle externo (CEQ)" e a janela não abriu (no início só li o botão "OK"; na repetição depois do `Logout` li o texto "Faça o login para registrar resultados do CEQ."). Vizinho (depois do `Logout`): Salvar, **Corrigir** e **Excluir** devolvem o mesmo ERRO\|0, F468 continua 108, linha não apagada, nenhuma trilha CEQ_* |
| OC-0034 | **RESOLVIDA** | `DistintosCEQ` (CAP) = conjunto das linhas reais: rodada 3, analito 25, avaliação 3. Ficaram fora "1/2/3", "RBC/WBC/HGB/HCT" e "OK", que só existem na simulação. Controllab vazio |
| OC-0035 | **RESOLVIDA** (texto) | Li o GUIA como usuário. O §3 cita o EQUIPAMENTO e o âmbar. O §5–6 cita a janela "Registrar resultado do CEQ", o login, o valor-alvo, "Salvar e próximo", "Corrigir/Excluir selecionado", "não digite nelas", o Z3, o Consolidar, o ATUALIZAR, a FÓRMULA SOBRESCRITA e "o R4 vale para todos os analitos". Nenhuma instrução manda digitar nas abas. Todos os botões citados existem: "Registrar resultado do CEQ" nos dois Resumos; "Consolidar na EQA_Base" nas duas abas de dados (macro `AtualizarEQABase`), texto que bate com o Z3; ATUALIZAR no Painel, DB, DB_ORGANIZADO e Digitar Resultados; e os botões do frmCEQ |
| OC-0036 | **RESOLVIDA** | 84 textos "x,xx / y,yy" na AT; AT14 "2,49 / 0,04" (o valor do critério do parecer), AT17 "1,73 / 0,02"; nenhum "004 / 000" |
| OC-0041 | **RESOLVIDA** | Painel protegido. K3 2026 + K4 S1 → digitar De/Até 01/09–31/10/2026 → K4 "Personalizado", K3 2026 → escolher K4 "Out" → G3/G4 01/10–31/10/2026. Vizinho: De/Até 01/12/2025–31/01/2026 → K4 "Personalizado", K3 vazio, nenhum "Ano 0" na Estatística; K3 2026 + "Ano inteiro" → 01/01–31/12/2026 |
| OC-0050 | **RESOLVIDA** | Principal A5 "Total: 48.391 · participam: 47.709 · X vermelho: 0 · inativados: 0 · manuais: 0 · substituídos por manual: 0". Contei na tblCQ_Final da cópia (openpyxl, fora do Excel): 48.391 linhas e 47.709 SIM |
| OC-0052 | **RESOLVIDA** | Mapa real do CAP da Hema (29 nomes): um par por prefixo literal ("MCH" → "MCHC"). `AvisoAnalitoCEQ('CAP','MCH')` = "AZUL\|No QC_INI: MCH. Este resultado entra no bias deste analito. Atenção: existe também “MCHC” (MCHC)." O par de palavra continua: "Lymphs" → “Lymphs Absolute” (LYMPH#) |
| OC-0054 | **RESOLVIDA** | EQA.CAP_Dados linha 2 (Basophils → BASO%, 4,8), seleção em AB1. Escrever o mesmo valor: nenhuma trilha, Z3 "base em dia". Escrever 5,8: uma CEQ_ALTERADO "Seu Resultado: 4,8 → 5,8", sem "?", e Z3 "⚠ HÁ EDIÇÃO NÃO CONSOLIDADA — clique em Consolidar na EQA_Base". Vizinho com a célula selecionada antes (linha 468 de teste): "Seu Resultado: 104 → 106" e o mesmo Z3. OC-0053 de carona: colar 2,55 em P2 → Z3 "FÓRMULA SOBRESCRITA" e CEQ_FORMULA_SOBRESCRITA; o G de BASO% ficou 2,5085549222470034 nos três momentos |
| OC-0055 | **RESOLVIDA** | Rodada "TESTE-AUDITORIA C-A 2025", ano 2025: descrição "CAP TESTE-AUDITORIA C-A 2025 · Red Blood Cell Count · TA-55 (resultado 110)", lista da rodada e trilha (CEQ_REGISTRADO, CEQ_EXCLUIDO) sem "2025 2025". Vizinho: a mensagem de duplicidade "Já existe este registro (CAP TESTE-AUDITORIA C-A 2025 · Red Blood Cell Count · TA-71) na linha 468. Para corrigir, selecione-o em “Registros desta rodada” e clique em Corrigir selecionado." também sem repetição |

Também conferi, de carona, a OC-0049 (que não está nesta lista): as 36 linhas sem analito da Estatística (98 a 133)
não mostram texto de A a AU.

## Ocorrências novas

### R4-H01 — Painel (modo TODOS, Hema): o aviso da janela de 180 corridas manda olhar a Estatística, mas ela mostra outro recorte

- **Prioridade:** Médio · **Tipo:** usabilidade · **Tela/módulo:** Painel A39 (texto do ADR-090, `mEstatistica`
  comum) × aba Estatística da Hema.
- **Ação:** Painel com WBC, H3 vazio (todos os lotes), K3 2026 + "Ano inteiro"; ler o aviso e abrir a aba
  Estatística, como o aviso manda.
- **Esperado:** achar na Estatística o n do período inteiro que o aviso cita (N1 374 · N2 400 · N3 340), como acontece
  na Bio (Glicose 8974: 174/161 nos dois lugares).
- **Observado (`Hematologia_F1_comp.json`, `tela`):**
  - O Painel mostra n 120/158/112, sem Média, DP, CV e ET (ADR-086: o modo TODOS não mistura lotes), e mostra Sigma
    2,48/3,68/5,56 com Status REJEITADO nos três níveis.
  - O A39 diz "Período inteiro: N1 374 · N2 400 · N3 340 — veja a aba Estatística".
  - A Estatística da Hema tem filtro próprio: Ano De/Até B3:D3 = 2025/2025 e lote B4 = 522611 "(refletido do Painel)",
    com o H3 do Painel vazio. Ela mostra WBC n 50/51/51.
  - Nenhum lugar da Estatística mostra 374/400/340 nem o período 2026.
- **Dificuldade:** o usuário segue a instrução e acha outro número, de outro ano e de um lote só.
- **Consequência:** o aviso, que deveria dar transparência ao corte, aponta para um número que não confere. Isso gera
  dúvida sobre qual estatística vale, justamente quando o Painel mostra REJEITADO.
- **Sugestão:** na Hema, trocar o fim do texto por algo que se verifica ("…escolha um lote no H3 para ver a estatística
  do período inteiro" ou "a aba Estatística é por lote e ano") ou omitir o "veja a aba Estatística" quando o H3 está
  vazio. Decisão de texto do Orquestrador; o número AF1 está certo (conferido).
- **Evidência:** `Hematologia_F1.json` (`O12.original_WBC_TODOS_2026`, `O12.estatistica_WBC`), `Hematologia_F1_comp.json`
  (`tela.painel_A6_J9_texto`, `tela.estat_cabecalho`, `tela.estat_WBC`).
- **Relação:** o defeito da OC-0012 (aviso só do gráfico) não ocorre mais; este é um efeito do texto novo, só na Hema.

### R4-01 (extensão para a Hema) — Digitar Resultados: a lista do MOTIVO tem um item só; cada motivo digitado é recusado

- **Prioridade:** Alto · **Tipo:** funcional · **Tela/módulo:** aba Digitar Resultados (tblResultados_Manuais),
  validação de dados criada por `camada_dados.validacoes` (`lista('MOTIVO', ','.join(MOTIVOS_MANUAL), …)`). É a mesma
  causa do R4-01 da Bio, que lá ficou como "não testada na Hema".
- **Ação:** escolher ou digitar o MOTIVO de um manual. O GUIA §3 manda preencher o motivo e diz que ele é
  **obrigatório para substituir** (ADR-082).
- **Esperado:** a lista oferece os 4 motivos e aceita qualquer um.
- **Observado (`Hematologia_F1.json`, `listas_validacao`):**
  - `Formula1` = "Falha do interfaceamento,Resultado não transmitido,Equipamento em contingência,Outro". O tipo é
    lista, o estilo é **Parar** e a mensagem é "Escolha o motivo da lista.".
  - `Validation.Value` dá Falso para "Falha do interfaceamento", "Resultado não transmitido", "Equipamento em
    contingência" e "Outro". Só o texto inteiro dá Verdadeiro.
  - Na Hema, EQUIPAMENTO ("XN1000") e MATRIZ ("SANGUE TOTAL") estão certos porque têm um item só, sem vírgula.
  - LOTE e ANALITO (`=lstLotes`, `=lstAnalitos`) e NÍVEL (1 a 3) estão certos.
- **Dificuldade:** o usuário digita o motivo e recebe "Escolha o motivo da lista."; a seta oferece uma única opção com
  os quatro motivos colados.
- **Consequência:** o motivo da substituição, que é a justificativa de auditoria do ADR-082, fica impossível de
  registrar certo pela tela. Quem usa o menu grava um texto sem sentido no Audit_Log e no I04. Não muda estatística.
- **Sugestão:** a mesma do R4-01. Criar a lista com o separador do idioma ou, melhor, apontá-la para um intervalo/nome,
  e reinstalar nas duas produções.
- **Evidência:** `Hematologia_F1.json` (`listas_validacao.MOTIVO`); fonte `camada_dados.py:537`.
- **Teste de digitação física:** NAO_EXECUTADO (sem interação com a tela); a avaliação é a do próprio Excel.

## Dúvidas de regra de negócio

1. **OC-0021:** como na Bio, o ADR-090 acentua o DETALHE por uma lista fixa de palavras. O texto do IMPLAUSIVEL
   (ADR-088) escapou. Os textos da aba QA_INTEGRACAO ("Resultado manual com campo obrigatorio ausente", "campos
   obrigatorios ausentes/invalidos") e da trilha continuam sem acento de propósito (paridade VBA × PQ). A pergunta é a
   mesma da Bio: acentuar na origem ou manter a lista atualizada a cada texto novo?
2. **Âmbar na Hema:** a instrução A3 da Digitar Resultados (texto comum aos dois produtos) diz "EQUIPAMENTO vazio
   recebe o do Painel (fica em âmbar): confira". Na Hema o XN1000 entra com fundo branco e sem nota, por decisão do
   ADR-090 (um equipamento só). O usuário da Hema lê uma promessa que não acontece. Vale um texto por produto?
3. **Manual implausível e o RUN do dia:** o manual −1 digitado às 03:29 de 09/10 recebeu o RUN 26100901. Antes, o dia
   tinha as corridas 01 (08:23), 02 (10:32) e 03 (13:29). O IMPLAUSIVEL conta como corrida real (ADR-088,
   `EhCorridaReal`), então um erro de digitação abre uma corrida e renumera as do dia, mesmo ficando fora da
   estatística, do gráfico e do Westgard. Isso é o desejado? (Não medi a renumeração dos reais; a regra está no
   código, `mEtlCQ.bas:958`.)
4. **Trilha CEQ:** o Z3 diz "clique em Consolidar na EQA_Base", e o botão, nas abas de dados, chama-se "Consolidar na
   EQA_Base". Confere. Já o GUIA §5 diz "o botão Consolidar da EQA_Base", o que sugere um botão na aba EQA_Base, onde
   ele não existe. Ajustar para "o botão Consolidar na EQA_Base (nesta mesma aba)"?

## Impressões de uso (Relatório A)

- **Janela de 180 corridas (TODOS):**
  - O aviso agora diz que o corte vale para o gráfico **e** para os indicadores, e dá o n do período inteiro, que
    confere.
  - Na Hema o ponteiro "veja a aba Estatística" não leva ao número (R4-H01).
  - O Painel no modo TODOS mostra n e Sigma, mas não Média, DP, CV e ET, e mesmo assim dá REJEITADO nos três níveis.
    Para quem não conhece o ADR-086, fica a dúvida de onde vem o REJEITADO.
- **Manuais:**
  - O caminho pelo botão leva direto à linha livre.
  - O STATUS e o DETALHE dizem exatamente o campo que falta, e o conflito diz "(4 min)".
  - O RUN em branco nos não incluídos ficou limpo.
  - A mensagem do ATUALIZAR separa os manuais incompletos dos erros de governança e diz o que fazer.
  - Sobra o texto sem acento do IMPLAUSIVEL ao lado dos acentuados.
  - Sobra também a lista do MOTIVO, que não aceita nenhum motivo (R4-01 na Hema).
- **CEQ:**
  - Sem login, a janela nem abre e a caixa diz por quê.
  - As listas mostram só o que é real (CAP 3 rodadas, 25 analitos).
  - O aviso do MCH cita o MCHC.
  - A descrição e a duplicidade não repetem o ano e dizem a linha e o botão.
  - O sem-alvo avisa a consequência antes de gravar e não mexe no bias.
- **Painel K3/K4:** "Personalizado" depois de digitar De/Até e K3 vazio ao atravessar o ano. Fica claro.
- **Formatos:** AT "2,49 / 0,04" e A5 "Total: 48.391 · participam: 47.709" estão legíveis e conferem.

## Arquivos

- Script: `_arquitetura/agentes/ferramentas/reteste04_usuario_f1_hema.py`.
- Saídas (fora do repositório): `%TEMP%\qcw\auditoria\2026-11-14_reteste04\Hematologia_F1{,_comp}.{json,log}`.
- O Audit_Log da cópia grava o nome, o login do Windows e a máquina de quem roda; esses dados foram omitidos deste
  relatório.
