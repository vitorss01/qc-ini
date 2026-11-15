# Reteste 04 — Usuário Simulado — Bioquímica, frente F1

Data: 14/11/2026 · Agente: AGENTE_USUARIO_SIMULADO · Modo: RETESTE (16 ocorrências) · Produção: entrega 41.

## Ambiente e método

| Item | Valor |
|---|---|
| Produção copiada (nunca aberta) | `QC_Bioquimica.xlsm` sha256 curto **8ec815ab653d**, igual antes e depois de cada uma das 3 execuções (`producao_intacta.igual = true`) |
| Cópias (novas, nenhuma salva) | `%TEMP%\qcw\auditoria\2026-11-14_reteste04\Bioquimica_F1.xlsm` (fluxo principal), `Bioquimica_F1_comp.xlsm` (vizinhos do OC-0012, OC-0019 e OC-0052), `Bioquimica_F1_val.xlsm` (validação de dados da Digitar Resultados) |
| Excel | um por vez, abertos e encerrados pelo `qa_final.QA`: PIDs 2020, 23060 e 7148. Outros Excel que apareceram (3076, 20232: tarefa agendada) não foram tocados. Nenhuma captura de tela; a caixa do `AbrirCEQ` foi lida pelo texto das janelas do meu PID |
| Script | `_arquitetura/agentes/ferramentas/reteste04_usuario_f1.py` (modos: padrão, `--complemento`, `--validacao`), documentado no cabeçalho. Reaproveita a infraestrutura (`qa_final.QA`, `repro_comum`) e as funções de leitura/vigia do `valida_adr090_091.py`; os oráculos novos (janela de 180 corridas, vizinhos, listas de validação) são meus |
| Saídas | `Bioquimica_F1.json/.log`, `Bioquimica_F1_comp.json/.log`, `Bioquimica_F1_val.json/.log` na mesma pasta do `%TEMP%` |
| Usuário | `TESTE_AUDITORIA_R4`, papel ANALISTA, criado só na cópia; login pelo `mSeguranca.DoLogin` |
| Dados de teste | manuais MAN_0001..0007 com MOTIVO "Outro: TESTE-AUDITORIA", valores 123,456 / 111,111 / 12,345 / −1 / "abc"; CEQ rodada "TESTE-AUDITORIA C-A 2025", amostras TA-31/55/71/81, arquivo TESTE-AUDITORIA.pdf; 2 linhas de mapa "TESTE-AUDITORIA MC"/"TESTE-AUDITORIA MCX" na EQA_Base W:Y (apagadas logo depois) |
| tblCQ_Final da cópia | 35.906 linhas; o ATUALIZAR não trouxe resultado novo (último recebido 13/11/2026 08:39:01) |

**Caminho usado.** Filtros do Painel pelas células (L4, C3, H3, K3, K4, G3/G4) com os eventos ligados; na OC-0041 o
Painel foi protegido como na tela. Manuais: navegação pelo botão (`mIntegracao.IrDigitarResultados`), digitação célula
a célula com os eventos ligados e **ATUALIZAR DADOS pelo núcleo do botão** (`mIntegracao.AtualizarDadosAutomatico`,
lendo o DB_SEAC só para leitura). CEQ: o núcleo que os botões da janela chamam (`mCEQForm.SalvarCEQ`, `CorrigirCEQ`,
`ExcluirCEQ`, `ConfirmacoesCEQ`, `AvisoAnalitoCEQ`, `DescricaoCEQ`, `RegistrosDaRodada`, `DistintosCEQ`) e o
`mCEQForm.AbrirCEQ` real. Edição direta da tabela do CAP com a aba liberada, como o administrador.

**Não executado por automação (sem captura de tela, sem clique em caixa):** o clique "Não"/"Sim" na caixa "Gravar mesmo
assim?" do frmCEQ e a digitação física numa célula com validação de dados. Para a validação, usei a avaliação do próprio
Excel (`Range.Validation.Value`), que é o critério que a caixa "Parar" aplica.

## Resultado por ocorrência

| OC | Estado | Evidência (cópia `Bioquimica_F1.xlsm` salvo indicação) |
|---|---|---|
| OC-0012 | **RESOLVIDA** | Glicose · DIMENSION 1 · 8974 · "Tudo (sem filtro)": `Eng_Saida!M1` = 39, `AF1` = "N1 174 · N2 161", A39 "⚠ 39 corrida(s) antigas do período fora do gráfico **e dos indicadores do Painel** (máx. 180). Período inteiro: N1 174 · N2 161 — veja a aba Estatística." Oráculo próprio (RUNs com resultado que participa, as 180 mais recentes): 219 RUNs, 39 cortadas, janela N1 140 · N2 131 = `engPainel` 140/131; período N1 174 · N2 161 = AF1 = Estatística (n 174/161). Vizinhos em `_comp`: mês com dados (Out/2025) → M1 0, sem aviso, n 24/22 = independente; H3 vazio + Tudo → M1 164, AF1 "N1 264 · N2 249" = independente, janela 129/130 = independente (o oráculo conta 163 corridas cortadas; a 164ª é uma corrida só com X vermelho, que o motor também conta, `mEstatistica` D10) |
| OC-0019 | **RESOLVIDA** (o que a ocorrência pedia) — ver **R4-01** | Painel em DIMENSION 2, manual sem EQUIPAMENTO: célula "DIMENSION 2", fundo 10284031 (RGB 255,235,156), fonte 22428 (RGB 156,87,0), nota "Preenchido pelo Painel com DIMENSION 2: confira antes de ATUALIZAR DADOS."; escrever "DIMENSION 1" por cima tira cor (16777215) e nota. A3 cita EQUIPAMENTO e "(fica em âmbar): confira". Vizinho: manual com o equipamento automático mantido (MAN_0002) entra ATIVO no DIMENSION 2, como o usuário deixou. **Mas** a correção que o aviso pede ("digite o certo", GUIA §3) é barrada pela validação de dados da coluna: ver R4-01 |
| OC-0021 | **PERSISTE** (parcial) | RUN vazio (Text "") nos MANUAL_INCOMPLETO e no CONFLITO_MANUAL; DETALHE "" no ATIVO; nenhum "0" nas 300 linhas. DETALHE acentuado nos textos da lista do ADR-090 ("campos obrigatórios ausentes/inválidos: NIVEL/HORA/RESULTADO"). **Continua sem acento** o DETALHE de dois status: IMPLAUSIVEL (MAN_0005: "resultado <= 0: implausivel para material de controle (fora da estatistica e do grafico)") e a nota do ADR-093 (MAN_0007: "… digitado num lote de URINA: lancado como 'CRE2' …"). O ADR-090 acentua só uma lista fixa de palavras (SUBSTITUTE) |
| OC-0023 | **RESOLVIDA** | 1º ATUALIZAR: sem NÍVEL (MAN_0001), sem HORA (MAN_0004) e RESULTADO "abc" (MAN_0006) → só RESULTADO_MANUAL_RECUSADO; conflito +5 min (MAN_0003) → RESULTADO_MANUAL_EM_CONFLITO. Completar o NÍVEL e atualizar → ATIVO, RUN 26101401, RESULTADO_MANUAL_INCLUIDO (uma vez). Vizinhos: −1 (MAN_0005) → RESULTADO_MANUAL_FORA_DA_ESTATISTICA (ADR-092); o 2º ATUALIZAR não repetiu trilha dos outros manuais |
| OC-0024 | **RESOLVIDA** | ResumoQA (e a mensagem do ATUALIZAR): "QA da integração: 7 erro(s)…", "• 3 resultado(s) manual(is) incompleto(s): veja o DETALHE na aba Digitar Resultados — não entraram na estatística." e "• 4 erro(s) de governança…". Os 4 são E01 **reais** (Lactato N2 inativado sem motivo em 13/10, BIO-366027/033/054/164); nenhum "Há ERRO DE GOVERNANÇA". Vizinho: sem manual incompleto (`_val`), a linha do E04 não aparece e fica só a de governança |
| OC-0031 | **RESOLVIDA** (função; clique na caixa não executado) | `ConfirmacoesCEQ` sem alvo: "Sem o valor-alvo, esta amostra NÃO entra no bias (nem no ET e no Sigma). Ela conta só na verificação dos limites."; com alvo: "". `SalvarCEQ` sem alvo: "OK\|547\|Gravado fora do bias (sem valor-alvo)."; G de Proteína total N1/N2 3,3793363600341007 → **igual bit a bit**; com alvo 100 → 5,034502270025575. Rótulos "Valor-alvo ¹" e nota no `lblSub` |
| OC-0033 | **RESOLVIDA** | Arquivo recém-aberto: `SalvarCEQ` "ERRO\|0\|Sessão encerrada: faça o login de novo. Nada foi gravado.", 545 → 545 linhas, nenhuma trilha; `AbrirCEQ` → caixa "Controle externo (CEQ)" "Faça o login para registrar resultados do CEQ.", janela não abriu. Vizinho (depois do `Logout`): Salvar, **Corrigir** e **Excluir** devolvem o mesmo ERRO\|0, F548 continua 108, linha não apagada, nenhuma trilha CEQ_* (o N-02 do reteste 03 também não ocorre mais) |
| OC-0034 | **RESOLVIDA** | `DistintosCEQ` (CAP) = conjunto das linhas reais: rodada 3, analito 31, avaliação 2; "A/B/C", "OK", "Glicose", "Creatinina", "Ureia"… da simulação fora. Controllab vazio |
| OC-0035 | **RESOLVIDA** (texto) | Li o GUIA como usuário: §3 cita EQUIPAMENTO e o âmbar; §5–6 citam a janela "Registrar resultado do CEQ", login, valor-alvo, Salvar e próximo, Corrigir/Excluir selecionado, "não digite nelas", Z3, Consolidar, ATUALIZAR, FÓRMULA SOBRESCRITA e o R4. Nenhuma instrução manda digitar nas abas. Todos os botões citados existem (Resumos do CAP/Controllab, abas de dados, frmCEQ, Painel/DB/Digitar Resultados) |
| OC-0036 | **RESOLVIDA** | 62 textos "x,xx / y,yy" na AT; AT14 "3,69 / 0,26", AT16 "2,37 / 0,17"; nenhum "004 / 000" |
| OC-0041 | **RESOLVIDA** | Painel protegido. K3 2026 + K4 S1 → digitar De/Até 01/09–31/10/2026 → K4 "Personalizado", K3 2026 → escolher K4 "Out" → G3/G4 01/10–31/10/2026. Vizinho: De/Até 01/12/2025–31/01/2026 → K4 "Personalizado", K3 vazio, Estatística D4 vazia, nenhum "Ano 0"; K3 2026 + "Ano inteiro" → 01/01–31/12/2026 |
| OC-0044 | **RESOLVIDA** | 96 linhas de analito com A:B centralizadas na vertical (−4108), inclusive Lactato N2 e Cálcio (33 pt) |
| OC-0050 | **RESOLVIDA** | Principal A5 "Total: 35.906 · participam: 35.763 · X vermelho: 4 · inativados: 11 · manuais: 0 · substituídos por manual: 0"; 35.906 linhas e 35.763 SIM contados na tblCQ_Final |
| OC-0052 | **RESOLVIDA** | O mapa real da Bio (37 nomes do CAP) não tem par por prefixo literal. Testei com um par sintético na cópia (`_comp`): "TESTE-AUDITORIA MC" (→ Glicose) e "TESTE-AUDITORIA MCX" (→ Ureia): `AvisoAnalitoCEQ` = "AZUL\|No QC_INI: Glicose. … Atenção: existe também “TESTE-AUDITORIA MCX” (Ureia)."; o nome maior não cita par. O par de palavra continua ("Lactate" → “Lactate Dehydrogenase”); "Glucose, serum" sem par |
| OC-0054 | **RESOLVIDA** | EQA.CAP_Dados linha 2 (ALT (SGPT), 120), seleção em AB1: escrever o mesmo valor → nenhuma trilha, Z3 "base em dia"; 121 → uma CEQ_ALTERADO "Seu Resultado: 120 → 121", sem "?". Vizinho com a célula selecionada antes (linha 548 de teste): "Seu Resultado: 104 → 106" e Z3 "HÁ EDIÇÃO NÃO CONSOLIDADA" |
| OC-0055 | **RESOLVIDA** | Rodada "TESTE-AUDITORIA C-A 2025", ano 2025: descrição "CAP TESTE-AUDITORIA C-A 2025 · Protein, total, serum · TA-55 (resultado 110)", lista da rodada e trilha (CEQ_REGISTRADO, CEQ_EXCLUIDO) sem "2025 2025". Vizinho: a mensagem de duplicidade "Já existe este registro (CAP TESTE-AUDITORIA C-A 2025 · … · TA-71) na linha 548…" também sem repetição |

Fluxos vizinhos sem defeito: ADR-093 (MAN_0007, lote 6354 + "Creatinina" → ANALITO CRE2, ATIVO, trilha INCLUIDO com a
nota); OC-0053 (colunas O/P/S..X com validação `=1=0`, P2 = 2,55 → Z3 "FÓRMULA SOBRESCRITA" e CEQ_FORMULA_SOBRESCRITA;
G de ALT inalterado 6,185176166653275 nos três momentos).

## Ocorrências novas

### R4-01 — Digitar Resultados: as listas de EQUIPAMENTO, MATRIZ e MOTIVO têm um item só ("DIMENSION 1,DIMENSION 2"); o valor certo digitado é recusado e o item do menu grava um equipamento inexistente que entra ATIVO

- **Prioridade:** Alto · **Tipo:** funcional · **Tela/módulo:** aba Digitar Resultados (tblResultados_Manuais), validação
  de dados criada por `camada_dados.validacoes` (`lista('EQUIPAMENTO', ','.join(...))`, idem MATRIZ e MOTIVO).
- **Ação:** depois do âmbar da OC-0019, corrigir o EQUIPAMENTO digitando "DIMENSION 1" (o que o GUIA §3 manda: "confira
  (ou digite o certo)"), ou escolher no menu da célula.
- **Esperado:** a lista oferece "DIMENSION 1" e "DIMENSION 2" (e "SORO"/"URINA", e os 4 motivos), e aceita qualquer um.
- **Observado (`Bioquimica_F1_val.json`, `_comp.json`):**
  - `Validation.Formula1` = "DIMENSION 1,DIMENSION 2", tipo lista, estilo **Parar**, mensagem "Escolha o equipamento da
    lista.". No Excel em português o separador de lista é ";", então a lista tem **um item só**, o texto inteiro.
  - `Validation.Value`: "DIMENSION 1" = Falso, "DIMENSION 2" = Falso; "DIMENSION 1,DIMENSION 2" = Verdadeiro.
  - O mesmo na MATRIZ ("SORO,URINA": "SORO" e "URINA" recusados) e no MOTIVO ("Falha do interfaceamento,Resultado não
    transmitido,Equipamento em contingência,Outro": cada motivo recusado). LOTE e ANALITO (`=lstLotes`, `=lstAnalitos`)
    estão certos.
  - Consequência medida: manual com o único item do menu (EQUIPAMENTO "DIMENSION 1,DIMENSION 2", Glicose N1, lote 8977) →
    ATUALIZAR → MAN_0001 **ATIVO, PARTICIPA SIM, RUN 26101401**, equipamento "DIMENSION 1,DIMENSION 2" na tblCQ_Final;
    DETALHE vazio e **nenhum achado no QA**. O resultado não aparece em nenhuma das duas séries do Painel.
- **Dificuldade:** o usuário vê o âmbar, tenta digitar o equipamento certo e recebe "Escolha o equipamento da lista."; a
  seta mostra uma única opção, que parece certa. O mesmo ao escolher o MOTIVO (obrigatório para substituir, ADR-082).
- **Consequência:** a correção que a OC-0019 pede fica impossível pela digitação; o atalho que sobra (trocar o L4 do
  Painel antes de digitar) não está no GUIA. Quem usa o menu grava um equipamento que não existe, com o resultado ATIVO
  fora das séries, sem aviso (agrava o G02 do ciclo 01). É a mesma causa do G18/G32 (texto com separador gravado pelo
  COM, lido pelo idioma do Excel).
- **Sugestão:** criar as listas com o separador do idioma (`Application.International(xlListSeparator)`) ou, melhor,
  apontá-las para um intervalo/nome (como LOTE e ANALITO), e reinstalar nas duas produções; a Hema usa o mesmo código
  para MATRIZ e MOTIVO (não testada aqui). Tratar no QA o EQUIPAMENTO fora da lista do produto (G02).
- **Evidência:** `Bioquimica_F1_val.json` (seções `listas` e `consequencia`), `Bioquimica_F1_comp.json`
  (`O19_validacao`), `Bioquimica_F1.json` (`O19_validacao_DIMENSION1` = false); fonte `camada_dados.py:532-537`.
- **Teste de digitação física:** NAO_EXECUTADO (sem interação com a tela); a avaliação é a do próprio Excel.

### R4-02 — Estatística B5/D5 mostram "00/01/1900" com o Painel em "Tudo (sem filtro)"

- **Prioridade:** Baixo · **Tipo:** usabilidade · **Tela:** aba Estatística, cabeçalho "PERÍODO ANALISADO".
- **Ação:** Painel em "Tudo (sem filtro)" (G3/G4 vazios) e abrir a Estatística, que o aviso da OC-0012 manda olhar.
- **Esperado:** De/Até vazios (o H5 já diz "histórico inteiro do lote").
- **Observado:** B5 `=Painel!G3` e D5 `=Painel!G4`, formato dd/mm/aaaa, mostram "00/01/1900"; ao lado, H5 "histórico
  inteiro do lote" e J5 vazio.
- **Consequência:** data sem sentido no cabeçalho de quem confere ou imprime a Estatística.
- **Sugestão:** `=IF(Painel!G3="","",Painel!G3)` (e G4), como já se fez na D4 (ADR-092).
- **Evidência:** `Bioquimica_F1.json` (`O12.estatistica_cabecalho`), `Bioquimica_F1_comp.json` (`estat_B5_D5_tudo`).

## Dúvidas de regra de negócio

1. **OC-0019, risco residual:** o manual com o equipamento automático mantido (MAN_0002, DIMENSION 2) entra ATIVO; o
   âmbar é só aviso. O item 3 do G01 (EQUIPAMENTO obrigatório na Bio, vazio = MANUAL_INCOMPLETO) segue sem decisão do
   Especialista/RT.
2. **OC-0021:** o ADR-090 deixou o `MOTIVO_EXCLUSAO_AUTOMATICA` sem acento de propósito (paridade VBA × PQ) e acentua só
   na tela por uma lista fixa. Textos novos (ADR-088, ADR-093) escapam da lista. Acentuar na origem ou manter a lista
   atualizada a cada texto novo?
3. **Trilha CEQ:** o Z3 diz "clique em Consolidar na EQA_Base" e a trilha diz "Consolidar na EQA_Base ou ATUALIZAR
   DADOS". Unificar o texto do Z3 com o da trilha?

## Impressões de uso (Relatório A)

- **Janela de 180 corridas:** o aviso agora diz o que muda (gráfico **e** indicadores) e dá o n do período inteiro, e a
  Estatística confirma o número (174/161). Para quem vê "REJEITADO" no J7 com CV de 19 %, o aviso é a pista certa.
- **Manuais:** o âmbar com nota é claro e chama atenção. O problema é o passo seguinte: com a validação das listas
  (R4-01), o usuário não consegue digitar o equipamento nem escolher um motivo; a única opção do menu grava lixo. Na
  prática, quem precisa lançar um manual hoje só consegue deixando o EQUIPAMENTO vazio com o Painel no equipamento
  certo e colando/escrevendo o MOTIVO por fora da validação. Isso precisa ir para o laboratório antes de segunda.
- **Status e DETALHE:** MANUAL_INCOMPLETO diz o campo que falta (NIVEL, HORA, RESULTADO); CONFLITO diz "coincide com
  resultado do interfaceamento (5 min)". O RUN em branco ficou limpo. Os textos sem acento do IMPLAUSIVEL e da troca para
  CRE2 destoam ao lado dos acentuados.
- **Mensagem do ATUALIZAR / ResumoQA:** separa manuais incompletos de governança; dá para saber o que fazer.
- **CEQ:** sem login a janela nem abre, e a mensagem diz por quê. As listas ficaram só com o que é real. A descrição e a
  duplicidade não repetem o ano e dizem a linha e o botão. O sem-alvo avisa a consequência antes de gravar.
- **Painel K3/K4:** depois de digitar De/Até, o rótulo "Personalizado" evita a confusão do "S1"; atravessando o ano o
  K3 fica vazio e a Estatística não mostra mais "Ano 0".
- **GUIA:** cobre a janela do CEQ e os avisos. A numeração das seções está fora de ordem (§6, §7, §6b, §6c, §7b, §8),
  o que atrapalha achar o item; e o §3 promete "digite o certo", o que hoje a validação impede (R4-01).

## Arquivos

- Script: `_arquitetura/agentes/ferramentas/reteste04_usuario_f1.py`.
- Saídas (fora do repositório): `%TEMP%\qcw\auditoria\2026-11-14_reteste04\Bioquimica_F1{,_comp,_val}.{json,log}`.
- O Audit_Log da cópia grava nome, login do Windows e máquina de quem roda; esses dados foram omitidos deste relatório.
