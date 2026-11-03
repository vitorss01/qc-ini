# Parecer UX/UI: ciclo 2026-10-31_ciclo01

Agente UX/UI (`.claude/agents/ux-ui.md`). Entrada: `rota_ux.json` (51 ocorrências, U01..U51 na ordem do arquivo),
evidência nos relatórios `<produto>_<frente>.md`. Referência visual: ADR-087 (critério "nada cortado ou sobreposto"),
`visual_telas.py`, `testes/auditoria_abas.py`, `testes/auditoria_janelas.py`, GUIA_DE_USO. Decisões de tela que
valem como contexto: ADR-077/078 (ajuste e devolução do filtro), ADR-082 (manual/SUBSTITUI_ID), ADR-086 (H3 vazio).

Não editei código e não abri Excel. As células e fórmulas citadas vêm do `src/` e dos instaladores
(`camada_dados.py`, `instalar_adr064.py`, `instalar_adr078.py`, `instalar_adr086.py`, `ceq_form_spec.py`,
`mCEQForm.bas`, `mIntegracao.bas`, `etapa1_multilote/painel.py`). Quando a célula exata só existe no `.xlsm`, está
escrito "a confirmar pelo implementador".

**Agrupamento:** as 51 viraram 32 pareceres (G01..G32). A coluna "Origem" diz quais U e quais OC dos relatórios.

**Armadilhas que valem para todas as correções:** aba protegida com `UserInterfaceOnly` (escrever pelo VBA com
`mSeguranca.LiberarEscrita`/`RestaurarProtecao`); nada de `Names.Add` com TRUE (usar `N(...)<>0`, como o
`engTodos`); nenhuma fórmula varrendo o banco (ADR-050: contagens em VBA com cache); textos inteiros (o
`auditar_aba` com 0 achados é o portão); fontes VBA em UTF-8.

**Causa comum achada no parecer (G18 e G32):** `TEXT(x;"0.00")` e `TEXT(x;"#,##0")` gravados em fórmula pelo COM
são lidos com o separador do Excel pt-BR (ponto = milhar, vírgula = decimal). Por isso aparecem "004 / 000" e
"44392,0". `FIXED(x;n)` respeita o idioma e já é usado na própria aba DB (`FIXED(ROWS(...),0)`). Sugiro ao Tester uma
varredura de `TEXT(` com código numérico em todos os instaladores.

---

## Quadro-resumo

| G | Título curto | Produto | Origem | Veredito | Prior. final | Autorização? |
|---|---|---|---|---|---|---|
| G01 | EQUIPAMENTO do manual preenchido sozinho | Bio | U01 (OC-3 F1) | PROCEDE | Alto | Parcial: destaque não; obrigatório sim |
| G02 | Manual fora do cadastro fica ATIVO e invisível | Bio+Hema | U02, U27 | PROCEDE | Médio | Alerta não; mudar status sim |
| G03 | RUN/DETALHE mostram 0; texto sem acento | Bio+Hema | U03, U32 | PROCEDE | Baixo | Não |
| G04 | I01 sem a matriz (CL de URINA) | Bio | U04 | PROCEDE | Baixo | Não |
| G05 | Audit_Log "INCLUIDO" para manual que não entrou | Bio+Hema | U05, U31 | PROCEDE | Baixo | Não |
| G06 | "ERRO DE GOVERNANÇA" para manual incompleto | Hema (mesma função na Bio) | U28 | PROCEDE | Médio | Não |
| G07 | Status REJEITADO = pior do período | Hema (vale p/ Bio) | U29 | PROCEDE_PARCIAL | Médio | Rótulo não; trocar o critério sim |
| G08 | MOTIVO "Outro" sem complemento | Hema (vale p/ Bio) | U30 | PROCEDE_PARCIAL | Baixo | Não |
| G09 | ANALISTA não corrige nem exclui CEQ | Bio+Hema | U06, U33 | PROCEDE | Alto | Sim (quem corrige/exclui) |
| G10 | Correção só vale após Consolidar; "base em dia" | Bio+Hema | U07, U34 | PROCEDE | Alto | Não |
| G11 | Nome fora do mapa / Controllab / analito novo | Bio+Hema | U08, U35 | PROCEDE | Alto | Aviso não; mapa e tela de mapa sim |
| G12 | "Lymphs" cai no LYMPH% sem aviso | Hema | U36 | PROCEDE | Alto | Não |
| G13 | CEQ sem valor-alvo aceito calado | Bio+Hema | U09, U37 | PROCEDE | Médio | Aviso não; obrigatório sim |
| G14 | Edição/consolidação/provedor sem trilha | Bio+Hema | U10, U38 | PROCEDE | Médio | Não |
| G15 | SalvarCEQ sem sessão nem papel | Bio+Hema | U11, U39 | PROCEDE | Médio | Sessão não; papel mínimo sim |
| G16 | Listas da janela com a simulação | Bio+Hema | U14, U41 | PROCEDE | Baixo | Não |
| G17 | GUIA §5–6 manda digitar nas abas bloqueadas | Bio+Hema | U15, U40 | PROCEDE | Médio | Não (depende de G09/G10/G11) |
| G18 | Coluna AT "004 / 000" | Bio+Hema | U16, U42 | PROCEDE | Médio | Não |
| G19 | Estatística no H3 vazio diz "refletido do Painel" | Bio (vale p/ Hema) | U12 | PROCEDE | Médio | Rótulo não; esvaziar C:F/H sim |
| G20 | Período não volta ao retornar de outro lote | Bio+Hema | U13, U20, U48 | PROCEDE_PARCIAL | Médio | Não (emenda de tela ao ADR-078) |
| G21 | "Lançamentos no lote" e A39 com o lote em uso | Bio+Hema | U17, U46 | PROCEDE | Médio | Não (lote em uso = ação do lab.) |
| G22 | Visões MENSAL/TRIMESTRAL mortas, "Ano 0" | Bio | U18 | PROCEDE | Médio | Não |
| G23 | K3/K4 com rótulo antigo após De/Até | Bio+Hema | U19, U49 | PROCEDE | Médio | Não |
| G24 | LJ suavizado, eixo Y sem números, DP sem rótulo | Bio+Hema | U21, U50 | PROCEDE | Médio | Não (avisar o usuário) |
| G25 | "REJEITADO" cortado em J7:J8 | Bio | U22 | PROCEDE | Baixo | Não |
| G26 | Nome do analito no pé da célula | Bio | U23 | PROCEDE | Baixo | Não |
| G27 | Lista do H3 com 22 lotes; troca o analito sozinho | Bio | U24 | PROCEDE | Médio | Não (emenda de tela ao ADR-077) |
| G28 | "Últ. violação" não é do período | Bio+Hema | U25, U47 | PROCEDE | Médio | Rótulo não; calcular no período sim |
| G29 | Ponto de violação cortado no topo do LJ | Bio | U26 | PROCEDE | Baixo | Não |
| G30 | Ano De/Até da Estatística corta o lote; Até < De | Hema | U43, U51 | PROCEDE | Médio | Aviso não; ET do lote inteiro sim |
| G31 | Linhas vazias da Estatística com 0 e SEM EP | Hema | U44 | PROCEDE | Baixo | Não |
| G32 | "Total: 44392,0" | Hema (mesma fórmula na Bio) | U45 | PROCEDE | Baixo | Não |

Nenhuma ocorrência foi julgada NAO_PROCEDE: todas têm evidência reproduzida e efeito real sobre quem usa o sistema.
Três são PROCEDE_PARCIAL porque o comportamento segue o que foi especificado (G07: GUIA §4b; G20: ADR-078 §6;
G08: lista do ADR-082), mas a tela não deixa isso claro ou a especificação deixou o usuário sem saída.

**Ordem sugerida:** G09+G10+G11+G12+G13 (fecham o CEQ) → G17 (GUIA) · G21+G19+G20+G27 (todas giram em torno do
"lote em uso"; fazer juntas) · G01+G02+G06 · demais.

---

## Bloco F1: entrada manual de CIQ (aba Digitar Resultados, QA, Painel)

### G01: EQUIPAMENTO vazio recebe o equipamento do Painel sem aviso (Bio; U01, OC-3 Bio F1)

- **Veredito:** PROCEDE.
- **Justificativa:** na Bioquímica há dois DIMENSION. O `PrepararLinha` grava `EquipPadrao()` (= `selEquipamento`,
  o equipamento que estiver no Painel) no primeiro toque da linha, e as instruções da aba (A3) e o GUIA §3 não citam
  o equipamento. Quem digita com o Painel no DIMENSION 2 um resultado que rodou no DIMENSION 1 não vê nada: o
  resultado entra na série e no Westgard do equipamento errado. É risco operacional, não preferência. Na Hema há um
  equipamento só, e o risco é nulo.
- **Prioridade final:** Alto (associação errada entre equipamentos; não é Crítico porque o valor vem de uma omissão
  do usuário e o campo fica visível na linha).
- **Correção proposta:**
  1. *(sem autorização)* `camada_dados.py`, instrução 1 da aba Digitar Resultados: "1) Preencha DATA, HORA,
     **EQUIPAMENTO**, LOTE, NÍVEL, ANALITO, RESULTADO e MOTIVO na próxima linha livre — o ID (MAN_0001…) é gerado
     sozinho. EQUIPAMENTO vazio recebe o do Painel: confira. Para SUBSTITUIR…" (o resto igual). GUIA §3.1 com o
     mesmo campo.
  2. *(sem autorização)* `mIntegracao.PrepararLinha`, quando **preencher** o EQUIPAMENTO sozinho e o produto tiver
     mais de um equipamento: pintar a célula de âmbar (o mesmo âmbar do ADR-078, fonte RGB(156,87,0), fundo
     RGB(255,235,156)) e pôr a nota da célula: "Preenchido automaticamente com DIMENSION 2 (o equipamento do Painel).
     Confira antes de ATUALIZAR." Quando o usuário digitar no EQUIPAMENTO, a cor e a nota saem.
  3. *(com autorização: Especialista/RT)* na Bioquímica, **não** preencher o equipamento sozinho e tratar o vazio
     como MANUAL_INCOMPLETO ("falta EQUIPAMENTO"). É a dúvida 3 da F1 Bio.
- **Autorização:** parcial: itens 1 e 2 não; item 3 sim.
- **Critério de aceite:** com o Painel em DIMENSION 2, digitar um manual sem EQUIPAMENTO → a célula fica âmbar com a
  nota acima (texto exato); digitar "DIMENSION 1" por cima → cor e nota somem. A instrução da aba cita EQUIPAMENTO e
  o `auditar_aba` da Digitar Resultados dá 0 CORTADO.

### G02: Manual com analito, lote ou equipamento fora do cadastro fica ATIVO e não aparece em lugar nenhum (Bio+Hema; U02, U27)

- **Veredito:** PROCEDE.
- **Justificativa:** a tela diz "ATIVO = entrou na estatística", mas o resultado não está no Painel. Digitado, a
  validação recusa. **Colado**, passa (colar invalida a validação de dados do Excel, e colar de planilha de
  contingência é o caso típico). O usuário acredita que lançou. A Hema já trata o equipamento (A10); falta o resto.
- **Prioridade final:** Médio.
- **Correção proposta:**
  1. *(sem autorização)* **Barrar o colado na hora.** No `PrepararLinha` (já roda a cada mudança da linha), conferir
     ANALITO contra `lstAnalitos`, LOTE contra `lstLotes` e EQUIPAMENTO contra a lista do produto, sem diferenciar
     maiúsculas. Fora da lista: célula em vermelho claro e mensagem única por colagem:
     "O valor colado em ANALITO (“Glicoze”) não está no cadastro. Esta linha não vai aparecer no Painel. Escolha o
     nome na lista da célula." (trocar ANALITO/LOTE/EQUIPAMENTO e o valor). Não corrigir sozinho: grafia só de
     maiúsculas é a OC-9/OC-HF1-1 (Crítico, rota do CQ), e a normalização é decisão dele.
  2. *(sem autorização)* **QA novo, ALERTA A11** "Resultado manual fora do cadastro (não aparece no Painel)", um por
     ID, com o detalhe "ANALITO 'Glicoze' não está no cadastro" / "LOTE '99999' não está no cadastro (aba
     Configuração)" / "EQUIPAMENTO 'DIMENSION 9' não é deste QC". Na Hema o equipamento continua no A10.
  3. *(sem autorização)* instrução 3 da aba: "ATIVO = aceito; aparece no Painel se analito, lote e equipamento forem
     do cadastro (senão o QA avisa A11)".
  4. *(com autorização)* virar MANUAL_INCOMPLETO em vez de ATIVO muda o que entra na base estatística: decisão do
     Especialista.
- **Autorização:** itens 1–3 não; item 4 sim.
- **Critério de aceite:** colar "Glicoze", "99999" e "DIMENSION 9" (cópia, `TESTE-AUDITORIA`) → cada célula fica
  vermelha com a mensagem acima; depois do ATUALIZAR, o QA tem um A11 por ID com o detalhe certo; digitar os mesmos
  valores continua recusado pela validação.

### G03: Colunas RUN e DETALHE mostram "0"; DETALHE sem acento (Bio+Hema; U03, U32)

- **Veredito:** PROCEDE. O `XLOOKUP` de célula vazia devolve 0, e "0" em DETALHE parece um código.
- **Prioridade final:** Baixo.
- **Correção proposta:** só formato, sem tocar na fórmula: `NumberFormat` da coluna RUN = `0;-0;;@` e da coluna
  DETALHE = `General;-General;;@` (o zero some, texto e número continuam). Alternativa: em DETALHE, `XLOOKUP(...)&""`.
  Acentuar os textos de `MOTIVO_EXCLUSAO_AUTOMATICA` gerados no `mEtlCQ` ("obrigatórios", "lançamento", "ambíguo"):
  fonte em UTF-8, conferir que nenhum teste compara o texto sem acento.
- **Autorização:** não.
- **Critério de aceite:** linha ATIVO com DETALHE vazio; linha MANUAL_INCOMPLETO/CONFLITO_MANUAL com RUN vazio;
  DETALHE com acentos; os testes de `qa_etl`/`qa_manual_substitui` continuam passando.

### G04: I01 lista NA/K/CL/CRE2/PHOS sem dizer que são da urina (Bio; U04)

- **Veredito:** PROCEDE. "CL … não aparecem no Painel" faz o usuário achar que o Cloro sérico sumiu.
- **Prioridade final:** Baixo.
- **Correção proposta:** `mEtlQA` (chamada `GrupoAnalito … "I01"`): agrupar por ANALITO|MATRIZ e escrever
  "CL (URINA): 3 resultado(s) guardados na base; não aparecem no Painel até o analito ser cadastrado/mapeado."
  (com acentos).
- **Autorização:** não.
- **Critério de aceite:** no QA da cópia, os I01 de NA, K e CL dizem "(URINA)"; nenhum I01 de matriz SORO para analito
  cadastrado.

### G05: Audit_Log grava RESULTADO_MANUAL_INCLUIDO para linha incompleta ou em conflito (Bio+Hema; U05, U31)

- **Veredito:** PROCEDE. A trilha diz "incluído" para algo que não entrou (chave "0|0|GLICOSE", RUN vazio).
- **Prioridade final:** Baixo.
- **Correção proposta:** no `mIntegracao` (auditoria entre atualizações), escolher a ação pelo status do manual:
  MANUAL_INCOMPLETO → `RESULTADO_MANUAL_RECUSADO` (detalhe: o campo que falta); CONFLITO_MANUAL →
  `RESULTADO_MANUAL_EM_CONFLITO` (detalhe: o ID automático que vale); ATIVO → `RESULTADO_MANUAL_INCLUIDO`. Quando a
  mesma linha passar depois a ATIVO, gravar a `INCLUIDO` nessa hora. A chave sem nível ("0|") fica com o nível vazio,
  não 0.
- **Autorização:** não (vocabulário da trilha, sem mudar regra).
- **Critério de aceite:** na cópia, manual sem MOTIVO → só `RESULTADO_MANUAL_RECUSADO`; completá-lo e atualizar →
  `RESULTADO_MANUAL_INCLUIDO` com RUN preenchido.

### G06: "Há ERRO DE GOVERNANÇA" quando o único erro é um manual incompleto (Hema; U28; mesma função na Bio)

- **Veredito:** PROCEDE. `mIntegracao.ResumoQA` acrescenta a frase de governança para **qualquer** ERRO. O
  usuário procura uma inativação que não existe.
- **Prioridade final:** Médio.
- **Correção proposta:** `ResumoQA` passa a contar por grupo de código e montar uma linha por grupo presente:
  - primeira linha (igual): "QA da integração: e erro(s), a alerta(s), i informativo(s).";
  - E04: "• n resultado(s) manual(is) incompleto(s): veja o DETALHE na aba Digitar Resultados — não entraram na
    estatística.";
  - E12: "• n manual(is) com SUBSTITUI_ID inválido: nada foi substituído — veja o DETALHE na aba Digitar Resultados.";
  - E01, E02, E03, E10, E11: "• n erro(s) de governança: veja a aba QA_INTEGRACAO (ex.: inativação sem motivo, ou ID
    que não é do analito informado — essa não foi aplicada).";
  - E05 a E09: "• n erro(s) nos dados recebidos: veja a aba QA_INTEGRACAO."
- **Autorização:** não.
- **Critério de aceite:** cópia com 1 E04 e nenhum outro ERRO → a mensagem não contém "governança" e contém a linha
  do E04; cópia com E01 → contém a linha de governança. A mensagem cabe na MsgBox sem rolar (≤ 1024 caracteres).

### G07: Status do Painel REJEITADO nos 3 níveis por uma regra antiga, mesmo após corridas OK (Hema; U29; vale p/ Bio)

- **Veredito:** PROCEDE_PARCIAL. O GUIA §4b define o Status como "veredito do Westgard no período", ou seja, o pior
  do período: o sistema faz o que está escrito. Mas a tela não diz isso, e quem acabou de lançar uma corrida OK lê
  "REJEITADO" como se fosse ela.
- **Prioridade final:** Médio.
- **Correção proposta:**
  1. *(sem autorização)* cabeçalho da coluna (J6, a confirmar) = "Status no período" (com quebra; altura da linha 6
     não muda, conferir no `auditar_aba`).
  2. *(sem autorização)* acrescentar ao texto do A39, antes do "Filtro …": "Status = pior corrida do período. Última
     corrida (RUN 26103002): N1 OK · N2 OK · N3 OK." montado do veredito por corrida que o motor já publica (coluna a
     confirmar pelo implementador no Calc/Eng_Saida; nenhuma fórmula nova varre o banco).
  3. *(com autorização: Especialista)* trocar o critério da coluna J para a última corrida (dúvida 4 da F1 Hema).
- **Autorização:** itens 1 e 2 não; item 3 sim.
- **Critério de aceite:** com Período "Tudo" e a corrida manual OK lançada, o Painel mostra J7:J9 "REJEITADO" sob
  "Status no período" e o A39 diz a última corrida com OK nos 3 níveis; `auditar_aba` do Painel com 0 achados.

### G08: MOTIVO só aceita os 4 itens; "Outro" sem complemento (Hema; U30; vale p/ Bio)

- **Veredito:** PROCEDE_PARCIAL. A lista fechada é boa para o caso comum e foi escolha do projeto; o problema é só o
  "Outro", que deixa a trilha da substituição (ADR-082) com "Outro" e nada mais. O ETL já aceita texto livre (colado
  entra no I04).
- **Prioridade final:** Baixo.
- **Correção proposta:** `camada_dados.py`, validação do MOTIVO com estilo **Aviso** (2) em vez de Parar, como já é o
  ANALITO da aba Inativar. Título "Motivo fora da lista"; mensagem (≤ 225 caracteres): "Escolha um motivo da lista.
  Para outro motivo, escreva “Outro: ” e a explicação (ex.: Outro: queda de energia no SIPEC) e clique em Sim."
  Atualizar a dica do SUBSTITUI_ID se citar o motivo.
- **Autorização:** não (o MOTIVO continua obrigatório; só passa a aceitar o complemento).
- **Critério de aceite:** digitar "Outro: TESTE-AUDITORIA" → aviso com Sim/Não; Sim grava; o I04 e o MOTIVO_MANUAL
  mostram o texto inteiro; um item da lista continua sem aviso.

---

## Bloco F2: Controle Externo (frmCEQ, abas EQA.*, GUIA)

### G09: ANALISTA não consegue corrigir nem excluir um CEQ; a mensagem manda editar uma aba bloqueada (Bio+Hema; U06, U33)

- **Veredito:** PROCEDE. A mensagem `ERRO|5 … Para corrigir, edite a linha existente.` manda o usuário para células
  `Locked` numa aba protegida. Não há caminho de correção para quem registra; a saída improvisada (registrar outra
  amostra) cria duplicata que distorce o bias.
- **Prioridade final:** Alto.
- **Correção proposta (na janela, sem mudar o modelo de dados):**
  - `frmCEQ`: dois botões na linha dos botões (Top 420, Height 27): **"Corrigir selecionado"** (Left 12, Width 120)
    e **"Excluir selecionado"** (Left 138, Width 120). O `btnSalvar` continua em 318. Os dois ficam desabilitados
    sem seleção em `lstRegs`.
  - **Corrigir:** carrega a linha escolhida nos campos, troca o título do botão para "Salvar correção" e mostra em
    `lblMsg` (azul): "Corrigindo o registro da linha 547. Altere os campos e clique em Salvar correção." O
    `SalvarCEQ` ganha o modo correção (linha alvo): mesma validação; a duplicidade ignora a própria linha; grava
    A..N, Q, R na mesma linha; consolida (`AtualizarEQABase`); registra `CEQ_ALTERADO` com antes → depois (G14).
  - **Excluir:** confirmação `vbYesNo`, botão padrão **Não**: "Excluir o registro CAP C-A 2025 · Glucose, serum ·
    CHM-01 (resultado 110)? Ele sai do bias, do ET e do Sigma. A exclusão fica no Audit_Log." Sim → limpa A..N, Q, R
    da linha (a linha da tabela fica, vazia, como as já previstas pelo `livre`), consolida e registra `CEQ_EXCLUIDO`
    com os valores apagados.
  - **Mensagem de duplicidade (texto exato):** "Já existe este registro (CAP C-A 2025 · Glucose, serum · CHM-01) na
    linha 547. Para corrigir, selecione-o em “Registros desta rodada” e clique em Corrigir selecionado."
  - Passa pela `auditoria_janelas.py` (sem sobreposição, textos inteiros).
- **Autorização:** **sim**: quem pode corrigir e excluir (dúvida 4 da F2). Sugestão de padrão para decidir: chave
  `PAPEL_MIN_CORRIGIR_CEQ` na Cfg_Integracao, como o `PAPEL_MIN_INATIVAR`. O mecanismo da janela não muda regra.
- **Critério de aceite:** ANALISTA (cópia) registra 110, seleciona e corrige para 105 → a EQA_Base, o G e o Sigma da
  Estatística mudam para os valores do cálculo independente (G 3,30 na Glicose Bio; 2,4002 no RBC Hema); excluir →
  o analito volta ao bias anterior; Audit_Log com `CEQ_ALTERADO` e `CEQ_EXCLUIDO`; `auditoria_janelas` sem achado.

### G10: Correção na tabela só chega ao bias depois do "Consolidar"; Z3 diz "base em dia" e Z2 fica parado (Bio+Hema; U07, U34)

- **Veredito:** PROCEDE. A aba de dados mostra o bias novo e a Estatística o antigo, e nada na tela avisa. O Z3
  compara só a quantidade de linhas; o Z2 é uma UDF sem dependência e nunca recalcula.
- **Prioridade final:** Alto.
- **Correção proposta:**
  1. **Z2:** trocar `=mEQA.CarimboEQA()` por referência direta ao carimbo que a consolidação grava
     (`="consolidado: "&EQA_Base!$Z$1`, conferir a célula do `CARIMBO_LIN/COL`). Recalcula sozinho.
  2. **Marca de edição:** `Worksheet_Change` das abas EQA.CAP_Dados/EQA.Controllab_Dados, quando a mudança cruza as
     colunas digitadas da tabela, grava `Now` numa célula oculta da aba (ex.: Z4, valor). O `AtualizarEQABase`
     grava também o instante numérico da consolidação (ex.: EQA_Base!AA1).
  3. **Z3:** `=IF(N(Z4)>N(EQA_Base!$AA$1),"⚠ HÁ EDIÇÃO NÃO CONSOLIDADA — clique em Consolidar na EQA_Base","base em
     dia")`, com a formatação âmbar do ADR-078 quando começar com "⚠". Mantém o teste atual de quantidade como
     segundo motivo.
  4. **Consolidar sozinho:** `Worksheet_Deactivate` dessas abas consolida quando houver marca pendente, com a janela
     de etapas (`frmCarregando`): "Consolidando o CEQ na EQA_Base…". O ATUALIZAR DADOS também consolida quando houver
     marca pendente.
- **Autorização:** não (não muda cálculo; só faz a correção valer quando o usuário espera).
- **Critério de aceite:** ADM edita 110 → 105 na tabela → Z3 âmbar com o texto acima; sair da aba → janela de etapas,
  Z2 com o carimbo novo, Z3 "base em dia" e a Estatística com G/ET/Sigma recalculados (mesmos valores do G09).

### G11: CEQ com nome fora do mapa, do Controllab ou de analito novo é gravado "OK" e nunca entra no bias (Bio+Hema; U08, U35)

- **Veredito:** PROCEDE. O registro aparece na lista como se estivesse certo, e o analito fica SEM EP em silêncio.
  Na Hema, R4 = Controllab zera o bias dos 28 analitos de uma vez.
- **Prioridade final:** Alto.
- **Correção proposta:**
  1. *(sem autorização)* **Aviso ao escolher/digitar o analito** (`cboAnalito_Change`/`AfterUpdate`): chamar
     `mEQA.MapearAnalito(provedor, analito)` e escrever em `lblMsg`:
     - mapeado (azul): "No QC_INI: Glicose. Este resultado entra no bias deste analito.";
     - não mapeado (âmbar): "⚠ “Glucose” não está no mapa do CAP: este resultado NÃO entra no bias, no ET nem no
       Sigma. Use o nome da lista (como no relatório) ou peça ao administrador para mapear."
  2. *(sem autorização)* **Confirmação ao gravar** sem canônico, `vbYesNo`, padrão **Não**: "O analito “Hemoglobin
     A1c” não está no mapa do CAP. O registro será gravado, mas NÃO entra no bias, no ET nem no Sigma até o
     administrador incluir o nome no mapa (EQA_Base, colunas W:Y). Gravar mesmo assim?" Sim grava e devolve no aviso
     do `OK|linha|aviso` (o `SalvarCEQ` já tem esse terceiro campo) "Gravado fora do bias (analito não mapeado).".
  3. *(sem autorização)* Hema R4: ao trocar para um provedor sem nenhuma entrada no mapa, `MsgBox` "O provedor
     Controllab ainda não tem analitos no mapa: com ele, nenhum analito terá bias, ET nem Sigma. Trocar mesmo
     assim?" (padrão Não).
  4. *(com autorização)* incluir o mapa do Controllab e uma tela (ou procedimento) para mapear o primeiro CEQ de um
     analito novo; quem pode mapear (dúvida 7 da F2 Bio; ADR-071: provedores nunca se misturam).
- **Autorização:** itens 1–3 não; item 4 sim.
- **Critério de aceite:** "Glucose" (CAP Bio), "Red Blood Cell" (CAP Hema) e qualquer nome no Controllab → aviso
  âmbar ao sair do campo e confirmação ao gravar; Não → nada gravado; nome da lista → aviso azul com o canônico;
  `auditoria_janelas` mede os dois textos mais longos dentro do `lblMsg` (576 × 27 pt, quebra em 2 linhas).

### G12: "Lymphs" (de LYMPH#) entra calado no bias do LYMPH% e derruba o Sigma dele (Hema; U36)

- **Veredito:** PROCEDE. Uma palavra a menos ("Absolute") muda o Sigma de outro analito de 4,33 para 1,70, com
  unidade incompatível (x 10E9/L num analito em %). O mapa tem 7 pares desse tipo.
- **Prioridade final:** Alto (associação errada entre analitos; mantenho Alto, como o relatório, porque a origem é
  digitação; o Orquestrador pode elevar a Crítico pelo §4 das regras).
- **Correção proposta:**
  1. O aviso azul do G11 já mostra o canônico ("No QC_INI: LYMPH%"), o que torna a troca visível antes de gravar.
  2. **Conferência de unidade ao gravar:** se a unidade digitada diferir da unidade já registrada para aquele
     canônico (mesma tabela, linhas de uso SIM), `vbYesNo` com padrão **Não**: "A unidade “x 10E9/L” é diferente da
     já registrada para LYMPH% (“%”). Se o resultado é do LYMPH#, use o nome “Lymphs Absolute”. Gravar mesmo
     assim?"
  3. Quando o nome digitado for prefixo de outro nome do mapa ("Lymphs" × "Lymphs Absolute"), acrescentar ao aviso
     azul: "Atenção: existe também “Lymphs Absolute” (LYMPH#)."
- **Autorização:** não (aviso e confirmação; a gravação continua possível).
- **Critério de aceite:** registrar "Lymphs" com x 10E9/L → aviso azul com LYMPH% e a menção a "Lymphs Absolute";
  ao gravar, a confirmação de unidade; Não → nada muda no LYMPH% (80 células iguais às de antes).

### G13: CEQ sem valor-alvo aceito sem aviso; conta nos limites e não no bias (Bio+Hema; U09, U37)

- **Veredito:** PROCEDE. O usuário acha que registrou a amostra inteira, e o bias sai com uma amostra a menos.
  Tornar o alvo obrigatório não é bom: há amostras sem alvo publicado (não avaliadas) que o laboratório pode querer
  guardar.
- **Prioridade final:** Médio.
- **Correção proposta:** rótulo `lblAlvo` = "Valor-alvo ¹" e `lblSub` = "Campos com * são obrigatórios. ¹ Sem o
  valor-alvo, a amostra não entra no bias. O resultado vai para a aba do provedor e a EQA_Base é consolidada em
  seguida." (medir no `auditoria_janelas`; se não couber em 1 linha, `lblSub` com 2 linhas e a altura certa). Ao
  gravar sem alvo, `vbYesNo` padrão **Não**: "Sem o valor-alvo, esta amostra NÃO entra no bias (nem no ET e no
  Sigma). Ela conta só na verificação dos limites. Gravar assim?"
- **Autorização:** não para o aviso. Alvo obrigatório, ou a amostra sem alvo deixar de contar na coluna S: sim
  (Especialista; dúvida 3 da F2 Hema).
- **Critério de aceite:** gravar sem alvo → confirmação; Não → nada gravado; Sim → `OK` e aviso; textos inteiros no
  `auditoria_janelas`.

### G14: Edição, consolidação, exclusão e troca de provedor sem trilha no Audit_Log (Bio+Hema; U10, U38)

- **Veredito:** PROCEDE. Não há como responder "qual era o CEQ antes, quem mudou e quando", nem quem trocou o
  provedor (na Hema, dos 28 analitos).
- **Prioridade final:** Médio.
- **Correção proposta:** `mAuditoria.RegistrarLog` nos pontos: `CEQ_ALTERADO` (janela G09 e `Worksheet_Change` da
  tabela: "prov | rodada | ano | analito | amostra | campo: antes → depois | aba linha"); `CEQ_EXCLUIDO` (valores
  apagados); `EQA_BASE_CONSOLIDADA` (linhas, sem canônico, quem); `PROVEDOR_ANALITO_ALTERADO` (Bio, Painel M4:
  analito, antes → depois) e `PROVEDOR_EQ_ALTERADO` (Hema, R4: antes → depois, "vale para todos os analitos").
- **Autorização:** não.
- **Critério de aceite:** na cópia, editar, consolidar, excluir e trocar o provedor → uma linha de cada no Audit_Log
  com antes/depois e usuário.

### G15: SalvarCEQ grava sem sessão nem papel (Bio+Hema; U11, U39)

- **Veredito:** PROCEDE. Depois do logout, `OK|553|`. A inativação já confere (`PodeInativar`, ADR-081).
- **Prioridade final:** Médio.
- **Correção proposta:** no `AbrirCEQ`: `If Not SessaoAtiva() Then MsgBox "Faça o login para registrar resultados do
  CEQ.", vbExclamation: Exit Sub`. No `SalvarCEQ` (primeira linha): devolver `"ERRO|0|Sessão encerrada: faça o
  login de novo. Nada foi gravado."`. Papel mínimo configurável `PAPEL_MIN_CEQ` (Cfg_Integracao), mesma rotina do
  `PAPEL_MIN_INATIVAR`, com a mensagem "Seu perfil (VISUALIZADOR) não pode registrar CEQ. Peça a um usuário com
  perfil ANALISTA ou superior."
- **Autorização:** sessão não; o papel mínimo e o valor padrão sim (dúvida 4 da F2).
- **Critério de aceite:** logout + `AbrirCEQ` → mensagem, janela não abre; logout + `SalvarCEQ` → `ERRO|0` e nada
  gravado; com login, fluxo normal.

### G16: Listas da janela misturam a simulação com o CAP real (Bio+Hema; U14, U41)

- **Veredito:** PROCEDE. Rodadas "A/B/C" e "1/2/3", nomes "RBC"/"Glicose" da simulação e avaliação "OK" confundem a
  escolha.
- **Prioridade final:** Baixo.
- **Correção proposta:** `mCEQForm.DistintosCEQ` pula as linhas cujo Arquivo fonte (coluna `CQ_ARQ` = 18) dá
  `mEQA.UsoAnalitico(...) = "NAO"`, a mesma regra que marca o Uso na EQA_Base. Sem mudar os dados.
- **Autorização:** não.
- **Critério de aceite:** listas de rodada, analito e avaliação sem "A/B/C", "1/2/3", "RBC", "Glicose", "OK" da
  simulação; o registro de uma rodada real continua igual.

### G17: GUIA §5–6 manda digitar o CEQ nas abas bloqueadas e não cita a janela, o Consolidar nem a correção (Bio+Hema; U15, U40)

- **Veredito:** PROCEDE. O manual é o que o laboratório segue quando a tela não é óbvia, e hoje ele leva à proteção
  da aba.
- **Prioridade final:** Médio.
- **Correção proposta:** reescrever o item CEQ do §6 depois de G09/G10/G11 (rascunho):
  "**CEQ:** no resumo do CAP ou do Controllab, clique em **Registrar resultado do CEQ**. Preencha na ordem do
  relatório: rodada → analito e amostra → seu resultado, **valor-alvo**, SD e nº de laboratórios → avaliação. Use o
  nome do analito da lista: a janela mostra para qual analito do QC_INI o resultado vai, e avisa quando ele não entra
  no bias. Sem valor-alvo, a amostra não entra no bias; sem SD e nº de laboratórios, a incerteza não sai. **Salvar e
  próximo** grava e consolida (15 a 25 s). **Corrigir ou excluir:** selecione o registro em “Registros desta rodada”
  e clique em Corrigir selecionado ou Excluir selecionado (fica no Audit_Log)." No §5 (Hema): "R4 vale para todos os
  analitos ao mesmo tempo."
- **Autorização:** não (texto; o conteúdo depende das decisões de G09/G10/G11).
- **Critério de aceite:** o Usuário Simulado segue só o GUIA e registra, corrige e exclui um CEQ na cópia sem bater na
  proteção.

---

## Bloco F3: Estatística (textos e acabamento)

### G18: Coluna AT mostra "004 / 000" em vez de "3,69 / 0,xx" (Bio+Hema; U16, U42)

- **Veredito:** PROCEDE. A memória da incerteza fica ilegível e enganosa. A causa é o `TEXT(…,"0.00")` em pt-BR (ver
  a causa comum no topo).
- **Prioridade final:** Médio.
- **Correção proposta:** `instalar_adr064.py`, coluna "RMS viés % / u(Cref) %": trocar os dois `TEXT(x,"0.00")` por
  `FIXED(x,2,TRUE)`. Não precisa de VBA.
- **Autorização:** não.
- **Critério de aceite:** Bio AT14 = "3,69 / " + u(Cref) com 2 casas e vírgula; AT20 = "24,19 / …"; Hema AT14 =
  "2,49 / …", AT23 = "1,30 / …" (valores do `vies_nordtest`).

### G31: Linhas vazias da Estatística mostram "0" e "SEM EP" (Hema; U44)

- **Veredito:** PROCEDE. Ruído visual em 36 linhas.
- **Prioridade final:** Baixo.
- **Correção proposta:** A14:A133 da Hema = `=IF(<origem>="","",<origem>)` (hoje `=Analitos!$A$n`, a referência ao
  catálogo); G, R, S, T testam `OR($A14="",$A14=0)`, como já fazem K e AF.
- **Autorização:** não.
- **Critério de aceite:** A98:T133 vazias; as 84 linhas com analito inalteradas (comparar `Value2`).

### G32: Cabeçalho da Principal - Resultados com "Total: 44392,0" (Hema; U45; a mesma fórmula está na Bio)

- **Veredito:** PROCEDE. Mesma causa do G18: `TEXT(…,"#,##0")` em pt-BR.
- **Prioridade final:** Baixo.
- **Correção proposta:** `camada_dados.py`, cabeçalho da 'Principal - Resultados': `FIXED(ROWS(tblCQ_Final),0)` e
  `FIXED(COUNTIFS(...,"SIM"),0)` no lugar dos dois `TEXT(…,"#,##0")` (como a aba DB já faz). Nos dois produtos.
- **Autorização:** não.
- **Critério de aceite:** A5 = "Total: 44.392  ·  participam: 43.710  ·  …".

---

## Bloco F4: Painel, períodos, lotes e gráficos

### G21: "Lançamentos no lote" e o A39 do H3 vazio mostram o lote em uso, que não é o da tela (Bio+Hema; U17, U46)

- **Veredito:** PROCEDE. O usuário não sabe qual lote está vendo nem para onde vão os manuais; o A39 dá o alvo de um
  lote sem ponto no período. Parte é dado desatualizado (Hema 522611 vencido desde 31/03/2026; Bio 8974 com a
  Glicose rodando no 8977), mas a tela também não ajuda.
- **Prioridade final:** Médio.
- **Correção proposta:**
  1. **C4** (`painel.py`): trocar "  ·  Lançamentos no lote: "&loteAtivo por
     "  ·  Lote em uso (digitação): "&loteAtivo&IF(<validade do lote em uso><TODAY()," ⚠ vencido","") (célula da
     validade na Configuração a confirmar; C4 já tem "reduzir para caber", conferir ≤ 30 %).
  2. **A39** (`instalar_adr078.a39`): no modo TODOS, o primeiro trecho vira
     `IF(engTodos,"TODOS os lotes do período: "&engLotesPer&" (cada ponto com o alvo do seu lote). ", <fórmula atual>)`
     (`engLotesPer` = Eng_Saida O1, que já tem "8977 (55) · 5659 (18)").
  3. **Laboratório:** atualizar o lote em uso na Configuração (Hema: 625311?; Bio: confirmar 8974). Não é código.
- **Autorização:** não (texto). A atualização do lote em uso é ação do laboratório, levada pelo Orquestrador.
- **Critério de aceite:** H3 = 8977 → C4 "Lote em uso (digitação): 8974"; H3 vazio na HbA1c jun/2026 → A39 começa
  com "TODOS os lotes do período: 8589 (…) · 8590 (…)" e não cita o alvo do 8974; `auditar_aba` com 0 achados.

### G19: Com H3 vazio, a Estatística calcula o lote em uso e diz "refletido do Painel" (Bio; U12; vale p/ Hema)

- **Veredito:** PROCEDE. O rótulo contradiz o Painel, que está em TODOS. ET, margem e orçamento passam a ser lidos
  como se fossem da seleção.
- **Prioridade final:** Médio.
- **Correção proposta:**
  1. *(sem autorização)* o rótulo ao lado do lote (Bio I4, Hema C4) vira fórmula:
     `=IF(engTodos,"(Painel em TODOS os lotes: esta aba mostra o lote em uso)","(refletido do Painel)")`, com fonte
     âmbar quando `engTodos`. O valor do lote (H4/B4 = `loteAnalise`) não muda, para não afetar o VBA que o lê.
  2. *(com autorização)* esvaziar C:F/H/N:Q no modo TODOS, como o Painel (ADR-086 decisão 4), muda a base do plano
     de CQ: decisão do Especialista. Depende também da OC-HF4-1 (a Estatística não recalcula na troca de lote; rota
     Tester/CQ), que deve ser resolvida antes ou junto.
- **Autorização:** item 1 não; item 2 sim.
- **Critério de aceite:** H3 vazio → Bio I4 / Hema C4 com o texto âmbar acima; H3 = 8977 → "(refletido do Painel)";
  textos inteiros no `auditar_aba` da Estatística.

### G20: Ao voltar de outro lote, o período escolhido não volta quando o lote de partida não é o lote em uso (Bio+Hema; U13, U20, U48)

- **Veredito:** PROCEDE_PARCIAL. Segue o ADR-078 §6 ("só é guardado se o lote que sai é o lote em uso"), mas com o
  lote em uso desatualizado quem trabalha no lote corrente perde a escolha toda vez. O motivo do ADR ("não guardar o
  filtro de outro lote") se resolve guardando o filtro **junto com o lote de origem**.
- **Prioridade final:** Médio.
- **Correção proposta:** `mLotes` (ajuste do ADR-077/078): antes de **qualquer** ajuste automático, guardar em
  Eng_Saida!AA1 "lote de origem|De|Até|K3|K4|equipamento|analito" (hoje só quando a origem é o lote em uso). Ao
  escolher em H3 o lote de origem, devolver esse filtro (a rede de segurança do ADR-077 continua depois). Mexer no
  filtro à mão continua esquecendo o guardado (`EsquecerFiltroRotina`, inalterado). Atualizar o GUIA §4b: "Ao voltar
  ao lote de onde saiu, o filtro de antes volta sozinho."
- **Autorização:** não (comportamento de tela; o Orquestrador registra a emenda ao ADR-078 §6).
- **Critério de aceite:** Bio 8977 Out → 8974 (ajusta Tudo) → 8977 volta a "Out" (De/Até de outubro); Hema 619611
  Set → 522611 → 619611 volta a Set; partindo do lote em uso, igual a hoje; `qa_varredura_lotes` V01–V04 passam.

### G27: Lista do H3 com os 22 lotes de todos os analitos; escolher lote de outro analito troca o analito sozinho (Bio; U24)

- **Veredito:** PROCEDE. Só 4 dos 22 lotes têm Glicose no DIM 1; um clique errado troca analito e período, e o aviso
  fica abaixo dos gráficos, fora da vista.
- **Prioridade final:** Médio.
- **Correção proposta:**
  1. **Lista do H3 por analito e equipamento:** um intervalo oculto (VBA, com o cache por versão do banco do
     `NResultadosLote`, ADR-050) com os lotes que têm resultado do analito no equipamento **mais** os lotes com
     Média/DP cadastrada para o analito na LotesStore (o lote novo, "AGUARDANDO 1ª CORRIDA", do ADR-078, tem de
     continuar escolhível). Refeito na troca de analito/equipamento. Validação do H3 `=lstLotesAnalito`, estilo
     **Aviso**: "Este lote não tem resultado de Glicose no DIMENSION 1. Clique em Sim para abrir mesmo assim."
  2. **Antes de trocar o analito** (ajuste do ADR-077), `vbYesNo` padrão **Não**: "O lote 8588 não tem resultado de
     Glicose. Ele tem resultados de Hemoglobina glicada. Trocar o analito do Painel para Hemoglobina glicada?" Não →
     H3 volta ao lote anterior, nada muda.
  3. O H4 passa a trazer o aviso curto ("Ajustado: analito → Hemoglobina glicada; período → Tudo") em vez de "ver o
     aviso abaixo dos gráficos".
- **Autorização:** não (tela; o Orquestrador registra a emenda ao ADR-077: a troca de analito passa a pedir
  confirmação).
- **Critério de aceite:** Glicose DIM 1 → a lista do H3 tem só os lotes com Glicose no DIM 1 e os cadastrados para a
  Glicose; digitar 8588 → aviso; Sim → confirmação da troca de analito; Não → H3 volta ao 8977 e C3/K4 não mudam;
  `qa_varredura_lotes` V01–V04 passam.

### G23: Depois de digitar De/Até, Ano (K3) e Período (K4) mantêm o rótulo antigo (Bio+Hema; U19, U49)

- **Veredito:** PROCEDE. A tela diz "S1" e mostra setembro–outubro. Quem imprime leva o período errado.
- **Prioridade final:** Médio.
- **Correção proposta:** acrescentar "Personalizado (De/Até)" ao fim da lista `PERIODOS` (`painel.py`). No
  `Worksheet_Change` do Painel (Bio e Hema), quando o alvo cruza G3:G4 (digitação do usuário; o
  `AplicarFiltroAnoMes` já escreve G3/G4 com os eventos desligados): com os eventos desligados, K4 = "Personalizado
  (De/Até)" e K3 = ano de G3 se G3 e G4 forem do mesmo ano, senão K3 vazio. `AplicarFiltroAnoMes`: rótulo que começa
  com "Personalizado" sai sem tocar em G3/G4.
- **Autorização:** não.
- **Critério de aceite:** K4 "S1" + digitar 01/09–31/10/2026 → K4 "Personalizado (De/Até)", K3 2026, gráfico e
  números iguais aos de antes da correção; escolher "Out" depois → G3/G4 de outubro.

### G24: LJ com linha suavizada, sem valores no eixo Y e sem rótulo nas linhas de DP; no H3 vazio, os degraus viram ondas (Bio+Hema; U21, U50)

- **Veredito:** PROCEDE. A suavização desenha valores que não existem (a linha passa acima dos pontos) e transforma
  os degraus do ADR-086 em senoides; sem números no eixo e sem ±k DP não dá para ler a quantos DP está o ponto, que é
  a função do Levey-Jennings. Os números estão certos (G-PONTOS/G-LIMITES aprovados): é só desenho.
- **Prioridade final:** Médio.
- **Correção proposta:** instalador de formato dos gráficos do Painel (e da tela "Ver todos os gráficos", se usar os
  mesmos):
  - `Series.Smooth = False` em todas as séries (Resultado e as 7 de limite de cada nível);
  - eixo Y: `TickLabelPosition = xlTickLabelPositionNextToAxis`, `TickLabels.NumberFormat` com as casas do analito
    (Calc `dec`), fonte Segoe UI 7, cinza;
  - rótulo na ponta direita de cada linha de limite: `DataLabel` no último ponto com valor (no modo normal, a borda
    da linha 185; no modo TODOS, o último degrau), textos "+3DP", "+2DP", "+1DP", "Média", "−1DP", "−2DP", "−3DP",
    `Position = xlLabelPositionRight`, Segoe UI 7, cor da própria linha;
  - `mUI.AjustarGraficos` continua encaixando os gráficos (a área de plotagem encolhe com os rótulos; conferir).
- **Autorização:** não (formato). Avisar o usuário: o ADR-078 registrou o pedido de "manter a formatação do gráfico".
- **Critério de aceite:** nas imagens exportadas (`Painel_out_8977`, `Painel_set_619611`, `Painel_todos_*`), linha
  reta entre os pontos, degraus retos no H3 vazio, números no eixo Y e os 7 rótulos legíveis; `auditar_aba` sem
  "eixo cortando valores" nem sobreposição.

### G28: "Últ. violação" e "Classificação" sob "violações no período" vêm do histórico do lote (Bio+Hema; U25, U47)

- **Veredito:** PROCEDE. Com 0 violações no período, a linha mostra "6x · Erro Sistemático" de agosto; no H3 vazio,
  de um lote que não está na tela. O usuário investiga a corrida errada.
- **Prioridade final:** Médio.
- **Correção proposta:**
  1. *(sem autorização)* separar o título: L5:R5 "WESTGARD — violações no período" e S5:U5 "Histórico do lote";
     cabeçalhos S6 "Últ. violação do lote" e T6 "Classificação (lote)" (quebra de linha; altura da linha 6 sem
     mudar). No modo TODOS, S7:T9 vazios (`IF(engTodos,"",…)` sobre `engPainel` col. 19/20).
  2. *(com autorização: Especialista)* calcular a última violação **do período** a partir dos pontos do motor, como
     M:Q (dúvida 7 da F4 Bio; dúvida 3 da F4 Hema).
- **Autorização:** item 1 não; item 2 sim.
- **Critério de aceite:** WBC 619611 Set/2026 → S/T sob "Histórico do lote", com "violações no período" = 0 ao lado;
  H3 vazio → S/T vazios; `auditar_aba` com 0 achados.

### G22: Visões MENSAL/TRIMESTRAL da Estatística inalcançáveis e quebradas, "Ano 0" (Bio; U18)

- **Veredito:** PROCEDE. Campos que parecem editáveis e não são, e "Ano 0" passa a impressão de erro. O cálculo com
  ano e parte válidos está certo.
- **Prioridade final:** Médio.
- **Correção proposta:** D4 = `=Painel!K3` (o I3 do Painel é a célula antiga do ano, anterior ao ADR-054). Rótulo da
  Visão (A4, a confirmar) = "Visão (segue o Painel)". "Período efetivo" (H5:J5): quando a janela vier vazia,
  escrever "histórico inteiro do lote" em vez de deixar vazio. B4 continua "PERSONALIZADO" e bloqueada.
- **Autorização:** não.
- **Critério de aceite:** a Estatística mostra "Ano 2026 (refletido do Painel)"; nenhum "Ano 0"; Período efetivo =
  G3:G4 do Painel; com G3:G4 vazios, "histórico inteiro do lote".

### G30: Ano De/Até da Estatística corta o lote sem aviso; Até menor que De aceito (Hema; U43, U51)

- **Veredito:** PROCEDE. Trocar para o lote corrente sem mexer no ano zera a Estatística ("o sistema perdeu os
  dados"); 533811 com 2025 dá n = 8 em vez de 76. "Ano Até 2025" com "Ano De 2026" usa só 2026 sem dizer.
- **Prioridade final:** Médio (U51 sozinho seria Baixo).
- **Correção proposta:**
  1. D3: validação personalizada `=D3>=B3`, estilo Parar, mensagem "Ano Até deve ser igual ou maior que Ano De.".
  2. No `AtualizarEstatisticaAba` (VBA, contagem com cache): se o lote de B4 tem corridas fora de B3..D3, escrever no
     início do texto de F4, em vermelho: "⚠ O lote 619611 tem corridas em 2026; o filtro de ano (2025–2025) não
     cobre o lote: n = 0. Ajuste Ano De/Até." (com os anos e o n reais). Sem alterar o filtro do usuário.
  3. *(com autorização)* o ET usar o lote inteiro, sem o filtro de ano: Especialista (OC-H5).
- **Autorização:** itens 1 e 2 não; item 3 sim.
- **Critério de aceite:** D3 = 2025 com B3 = 2026 → recusado com a mensagem; lote 619611 com Ano 2025 → aviso
  vermelho com "n = 0"; com 2026, sem aviso; `auditar_aba` da Estatística com 0 CORTADO.

### G25: "REJEITADO" cortado em J7:J8 (Bio; U22)

- **Veredito:** PROCEDE. 48 pt de texto em 47 pt. O ajuste do `visual_telas` não pegou porque depende do dado (na
  instalação o status era OK).
- **Prioridade final:** Baixo.
- **Correção proposta:** `ShrinkToFit = True` fixo em J7:J9 (Bio e Hema) no instalador, sem mexer na largura (há
  gráficos e botões alinhados às colunas). Medir no `auditar_aba` todos os textos de status do GUIA §4b ("SEM
  CORRIDA DESTE ANALITO", "AGUARDANDO 1ª CORRIDA" etc.) escritos à força na cópia; se algum precisar de mais de 30 %
  de redução, usar a forma curta já existente ou quebra.
- **Autorização:** não.
- **Critério de aceite:** com cada status escrito em J7, `auditar_aba` com 0 CORTADO e redução ≤ 30 %.

### G26: Nome do analito no pé da célula na Estatística (Bio; U23)

- **Veredito:** PROCEDE. Nas linhas altas, o nome parece da linha de baixo.
- **Prioridade final:** Baixo.
- **Correção proposta:** `VerticalAlignment = xlCenter` em A14:B133 da Estatística (Bio e Hema), no passo de grades
  do `visual_telas`.
- **Autorização:** não.
- **Critério de aceite:** na imagem `Estatística.png`, Lactato N2, Ácido úrico N2 e Cálcio com o nome na altura dos
  números.

### G29: Ponto de violação cortado na borda superior do LJ (Bio; U26)

- **Veredito:** PROCEDE. No modo TODOS, a moldura do eixo Y usa os degraus com margem de 5 %, mas os pontos sem
  margem (`instalar_adr086.eixo_y`): o máximo do eixo fica igual ao ponto de 1_3s, justamente o mais importante, que
  sai metade para fora e fica difícil de dar duplo clique.
- **Prioridade final:** Baixo.
- **Correção proposta:** Calc linha 1, moldura do eixo Y: nos termos dos dados, `AGGREGATE(4,6,d)+0.05*faixa` (máx.) e
  `AGGREGATE(5,6,d)-0.05*faixa` (mín.), com `faixa` = amplitude já usada. No modo normal, a mesma margem quando um
  ponto passa de Média ± 3,3·DP. Instalador idempotente (a 2ª execução dá 0 mudanças, como o ADR-086).
- **Autorização:** não (formato do eixo).
- **Critério de aceite:** HbA1c DIM 1 jun/2026 H3 vazio → `eixo.y_max` > 10,53 (≥ 10,53 + 5 % da amplitude) e o
  marcador inteiro dentro da área de plotagem; modo normal sem mudança quando todos os pontos estão dentro de
  ±3,3 DP.

---

## Pendências para o Orquestrador

1. **Decisões que travam correções (autorização):** papel para corrigir/excluir CEQ (G09) e para registrar (G15);
   mapa do Controllab e tela de mapeamento (G11); alvo obrigatório e contagem em S (G13); EQUIPAMENTO obrigatório na
   Bio (G01); manual fora do cadastro como INCOMPLETO (G02); critério do Status (G07); última violação do período
   (G28); Estatística vazia no modo TODOS (G19); ET do lote inteiro (G30).
2. **Emendas de tela a registrar:** ADR-078 §6 (G20) e ADR-077 (G27). Não mudam regra de negócio.
3. **Ação do laboratório:** atualizar o lote em uso nos dois QCs (G21), que também reduz G19, G20 e G27.
4. **Ligação com ocorrências de outras rotas:** G02 com a OC-9/OC-HF1-1 (minúsculas, Crítico); G19 com a OC-HF4-1
   (Estatística não recalcula na troca de lote, Crítico).
5. **Varredura sugerida ao Tester:** todo `TEXT(` com código numérico em instaladores e fórmulas (causa de G18 e G32).
