# Parecer UX/UI: reteste 07 (2026-11-16_reteste07)

Data: 17/11/2026 · Agente: AGENTE_UX_UI · Modo: reteste · Base: relatórios `Bioquimica_F1.md`, `Hematologia_F1.md` e
`00_PARECER.md` deste ciclo, ADR-097, GUIA §3 e o código atual de `src/{bio,hema}/mIntegracao.bas` (leitura apenas).
Não abri Excel, não editei código e não capturei tela. As evidências do Usuário Simulado bastaram.

**Consolidação.** A R7-01 da Bio e a da Hema são o mesmo defeito, no mesmo código (`mIntegracao` é igual nos dois
produtos) e com os mesmos números de comportamento. Proponho **uma OC "Ambos"** (a P-01 do parecer do Orquestrador).

## Resumo

| Item | Veredito | Prioridade final | Autorização | Onde corrige |
|---|---|---|---|---|
| R7-01 Bio: recortar e colar um manual integrado troca o dono sem rastro | PROCEDE | Médio | Não (fechar a D-01 com "vale quem lançou") | `mIntegracao.PrepararLinha` (dois produtos), GUIA §3, ADR-097 §5 |
| R7-01 Hema: idem | PROCEDE (mesma OC, "Ambos") | Médio | Idem | Idem |

---

## R7-01 (Bio e Hema): recortar e colar um manual já integrado troca o USUARIO_MANUAL para quem moveu

### Veredito: PROCEDE

É um defeito de trilha, não uma preferência.

- **Para o usuário real, mover uma linha não é lançar.** Quem recorta uma linha na bancada quer arrumar a tabela, ou
  tirar a linha do meio de outras, ou juntar as do dia. Ninguém espera que isso transfira para si a autoria de um
  resultado de CQ que outra pessoa dosou e digitou. Se perguntado depois numa auditoria ("quem lançou o MAN_0003?"),
  o sistema dá o nome errado e não há linha no Audit_Log que permita reconstituir a verdade.
- **O usuário não tem como perceber.** USUARIO e REGISTRADO_EM são travados e ficam fora da seleção ID…MOTIVO. Nenhum
  aviso aparece. Na tela tudo continua ATIVO, com o mesmo valor. É o pior tipo de erro de trilha: silencioso.
- **O próprio ADR-097 §1 diz que "o carimbo pertence ao lançamento daquela linha".** O §5 contradiz esse princípio para
  o recorte, e o critério 3 da OC-0078 seguiu o §1. A causa é mecânica, não uma decisão: o destino chega com o ID e
  com os dados, mas sem carimbo (colunas fora da seleção). O `LancamentoNovo` corretamente responde "não é novo", e o
  `Carimbar`, que só preenche o vazio, grava quem está logado.
- **A correção é viável sem tocar no motor.** A tblCQ_Final já guarda, para cada manual integrado, o USUARIO_MANUAL e
  o RECEBIDO_EM (= REGISTRADO_EM da digitação, comentário da linha 501 do `mIntegracao`). Um ID MAN_ nunca é
  reaproveitado (D05, `ProximoIdManual`), então um MAN_ que chega numa linha e já existe na tblCQ_Final é o mesmo
  lançamento. E o `mIntegracao` já tem a busca pronta (`LinhaCQ(AcharTabela(FONTE_CQ), id)`, usada na aba Inativar).
- **Não é efeito colateral.** A base (entrega 43) faz igual, e o resto da OC-0078 está resolvido.

### Prioridade final: Médio

Mantenho Médio. A trilha fica errada e sem rastro, mas não muda número, status, LJ nem Westgard, e mover uma linha
manual é raro. Não sobe para Alto porque nenhum resultado sai do CQ e nada trava o trabalho. Não desce para Baixo porque
é a atribuição de um resultado de CQ, que a auditoria da qualidade pode cobrar.

### Correção proposta

Três partes, nesta ordem de valor. A 1 resolve, a 2 explica para o laboratório, a 3 alinha os documentos. Nenhuma muda
fórmula, critério estatístico, motor (`mEtlCQ`), proteção da aba ou cálculo.

**1. O destino recupera o carimbo do lançamento original (`mIntegracao.PrepararLinha`, ramo `TB_MANUAL`, Bio e Hema).**
Logo antes do `Carimbar Celula(lo, r, "REGISTRADO_EM"), Celula(lo, r, "USUARIO")`:

- condição (todas):
  1. USUARIO **e** REGISTRADO_EM da linha estão vazios;
  2. o ID **não** foi gerado nesta chamada (ele já estava na linha ao entrar no `PrepararLinha`, ou seja, veio com a
     colagem/recorte ou foi digitado);
  3. o ID é MAN_ no padrão (`IdManualPadrao`) e não está repetido em outra linha da tabela manual (o ramo do `idRep`
     do ADR-096 já trata a cópia: ela ganha ID novo e carimbo de quem colou, e não chega aqui);
  4. `LinhaCQ(AcharTabela(FONTE_CQ), id) > 0` e, nessa linha da tblCQ_Final, USUARIO_MANUAL e RECEBIDO_EM estão
     preenchidos;
- ação: gravar REGISTRADO_EM = RECEBIDO_EM e USUARIO = USUARIO_MANUAL daquela linha da tblCQ_Final. Se qualquer uma
  das condições falhar, segue o `Carimbar` de hoje (quem está logado, agora);
- sem mensagem ao usuário. Não há nada que ele precise decidir: a linha simplesmente continua sendo de quem a lançou;
- armadilhas: escrever dentro do `LiberarEscrita`/`RestaurarProtecao` e com `EnableEvents` desligado, como o resto do
  evento (as duas colunas são travadas); sem `Names.Add`; nenhuma fórmula de tela muda. A busca roda só nas linhas que
  chegam sem carimbo e com ID, então não pesa na colagem de muitas linhas nem na passada do ATUALIZAR;
- efeito no `PrepararEntradas` (ATUALIZAR): uma linha com ID integrado e sem carimbo (por exemplo, recortada com os
  eventos desligados) também recupera o dono, em vez de ganhar o de quem atualizou. É o comportamento certo.

**Limite que continua (e passa a ser o único):** recortar um manual **ainda não integrado** (antes do primeiro
ATUALIZAR) leva o carimbo de quem moveu, porque o lançamento original ainda não está em lugar nenhum fora da linha de
origem, que o Excel já limpou. Fica registrado no ADR. O efeito é menor: o Audit_Log ainda não tinha nenhum dono para
esse ID, então nada é trocado em silêncio.

**Cinto de segurança (opcional, recomendado se a parte 1 não for feita agora):** no `AuditarMudancas`/`Registrar`
(`mIntegracao`, trilha, não o motor), incluir o dono do manual no estado comparado, para que uma troca de
USUARIO_MANUAL de um manual existente grave uma linha `RESULTADO_MANUAL_ALTERADO` com a nota
"lancado por <antes> -> <depois>". Isso não corrige a atribuição, mas deixa rastro. Com a parte 1 feita, essa linha
só apareceria no caso-limite acima, e por isso fica opcional.

**2. GUIA §3, passo 1 (texto exato, logo depois de "Para corrigir um manual, corrija os campos na própria linha."):**

> Não é preciso mudar um manual de lugar: a ordem das linhas não muda nada, porque o gráfico e a estatística usam a
> data e a hora. Se você recortar e colar uma linha, o ID vai junto e o lançamento continua no nome de quem o digitou.
> Só o manual ainda não atualizado passa para o nome de quem moveu.

O primeiro período é o que mais ajuda na bancada: tira o motivo para recortar.

**3. Documentos.** Novo ADR (ou adendo do ADR-097) trocando o §5 por: "Linha movida: o destino de um recorte mantém o
carimbo do lançamento original quando o manual já foi integrado (recuperado da tblCQ_Final); antes do primeiro
ATUALIZAR, recebe o de quem moveu (limite)". E o critério 3 da OC-0078 passa a citar esse texto. Assim a D-01 fica
fechada com uma só resposta: **vale quem lançou**.

### Autorização: não

- A parte 1 aplica o princípio que o ADR-097 §1 já declara ("o carimbo pertence ao lançamento daquela linha") e que o
  critério 3 da OC-0078 já pedia. Não muda o cálculo, a regra do motor, a regra "DATA/HORA posterior ao lançamento", as
  colunas travadas nem a trilha das outras ações.
- A D-01 é uma divergência entre dois textos do próprio projeto, e o Orquestrador a fecha. Minha recomendação é "vale
  quem lançou", que é a leitura de atribuição que uma auditoria da qualidade espera.
- **Só dependeria de autorização** a decisão contrária: oficializar que quem move uma linha passa a ser o dono do
  lançamento. Isso redefine a trilha de autoria e iria ao Especialista em CQ e ao usuário.

### Critério de aceite (reteste: Usuário Simulado F1 com `reteste07_usuario_f1.py` ou o irmão do ciclo seguinte; Tester com `qa_adr097`)

Nos dois produtos, com dois usuários de teste (U1 lança, U2 move), dados sintéticos marcados como nos reteste 06 e 07:

1. **Recorte de manual integrado:** U1 digita MAN_x e roda o ATUALIZAR (USUARIO_MANUAL = U1, RECEBIDO_EM = t0). U2
   recorta ID…MOTIVO para uma linha vazia. Logo depois do recorte, o destino tem o mesmo ID, USUARIO = U1 e
   REGISTRADO_EM = t0 (`Value2` igual ao RECEBIDO_EM da tblCQ_Final, diferença ≤ 1 s). A origem fica sem ID e sem
   carimbo (como já é).
2. **ATUALIZAR depois do recorte:** o MAN_x continua com USUARIO_MANUAL = U1 e RECEBIDO_EM = t0. Valor, DATA_HORA,
   status, RUN e ponto do LJ (Eng_Saida) iguais aos do ATUALIZAR anterior. O Audit_Log não ganha linha "manual lancado
   por U2" para o MAN_x. O n e a média do Painel e da Estatística batem com o oráculo, como no reteste 07.
3. **Recortar e depois corrigir:** U2 recorta o MAN_x e troca o RESULTADO no destino. O ID e o carimbo continuam de U1,
   e o Audit_Log grava RESULTADO_MANUAL_ALTERADO com "manual lancado por U1 em <t0>".
4. **Vizinhos que não podem mudar (todos já aprovados no reteste 07):** Ctrl+C/Ctrl+V de ID…MOTIVO dá ID novo e o
   carimbo de quem colou; Ctrl+D mantém o dono da linha de origem e dá IDs novos às de baixo; o Delete de ID…MOTIVO e o
   de DATA…MOTIVO limpam ID e carimbo; o manual digitado numa linha apagada fica com o nome e a hora de quem digitou;
   apagar só o ID de uma linha com dado dá ID novo e mantém o carimbo; o ADR-082 (SUBSTITUI_ID) funciona.
5. **Caso-limite documentado:** recortar um manual ainda não integrado deixa o carimbo de quem moveu, e o ATUALIZAR o
   aceita (ATIVO, sem E04). O ADR descreve exatamente esse caso, e só ele.
6. **Texto:** o GUIA §3 contém o parágrafo da parte 2, com as palavras inteiras e sem cortar a linha na largura do
   arquivo; o ADR novo (ou adendo) substitui o §5 do ADR-097; o critério 3 da OC-0078 no `OCORRENCIAS.md` usa o mesmo
   texto.
7. **Tester:** o caso de recorte do `qa_adr097` (C08) passa a exigir o carimbo original no destino de um manual
   integrado e aceita o de quem moveu só para o não integrado. A regressão direcionada (`qa_adr097`,
   `qa_manual_substitui`, `qa_etl`, `qa_inativar`) passa sem falha nos dois produtos, e as produções ficam intactas
   (hash igual antes e depois).

---

## Observações fora das ocorrências (sem OC)

- **Texto do aviso no Ctrl+D (acabamento, Baixo, só se mexerem no `EntradaMudou` por causa da parte 1):** o aviso diz
  "As 2 linhas **coladas** traziam IDs…" também no Ctrl+D. Proposta: "As 2 linhas copiadas traziam IDs que já existem
  em outras linhas e receberam IDs novos (MAN_0010 a MAN_0011)." O resto da mensagem não muda. Não justifica entrega
  própria.
- **Dúvida 2 do Usuário Simulado (quem aparece no RESULTADO_MANUAL_RETIRADO):** não recebi como ocorrência. Do ponto de
  vista do uso, registrar quem rodou o ATUALIZAR é aceitável enquanto a retirada for "o sistema constatou que a linha
  saiu". Se o laboratório quiser "quem apagou", seria preciso gravar o Delete no momento em que acontece, o que é uma
  mudança de trilha e depende de decisão do Orquestrador.
