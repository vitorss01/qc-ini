# Parecer UX/UI: ciclo 2026-11-15_reteste06

Agente: avaliador de UI/UX (ADR-081/087). Data: 15/11/2026. Modo: reteste (fluxos vizinhos da OC-0073).

**De onde vêm os dados.** O Excel não foi aberto e nenhum arquivo foi alterado além deste parecer. Li as regras comuns,
os relatórios das frentes (`Bioquimica_F1.md`, `Hematologia_F1.md`, `00_PARECER.md`), o ADR-096, o ADR-070, o ADR-087
(critério "nada cortado nem sobreposto"), o GUIA §3, `camada_dados.py` (fórmula do DETALHE e instruções da aba) e,
só para leitura, as fontes citadas nas evidências: `mIntegracao.PrepararLinha`/`Carimbar`/`EntradaMudou`,
`mEtlQA.TextoA01`, `mIntegracao.ResumoQA`, `mEtlCQ` (teste "DATA/HORA posterior ao lancamento") e
`mEstatistica.AtualizarEstatisticaAba` da Hema. Produção auditada pelas frentes: Bio bd5314426ab8, Hema 49f68fba99d5;
base (entrega 42): Bio b20fe00f7e51, Hema 7bab8fc008f6.

**Consolidação.** As 4 ocorrências da Bio (R6-01 a R6-04) e as 4 confirmações da Hema são o mesmo defeito, com a mesma
causa e o mesmo código (`src/bio` e `src/hema` têm o mesmo `mIntegracao`). Proponho registrar **uma OC "Ambos" para cada
par**. A R6H-05 é a OC-0048 (G30 do ciclo01), já aberta: não abrir OC nova, juntar a evidência.

## Resumo

| Item | Veredito | Prioridade final | Autorização | Onde corrige |
|---|---|---|---|---|
| R6-01 (Bio+Hema) carimbo residual na linha apagada ou movida | PROCEDE | **Alto** | Não (a regra do motor não muda) | `mIntegracao.PrepararLinha` (dois produtos) |
| R6-02 (Bio+Hema) copiar a linha inteira esbarra na proteção | PROCEDE_PARCIAL | Médio | Não | GUIA §3, ADR-096; opcional: mensagem de entrada no ID |
| R6-03 (Bio+Hema) ID inválido: DETALHE "atualize os dados" | PROCEDE | Médio | Não | `camada_dados._formula_detalhe` (só a tela) |
| R6-04 (Bio+Hema) colar e trocar só o resultado; A01 "inative um deles" | PROCEDE | **Médio** (sobe de Baixo) | Não para os textos; Sim se quiserem mudar a chave | aviso do `EntradaMudou`, `TextoA01` + `QA_INTEGRACAO.m`, `ResumoQA` |
| R6H-05 (Hema) Estatística abre em 2025, n 0 | PROCEDE (duplicata da OC-0048) | Médio | Não (aviso); Sim só para o ET ignorar o ano (G30 item 3) | `AtualizarEstatisticaAba` da Hema + validação do D3 |

Ordem sugerida: R6-01 primeiro (é o único que tira do CQ um resultado válido sem saída para o usuário), depois R6-03 e
R6-04 juntos (são textos da mesma aba e da mesma mensagem), R6-02 (documentação), R6H-05 com a OC-0048.

---

## R6-01: apagar ou mover uma linha manual deixa USUARIO e REGISTRADO_EM (Bio e Hema)

- **Veredito:** PROCEDE. É defeito, não preferência.
  - O `Carimbar` só preenche o vazio, e o `PrepararLinha` sai sem fazer nada quando a linha fica sem ID e sem dado. Com
    L:M travadas, o usuário não consegue apagar o carimbo.
  - O botão Digitar Resultados leva justamente para essa linha (a DATA está vazia). O caminho é o mais natural possível:
    apagar o erro hoje e digitar o manual de amanhã onde o botão parar.
  - O motor recusa com "DATA/HORA posterior ao lançamento", e o usuário olha a data/hora, que está certa. Ninguém na
    bancada vai deduzir que o problema é um carimbo travado de outro dia.
  - A trilha fica errada: "manual lancado por <outro usuário> em <outro dia>".
  - O ADR-096, o A01 e o GUIA mandam "apague a linha", o que aumenta a frequência.
  - O próprio motor já trata REGISTRADO_EM vazio como "agora" (`If IsEmpty(reg) Then regc = agora`), e a aba Inativar
    já limpa o carimbo ao apagar (ADR-070). Limpar o carimbo da linha vazia é aplicar a mesma regra, não criar outra.
- **Prioridade final:** Alto (risco operacional relevante: resultado de CQ válido fora da estatística, do LJ e do
  Westgard, sem saída visível; trilha atribuída a outra pessoa).
- **Correção proposta (implementação, dois produtos):**
  1. `PrepararLinha`, ramo `TB_MANUAL`: antes do `Exit Sub` de "sem dado e sem ID", se também o `SUBSTITUI_ID` estiver
     vazio, apagar (`ClearContents`) USUARIO e REGISTRADO_EM da linha. "Sem dado" = ID_REGISTRO, DATA, HORA, LOTE,
     NIVEL, ANALITO, RESULTADO e SUBSTITUI_ID vazios. EQUIPAMENTO, MATRIZ, UNIDADE e MOTIVO não contam (os dois
     primeiros são preenchidos sozinhos e não fazem um lançamento).
  2. Como o `EntradaMudou` chama o `PrepararLinha` para cada linha tocada, o Delete de A:K, o recorte (a origem) e o
     Delete com várias áreas (Ctrl) já ficam cobertos. O `PrepararEntradas` cobre os resíduos que já existem nos
     arquivos.
  3. Na entrega, o Tester conta numa cópia da produção quantas linhas vazias têm carimbo antes de instalar e anota o
     número no ADR. Não é registro de resultado (a retirada do manual já está no Audit_Log como
     RESULTADO_MANUAL_RETIRADO), mas a limpeza fica documentada.
  4. Complemento (mesmo ponto, recomendado): quando o `EntradaMudou` dá ID novo a uma linha colada (ADR-096), apagar
     USUARIO e REGISTRADO_EM antes do `PrepararLinha`, para a linha colada nunca herdar o carimbo da origem (hoje só
     acontece pela COM, T15, mas é o mesmo erro de trilha).
  - Não muda: a regra "DATA/HORA posterior ao lançamento", o carimbo de linha com dado (corrigir o valor de um manual
    mantém o carimbo e continua indo para o Audit_Log como RESULTADO_MANUAL_ALTERADO), as colunas travadas.
  - Armadilhas: escrever com `LiberarEscrita`/`RestaurarProtecao` como o resto do evento; `EnableEvents` desligado
    durante a limpeza (já é assim no `EntradaMudou` e no `PrepararEntradas`).
- **Autorização:** não. É a leitura do ADR-070 aplicada à aba manual; a regra de negócio não muda. A dúvida do Usuário
  Simulado ("o REGISTRADO_EM é desta linha ou do primeiro carimbo da célula?") tem resposta pela trilha: é desta linha.
  Se o Orquestrador quiser, o Especialista em CQ confirma sem bloquear a correção.
- **Critério de aceite (reteste, Bio e Hema, com `reteste06_carimbo_residual.py` e o irmão da Hema):**
  1. Digitar A, apagar A:K: na mesma linha, USUARIO e REGISTRADO_EM ficam vazios (`Value2` vazio).
  2. Botão Digitar Resultados para nessa linha; digitar B: USUARIO = usuário da sessão atual e REGISTRADO_EM ≥ instante
     da digitação de B.
  3. ATUALIZAR: B fica ATIVO com RUN, sem E04, entra no Painel (n +1) e no Eng_Saida; Audit_Log "manual lancado por
     <usuário de B> em <instante de B>". O controle C continua ATIVO.
  4. Recortar A:K para outra linha: a origem fica sem carimbo; a linha de destino tem carimbo; o ID se mantém.
  5. Delete com duas áreas (Ctrl) de duas linhas: as duas ficam sem carimbo.
  6. Linha só com o ID apagado (dados mantidos): ganha ID novo e **mantém** o carimbo (não é linha vazia).
  7. Cópia com um carimbo órfão injetado numa linha vazia: depois do `PrepararEntradas` (ATUALIZAR), ele some; as linhas
     com dado têm USUARIO/REGISTRADO_EM idênticos aos de antes (comparar todas).
  8. Base (entrega 42) com o mesmo roteiro continua reprovando (prova de que o teste pega o defeito).
  9. Regressão: `qa_adr095` P01, `qa_inativar` N01-N12, `qa_manual_substitui`, `qa_etl` sem falha.

## R6-02: copiar a linha inteira é recusado pela proteção (Bio e Hema)

- **Veredito:** PROCEDE_PARCIAL.
  - A recusa é o comportamento correto: USUARIO, REGISTRADO_EM, RUN, STATUS e DETALHE são do sistema e têm de ficar
    travados. Destravar L:P para a colagem passar abriria a trilha à edição (e pela COM a linha inteira já mostrou que
    leva o carimbo da origem). Isso não muda.
  - O que procede é a instrução: o GUIA §3 diz "copiar uma linha e colar embaixo", e o gesto natural (clicar no número
    da linha, Ctrl+C, Ctrl+V) dá a caixa nativa do Excel que fala em "planilha protegida" e "senha". Quem está na
    bancada lê "senha", acha que não tem permissão e desiste do atalho, ou chama alguém. A mensagem nativa não pode ser
    trocada pelo VBA, então a correção é ensinar o gesto certo.
- **Prioridade final:** Médio (o GUIA leva a um erro com texto que assusta; sem risco para os dados).
- **Correção proposta:**
  1. GUIA §3, passo 1, trocar a frase "Pode copiar uma linha e colar embaixo para aproveitar a digitação: …" por:
     > Para aproveitar a digitação, selecione a linha **de ID até MOTIVO** (arraste do ID ao MOTIVO), copie (Ctrl+C) e
     > cole (Ctrl+V) na coluna ID da próxima linha livre. A linha colada **ganha um ID novo** sozinha (um aviso diz
     > qual): troque o que for diferente antes de atualizar (ADR-096). Não copie a linha inteira (clicando no número da
     > linha): as colunas de USUARIO a DETALHE são do sistema, e o Excel recusa com um aviso de planilha protegida. Não
     > é preciso senha: basta copiar de ID até MOTIVO.
  2. ADR-096 §1: acrescentar a mesma regra ("copiar de ID até MOTIVO; a linha inteira é recusada pela proteção, de
     propósito").
  3. Opcional (pequeno, recomendado): mensagem de entrada (validação "Qualquer valor", só `InputTitle`/`InputMessage`,
     sem lista e sem nome definido) nas células de ID_REGISTRO da tabela, aplicada pelo instalador com a aba
     desprotegida e reprotegida:
     - título: `Copiar uma linha`
     - texto: `Para repetir uma linha: selecione de ID até MOTIVO, copie e cole no ID da próxima linha livre. A linha inteira não cola (colunas do sistema).`
     Se a coluna já tiver validação, só acrescentar a mensagem de entrada, sem trocar o tipo.
  - Não mexer nas instruções A3/A4 da aba: já são longas, sem quebra, e o ADR-087 proíbe texto cortado.
- **Autorização:** não.
- **Critério de aceite:**
  1. O GUIA §3 e o ADR-096 têm o texto acima (conferência por texto).
  2. Com a proteção do usuário (sem UserInterfaceOnly), colar A:K continua ACEITO e a linha inteira continua RECUSADA
     (`P1c_protecao_do_usuario`): a proteção não pode ter sido afrouxada.
  3. Se o item 3 for feito: a célula do ID mostra a mensagem de entrada; `auditar_aba` da Digitar Resultados com 0
     CORTADO e 0 sobreposto; as colagens P1/P1b/P1c continuam com ID novo e o aviso do ADR-096.

## R6-03: ID fora do padrão mostra "atualize os dados" no DETALHE (Bio e Hema)

- **Veredito:** PROCEDE. A mensagem do ATUALIZAR manda "veja o DETALHE", e o DETALHE manda "atualize os dados": o
  usuário entra num ciclo sem saída. O motor está certo (MAN_INVALIDO_L<n>, E04), mas o XLOOKUP da tela procura o ID
  digitado, que nunca vai existir na tblCQ_Final. Ainda pior: o motivo só está na QA_INTEGRACAO, que o usuário de
  bancada não abre. O caso real é digitar ou colar o número do interfaceamento no ID.
- **Prioridade final:** Médio.
- **Correção proposta (só a fórmula da tela; `camada_dados._formula_detalhe`, reinstalada na coluna DETALHE dos dois
  produtos):** trocar o valor do `IFERROR` final. Hoje:
  `IFERROR(<x>,"atualize os dados")`. Proposto:

  ```
  =IF([@ID_REGISTRO]="","",IFERROR(<x>,
     IF(AND(LEFT(UPPER(TRIM([@ID_REGISTRO])),4)="MAN_",ISNUMBER(-MID(TRIM([@ID_REGISTRO]),5,20))),
        "atualize os dados",
        "ID fora do padrão: apague-o (um novo é gerado)")))
  ```

  - O texto tem 46 caracteres e cabe na largura do DETALHE (50), sem depender do transbordo para o SUBSTITUI_ID.
  - Aparece já na digitação (antes do ATUALIZAR), o que é melhor ainda: o usuário corrige na hora.
  - STATUS e RUN continuam "—": não inventar um status que o motor não grava.
  - O ID inválido continua não sendo trocado sozinho (decisão do ADR-096: trocar mudaria a regra). Apagar o ID gera o
    MAN_ novo pelo `PrepararLinha`, como já acontece.
  - O texto já tem acento: não passa pela lista `ACENTOS_DETALHE` e não fere a trava `SEM_ACENTO_*` do `qa_adr095`
    (é literal da tela, não do motor).
  - Sem `Names.Add`; a coluna é calculada da tabela e o instalador já a reescreve com a aba desprotegida.
- **Autorização:** não.
- **Critério de aceite:**
  1. Duas linhas com ID `TESTEAUD9` (P6a/P6b): antes e depois do ATUALIZAR, DETALHE = "ID fora do padrão: apague-o (um
     novo é gerado)"; STATUS e RUN "—"; motor igual (MAN_INVALIDO_L13/L14, MANUAL_INCOMPLETO, 2 E04); a mensagem do
     ATUALIZAR continua "veja o DETALHE".
  2. Apagar o ID das duas: ganham MAN_ novos; ATUALIZAR: ATIVO com RUN; DETALHE com o texto do motor.
  3. ID `MAN_0099` válido ainda não atualizado: DETALHE "atualize os dados" (não muda).
  4. `man_7` e ` MAN_0007 ` (espaços): tratados como válidos (o `PrepararLinha` já normaliza).
  5. `auditar_aba` da Digitar Resultados: o DETALHE sem CORTADO; `qa_adr095` D01 e `qa_adr090` sem falha.

## R6-04: colar e trocar só o resultado; aviso e A01 enganam (Bio e Hema)

- **Veredito:** PROCEDE.
  - A regra (mesma chave e mesmo instante = repetido, vale o primeiro) é do motor e não está em discussão aqui.
  - O aviso do ADR-096 lista "nível, hora, resultado" como coisas que basta mudar; trocar só o resultado não basta. O
    usuário segue a tela e perde o resultado.
  - O A01 desse ramo ("confirme e inative um deles") é o mesmo tropeço que o ADR-096 corrigiu no ramo do ID: a seta
    aponta para o MAN_0001, que é o que vale. Inativá-lo tira o resultado bom do CQ.
  - A mensagem do ATUALIZAR não conta essas linhas: só quem olha o STATUS linha a linha descobre.
  - Na Hema, repassar o mesmo controle em seguida é rotina; o caso não é raro.
- **Prioridade final:** Médio. Subo de Baixo: o texto de governança orienta uma ação (inativar) que pode tirar da
  estatística o resultado válido, e o aviso induz à perda do resultado repetido. É problema de usabilidade com efeito
  operacional, não acabamento.
- **Correção proposta (textos; a regra não muda):**
  1. Aviso do `EntradaMudou` (colagem com ID novo), trocar a última frase
     "Mude o que for diferente (nível, hora, resultado) antes de clicar em ATUALIZAR DADOS." por:
     > Mude o que for diferente antes de clicar em ATUALIZAR DADOS. Atenção: com a mesma data, hora, nível e analito de
     > outra linha, ela fica como repetida (CONFLITO_MANUAL), mesmo com outro resultado. Se for uma nova dosagem, use a
     > hora real dela.
  2. `mEtlQA.TextoA01`: novo `ElseIf` antes do `Else`, só para o motivo que começa com
     "resultado manual repetido" (o `Else` genérico fica como está para os outros motivos). Texto (sem acento, como os
     outros literais do motor; `rel` = ID que vale, `idLinha` = ID desta linha):
     > `; resultado repetido na aba Digitar Resultados (mesma data, hora, nivel e analito de <rel>): se for outra dosagem, corrija a HORA da linha <idLinha> para a hora real; se for digitacao dobrada, apague a linha <idLinha>. Nao inative o <rel>: isso tiraria o resultado que vale.`
     O mesmo texto em `pq/QA_INTEGRACAO.m` (o `qa_equiv_qa_adr094` compara os dois).
  3. `ResumoQA`: contar esses A01 (como já conta o ID repetido desde o ADR-096, mesmo sem ERRO) e acrescentar:
     > • N linha(s) da Digitar Resultados repete(m) data, hora, nível e analito de outra e ficou(aram) fora da
     > estatística (CONFLITO_MANUAL). Se for outra dosagem, corrija a HORA; se for digitação dobrada, apague a linha.
  4. GUIA §3, no item CONFLITO_MANUAL: acrescentar "ou repete data, hora, nível e analito de outro manual (vale o
     primeiro): corrija a HORA se for outra dosagem, ou apague a linha repetida. Não inative o outro."
- **Autorização:** não para os textos. **Sim** se o laboratório quiser que dois manuais no mesmo minuto com valores
  diferentes valham os dois (muda a chave do motor): vai ao Especialista em CQ e ao usuário (dúvida 2 dos relatórios).
- **Critério de aceite:**
  1. P2c (colar e trocar só o RESULTADO): o aviso tem a frase "mesmo com outro resultado" e "use a hora real"; a
     linha continua CONFLITO_MANUAL (regra intacta).
  2. QA_INTEGRACAO: o A01 do MAN_0006 (e dos colados sem edição) tem "Nao inative o MAN_0001" e "corrija a HORA da
     linha MAN_0006"; nenhum A01 desse ramo contém "inative um deles".
  3. A mensagem do ATUALIZAR tem o bullet, e o N é igual ao número de linhas da tblCQ_Final com "resultado manual
     repetido (mesma chave e instante)" no mesmo momento (no 1º ATUALIZAR da Hema, MAN_0002/3/4/6/11/12/13 = 7).
  4. Seguindo o texto: corrigir a HORA do MAN_0006 para outro minuto → ATIVO com corrida própria; apagar a linha →
     sai da tblCQ_Final com RESULTADO_MANUAL_RETIRADO; MAN_0001 continua ATIVO.
  5. `qa_equiv_qa_adr094` (VBA = PQ) e `qa_etl` E09d/e/f sem falha; os outros ramos do A01 ("coincide…", "ID_REGISTRO
     manual repetido") com o texto de antes.
  6. Caixa do aviso inteira (texto sem corte na MsgBox; conferir pelo texto lido da janela do PID, como a frente fez).

## R6H-05: Estatística da Hema abre com Ano De/Até = 2025 e mostra n 0 (Hema)

- **Veredito:** PROCEDE, e é a **OC-0048** (G30 do `relatorios/2026-10-31_ciclo01/parecer_ux.md`), ainda aberta:
  mesma causa (o filtro de ano corta o lote sem aviso). Não abrir OC nova; anexar esta evidência à OC-0048.
  - Quem confere o lançamento pela aba Estatística vê n 0 e conclui que "não entrou" ou que "o lote não tem dado". O
    Painel está em nov/2026 e a Estatística em 2025, e nada na aba liga uma coisa à outra.
  - Não proponho trocar o ano sozinho: o ano é escolha do usuário (relatório anual de 2025 é uso legítimo), e mudar o
    filtro sem pedir confunde mais.
- **Prioridade final:** Médio.
- **Correção proposta (a mesma do G30, itens 1 e 2, com o texto fechado):**
  1. D3: validação personalizada `=D3>=B3`, estilo Parar, mensagem: `Ano Até deve ser igual ou maior que Ano De.`
     (hoje o código força `anoAte = anoDe` sem dizer).
  2. `AtualizarEstatisticaAba` (Hema): se o lote do B4 tiver resultado fora de B3..D3 e n = 0 dentro, escrever numa
     célula livre do cabeçalho (o G30 sugeriu F4; o implementador escolhe a que o `auditar_aba` aprovar sem corte), em
     vermelho:
     > ⚠ O lote 625311 não tem resultado em 2025–2025 (tem em 2026). Ajuste o Ano De/Até (B3/D3).
     Com o lote parcialmente coberto (n > 0, mas há resultado fora do intervalo):
     > ⚠ O lote 625311 também tem resultado em 2026, fora do Ano De/Até: os números abaixo não são o lote inteiro.
     Com tudo coberto: célula vazia. Os anos e o lote saem dos dados reais.
  3. (com autorização, Especialista em CQ, já registrado no G30) o ET usar o lote inteiro, sem o filtro de ano.
- **Autorização:** não para 1 e 2; sim para 3.
- **Critério de aceite:**
  1. Arquivo como vem (B3 = D3 = 2025, B4 = 625311): aviso vermelho com "não tem resultado em 2025–2025 (tem em 2026)";
     números do bloco iguais aos de hoje (n 0).
  2. B3 = D3 = 2026: sem aviso; números iguais ao oráculo do relatório B (N1 41, N2 64/65, N3 34).
  3. D3 = 2025 com B3 = 2026: recusado com a mensagem.
  4. `auditar_aba` da Estatística da Hema: 0 CORTADO, 0 sobreposto; Bio sem mudança.
