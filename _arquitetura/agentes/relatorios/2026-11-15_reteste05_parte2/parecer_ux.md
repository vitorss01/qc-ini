# Parecer UX/UI: ciclo 2026-11-15_reteste05_parte2 (reteste)

**Avaliador:** UX/UI (ADR-081) · **Data:** 15/11/2026 · **Entrada:** 5 ocorrências das frentes F1 (Bio R5P2-01 e
R5P2-02; Hema R5P2-H01, R5P2-H02 e R5P2-H03), deduplicadas no `00_PARECER.md` em P-01, P-02 e P-03 · **Produções
avaliadas:** Bio b20fe00f7e51, Hema 7bab8fc008f6.

**Como trabalhei.** Não abri o Excel e não editei código. Usei a evidência dos relatórios (`00_PARECER.md`,
`Bioquimica_F1.md`, `Hematologia_F1.md` e os JSON citados neles) e li estas fontes, só para leitura:
- `src/comum/mEtlCQ.bas` (808, 816, 1243-1244);
- `src/comum/mEtlQA.bas` (A01 e `TextoA01`);
- `src/bio|hema/mIntegracao.bas` (`EntradaMudou`, `PrepararLinha`, `ProximoIdManual`, `ResumoQA`);
- `camada_dados.py` (`ACENTOS_DETALHE`, `SEM_ACENTO_PROIBIDO`, `validacoes`, travas das colunas);
- `instalar_adr087.py` (`FILTROS`, `converter`), `instalar_integracao.py` (Painel L4);
- `testes/qa_adr095.py` (D01, L02);
- ADR-087, ADR-095 e o GUIA §3.

| Ocorrência (entrada) | Unificada em | Veredito | Prioridade final | Autorização |
|---|---|---|---|---|
| Bio R5P2-01 e Hema R5P2-H01: "é" sem acento no DETALHE | P-01 | PROCEDE | Baixo (mantida) | não |
| Bio R5P2-02 e Hema R5P2-H02: recusa com a mensagem genérica do Excel | P-02 | PROCEDE | Baixo (mantida) | não |
| Hema R5P2-H03: linha colada com o ID junto aparece ATIVA | P-03 | PROCEDE | **Alto** (sobe de Médio) | não (ver nota) |

---

## P-01: DETALHE com o verbo "é" sem acento (Ambos)

**Veredito:** PROCEDE · **Prioridade final:** Baixo (mantida) · **Precisa de autorização:** não (só a tela muda; o
`MOTIVO_EXCLUSAO_AUTOMATICA` continua igual nos dois motores, como manda o ADR-095 §2).

### Avaliação do ponto de vista do usuário

- É um defeito de texto, não uma preferência. "HEM-212660 e retransmissão de HEM-212425" se lê como uma lista
  ("este e aquele"), e não como a razão da recusa. Quem está com pressa entende "o manual tem a ver com os dois IDs"
  e não "informe o ID original". "MAN_0002 e resultado manual" passa melhor, mas é o mesmo erro.
- Mantenho Baixo: STATUS, MOTIVO, RUN e números são iguais aos da base. O caso só aparece quando a substituição é
  recusada, e o resto do texto ("informe o ID original", "aceita só resultado do interfaceamento") já orienta.
- Por que sobrou: o `ACENTOS_DETALHE` troca só palavras, e " e " não pode ser trocado sozinho, porque quebraria a
  conjunção ("da estatística e do gráfico"). A saída é um par com contexto. Conferi os literais dos dois motores: só
  `mEtlCQ.bas:808/816` e `DB_CQ_FINAL.m:372` têm " e resultado manual" ou " e retransmiss". Nenhum outro texto usa a
  conjunção antes dessas palavras, então os pares não pegam conjunção nenhuma.

### Correção proposta (camada da tela; motores intactos)

1. Em `camada_dados.ACENTOS_DETALHE`, acrescentar três pares (a ordem não importa: " e retransmiss" casa tanto antes
   quanto depois do par `retransmissao`):

   | De | Para | Texto de origem |
   |---|---|---|
   | `' e resultado manual'` | `' é resultado manual'` | `mEtlCQ.bas:808`, `DB_CQ_FINAL.m:372` |
   | `' e retransmiss'` | `' é retransmiss'` | `mEtlCQ.bas:816` e o mesmo texto no PQ |
   | `' 1a linha'` | `' 1ª linha'` | `mEtlCQ.bas:1244` (hoje não chega à tela; ver P-03) |

   A fórmula do DETALHE passa de 17 para 20 `SUBSTITUTE` aninhados. Isso fica muito abaixo do limite do Excel (64
   níveis, 8.192 caracteres). O "ª" (U+00AA) vai no `camada_dados.py` em UTF-8 e chega pela COM como Unicode. O
   texto não muda de tamanho, então não há risco novo de corte na coluna.
2. Trava D01 (`qa_adr095`): acrescentar em `camada_dados` uma lista `SEM_ACENTO_CONTEXTO = [' e resultado manual',
   ' e retransmiss', ' 1a linha']`. O `sem_acento()` reprova se algum desses trechos sobrar depois da troca. A
   `trava_sintetica` ganha o texto `'MAN_0009 e resultado manual (texto de teste)'`, e a trava tem de pegá-lo.
3. **Não** recomendo reescrever o texto na origem. Isso mexeria nos dois motores, no `qa_sync` S6 e no detalhe do E12,
   que é recortado do MOTIVO (`Mid$(..., pSub)`), contra a decisão do ADR-095 §2. O ganho seria o mesmo.

### Critério de aceite (reteste)

1. **Bio e Hema, M10** (SUBSTITUI_ID = MAN_0002): `DETALHE.Text` contém "MAN_0002 é resultado manual (SUBSTITUI_ID
   aceita só resultado do interfaceamento)".
2. **Hema, M13** (SUBSTITUI_ID = HEM-212660): `DETALHE.Text` contém "HEM-212660 é retransmissão de HEM-212425 (informe
   o ID original)".
   - Bio: o caso não pode ser provocado, porque não há DUPLICIDADE_ORIGEM. Fica coberto pelo D01, já que a fórmula é
     a mesma nos dois produtos. Registrar como NAO_EXECUTADO, com esse motivo.
3. **Regressão da conjunção:** nos 17 manuais de cada produto, nenhum outro " é " aparece no DETALHE. O texto do
   IMPLAUSIVEL continua "fora da estatística e do gráfico…", com a conjunção sem acento.
4. **Motores intactos:** a `MOTIVO_EXCLUSAO_AUTOMATICA` da tblCQ_Final dos manuais é idêntica à da base, byte a byte,
   e o `qa_sync` S6 continua verde.
5. **D01 verde:** os 3 pares novos estão na fórmula (`acentos_fora_da_formula` vazio), `sobram_sem_acento` vazio e a
   trava sintética pega os 3 textos.

---

## P-02: Registros e Painel L4 recusam com a mensagem genérica do Excel (Ambos)

**Veredito:** PROCEDE · **Prioridade final:** Baixo (mantida) · **Precisa de autorização:** não (só a mensagem de
erro; lista, tipo de validação e estilo Parar ficam iguais).

### Avaliação do ponto de vista do usuário

- É um problema pequeno, mas real: a mesma ação ("digitei e o Excel recusou") tem duas respostas diferentes no
  sistema. Na Digitar Resultados, a caixa diz o que fazer ("Escolha … na seta da célula"). No Registros e no Painel,
  aparece a caixa padrão do Excel ("Este valor não corresponde às restrições de validação de dados…"), que não diz
  qual é a lista nem como sair dela.
- No filtro, isso pesa um pouco mais. Quem digitou "S" no H4 querendo "Sim" ou quer limpar o filtro fica preso em
  "Repetir/Cancelar", sem saber que basta apagar a célula ou clicar em Limpar filtros.
- Mantenho Baixo: a recusa funciona, nada errado é gravado e os filtros só mudam a vista.
- Pela leitura do código, a falta é maior do que a evidência mostra. O laço do `instalar_adr087.FILTROS` (linhas
  192-197) cria **todas** as listas de filtro (Lote B4, Analito C4, Equipamento G4, Calibrado? H4 e Status J4) sem
  `ErrorMessage`. Só a validação de data (E4/F4) tem mensagem. A tblRegistros (colunas Nível, Foi calibrado? e
  Status, recriadas no `converter`) e o L4 do Painel (`instalar_integracao`, linhas 243-244) também estão sem.
  Proponho corrigir todas de uma vez, para não abrir um reteste por célula.

### Correção proposta

Textos exatos (todos com menos de 225 caracteres, o limite do Excel), na mesma forma da Digitar Resultados:

| Onde | ErrorTitle | ErrorMessage |
|---|---|---|
| Registros B4 (Lote) | `Filtro de lote` | `Escolha o lote na seta da célula. Para ver todos, apague o filtro (tecla Delete) ou clique em Limpar filtros.` |
| Registros C4 (Analito) | `Filtro de analito` | `Escolha o analito na seta da célula. Para ver todos, apague o filtro (tecla Delete) ou clique em Limpar filtros.` |
| Registros G4 (Equipamento) | `Filtro de equipamento` | `Escolha o equipamento na seta da célula. Para ver todos, apague o filtro (tecla Delete) ou clique em Limpar filtros.` |
| Registros H4 (Calibrado?) | `Filtro Calibrado?` | `Escolha Sim ou Não na seta da célula. Para ver todos, apague o filtro (tecla Delete) ou clique em Limpar filtros.` |
| Registros J4 (Status) | `Filtro de status` | `Escolha o status na seta da célula. Para ver todos, apague o filtro (tecla Delete) ou clique em Limpar filtros.` |
| tblRegistros, coluna "Foi calibrado?" (F7 até a última linha) | `Foi calibrado?` | `Escolha Sim ou Não na seta da célula.` |
| tblRegistros, colunas "Nível" e "Status" (D e J), se tiverem lista | — | `Escolha o nível na seta da célula.` / `Escolha o status na seta da célula.` |
| Painel L4 (Bio) | `Equipamento` | `Escolha o equipamento na seta da célula.` (o mesmo texto da Digitar Resultados) |

Como fazer (armadilhas):
1. **Idempotente.** No `instalar_adr087`, gravar `ErrorTitle`, `ErrorMessage` e `ShowError = True` **sempre**, fora
   do `if lst and f1 != lst`. Do jeito que está, a produção, que já tem a lista certa, nunca receberia a mensagem.
   Para a tblRegistros, usar um passo próprio que rode em toda instalação, porque o `converter` só roda uma vez e
   sai cedo se a tabela já existe. Esse passo grava só as propriedades da mensagem na `DataBodyRange` da coluna, sem
   `Delete`/`Add`, sem tocar no `Formula1` e sem mexer nas linhas de dados.
2. **Painel L4:** no `instalar_integracao`, logo depois do `l4.Validation.Add(...)`.
3. **Aba protegida:** o instalador já faz `Unprotect`. Proteger de novo no fim, como hoje.
4. Nenhum `Names.Add` novo. As listas continuam com o separador do Excel (`cd.sep_lista`), sem `x12ac:list`.
5. A linha nova da tblRegistros, que "aparece sozinha", herda a validação da coluna. Conferir no aceite.

### Critério de aceite (reteste)

1. **Bio e Hema:** em B4, C4, G4, H4 e J4 do Registros, e na coluna "Foi calibrado?" (F7 e a última linha da
   tblRegistros), `Validation.ErrorMessage` é igual ao texto da tabela acima, e `ShowError = True`. Bio: idem no
   Painel L4.
2. **Lista inalterada:** `Validation.Formula1`, `Type` (3) e `AlertStyle` (1, Parar) iguais aos da base. 0
   `x12ac:list` no XML. "Talvez"/"S" no H4 e "DIMENSION 9" no L4 continuam com `Validation.Value = Falso`, e Sim,
   Não, DIMENSION 1 e DIMENSION 2 continuam aceitos (L02 do `qa_adr095` verde).
3. **Filtros só mudam a vista:** o hash do Value2 da tblRegistros fica constante em H4 Sim/Não/vazio e De/Até, como
   neste ciclo.
4. **Idempotência:** rodar o instalador duas vezes sobre a cópia. A segunda vez não registra mudança e as mensagens
   continuam lá.
5. **A caixa na tela** (o texto aparece ao digitar "Talvez"): só numa sessão interativa. Com `Interactive = False`,
   registrar como NAO_EXECUTADO. Os itens 1 a 4 bastam para aprovar.

---

## P-03: linha manual colada com o ID junto aparece ATIVA, mas fica fora da estatística (Hema; código comum)

**Veredito:** PROCEDE · **Prioridade final:** **Alto** (a frente deu Médio) · **Precisa de autorização:** não, pela
correção proposta (ver a nota no fim desta seção).

### Avaliação do ponto de vista do usuário

- É um problema, e o mais sério da rodada. É o único ponto em que a tela contradiz o que o sistema fez. O GUIA §3
  manda o usuário "atualizar e conferir o STATUS", e o STATUS mostra ATIVO para uma linha que não entrou.
- Copiar uma linha e colar é o atalho natural de quem lança vários manuais parecidos: mesma data, equipamento e lote,
  e muda o nível ou o resultado. A frente testou a cópia sem mudar nada, e aí os números ficam certos (a cópia não
  entra). O caso realista, porém, é **colar e editar**. Pelo código (`mEtlCQ.bas:1243`), o teste "ID_REGISTRO manual
  repetido" vem antes de qualquer outro e não olha o conteúdo. Então o N2 digitado sobre a linha colada fica fora da
  estatística, do LJ e do Westgard, e a tela mostra ATIVO, com o RUN e o DETALHE do MAN_0007 (o XLOOKUP pelo ID acha
  a 1ª linha).
  - Resultado: um resultado de controle que o laboratório acredita ter lançado não é avaliado, sem nenhum sinal
    visível. O único aviso é um ALERTA A01 no meio de milhares na QA_INTEGRACAO, e o `ResumoQA` do ATUALIZAR só conta
    ERROs.
  - Por isso subo para **Alto**: é risco operacional relevante (REGRAS §4), e não só inconsistência de tela. Esse
    cenário de colar e editar é inferência pelo código e precisa ser executado no reteste (aceite, item 2).
- O texto do A01 ("confirme e inative um deles") piora a situação. Os dois têm o mesmo ID, e inativar o MAN_0007
  tiraria o resultado **bom**.
- Hoje já existe uma saída, só que ninguém a conhece: o ID_REGISTRO da Digitar Resultados é destravado
  (`camada_dados.py:494`), e apagar o ID faz o `PrepararLinha` carimbar um novo. A correção abaixo faz isso sozinha.

### Correção proposta

1. **Carimbar ID novo na colagem.** Nos dois `mIntegracao.bas` (`src/bio` e `src/hema`, em `EntradaMudou` ou
   `PrepararLinha`, ramo `TB_MANUAL`, hoje idênticos), a regra é esta:
   - quando a célula ID_REGISTRO da linha `r` estiver dentro do `Target` (colada ou digitada), e o ID normalizado já
     existir em outra linha da tblResultados_Manuais **fora** das linhas do `Target`, gravar
     `ProximoIdManual(lo)` na linha `r`;
   - com várias linhas coladas, cada uma recebe o seu ID em sequência. O `ProximoIdManual` já avança a marca e nunca
     reaproveita número (D05);
   - recortar e colar (mover) não gera repetido, porque a linha de origem fica vazia, então nada muda nesse caso;
   - nenhum `Names.Add` novo: a marca `qcUltimoIdManual` já existe e é numérica.
2. **Aviso na colagem** (só com `Application.Interactive`, pelo mesmo `celAviso` que já existe: uma caixa por
   colagem, nunca uma por linha). Título "Digitar Resultados", texto:
   `A linha colada trazia o ID MAN_0007, que já existe em outra linha. Ela recebeu o ID novo MAN_0018. Mude o que
   for diferente (nível, hora, resultado) antes de clicar em ATUALIZAR DADOS.`
   Se a colagem for de várias linhas, use `… As N linhas coladas receberam IDs novos (MAN_0018 a MAN_0020) …`.
3. **Mensagem do ATUALIZAR** (`ResumoQA`): mesmo com 0 ERRO, uma linha a mais quando houver CONFLITO_MANUAL com o
   MOTIVO começando por "ID_REGISTRO manual repetido". Ela cobre as repetições que já existirem ou que entrarem com
   os eventos desligados:
   `• N linha(s) da Digitar Resultados repete(m) o ID de outra (ex.: MAN_0007) e ficou(aram) fora da estatística.
   Apague o ID da linha repetida (um novo é gerado) e atualize de novo.`
4. **Texto do A01 para esse caso** (`TextoA01`, ramo `Else`, no VBA e no `QA_INTEGRACAO.m`, com o mesmo texto, sem
   acento, como os demais): quando o MOTIVO começar por "ID_REGISTRO manual repetido", trocar "confirme e inative um
   deles ou corrija o manual" por `; linha <n> da Digitar Resultados com ID repetido: apague o ID dela (um novo
   sera gerado) ou a linha. Nao inative o <ID>: isso tiraria o resultado que vale.` O `<n>` vem do sufixo `~L<n>`.
   A frase evita o verbo "e" sem acento, o mesmo tropeço do P-01.
5. **Não recomendo** casar STATUS/RUN/DETALHE da tela pela chave `~L<n>`. Seria uma fórmula com contagem acumulada
   por linha, e o `LET` tem histórico de falha via COM. Ela teria de copiar a regra "1ª linha válida por ID (senão a
   1ª)" do motor (`mEtlCQ.bas:1140`) e manteria dois registros com o mesmo ID na tela. O carimbo resolve na origem, e
   o item 3 cobre o que sobrar.

**Nota sobre autorização.** A regra do motor ("vale a 1ª linha; as outras ficam CONFLITO_MANUAL") não muda, e nenhum
critério estatístico muda. O carimbo só faz a colagem obedecer à regra que já vale para quem digita: o ID é gerado
sozinho e é único (GUIA §3, D05). A linha colada sem mudança continua fora da estatística, agora pelo motivo certo
("resultado manual repetido (mesma chave e instante)") e com o próprio ID na tela. A linha colada e editada passa a
valer como um manual novo, exatamente como se tivesse sido digitada. Se o Orquestrador entender que isso é mudança de
regra de entrada, ele pede autorização. Os itens 3 e 4, só de texto, podem seguir de qualquer forma.

### Critério de aceite (reteste, **nos dois produtos**; a Bio não foi executada neste ciclo)

1. **Colar sem editar** (cópia da linha do MAN_0007 na próxima linha livre, com os eventos ligados):
   - a linha colada fica com um ID novo, maior que todos (ex.: MAN_0018), e o MAN_0007 continua na linha original;
   - depois do ATUALIZAR, a tela da linha colada mostra CONFLITO_MANUAL e o DETALHE "resultado manual repetido (mesma
     chave e instante)", e o MAN_0007 continua ATIVO com o RUN dele;
   - a tblCQ_Final não tem nenhum `~L`, e o engPainel da série é idêntico ao de antes da colagem (n, média, DP).
2. **Colar e editar** (cenário novo): colar a linha do MAN_0007 e trocar NIVEL e RESULTADO por um valor sintético
   (ex.: 111,111). Depois do ATUALIZAR:
   - a linha tem ID novo, STATUS ATIVO e RUN próprio;
   - o n da série do novo nível sobe 1, com a média conferida por conta própria;
   - a série do MAN_0007 não muda.
3. **Várias linhas:** colar 3 linhas existentes de uma vez. Cada uma recebe um ID distinto e em sequência, e nenhum
   número é reaproveitado (a marca `qcUltimoIdManual` avança).
4. **Mover:** recortar e colar uma linha manual. O ID dela **não** muda.
5. **Repetido que já existe** (criado com os eventos desligados, como a frente fez):
   - a mensagem do ATUALIZAR traz a linha do item 3 da correção, com o ID;
   - o A01 traz o texto novo, com a linha `<n>`, e não diz mais "inative um deles";
   - apagar o ID da linha repetida gera um ID novo, e o próximo ATUALIZAR não acusa mais nada.
6. **Regressão:** os 17 cenários manuais deste ciclo dão o mesmo STATUS, MOTIVO e RUN da base, exceto o M19, que
   passa a ter ID novo. O `qa_sync` S6 continua verde.
7. **Caixa de aviso na colagem:** só numa sessão interativa. Sem ela, registrar NAO_EXECUTADO e conferir pelo menos
   o ID carimbado.

---

## Fora do escopo deste parecer

A D-04 ("Outro" sem texto livre no MOTIVO) não me foi encaminhada como ocorrência. Pela tela, ela pede um lugar para
o texto livre. Como o ADR-082 trata o motivo como justificativa de auditoria, depende do Especialista em CQ e de
autorização.
