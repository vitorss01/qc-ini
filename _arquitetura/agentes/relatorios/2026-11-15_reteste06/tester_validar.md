# Tester/QA: modo VALIDAR, ciclo 2026-11-15_reteste06 (ADR-096, entrega 43)

Correção validada: ADR-096, commit 81b04a4, entrega 43 em produção desde 15/11. **Não fui eu quem a implementou**:
o Tester/QA só reproduziu (relatório `tester_reproduzir.md`, que trata de outra ocorrência, a P-01) e validou.

## Ambiente

| Item | Valor |
|---|---|
| Produção | `QC_Bioquimica.xlsm` sha256 `bd5314426ab80aa3…`; `QC_Hematologia.xlsm` sha256 `49f68fba99d5c748…`. Conferi antes e depois de cada rodada: **não mudou** (último `sha256sum` às 02:22 de 16/11) |
| Base de comparação | backup anterior à entrega 43: `_backup_pre_integracao_2026-11-15_1802` (Bio, sha `b20fe00f7e51503b`) e `_2059` (Hema), só leitura ou cópia |
| Cópias | todas novas, em `%TEMP%\qcw\auditoria\2026-11-15_reteste06\tester_validar\`, nunca salvas e apagadas no fim (`copia_apagada: true` em cada JSON) |
| Excel | um por vez, aberto e fechado pelo `qa_final.QA`. PIDs: Bio 19812 (15/11 23:40), Bio base 18316, Hema 18324 (16/11 01:52), complemento Bio 22128 e Hema 9212, suítes pelo `rodar_suites.py`. **Não toquei** nos PIDs 13084, 8640 e 15880 (outros processos e tarefa agendada do DB_SEAC). Nenhuma captura de tela |
| PID órfão | Uma rodada anterior do Tester neste ciclo (Hema, 15/11 23:54) foi interrompida e deixou o Excel **30332** aberto, sem nenhum Python o controlando, com a cópia `Hematologia_nova_235456.xlsm`. O log dela registra `EXCEL_PID 30332`. Encerrei só esse PID e apaguei a cópia antes de começar. A rodada da Hema foi refeita do zero |
| Scripts (meus) | `_arquitetura/agentes/ferramentas/valida_r6_adr096.py` (OC-0073/0021/0074/0075/0076, oráculo próprio), `valida_r6_xml.py` (validações e formatos lidos do XML, sem Excel), `valida_r6_complemento.py` (listas do Registros, causa) e `valida_r6_resumo.py`. **Não** reaproveitei `reteste06_usuario_f1*.py` nem a suíte `qa_adr095` para os critérios |
| Dados de teste | só nas cópias: usuário `TESTE_AUDITORIA_R6V` (ANALISTA); manuais de 14/11/2026, 05:11 a 07:42, com valores sintéticos de 3 decimais. **Bio**: Magnésio, lote 8977, DIMENSION 1 (2,446; 2,441; 5,806 = média do lote + 4 DP; 4,750; 2,448; 2,443; 4,747; 4,752) e GLUC 5659. **Hema**: BASO%, lote 625311, XN1000 (4,762; 4,757; 6,370 = + 4 DP; 4,791; 4,764; 4,759; 4,788; 4,793) e EO#. ID inválido `TESTE-AUD-R6V`. Escolhi analitos diferentes dos do Usuário Simulado (Glicose/ALT e WBC/BASO#) |
| Saídas | `valida_r6_<Produto>_nova.json/.log`, `valida_r6_Bioquimica_base.json`, `valida_r6_xml.json`, `complemento_<Produto>.json`, `resumo_<Produto>.txt` e `regressao*/`, todos na pasta acima (fora do repositório) |

## Vereditos

| ID | Produto | Veredito | Resumo |
|---|---|---|---|
| OC-0073 | Ambos | **VALIDADA** | Todos os ramos do critério conferidos com oráculo próprio, no VBA e no PQ |
| OC-0021 | Ambos | **VALIDADA** | "é resultado manual", "é retransmissão" e "1ª linha" na tela; conjunção intacta; MOTIVO do motor sem acento e igual no VBA e no PQ |
| OC-0074 | Ambos | **VALIDADA** (com ressalva) | Mensagens corretas em todas as validações pedidas; listas e valores aceitos iguais aos da base. Ressalva: os filtros de equipamento e de status ficam com a lista vazia e aceitam qualquer texto. Isso já existia e não é do ADR-096 (achado novo T6V-01) |
| OC-0075 | Bioquímica | **VALIDADA** | "01/09/2026 a último dado" e "primeiro dado a 30/09/2026"; Value2, fórmulas e grade iguais aos da base |
| OC-0076 | Bioquímica | **VALIDADA** | Com H3 vazio: "Período inteiro (todos os lotes): … resultados.", sem citar a Estatística; com lote, o texto de antes; números iguais ao oráculo |

### OC-0073: linha colada com o ID de outra (Ambos), VALIDADA

Roteiro próprio, na aba protegida, com os eventos ligados e `Interactive`. O vigia lê o texto das caixas do meu PID e
as fecha.

| Passo | Bioquímica | Hematologia |
|---|---|---|
| O: origem digitada | MAN_0001, N1 05:11 2,446 | MAN_0001, N1 05:11 4,762 |
| S1: colar A:K logo abaixo, sem editar | **MAN_0002** na hora, com o aviso "A linha trazia o ID MAN_0001, que já existe em outra linha. Ela recebeu o ID novo MAN_0002 (o ID é único e gerado pelo sistema). Mude o que for diferente (nível, hora, resultado) antes de clicar em ATUALIZAR DADOS." | igual (MAN_0002) |
| S2: colar **acima** da origem e editar a hora e o resultado | MAN_0003 + aviso; a origem continua MAN_0001 | igual |
| S3: colar **só valores** (PasteSpecial) e trocar para N2 com média + 4 DP | MAN_0004 + aviso | igual |
| S4: digitar " man_1 " (minúsculo, sem zeros, com espaços) numa linha vazia | MAN_0005 + aviso | igual |
| S5: FillDown de 3 linhas | a de cima fica com MAN_0006, e as outras duas ganham MAN_0007 e MAN_0008, numa caixa só: "As 2 linhas coladas traziam IDs que já existem…" | igual |
| S7: falso positivo (editar o RESULTADO, redigitar o próprio ID) | o ID não muda, sem aviso | igual |
| S6: mover (recortar A:K e colar) | o destino **mantém** MAN_0008, a origem fica vazia, sem aviso | igual |
| S8: ID fora do padrão repetido (`TESTE-AUD-R6V` em 2 linhas) | **não** é trocado. O motor dá `MAN_INVALIDO_L13`/`L14`, E04 "ID_REGISTRO (padrao MAN_0001)" | igual |
| S9: repetido criado com os **eventos desligados** | `MAN_0001~L15` CONFLITO_MANUAL "ID_REGISTRO manual repetido (vale a 1a linha)". O A01 diz: "…apague o ID da linha repetida (um novo sera gerado) ou a linha. Nao inative o MAN_0001: isso tiraria o resultado que vale." A mensagem do ATUALIZAR traz "• 1 linha(s) da Digitar Resultados repete(m) o ID de outra (ex.: MAN_0001) e ficou(aram) fora da estatística. Apague o ID d…" | igual |
| Correção do S9 (apagar o ID) | ganha MAN_0011 → ATIVO, RUN 26111405. O A01 e a linha da mensagem somem (alertas 135 → 134) | MAN_0012 → ATIVO, RUN 26111405 (593 → 592) |
| Original | MAN_0001 intacto em todos os passos (`origem_mudou = {}`), ATIVO, RUN 26111401 | igual |

**Depois do ATUALIZAR (VBA, 24,9 s na Bio e 36,2 s na Hema):** os colados e editados (S2, S3, S4, P, FD1, MOV) ficam
ATIVO, com PARTICIPA SIM, na `tblCQ_Final` e na tela.

**n e média, comparados com o oráculo** (Python sobre a `tblCQ_Final`: PARTICIPA = SIM, analito, lote, nível,
equipamento e janela). A tolerância é de 1e-9 relativo, e todas as diferenças ficaram ≤ 2e-16.

| | Antes | Depois do 1º ATUALIZAR | Depois da correção |
|---|---|---|---|
| Bio Painel (nov/2026, lote 8977) N1 | 12 / 2,6250 | **16** / 2,579875 (oráculo 16 / 2,579875) | 16 |
| Bio Painel N2 | 10 / 4,692 | **13** / 4,786385 (oráculo 13) | **14** / 4,783929 (oráculo 14) |
| Bio Estatística (H4 = 8977, nov) N1 / N2 | — | 16 / 13, média igual ao oráculo | 16 / 14 |
| Hema Painel (nov, lote 625311) N1/N2/N3 | 17 / 23 / 12 | **21 / 26** / 12 (oráculo igual) | 21 / **27** / 12 |
| Hema Estatística (ano 2026, B4 = 625311) | — | 42 / 65 / 33 (oráculo igual) | 42 / **66** / 33 |

As contas fecham. Bio N1: +4 (O, S2, P, FD1; o S1 fica de fora, veja abaixo). Bio N2: +3 (S3, S4, MOV), +1 com a
correção.

- **LJ (Eng_Saida):** os slots 138 a 141 (Bio) e 68 a 71 (Hema) trazem os IDs MAN_ e os valores digitados.
- **Westgard:** o S3 (+4 DP) dispara o 1-3s no N2. Na Bio, o slot 138 traz `N2_R13s = 1` e REJEITADO. Na Hema, o slot 68
  traz `N2_R13s = 1` e REJEITADO.
- **Reserva PQ** (MOTOR_DADOS = PQ, mesmo estado): **nenhuma diferença** de STATUS, PARTICIPA, MOTIVO, RESULTADO ou NIVEL
  em relação ao VBA. O texto do A01 é idêntico nos dois motores. A Bio levou 400 s. A Hema levou 1287 s e 48.405 linhas,
  e o motor de produção é o VBA (36 s). Isso não é defeito: é um ATUALIZAR completo pelo PQ, não o S6.

**O que não muda e já é conhecido** (registrado pelo Usuário Simulado e pelo UX; não é ocorrência nova):

- **S1, colar sem editar:** ganha ID novo, mas fica CONFLITO_MANUAL "resultado manual repetido (mesma chave e instante)",
  com o A01 antigo, "inative um deles". É a regra do motor e a P-04/D-02 do parecer.
- **S9, antes da correção:** a linha repetida mostra STATUS ATIVO na tela, porque o XLOOKUP acha a 1ª linha do ID. O
  alerta fica a cargo do A01 e da mensagem do ATUALIZAR, como o ADR-096 previu.
- **S8:** STATUS "—" e DETALHE "atualize os dados" (P-03).
- **S6:** a origem do recorte fica com USUARIO e REGISTRADO_EM. É o resíduo da P-01.

### OC-0021: DETALHE com acento (Ambos), VALIDADA

- **Tela real:**
  - D1, SUBSTITUI_ID = um MAN_, nos dois produtos: "SUBSTITUI_ID inválido: MAN_0001 **é** resultado manual (SUBSTITUI_ID
    aceita **só** resultado do interfaceamento) (este manual **não** substitui nada)";
  - D2, Hema, SUBSTITUI_ID = a retransmissão HEM-212660: "… HEM-212660 **é retransmissão** de HEM-212425 (informe o ID
    original) …". A Bio não tem DUPLICIDADE_ORIGEM na base, então esse caso foi provado pela fórmula;
  - D3, resultado 0: "resultado <= 0: implausível para material de controle (fora da estatística e do gráfico)", com a
    conjunção "e" intacta.
- **Fórmula da coluna DETALHE:** 20 SUBSTITUTE na produção contra 17 na base, com os 3 pares de contexto novos. Avaliei a
  fórmula numa aba de rascunho da cópia com os textos dos dois motores. Os 7 literais saem certos, inclusive "vale a
  **1ª linha**", e nenhum MOTIVO distinto da cópia sobra sem acento (Bio 3, Hema 2).
- **Na tela real, o "1ª linha" não aparece**, porque a linha repetida mostra o status da 1ª (veja acima). Por isso a
  prova é pela fórmula.
- **Motor:** o MOTIVO na `tblCQ_Final` continua sem acento ("SUBSTITUI_ID invalido: MAN_0001 e resultado manual…",
  "vale a 1a linha", "implausivel") e é **igual no VBA e no PQ** (diferenças = {}).

### OC-0074: mensagens das listas (Ambos), VALIDADA com ressalva

| Validação | Mensagem (Bio e Hema) | Lista/aceitos × base |
|---|---|---|
| Registros B4, filtro de lote | "Filtro de lote" / "Escolha o lote na seta da célula. Para ver todos, apague o filtro (tecla Delete) ou clique em Limpar filtros." | iguais |
| C4 analito, G4 equipamento, J4 status, H4 Calibrado? | o mesmo padrão ("Escolha o analito / o equipamento / o status / Sim ou Não na seta da célula. Para ver todos…") | iguais |
| tblRegistros: Nível, Foi calibrado?, Status | "Escolha o nível / Sim ou Não / o status na seta da célula." | iguais |
| Linha nova da tblRegistros (`ListRows.Add`) | herda as 3 validações e as mensagens, e recusa texto fora da lista | — |
| Painel L4 (Bio) | "Equipamento" / "Escolha o equipamento na seta da célula." | iguais (DIMENSION 1;DIMENSION 2) |
| Digitar Resultados (EQUIPAMENTO, MATRIZ, MOTIVO) | inalteradas (ADR-095) | iguais |

- Em todas: ShowError = True e AlertStyle = 1 (Parar).
- **XML de produção × backup:** as mesmas validações (sqref e fórmula). A única diferença é que o H4 saiu do grupo
  `F7:F9 H4` para ter a própria mensagem.
- **Aceitos e recusados** (`Validation.Value`), Bio nova × base: os mesmos para todos os itens e para os 7 textos
  inválidos.

**Ressalva (achado novo T6V-01, Baixo, já existia):**

- **Bio:** `lstRegEquip` (G4) e `lstRegStatus` (J4) apontam para uma única célula vazia, porque nenhuma linha da
  tblRegistros tem equipamento ou status preenchido.
- **Hema:** o mesmo acontece no G4. O J4 tem "Encerrado".
- Validação de lista com origem vazia aceita qualquer texto ("TESTE-AUDITORIA-R6V", "XN2000"…). A mensagem nova nunca
  aparece, e o filtro fica com um valor que não existe.
- É igual na base.
- **Causa:** `src/comum/mCalibracao.bas:589-592` (`EscreverLista`, ramo `n = 0`: grava "" e aponta o nome para a
  célula vazia).

**Observação (já existia, informativa):**

- Os itens de `lstRegLote` são texto: `mCalibracao.bas:604`, `out(i, 1) = "'" & a(i)`.
- Escolher na seta funciona. Digitar o número do lote (8977, 522611) vira número e é **recusado** com "Escolha o lote
  na seta da célula…", e a mensagem orienta certo.
- **Correção de leitura da minha rodada de 15/11:** "Bio recusa todo lote" era um artefato do meu teste, que escrevia
  o item como número. Com o apóstrofo, ele é aceito nos dois produtos.

### OC-0075: Estatística da Bio com só o De ou só o Até, VALIDADA

Testei Glicose e AST (TGO), lote 8977, DIMENSION 1, em 5 janelas.

| Janela no Painel | Linha 5 da Estatística | Base |
|---|---|---|
| só De 01/09/2026 | "01/09/2026 a **último dado**" | "01/09/2026 a 00/01/1900" |
| só Até 30/09/2026 | "**primeiro dado** a 30/09/2026" | "00/01/1900 a 30/09/2026" |
| De e Até | "01/09/2026 a 30/09/2026" | igual |
| só De 01/05/2026 | "01/05/2026 a último dado" | "01/05/2026 a 00/01/1900" |
| Tudo | "histórico inteiro do lote" | igual |

- **Formatos:** H5 `dd/mm/aaaa;;"primeiro dado"`, J5 `dd/mm/aaaa;;"último dado"`.
- **Inalterado em relação à base:** Value2 e fórmula de B5, D5, H5, I5 e J5, e o hash da grade A14:AU133, nas 10
  combinações.
- **n da grade × oráculo:** iguais, e a média também (rel = 0). Exemplos: Glicose só De, N1 58 / N2 56; AST só Até,
  74 / 68.
- Nenhum "1900" em nenhuma combinação.

### OC-0076: aviso do corte de 180 corridas na Bio, VALIDADA

| Cenário (Painel A39, "Tudo") | Texto | Oráculo |
|---|---|---|
| Glicose, DIMENSION 1, H3 vazio | "… ⚠ 164 corrida(s) antigas … (máx. 180). Período inteiro (todos os lotes): N1 264 · N2 249 resultados." Não cita a Estatística | N1 264, N2 249 (Eng AF1 igual); M1 = 164 = base |
| Bicarbonato, DIMENSION 2, H3 vazio | "… 213 corrida(s) … Período inteiro (todos os lotes): N1 294 · N2 310 resultados." | 294 / 310 |
| AST (TGO), DIMENSION 1, H3 vazio | "… 214 corrida(s) … (todos os lotes): N1 317 · N2 282 resultados." | 317 / 282 |
| Glicose, **lote 8974** no H3 | "… 39 corrida(s) … Período inteiro: N1 174 · N2 161 — veja a aba Estatística." É **idêntico à base** | 174 / 161 |

- O número de corridas antigas (M1) é igual ao da base em todos os cenários.
- **Hema:** o texto do ADR-095 continua, sem citar a Estatística. WBC e PLT com todos os lotes dão N1 515 · N2 545 ·
  N3 479, iguais ao oráculo.

## Regressão

Rodei as suítes em cópias novas, pelo `rodar_suites.py`, com um produto por vez e o sha da produção conferido antes e
depois.

- **Rodada completa:** 15/11, das 22:19 às 23:35, pelo Tester deste ciclo, sobre os mesmos sha de produção de hoje.
  Log em `tester_validar\regressao\rodar.log`.
- **qa_adr095 de novo:** 16/11 às 02:21, pasta `regressao_1611`.
- Os itens INFO não são falha: E12b é impossível nesta configuração, E13c está coberto noutra suíte, TEMPOS, M13/M19 não
  têm dado na base, e há a pendência anterior ao ADR-090.

| Produto | Suíte | PASS | FAIL | Tempo |
|---|---|---|---|---|
| Bioquímica | qa_adr095 (15/11) | 10 | 0 | 28 s |
| Bioquímica | qa_adr095 (16/11) | 10 | 0 | 29 s |
| Bioquímica | qa_etl | 64 | 0 (+3 INFO) | 739 s |
| Bioquímica | qa_manual_substitui | 18 | 0 (+2 INFO) | 984 s |
| Bioquímica | qa_adr090 | 21 | 0 | 245 s |
| Bioquímica | qa_adr087 | 14 | 0 | 191 s |
| Hematologia | qa_adr095 (15/11) | 8 | 0 | 28 s |
| Hematologia | qa_adr095 (16/11) | 8 | 0 | 24 s |
| Hematologia | qa_etl | 63 | 0 (+3 INFO) | 863 s |
| Hematologia | qa_manual_substitui | 19 | 0 (+1 INFO) | 1030 s |
| Hematologia | qa_adr090 | 22 | 0 (+1 INFO) | 251 s |
| Hematologia | qa_adr087 | 14 | 0 | 198 s |

O `qa_etl` reescreve `testes/resultados/etl_rastreabilidade_<Produto>.json`, que é versionado. A diferença é de 1 linha
por arquivo e aparece no `git status`. Não a reverti.

## Efeitos colaterais no fluxo vizinho

Nenhum efeito colateral do ADR-096 foi encontrado.

- **Digitação normal:** editar o resultado ou redigitar o próprio ID não troca o ID nem gera aviso. Mover mantém o ID.
- **SUBSTITUI_ID** (MAN_ e retransmissão), **IMPLAUSIVEL** e o **ID inválido** se comportam como na base.
- **Motor:** a `tblCQ_Final` da cópia é idêntica à da base antes de digitar (hash `9ce135c04bab3654` na Bio). VBA e PQ
  saem iguais.
- **Telas:** as listas da Digitar Resultados não mudaram. A linha nova da tblRegistros herda as validações. Na
  Estatística, só o formato de H5/J5 mudou. O A39 com lote ficou igual ao da base.

## Achados novos

| ID | Prioridade | Produto | Achado |
|---|---|---|---|
| T6V-01 | Baixo | Ambos | Filtros do Registros com a lista vazia (G4 equipamento; J4 status na Bio) aceitam qualquer texto, e a mensagem do ADR-096 nunca aparece. Já existia. Causa: `src/comum/mCalibracao.bas:589-592`. Sugestão para quem corrigir: com a lista vazia, gravar um item neutro, ou desligar o campo, ou dizer "Nenhum equipamento registrado" |
| T6V-02 | Baixo (informativo) | Ambos | O filtro de lote recusa o lote digitado, porque a lista é de texto (`mCalibracao.bas:604`). A mensagem orienta a usar a seta, e escolher na seta funciona. Já existia. Avaliar com o UX se vale aceitar o número digitado |
