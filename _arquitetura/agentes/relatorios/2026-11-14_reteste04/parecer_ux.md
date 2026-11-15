# Parecer de UX: ciclo 2026-11-14_reteste04

**Agente UX/UI, 14/11/2026, só leitura.** Este parecer não abriu o Excel, não editou código e não mudou fórmulas. A base são os relatórios das frentes desta pasta (`00_PARECER.md`, `Bioquimica_F1.md`, `Bioquimica_F2.md`,
`Hematologia_F1.md`, `Hematologia_F2.md`) e a leitura das fontes citadas nas evidências: `camada_dados.py`,
`src/comum/mCEQForm.bas`, `src/comum/mEtlCQ.bas`, `src/bio/mEstatistica.bas`, `instalar_adr078.py`,
`instalar_adr092.py` e `ceq_form_spec.py`. As referências de padrão são o ADR-087 (nada cortado ou sobreposto, janela
do CEQ), o ADR-090, o ADR-091 e o ADR-092.

Nenhuma correção proposta aqui muda n, média, DP, CV, Westgard, bias, ET, Sigma ou incerteza. Os casos em que isso
poderia acontecer estão marcados.

## Resumo

| # | Ocorrência (P-xx do consolidado) | Produto | Veredito | Prioridade final | Autorização |
|---|---|---|---|---|---|
| 1 | OC-0021: DETALHE do IMPLAUSIVEL e nota do ADR-093 sem acento (P-08) | Bio | PROCEDE | Baixo | Não, pela lista de tela. Acentuar na origem depende da D-02 |
| 2 | Listas EQUIPAMENTO, MATRIZ e MOTIVO com um item só (P-01) | Bio | PROCEDE | **Alto** (sobe para Crítico se o reteste mostrar que o manual com equipamento inexistente muda algum número) | Não, para a lista. Sim, para um QA novo ou se o manual inválido deixar de participar |
| 3 | Estatística B5/D5 "00/01/1900" em "Tudo" (P-09a) | Bio | PROCEDE | Baixo | Não (só formato) |
| 4 | Janela do CEQ: Controllab mapeado com AR = CAP diz "entra no bias" (P-03) | Bio | PROCEDE | Médio | Não |
| 5 | Painel em lote com regra FORA_DO_CQ/SO_ANALITOS (P-04) | Bio | PROCEDE | Médio | Não |
| 6 | Estatística D4 "2026" em "Tudo" (P-09b) | Bio | PROCEDE | Baixo | Não |
| 7 | Aviso da janela de 180 aponta para a Estatística, que mostra outro recorte (P-05, OC-0012) | Hema | PROCEDE | Médio | Não |
| 8 | Lista do MOTIVO com um item só (P-02) | Hema | PROCEDE | **Alto** | Não |
| 9 | OC-0021: DETALHE do IMPLAUSIVEL sem acento (P-10) | Hema | PROCEDE | Baixo | Não (igual ao item 1) |
| 10 | Janela do CEQ: Controllab mapeado com R4 = CAP diz "entra no bias" (P-06) | Hema | PROCEDE | Médio | Não |
| 11 | Painel em lote FORA_DO_CQ (P-07) | Hema | PROCEDE | Médio (latente) | Não |

As correções em comum (2/8, 4/10, 5/11, 1/9, 3/6) devem sair numa entrega só, nos dois produtos. Assim o mesmo texto e
a mesma regra de tela valem nos dois.

---

## 1 e 9. OC-0021: DETALHE sem acento (Bio P-08, Hema P-10). PROCEDE, Baixo

**Avaliação.** É um defeito, não uma preferência. O próprio critério da OC-0021 pedia "DETALHE com acentos". O impacto,
porém, é de acabamento: o texto continua compreensível e não leva a erro. Mesmo assim o problema volta, porque a lista
fixa do ADR-090 (`camada_dados.ACENTOS_DETALHE`, 8 palavras) não acompanha os textos novos. Pela leitura de
`src/comum/mEtlCQ.bas`, os textos que hoje podem chegar ao DETALHE sem acento são estes:
- `TXT_IMPLAUSIVEL`, `TXT_LOTE_FORA`, `TXT_LOTE_MZSEM` e `TXT_LOTE_OUTRO`: "implausivel", "estatistica", "grafico" e
  "catalogo";
- `TextoMzCorr` (ADR-093): "lancado";
- os textos da substituição (linhas 51, 797, 808, 810 e 1249): "nao", "substituido", "so" e "ja".

**Correção proposta (sem mudar o que é gravado):**
1. Completar `ACENTOS_DETALHE` com: `implausivel→implausível`, `estatistica→estatística`, `grafico→gráfico`,
   `lancado→lançado`, `catalogo→catálogo` e `substituido→substituído`. As palavras curtas entram **com os espaços**,
   para não atingir "janela", "soro" ou "nação": `" nao "→" não "`, `" ja "→" já "` e `" so "→" só "`. Ficam 17
   SUBSTITUTE aninhados, bem abaixo do limite do Excel.
2. Acrescentar uma **trava contra textos novos** na suíte (por exemplo, `qa_adr090` O21). O teste:
   - extrai todos os textos literais que os dois motores podem gravar no `MOTIVO_EXCLUSAO_AUTOMATICA`
     (`mEtlCQ.bas` e `pq/DB_CQ_FINAL.m`);
   - aplica a lista;
   - **reprova** se sobrar alguma palavra de um dicionário de formas sem acento (`nao`, `ja`, `so`, `estatistica`,
     `grafico`, `implausivel`, `lancado`, `catalogo`, `substituido`, `periodo`, `nivel`, `codigo`, `numero`,
     `historico`).
   Assim, um texto novo sem acento falha na suíte antes de chegar à tela.
3. O `MOTIVO_EXCLUSAO_AUTOMATICA` da tblCQ_Final continua sem acento, igual nos dois motores, porque o `qa_sync` S6 os
   compara. Acentuar na origem (D-02) muda o texto gravado nos dois motores e pode afetar quem lê essa coluna. Por
   isso **depende de decisão do Orquestrador** e **não é recomendado agora**: a lista mais a trava resolvem a tela.

**Critério de aceite (reteste):**
- Na cópia, lançar manuais: −1 (IMPLAUSIVEL); "Creatinina" no lote 6354 (ADR-093, Bio); um em lote FORA_DO_CQ de teste;
  e um SUBSTITUI_ID inexistente. Depois, ATUALIZAR.
- O DETALHE mostra "implausível", "estatística", "gráfico", "lançado", "catálogo" e "não". Nenhuma palavra do
  dicionário aparece sem acento em nenhuma das 300 linhas da Digitar Resultados, nos dois produtos.
- Códigos em maiúsculas (IMPLAUSIVEL, FORA_DO_CQ, SO_ANALITOS, tblLotesMaterial) e nomes de analito digitados ficam
  intactos.
- A tblCQ_Final continua igual entre os dois motores (S6 sem diferença nova).
- A trava reprova num caso sintético (um texto "fora da estatistica" acrescentado só no teste).

---

## 2. Listas EQUIPAMENTO, MATRIZ e MOTIVO com um item só (Bio, P-01). PROCEDE, Alto

**Avaliação.** É um problema funcional grave do ponto de vista de quem digita.
- **A primeira tentativa falha.** O técnico vê o âmbar ("confira"), digita "DIMENSION 1", como o GUIA §3 manda, e
  recebe "Escolha o equipamento da lista.".
- **O menu induz ao erro.** Ele abre a seta e há **uma** opção, "DIMENSION 1,DIMENSION 2", que parece certa para quem
  está com pressa. O manual gravado com ela entra ATIVO, com PARTICIPA SIM e equipamento inexistente, sem achado no QA
  e fora das duas séries do Painel. A pessoa acha que lançou, e o ponto não aparece em lugar nenhum do gráfico.
- **O MOTIVO não pode ser escolhido**, e ele é obrigatório no SUBSTITUI_ID (ADR-082).

**Causa** (`camada_dados.py:532-537`): `','.join(...)` grava uma lista literal com vírgula pelo COM, e o Excel pt-BR a
lê com o separador ";". É a mesma família do G18/G32. LOTE e ANALITO funcionam porque apontam para nomes (`=lstLotes`,
`=lstAnalitos`).

**Correção proposta (a lista; sem autorização):**
1. Pôr os itens num **intervalo**, como já se faz em LOTE e ANALITO, e não numa lista literal:
   - **Onde:** um bloco de 3 colunas numa aba de configuração oculta e protegida (por exemplo, `Cfg_Integracao`, em
     colunas livres conferidas antes, como o ADR-090 fez com `Eng_Saida!AE1:AF1`). Cabeçalhos `LISTA_EQUIPAMENTO`,
     `LISTA_MATRIZ` e `LISTA_MOTIVO`, uma linha por item, vindos de `p['equipamentos']`, das matrizes do de/para e de
     `MOTIVOS_MANUAL`.
   - **Nomes:** `lstEquipManual`, `lstMatrizManual` e `lstMotivoManual`, criados **só** com
     `wb.Names.Add(nome, referencia)`, de dois argumentos, como o `nome()` do `instalar_adr078.py`. **Nunca** usar
     `Names.Add` com TRUE (armadilha registrada).
   - **Validação:** `Formula1 = "=lstEquipManual"` (idem os outros dois), estilo **Parar**, `IgnoreBlank = True`. A
     referência a intervalo não depende do idioma do Excel.
2. Mensagens de erro, texto exato:
   - EQUIPAMENTO: "Escolha o equipamento na seta da célula (um por linha da lista)."
   - MATRIZ: "Escolha a matriz na seta da célula."
   - MOTIVO: "Escolha o motivo na seta da célula. Ele é obrigatório quando o manual substitui um resultado."
   (Todas abaixo do limite de 225 caracteres do `ErrorMessage`.)
3. Reinstalar nas **duas** produções. A Hema usa o mesmo código. Corrigir, no mesmo passo, o fallback
   `Validation.Add(3, 1, 1, 'SIM,NÃO')` da `tblInativacao_NaoConformes` (`camada_dados.py:503`), que tem o mesmo
   defeito quando não há caixa de seleção nativa.
4. Antes de reinstalar, conferir numa **cópia** se a produção já tem manual com EQUIPAMENTO, MATRIZ ou MOTIVO que não
   esteja na lista nova (por exemplo, "DIMENSION 1,DIMENSION 2"). Se houver, levar ao Orquestrador. O instalador
   **não** corrige registro existente.
5. GUIA §3: trocar "confira (ou digite o certo)" por "confira; se o equipamento for outro, escolha-o na seta da
   célula (DIMENSION 1 ou DIMENSION 2)".

**Depende de autorização (não fazer sem ela):**
- um QA novo para EQUIPAMENTO fora da lista do produto, que mexe nos dois motores e no vocabulário do QA;
- qualquer regra que tire da estatística o manual com equipamento inválido.
Com a lista certa, a entrada pela tela já não produz esse valor. O que sobra é a colagem por cima da validação, que a
trava do QA cobriria (D-04).

**Prioridade:** Alto. Sobe para **Crítico** se o reteste mostrar que o manual "DIMENSION 1,DIMENSION 2" muda algum
número. A `EstatBasica` da Bio filtra por analito e lote, não por equipamento, então ele pode estar no n da
Estatística, no CVp ou no u(Rw). Isso seria "associação errada a equipamento" (REGRAS §4). Peço ao Tester/QA que
confira.

**Critério de aceite (reteste, na cópia, nos dois produtos):**
- `Validation.Formula1` das três colunas aponta para um nome ou intervalo, não para uma lista literal.
- O menu mostra **2** equipamentos (Bio), as matrizes uma por linha e **4** motivos.
- `Validation.Value` é **Verdadeiro** para "DIMENSION 1", "DIMENSION 2", "SORO", "URINA" e cada um dos 4 motivos,
  inclusive "Resultado não transmitido" e "Equipamento em contingência", com acento. É **Falso** para
  "DIMENSION 1,DIMENSION 2" e para "Outro motivo".
- Um manual com EQUIPAMENTO "DIMENSION 2" escolhido na lista, depois de ATUALIZAR, aparece na série do DIMENSION 2 do
  Painel. O âmbar some ao escolher.
- Um manual com SUBSTITUI_ID e MOTIVO "Falha do interfaceamento" é aceito, e o Audit_Log grava o motivo exato.
- `instalar_*` na 2ª execução: 0 mudanças.
- Nenhuma lista literal com "," nas validações da Digitar Resultados e da Inativação.
- A auditoria visual da aba continua sem corte nem sobreposição.

---

## 3 e 6. Cabeçalho da Estatística da Bio em "Tudo (sem filtro)" (P-09). PROCEDE, Baixo

**Avaliação.** "00/01/1900" é uma data sem sentido num cabeçalho que se confere e se imprime. "Ano 2026" ao lado de
"histórico inteiro do lote" contradiz o número. Os números estão certos (n 90/88 conferidos), por isso é Baixo.

**Cuidado que muda a correção sugerida pelas frentes.** O motor **lê** essas células.
`src/bio/mEstatistica.bas:2132-2136`, com B4 = PERSONALIZADO, faz `dIniPer = ws.Range("B5").Value` e
`If IsDate(...) Then anoDe = Year(...)`. Trocar o **valor** de B5/D5 de 0 para "" pode mudar `anoDe` e `anoAte`, e com
eles o filtro de ano da `EstatBasica`. Isso seria mudança de cálculo disfarçada de tela. Por isso:

**Correção proposta:**
- **B5 e D5 (item 3): só formato.** Fórmulas inalteradas (`=Painel!G3`, `=Painel!G4`). `NumberFormat` (a propriedade
  em inglês, não a `NumberFormatLocal`) passa a `dd/mm/yyyy;;`. O valor 0 fica em branco, e qualquer data aparece como
  hoje. Nenhum valor muda.
- **D4 (item 6):** `=IF(OR(Painel!K3="",AND(Painel!$G$3="",Painel!$G$4="")),"",Painel!K3)`. É a mesma condição
  `SEM_DEATE` do H5/J5 (ADR-092). Com B4 travado em PERSONALIZADO, o motor não usa a D4 (linha 2139 só roda quando a
  visão não é PERSONALIZADO), e o `JanelaInicio`/`JanelaFim` com PERSONALIZADO devolve B5/D5. A D4 é só exibição.
  Mesmo assim, o reteste confere os números.
- O rótulo "(refletido do Painel)" fica como está. Com a D4 vazia e o H5 dizendo "histórico inteiro do lote", a linha
  fica coerente, como já acontece no De/Até que atravessa o ano.

**Critério de aceite:**
- Painel K3 2026 + K4 "Tudo (sem filtro)", depois abrir a Estatística: B5, D5 e D4 aparecem **vazios** (`.Text` = "").
  H5 diz "histórico inteiro do lote".
- Set/2026: B5/D5 "01/09/2026" e "30/09/2026", D4 "2026".
- De/Até 01/12/2025–31/01/2026: D4 vazio, B5/D5 com as datas.
- C14:M133 da Estatística iguais **bit a bit** (`Value2`) antes e depois da mudança, nos estados "Tudo", Set/2026 e
  T3/2026.
- `auditoria_abas` sem "####" nem corte em A4:J5.

---

## 4 e 10. Janela do CEQ: Controllab mapeado com o provedor do analito = CAP (Bio P-03, Hema P-06). PROCEDE, Médio

**Avaliação.** Procede. O aviso azul existe para dar confiança ("vai para o bias deste analito"). Quando ele diz isso e
o número não muda, o analista conclui o contrário do que acontece e pode deixar de procurar o motivo de um Sigma que
"não reagiu". Hoje é latente (mapas do Controllab vazios), mas aparece no primeiro mapeamento. Médio está certo.

A regra fica intacta: CAP e Controllab continuam sem se misturar. A alternativa de recusar o mapa do Controllab
enquanto o provedor for CAP (D-06) muda o fluxo de cadastro e **não é recomendada**. O laboratório pode mapear antes de
trocar o provedor.

**Correção proposta (`mCEQForm.AvisoAnalitoCEQ` e `AvisoGravado`):**
- Depois de achar o canônico, ler o provedor **do analito** pela **mesma fonte** que a Estatística passa ao `BiasEQ`.
  Na Bio, `Especificações!AR` da linha do canônico. Na Hema, `Estatística!R4`, que é global. Nada de regra nova: se o
  AR vazio vale CAP para o `BiasEQ`, vale CAP aqui.
- Se o provedor digitado for **diferente**, devolver âmbar, e não azul. Textos exatos:
  - Bio: `AMBAR|⚠ No QC_INI: <canônico>. O bias deste analito usa o <prov. do analito> (Especificações, coluna AR):
    este resultado do <prov. digitado> fica gravado, mas NÃO entra no bias, no ET nem no Sigma.`
  - Hema: `AMBAR|⚠ No QC_INI: <canônico>. O bias da Hematologia usa o <R4> (aba Estatística, R4): este resultado do
    <prov. digitado> fica gravado, mas NÃO entra no bias, no ET nem no Sigma.`
  - O "Atenção: existe também …" do ADR-090 continua sendo acrescentado quando houver par.
- Retorno do Salvar (`AvisoGravado`), no mesmo caso: `Gravado fora do bias: o bias de <canônico> usa o <prov. do
  analito>, não o <prov. digitado>.`
- Tamanho: com "Hemoglobina glicada (HbA1c)" o texto da Bio tem 187 caracteres e o da Hema 178. São menores que o
  âmbar mais longo já medido (195, em `ceq_form_spec.py`). Acrescentar os dois ao conjunto de textos de tempo de
  execução do `ceq_form_spec`, para o `auditoria_janelas` medir.

**Critério de aceite (reteste, cópia):**
- Mapa de teste "TESTE-AUDITORIA GLU CTL" → Glicose (Bio) ou nome → HGB (Hema), com o AR/R4 = CAP: o aviso começa com
  "AMBAR|⚠ No QC_INI:" e contém "NÃO entra no bias". O `SalvarCEQ` devolve "OK|n|Gravado fora do bias: o bias de … usa
  o CAP, não o Controllab.".
- A Estatística continua sem mudança (0 células).
- Com o R4 = Controllab (Hema), o mesmo nome dá "AZUL|… entra no bias" e o número muda, como no RX deste ciclo.
- Nome do CAP mapeado com o AR/R4 = CAP continua azul. Nome não mapeado continua com o âmbar atual.
- `auditoria_janelas`: nenhum texto cortado no `lblMsg`.

---

## 5 e 11. Painel em lote com regra da tblLotesMaterial (Bio P-04, Hema P-07). PROCEDE, Médio

**Avaliação.** Procede, e na Bio **não é latente**: os lotes 8794 e 8353 estão no H3. A tela dá três sinais errados
juntos:
- "SEM CORRIDA" e "o lote não tem resultado": o lote tem 31 resultados de Glicose;
- "SEM MÉD/DP — digite a Média e o DP": induz a cadastrar alvo num lote que não é de CQ;
- nenhuma menção à regra que explica tudo.
A Hema tem o mesmo texto, latente (tabela vazia). O número está certo, por isso fica em Médio.

**Correção proposta (só tela; a regra ADR-091 não muda):**
1. **Motor:** ao montar o Painel, o motor já sabe o status LOTE_* das linhas do lote. Ele passa a gravar numa célula
   livre do `Eng_Saida`, conferida antes como no ADR-090 (por exemplo, `AG1`), o texto da regra que vale para
   `loteAnalise` + `selAnalito`. Fica vazio sem regra. Formatos:
   - `FORA_DO_CQ|<n>`;
   - `SO_ANALITOS|<analitos da regra>|<n>`.
   Aqui `<n>` é a quantidade de resultados LOTE_* do par. Se o `<n>` for caro, pode ser omitido.
   Nome `engRegraLote`, criado com `wb.Names.Add(nome, ref)`, de dois argumentos.
2. **J7:J9 (forma curta):** um ramo novo **depois** de DESATUAL./BASE VAZIA e **antes** de "SEM MÉD/DP":
   `IF(engRegraLote<>"","FORA DO CQ", …)`. "FORA DO CQ" tem 10 caracteres, menos que "SEM CORRIDA", e cabe nos 48 pt.
   Cor: a mesma regra neutra de "SEM CORRIDA". Não é falha do método, e por isso não vai em vermelho.
3. **Q4 (forma longa):** `IF(COUNTIF(J,"FORA DO CQ")>0,"LOTE FORA DO CQ (REGRA DO LOTE)", …)`, também antes de "SEM
   MÉDIA/DP". Acrescentar ao `STATUS_LONGO`.
4. **A39:** com `engRegraLote` preenchido, o texto substitui o "⛔ … SEM MÉDIA/DP … digite a Média e o DP" e o "O lote
   não tem resultado". Textos exatos:
   - FORA_DO_CQ: `⚠ LOTE 8794 · Glicose: lote FORA DO CQ (regra da tblLotesMaterial). Os 31 resultados dele ficam
     guardados, fora da estatística e do gráfico. Escolha outro lote no H3.`
   - SO_ANALITOS: `⚠ LOTE 8353 · Glicose: este lote só vale para Proteína C-reativa (regra da tblLotesMaterial). Os 31
     resultados de Glicose ficam guardados, fora da estatística e do gráfico. Escolha outro lote no H3.`
   - Sem `<n>`, trocar "Os 31 resultados …" por "Os resultados …".
   - O restante do A39 (FILTRO, janela de 180, período) continua sendo acrescentado depois.
5. **Suítes:** `qa_troca_lote`, `qa_varredura_lotes` e `qa_linhas_alvo`, que esperam "SEM MÉD/DP" em lotes sem
   média, precisam reconhecer o estado novo **só** nos lotes com regra. Registrar no ADR o porquê.
6. **Fora desta correção:** a troca automática de analito no lote SO_ANALITOS da Hema (D-13) não foi enviada a este
   parecer.

**Critério de aceite (reteste):**
- Bio, H3 8794 + Glicose e H3 8353 + Glicose:
  - J7:J8 "FORA DO CQ";
  - Q4 "LOTE FORA DO CQ (REGRA DO LOTE)";
  - A39 com o texto da regra e o número de resultados;
  - **nenhuma** ocorrência de "digite a Média e o DP" nem de "não tem resultado";
  - nenhum ponto no gráfico.
- Bio, H3 8353 + Proteína C-reativa: pontos e status normais, sem o texto da regra.
- Hema, regra de teste 533811 FORA_DO_CQ + WBC: o mesmo.
- Lote sem regra e sem Média/DP: continua "SEM MÉD/DP" e o texto antigo.
- `engPainel`, Calc e Estatística iguais bit a bit com e sem a mudança.
- `auditoria_abas`: J7:J9 dentro do "reduzir para caber" até 30 %, e A39 sem corte com o texto mais longo, de
  SO_ANALITOS, com o nome de analito mais longo.

---

## 7. Aviso da janela de 180 corridas na Hema aponta para outro recorte (P-05, efeito colateral da OC-0012). PROCEDE, Médio

**Avaliação.** Procede. Uma instrução que leva a um número diferente é pior do que não ter instrução. Isso acontece
justamente com o Painel em REJEITADO, quando o usuário mais precisa confiar no que lê. A Estatística da Hema tem filtro
próprio (Ano De/Até, lote B4), não segue o mês nem o modo TODOS do Painel, e por isso **nunca** garante o mesmo
recorte. Na Bio, o ponteiro funciona e fica.

**Correção proposta (texto do A39, por produto; `instalar_adr078.a39` recebe o produto):**
- **Hema**, com o H3 vazio:
  `⚠ 278 corrida(s) antigas do período fora do gráfico e dos indicadores do Painel (máx. 180). Período inteiro (todos
  os lotes): N1 374 · N2 400 · N3 340 resultados.`
- **Hema**, com o H3 preenchido:
  `… Período inteiro do lote: N1 … · N2 … · N3 … resultados.`
- Nos dois casos da Hema, **sem** "veja a aba Estatística".
- **Bio:** sem mudança.

O AF1 está certo e continua sendo a fonte.

**Observação para o Orquestrador** (não está nesta lista): a Estatística da Hema mostra lote B4 = 522611 "(refletido
do Painel)" com o H3 do Painel vazio. O rótulo promete um vínculo que não existe nesse estado. Vale abrir como Baixo
se o Usuário Simulado confirmar.

**Critério de aceite:**
- Hema, WBC, H3 vazio, K3 2026 + "Ano inteiro": A39 contém "Período inteiro (todos os lotes): N1 374 · N2 400 · N3
  340" e **não** contém "Estatística".
- Hema, lote 528211 em "Tudo": sem aviso de corte (M1 0), como hoje.
- Bio, Glicose 8974 "Tudo": o texto continua com "— veja a aba Estatística", e 174/161 conferem nos dois lugares.
- `auditoria_abas`: A39 sem corte.

---

## 8. Lista do MOTIVO com um item só (Hema, P-02). PROCEDE, Alto

**Avaliação.** É a mesma causa e a mesma correção do item 2. Na Hema só o MOTIVO quebra, porque EQUIPAMENTO e MATRIZ
têm um item só. O efeito é Alto: a substituição de um resultado (ADR-082) não consegue registrar um motivo válido pela
tela, e quem usa o menu grava no Audit_Log e no I04 um motivo sem sentido. Isso compromete a trilha de auditoria, que
é exatamente o que o MOTIVO existe para garantir.

**Correção proposta:** a do item 2, no mesmo instalador e na mesma entrega, com `lstMotivoManual` (4 itens).
EQUIPAMENTO e MATRIZ passam para nomes também, por uniformidade: um equipamento novo na Hema não reabriria o defeito.

**Critério de aceite:**
- `Validation.Value` Verdadeiro para "Falha do interfaceamento", "Resultado não transmitido", "Equipamento em
  contingência" e "Outro"; Falso para o texto inteiro com vírgulas.
- Manual com SUBSTITUI_ID real e MOTIVO "Resultado não transmitido": aceito. O Audit_Log grava o motivo com acento, e o
  resultado substituído fica SUBSTITUIDO_MANUAL, como na regra atual.
- EQUIPAMENTO "XN1000" e MATRIZ "SANGUE TOTAL" continuam aceitos.
- 2ª execução do instalador: 0 mudanças.

---

## Observações gerais

- **Ordem de entrega sugerida:**
  1. itens 2 e 8 (Altos, mesma correção);
  2. itens 5 e 11, que já acontecem na Bio;
  3. item 7;
  4. itens 4 e 10;
  5. itens 1, 9, 3 e 6, que são de acabamento e podem ir junto.
- **Dúvidas que as frentes endereçaram ao UX** (D-03, D-10, D-13, D-14 e D-15) **não vieram nesta lista** e não foram
  avaliadas. Posso dar parecer se o Orquestrador as enviar. A D-10 (instrução A3 da Hema promete âmbar) convém sair
  junto com o item 2, porque mexe no mesmo texto da Digitar Resultados.
- **Nenhum dado de paciente, senha ou caminho de rede** consta deste parecer.
