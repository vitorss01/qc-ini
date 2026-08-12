# Análise de escalabilidade, tamanho e performance — QC_Bioquimica.xlsm

Data da análise: 11/08/2026
Arquivo analisado: `QC_Bioquimica.xlsm` (produção), 1.789.049 bytes
Máquina de teste: i7-1255U (10 núcleos / 12 threads), 16 GB RAM, Excel 365 x64, Windows 11
Amostra de carga: janeiro/2026, dado real

**Nenhum arquivo de produção foi alterado.** Todos os testes rodaram sobre cópias
no diretório temporário da sessão.

Convenção de rótulos, usada em todo o documento:

- **MEDIDO** — número lido do arquivo ou cronometrado em teste físico
- **CALCULADO** — aritmética direta sobre números medidos
- **ESTIMADO** — extrapolação por curva ajustada, com intervalo
- **INFERIDO** — conclusão de leitura de código, sem medição direta

---

## VEREDITO EXECUTIVO

O que limita esta planilha **não é o tamanho do arquivo e não é o número de
linhas do Excel**. Nos dois casos a folga é de décadas.

O que limita são duas coisas, nesta ordem:

1. **Tetos de provisionamento fixo**, que são atingidos **em silêncio** — sem erro,
   sem aviso, com o sistema continuando a parecer correto. O primeiro deles
   chega por volta do **mês 6 de uso de um mesmo lote**, e o mais grave (o do
   banco) por volta do **mês 13 de histórico total**.
2. **Duas colunas de fórmula com custo quadrático** no `DB_Resultados`, que
   tornam o recálculo completo desconfortável a partir de **~30 meses** e
   inviável a partir de **~48 meses**.

> **Com base no mês de janeiro, esta planilha comporta com conforto
> aproximadamente 1,1 ano de dados na arquitetura atual — estimativa central
> 13 meses, intervalo plausível 11 a 15 meses — e não 5 anos.**
>
> **Corrigidos três pontos localizados (descritos na seção H), a mesma
> arquitetura passa a comportar 5 anos com folga confortável — estimativa
> central 7 anos, intervalo plausível 5 a 10 anos.**

A distinção importa: a resposta para "5 anos?" **hoje** é *provavelmente não*,
mas o motivo não é "Excel não aguenta". É que três limites foram dimensionados
para um ano de operação. Levantá-los é trabalho de dias, não de reescrita.

E o risco relevante aqui não é lentidão. É que **os três tetos truncam dados sem
avisar** — num sistema de controle de qualidade, um número plausível e errado é
pior do que um erro visível.

---

## A. DIAGNÓSTICO ATUAL — inventário técnico

### A.1 O arquivo por dentro (MEDIDO)

| Item | Valor |
|---|---:|
| Tamanho em disco | 1,706 MB (1.789.049 bytes) |
| Partes OOXML | 77 |
| Conteúdo expandido | 13,34 MB (fator de compressão 7,9×) |
| Abas | 20 (14 visíveis, 6 `veryHidden`) |
| Células ocupadas (todas as abas) | 240.052 |
| — com fórmula | 56.956 |
| — com valor | 183.096 |
| Entradas no `calcChain` | 56.910 |
| Intervalos nomeados | 59 |
| Fórmulas matriciais (`t="array"`) | 1 |
| Regras de formatação condicional | 13 |
| Validações de dados | 18 |
| Gráficos | 2 |
| Objetos de desenho / botões | 19 / 13 controles |
| Tabelas estruturadas | 0 |
| Tabelas dinâmicas | 0 |
| Conexões externas | 0 |
| Consultas Power Query | 0 |
| Módulos VBA | 37 (3.809 linhas, 457 KB expandidos) |
| Eventos `Workbook_*` / `Worksheet_*` | sim (`EstaPastaDeTrabalho` + 4 módulos de planilha) |

### A.2 Onde os bytes moram (MEDIDO)

| Categoria | KB comprimidos | % do arquivo |
|---|---:|---:|
| Planilhas (worksheets) | 1.339,0 | 77,1% |
| VBA (`vbaProject.bin`) | 197,6 | 11,4% |
| `calcChain.xml` | 149,4 | 8,6% |
| Gráficos e desenhos | 30,4 | 1,8% |
| Estilos, tema, textos, outros | 19,5 | 1,1% |

Uma única aba — `DB_Resultados`, com 1.002 KB — responde por **56% do arquivo
inteiro**. E ela está quase vazia de dado: 1.110 registros ocupando 165.079
células, das quais **45.056 são fórmulas pré-provisionadas para 15.000 linhas
que ainda não existem**.

### A.3 Por aba (MEDIDO)

| Aba | Estado | KB | Células | Fórmulas | Valores | Última linha | CF | DV |
|---|---|---:|---:|---:|---:|---:|---:|---:|
| DB_Resultados | visível | 1.002,0 | 165.079 | 45.056 | 120.023 | 15.003 | 0 | 5 |
| Resultados | visível | 134,2 | 36.066 | 4 | 36.062 | 3.003 | 0 | 0 |
| Calc | veryHidden | 52,6 | 8.879 | 8.650 | 229 | 182 | 0 | 0 |
| EQC_Dados | visível | 43,7 | 11.405 | 363 | 11.042 | 1.003 | 3 | 0 |
| Estatística | visível | 27,0 | 2.258 | 1.940 | 318 | 93 | 3 | 3 |
| Importar | visível | 18,0 | 6.868 | 0 | 6.868 | 204 | 0 | 0 |
| Analitos | visível | 11,7 | 1.198 | 482 | 716 | 43 | 0 | 1 |
| Registros | visível | 10,3 | 2.639 | 5 | 2.634 | 203 | 1 | 3 |
| RegistrosStore | veryHidden | 7,1 | 2.412 | 0 | 2.412 | 201 | 0 | 0 |
| Liberação | visível | 7,0 | 1.221 | 401 | 820 | 203 | 0 | 0 |
| demais 10 abas | — | 21,4 | 2.027 | 55 | 1.972 | ≤200 | 6 | 6 |

### A.4 O que cresce e o que não cresce (INFERIDO da estrutura + MEDIDO)

**Cresce com o histórico (grupo A — linear):**

- linhas de dado no `DB_Resultados` (colunas A:G) — **+1.110/mês**
- linhas da view `Resultados` — **+555/mês por nível**, enquanto o lote não muda
- linhas da `Liberação` — **+18/mês por lote**
- linhas do `EQC_Dados` — **+30 por ciclo de EQ** (hoje 3 ciclos/ano, 6 analitos)

**Praticamente fixo (grupo B — não cresce):**

- os 37 módulos VBA e os 6 formulários (457 KB) — constante
- os 2 gráficos: as séries apontam para `Calc!$X$3:$X$182`, **180 pontos fixos**
- a aba `Calc` inteira: 182 linhas, sempre
- `Analitos`, `Cfg_Especificacoes`, `DB_Especificacoes`, `Eng_Especificacoes`,
  `Usuarios`, `Configuração`, `Login`, `Início` — todas dimensionadas por
  analito ou por usuário, não por tempo
- estilos, tema, formatação condicional, validações, botões
- a aba `Estatística`: 1.940 fórmulas sobre faixas fixas de 1.000 linhas do
  `EQC_Dados` e sobre a UDF `EstatPeriodo`

**Cresce de forma NÃO linear (grupo C — o problema):**

- `DB_Resultados!BB` e `!BC` — 30.000 células com `COUNTIFS` de **faixa
  expansiva**. Detalhe na seção G.
- a aba `Calc` — 180 linhas × `AGGREGATE`/`COUNTIFS`/`SUMIFS` sobre os intervalos
  nomeados `r*`, que hoje têm 15.000 linhas. O custo é O(180 × N): a aba não
  cresce, mas **cada uma das suas 8.650 fórmulas fica mais cara** conforme o
  banco cresce.

---

## B. TAXA DE CRESCIMENTO — janeiro como unidade de carga

### B.1 O que janeiro contém (MEDIDO)

| Grandeza | Valor |
|---|---:|
| Registros no `DB_Resultados` | 1.110 |
| Período | 05/01/2026 a 30/01/2026 |
| Dias com corrida | 18 |
| Corridas (RUNs) | 18 — exatamente 1 por dia |
| Analitos distintos | 31 |
| Níveis | 2 (555 registros cada) |
| Lotes | 2 (`QC-897401`, `QC-897402`) |
| Registros por corrida | 61,7 (média) — 62 típico, 56 no dia 21 |
| Registros por dia | mín 56, máx 62, mediana 62 |

A carga é **notavelmente regular**: 31 analitos × 2 níveis = 62 registros, uma
corrida por dia útil. O desvio observado (o dia 21, com 56) corresponde a 3
analitos não medidos. Isso torna a projeção mais confiável do que o usual —
a variabilidade mensal vem quase só do número de dias úteis (18 a 23).

### B.2 Custo em bytes de cada mês (MEDIDO)

| Componente | KB por mês |
|---|---:|
| Dado bruto (colunas A:G do banco) | 14,9 |
| Fórmulas `BA:BC` correspondentes | 72,3 |
| **Total por mês** | **87,2** |
| **Total por ano** | **1,02 MB** |

Enquanto o provisionamento de 15.000 linhas não for ultrapassado, o arquivo
cresce só os 14,9 KB/mês do dado — as fórmulas já estão lá, pagas.

---

## C. PROJEÇÃO — três cenários

### Premissas

- **CONSERVADOR** — cada mês igual a janeiro: 1.110 registros/mês. É o cenário
  base, e é bem sustentado: a rotina de CQ de um laboratório é estável por
  definição normativa, não por acaso.
- **REALISTA** — 1.110/mês nos 12 primeiros meses, depois +8% ao ano, refletindo
  crescimento vegetativo do menu de analitos (o produto já saltou de 20 para 31
  analitos em uma revisão) e eventual segunda corrida diária em parte dos dias.
- **PESADO** — 2 corridas/dia e 3 níveis desde o início: 1.110 × 2 × 1,5 =
  3.330/mês. É o laboratório de grande porte, e é o que a Hematologia (NLV = 3)
  já representa hoje.

> Não há evidência de sazonalidade na estrutura da planilha — só um mês está
> registrado. **As projeções assumem estabilidade do volume mensal**, corrigida
> apenas pelo número de dias úteis.

### C.1 Cenário CONSERVADOR (CALCULADO sobre janeiro; tamanhos MEDIDOS)

| Período | Meses | Registros | Linhas no banco | Fórmulas no banco | Tamanho | Observação |
|---|---:|---:|---:|---:|---:|---|
| Janeiro | 1 | 1.110 | 1.113 | 45.056 | 1,70 MB | **observado** |
| 3 meses | 3 | 3.330 | 3.333 | 45.056 | 1,73 MB | **medido** |
| 6 meses | 6 | 6.660 | 6.663 | 45.056 | 1,77 MB | **medido** |
| 12 meses | 12 | 13.320 | 13.323 | 45.056 | 1,86 MB | **medido** |
| 24 meses | 24 | 26.640 | 26.643 | 79.929 | 2,79 MB | medido (com provisionamento estendido) |
| 36 meses | 36 | 39.960 | 39.963 | 119.889 | 3,84 MB | medido (idem) |
| 48 meses | 48 | 53.280 | 53.283 | 159.849 | 4,89 MB | medido (idem) |
| 60 meses | 60 | 66.600 | 66.603 | 199.809 | 5,94 MB | medido (idem) |

### C.2 Tamanho nos três cenários (ESTIMADO a partir da inclinação medida de 87,2 KB/mês)

| Horizonte | Conservador | Realista | Pesado |
|---|---:|---:|---:|
| 1 ano | 1,86 MB | 1,86 MB | 2,9 MB |
| 2 anos | 2,79 MB | 2,9 MB | 5,1 MB |
| 3 anos | 3,84 MB | 4,1 MB | 7,5 MB |
| 4 anos | 4,89 MB | 5,4 MB | 10,1 MB |
| 5 anos | 5,94 MB | 6,8 MB | 12,8 MB |

**Em nenhum cenário o tamanho do arquivo é um problema.**

---

## D. PERFORMANCE — medida, separada por dimensão

Os testes abrem, recalculam e salvam cópias reais em cada volume. O ponto
central da seção 5 do pedido está confirmado pelos números: **tamanho e
performance divergem completamente aqui**. O arquivo cresce 3,5× de 1 para 60
meses; o recálculo completo cresce 138×.

### D.1 As dimensões, cada uma por si (MEDIDO)

| Dimensão | 1 mês | 12 meses | 24 meses | 36 meses | 48 meses | 60 meses |
|---|---:|---:|---:|---:|---:|---:|
| **A.** Tamanho do arquivo | 1,70 MB | 1,86 MB | 2,80 MB | 3,86 MB | 4,91 MB | 5,96 MB |
| **B.** RAM do processo EXCEL | 424 MB | 421 MB | 585 MB | 925 MB | 1.434 MB | 2.028 MB |
| **C.** Tempo de abertura | 0,93 s | 0,94 s | 1,83 s | 4,00 s | 4,24 s | 6,46 s |
| **D.** Tempo de salvamento | 0,66 s | 0,73 s | 1,15 s | 2,00 s | 3,51 s | 2,07 s |
| **E.** Recálculo completo | 2,94 s | 5,67 s | 28,49 s | 71,84 s | 120,86 s | 220,66 s |
| **J.** Digitar um resultado | 0,03 s | 0,10 s | 0,02 s | 0,02 s | 0,02 s | 0,03 s |
| **H.** Pontos no gráfico L-J | 18 | 180 | 180 | 180 | 180 | 180 |

> As colunas de 24 a 60 meses são a **medição válida**, com os intervalos
> nomeados saneados (ver apêndice). O gráfico manteve os 180 pontos em todos os
> volumes: **o Painel não quebra por volume de histórico.**

### D.2 As rotinas que o usuário dispara (MEDIDO)

| Operação | 1 mês | 12 meses | 24 meses | 36 meses | 60 meses |
|---|---:|---:|---:|---:|---:|
| **F.** `AtualizarViewResultados` | 0,08 s | 0,15 s | 0,19 s | 0,19 s | 0,24 s |
| **F.** `AtualizarOperacao` | 0,04 s | 0,12 s | 0,17 s | 0,24 s | 0,57 s |
| **I.** Motor da Estatística, 1ª chamada | 0,01 s | 0,11 s | 0,22 s | 0,26 s | — |
| **I.** Motor da Estatística, com cache | 0,00 s | 0,00 s | 0,01 s | 0,00 s | — |

### D.2.1 A ação mais frequente do dia: trocar o analito no Painel (MEDIDO)

O analista revisa 31 analitos, um a um, pelo seletor do Painel. Medido com
cálculo **automático** — só a cadeia suja recalcula, que é o que acontece de
verdade (média de 6 trocas por volume):

| Histórico | 1 mês | 6 m | 12 m | 24 m | 36 m | 48 m | 60 m |
|---|---:|---:|---:|---:|---:|---:|---:|
| Por troca | 0,57 s | 0,61 s | **0,58 s** | 1,01 s | 1,57 s | 2,21 s | 2,68 s |
| Revisar os 31 | 0,3 min | 0,3 min | **0,3 min** | 0,5 min | 0,8 min | 1,1 min | 1,4 min |

O crescimento é **linear**, não quadrático: vem do termo O(180 × N) das fórmulas
matriciais do `Calc`, não das colunas `BB`/`BC`. Mesmo aos 60 meses, revisar o
menu inteiro custa **1,4 minuto de espera acumulada** — incômodo, não impeditivo.

> Isto sustenta a conclusão central da seção D.3: **o uso diário não é o
> gargalo em nenhum horizonte de 5 anos.** O que degrada é o recálculo completo,
> que é operação rara.

**G. Power Query:** não existe no arquivo — 0 conexões, 0 consultas, 0
`queryTables`. Não há o que medir hoje. Ver seção G.6 para o risco futuro.

**H. Dashboards e gráficos:** custo **constante**. As duas séries apontam para
`Calc!$X$3:$X$182` — 180 pontos, sempre, independentemente do histórico. Os
gráficos nunca vão pesar mais do que pesam hoje.

### D.3 A leitura correta desses números

A tabela D.1 mistura duas coisas que a operação real não mistura:

- **Recálculo completo** (`Ctrl+Alt+Shift+F9`) é o pior caso absoluto. Reconstrói
  a árvore de dependências inteira. Acontece raramente: numa mudança de versão,
  num script de manutenção, numa recuperação de arquivo.
- **O que acontece quando o analista digita um resultado** é a linha J: o Excel
  recalcula só as células sujas. **Manteve-se abaixo de 0,1 s em todos os
  volumes, inclusive em 60 meses.**

Essa é a diferença entre "a planilha é lenta" e "uma operação rara é lenta".
Na arquitetura atual, **o uso diário não degrada de forma perceptível até os 60
meses**. O que degrada é o recálculo completo e a abertura.

---

## E. ESTIMATIVA DE 40 MB

Inclinação medida entre 12 e 60 meses: **87,2 KB/mês** (1,02 MB/ano), já
incluindo a extensão das fórmulas `BA:BC`.

| Pergunta | Resposta |
|---|---|
| Quando chega a 40 MB? | **CALCULADO: mês 460 — cerca de 38 anos** |
| No cenário PESADO (3×)? | ESTIMADO: ~13 anos |
| 40 MB significa criar outro arquivo? | **Não.** |

**Os 40 MB não são o limite desta aplicação — nem de perto.** A planilha vai
atingir problemas operacionais sérios com **6 MB**, aos 60 meses, muito antes de
40 MB significarem qualquer coisa.

Distinção pedida explicitamente:

- **Limite técnico do Excel:** 1.048.576 linhas/aba; arquivo x64 limitado pela
  RAM disponível. Arquivos de 40–100 MB são rotineiros. Irrelevante aqui.
- **Limite operacional desta aplicação:** o recálculo completo e os tetos de
  provisionamento. Chegam **primeiro, e com folga**.

Usar MB como indicador de capacidade **nesta planilha específica seria enganoso**:
o arquivo é pequeno porque o dado é numérico e comprime bem; o custo está na
topologia das fórmulas, que o tamanho não mostra.

---

## F. VIABILIDADE DE 5 ANOS

### Resposta: **PROVAVELMENTE NÃO na arquitetura de hoje — SIM, com folga, após três correções pontuais.**

### F.1 Por que não, hoje

Não é performance. São três tetos que truncam dados **sem emitir aviso**:

| # | Teto | Provisionado | Consumo | Satura em | Sintoma |
|---|---|---:|---:|---|---|
| 1 | Fórmulas `BA:BC` + intervalos `r*` do `DB_Resultados` | 15.000 linhas | 1.110/mês | **13,5 meses de histórico total** | registros acima da linha 15.003 ficam fora de **toda** a análise — Painel, Calc, Estatística |
| 2 | View `Resultados` (`VIEW_ROWS = 3000`) | 3.000 linhas/nível | 555/mês/lote | **5,4 meses de um mesmo lote** | `If cont(t) < VIEW_ROWS` descarta o excedente, calado |
| 3 | Janela do `Calc` / gráfico Levey-Jennings | 180 RUNs | 18/mês/lote | **10 meses de um mesmo lote** | `AGGREGATE(15;…;$A3)` pega os 180 RUNs **mais antigos**: as corridas novas somem do gráfico |

Os tetos 2 e 3 são **por lote** e reiniciam na troca de lote — se o laboratório
troca de lote a cada 6 meses, quase não incomodam. O teto 1 **não reinicia
nunca**: é o histórico acumulado, e é o que fecha a porta em ~13 meses.

Tetos secundários, sem urgência (INFERIDO):

- `Liberação`: 200 linhas por lote ÷ 18/mês = **11 meses de um mesmo lote**
- `EQC_Dados`: 1.000 linhas. Hoje 6 analitos inscritos, 3 ciclos/ano = 90
  linhas/ano → 11 anos. **Se os 31 analitos forem inscritos**, passa a 465
  linhas/ano → **2,2 anos**.
- `Registros`: 200 por lote — não é restritivo.

### F.2 Números para 5 anos (60 meses), cenário conservador

| Pergunta | Resposta | Rótulo |
|---|---|---|
| Registros | 66.600 | CALCULADO |
| Linhas no `DB_Resultados` | 66.603 | CALCULADO |
| Fórmulas no banco | 199.809 (`BA`, `BB`, `BC`) | CALCULADO |
| Fórmulas no arquivo inteiro | ~212.000 | CALCULADO |
| Tamanho do arquivo | 5,94 MB | **MEDIDO** |
| RAM do Excel | 2,03 GB | **MEDIDO** |
| Abertura | 6,5 s | **MEDIDO** |
| Salvamento | 2,1 s | **MEDIDO** |
| Recálculo completo | 221 s (3 min 41 s) | **MEDIDO** |
| Digitar um resultado | 0,03 s | **MEDIDO** |
| Carga computacional do recálculo | ~1,1 × 10¹⁰ comparações de célula | CALCULADO |

### F.3 Por que 5 anos passa a ser viável após as correções

O termo quadrático vem **inteiro** das colunas `BB` e `BC`. Removido ele, sobra
a parte linear — abertura, salvamento, a varredura O(N) do motor da Estatística
(que já é memoizada) e o O(180 × N) do `Calc`. Todos crescem em linha reta e
todos estão hoje na casa de décimos de segundo.

ESTIMADO, com o termo n² eliminado: recálculo completo aos 60 meses cai para a
faixa de **8 a 20 s** (contra 406 s), e o horizonte confortável passa para
**~7 anos (intervalo 5 a 10)**, limitado então pela RAM e pelo tempo de abertura.

---

## G. PRINCIPAL GARGALO

### G.1 O gargalo, nomeado

`DB_Resultados!BB` e `DB_Resultados!BC`:

```
BB4: =IF(OR($E4="";$G4<>"Ativo");"";IF(COUNTIFS($E$4:$E4;$E4;$A$4:$A4;$A4;$G$4:$G4;"Ativo")=1;1;0))
BC4: =IF(OR($A4="";$G4<>"Ativo");"";IF(COUNTIFS($A$4:$A4;$A4;$G$4:$G4;"Ativo")=1;1;0))
```

A faixa `$E$4:$E4` é **expansiva**: ancorada no topo, solta embaixo. Na linha
4 ela varre 1 célula; na linha 66.603 varre 66.600. Somando as linhas:

```
custo ≈ Σ (3i + 2i) = 2,5 × n²
```

- n = 1.110 (hoje): 3,1 milhões de comparações — imperceptível
- n = 13.320 (1 ano): 443 milhões
- n = 66.600 (5 anos): **11,1 bilhões**

É exatamente o problema descrito no item 14 do pedido: *"se as fórmulas
recalculam todo o histórico a cada inserção"*. Aqui é pior do que "todo o
histórico" — é **a soma de todos os prefixos do histórico**.

Duas atenuações importantes, e ambas são reais:

1. O guarda `IF(OR($E4="";…))` faz curto-circuito. Linhas vazias custam quase
   nada — por isso as 15.000 linhas provisionadas com só 1.110 preenchidas
   recalculam em 2,9 s.
2. Acrescentar uma corrida **no fim** suja apenas as linhas novas: as faixas
   expansivas das linhas anteriores não incluem as posteriores. Por isso a linha
   J da tabela D.1 é plana. **Editar uma linha no meio do histórico**, ao
   contrário, suja tudo o que vem depois — e aí o custo aparece.

### G.2 O segundo gargalo: a aba `Calc`

180 linhas × ~50 fórmulas, várias delas operações matriciais sobre os intervalos
nomeados de 15.000 linhas:

```
B3: =SE(selAnalito="";"";SEERRO(AGREGAR(15;6;rRUN/((rAnalito=selAnalito)*(rFirst=1)*((""&rLote)=(""&loteAtivo)));$A3);""))
```

Cada uma das 180 linhas avalia a expressão sobre a faixa inteira. Custo
O(180 × N × k) — **linear em N, mas com constante alta**. É esse termo que
responde pelos ~3 s de piso do recálculo mesmo com a planilha praticamente
vazia. Aos 60 meses ele sozinho fica ~4,4× maior.

### G.3 Ordem dos gargalos por impacto provável

| # | Componente | Quando incomoda | Gravidade |
|---|---|---|---|
| 1 | **Integridade: os 3 tetos silenciosos** | 6 a 13 meses | **crítica — perde dado sem avisar** |
| 2 | Fórmulas `BB`/`BC` (n²) — recálculo completo | ~30 meses | alta |
| 3 | Fórmulas matriciais do `Calc` (O(180×N)) | ~36 meses | média |
| 4 | Memória (2 GB aos 60 meses) | ~48 meses | média |
| 5 | Tempo de abertura (10 s aos 60 meses) | ~54 meses | baixa |
| 6 | Tempo de salvamento (2,5 s) | não incomoda | desprezível |
| 7 | Dependências entre abas | não incomoda | desprezível |
| 8 | Formatação condicional (13 regras) | não incomoda | desprezível |
| 9 | Gráficos (180 pontos fixos) | nunca | nulo |
| 10 | Tabelas dinâmicas | não existem | nulo |
| 11 | Power Query | não existe ainda | ver G.6 |
| 12 | Tamanho do arquivo | ~38 anos | nulo |

Vale registrar o que **não** é gargalo, porque foi bem resolvido:

- `UpsertResultados` e `AtualizarViewResultados` carregam o banco em array e
  escrevem em bloco: **0,24 s com 66.600 linhas**.
- O motor `EstatPeriodo` varre o banco **uma vez** e memoiza por carimbo
  (`Agregar` + `mCarimbo`). Primeira chamada 0,26 s aos 36 meses; as seguintes,
  0,00 s.
- Os gráficos leem uma janela fixa.

### G.4 Zonas operacionais (derivadas dos testes)

| Horizonte | Registros | Status | Justificativa medida |
|---|---:|:---:|---|
| 6 meses | 6.660 | 🟢 | recálculo 3,4 s; abertura 0,9 s; **mas a view `Resultados` começa a truncar se o lote não mudar** |
| 12 meses | 13.320 | 🟢 | recálculo 5,7 s; tudo abaixo de 1 s no uso diário |
| 13,5 meses | 15.000 | 🔴 | **teto do banco: registros novos deixam de ser analisados, em silêncio** |
| 24 meses | 26.640 | 🟡 | recálculo 28 s; RAM 585 MB; uso diário ainda instantâneo |
| 36 meses | 39.960 | 🟠 | recálculo 72 s; RAM 925 MB; abertura 4,0 s |
| 48 meses | 53.280 | 🟠 | recálculo 121 s; RAM 1,4 GB; abertura 4,2 s |
| 60 meses | 66.600 | 🔴 | recálculo 221 s; RAM 2,0 GB; abertura 6,5 s |

A linha vermelha dos 13,5 meses é de natureza diferente das outras: **não é
performance, é perda de dado.** Ela vem antes de qualquer amarelo de velocidade.

### G.5 Ponto de saturação operacional

- **Saturação de integridade:** mês **13,5** (banco) — CALCULADO, e mês **5,4**
  de um mesmo lote (view) — CALCULADO. É o que manda.
- **Saturação de performance:** entre os meses **30 e 36**, quando o recálculo
  completo cruza a faixa dos 30–60 s — ESTIMADO por interpolação entre pontos
  medidos de 24 e 36 meses.

### G.6 Risco futuro: Power Query

Não existe hoje. Quando entrar, ele muda o quadro em dois pontos:

- Uma consulta que leia o `DB_Resultados` inteiro a cada atualização é O(N) e
  não incomoda; uma que faça *merge* de tabelas grandes sem *folding* é O(N²) e
  vira o novo gargalo número 1.
- O `queryTable` materializado no arquivo **duplica** o dado em disco. A
  inclinação de 87 KB/mês pode dobrar.

Recomendação, quando o momento chegar: consulta de **agregação** (resumo por
analito/mês), nunca de detalhe linha a linha, e carga para o Modelo de Dados
em vez de para célula.

---

## H. RECOMENDAÇÃO ARQUITETURAL

### H.1 Comparação das cinco opções

| Critério | A. Um arquivo, 5 anos | B. Um por ano | C. Operacional + histórico | D. Base externa + Excel de análise | E. Manter dado, reduzir fórmula |
|---|---|---|---|---|---|
| Performance | ruim aos 4–5 anos | ótima sempre | ótima | ótima | **ótima** |
| Complexidade de implantação | nula | baixa | média | alta | **baixa/média** |
| Manutenção | 1 arquivo | 5 arquivos a versionar | 2 arquivos | 2 sistemas | 1 arquivo |
| Risco | tetos silenciosos | divergência de versão entre anos | idem, menor | dependência de TI | **baixo, se validado pela suíte** |
| Auditoria (ISO 15189 §8.4) | trilha contínua | **trilha quebrada na virada** | contínua no histórico | ótima | **contínua** |
| Integridade | tetos truncam calado | boa | boa | ótima | **boa** |
| Consulta plurianual | nativa | **exige abrir vários** | boa | ótima | nativa |
| Backup | 1 arquivo pequeno | simples | simples | precisa de rotina | 1 arquivo |
| Escalabilidade | ~13 meses | indefinida | indefinida | indefinida | **~7 anos** |

### H.2 O que eu recomendo

**Opção E como caminho principal, com a Opção C como plano de longo prazo — e
não a Opção B.**

O argumento contra a B, que costuma ser a escolha instintiva: um arquivo por ano
**quebra a trilha de auditoria na virada do ano** e multiplica por cinco a
superfície de manutenção de um produto que vai ser vendido e vai receber
correções. Um bug corrigido em 2028 teria de ser propagado para quatro arquivos
já em uso em cada cliente. Para um sistema sob ISO 15189 §8.4, com histórico de
lote atravessando o ano-calendário, isso é caro e frágil.

**O trabalho concreto, em ordem de prioridade** — seguindo a regra do projeto,
Integridade > Auditoria > Arquitetura > Estabilidade > Testes > Performance:

1. **INTEGRIDADE (urgente, vale para o uso já em curso).** Fazer os três tetos
   **falarem**. Enquanto não forem levantados, pelo menos não podem truncar em
   silêncio: `AtualizarViewResultados` deve avisar quando `cont(t)` bater em
   `VIEW_ROWS`; a gravação deve recusar-se a passar da linha 15.003; o `Calc`
   deve sinalizar quando o lote passar de 180 corridas. Isto é pequeno e é o
   que separa "lento" de "errado".

2. **ARQUITETURA — a janela do `Calc`.** Trocar `AGGREGATE(15;…)` (k-ésimo
   *menor*) por `AGGREGATE(14;…)` (k-ésimo *maior*) com a ordem invertida, para
   que a janela de 180 pontos acompanhe as corridas **recentes** e não as mais
   antigas. Hoje, um lote com mais de 180 corridas para de mostrar o que está
   acontecendo agora — que é justamente para o que serve um Levey-Jennings.

3. **PERFORMANCE — eliminar o termo n².** Substituir `BB` e `BC` de fórmula por
   **valor gravado pelo `UpsertResultados`**: ele já percorre os registros e já
   mantém dicionários de chave; decidir ali se o registro é o primeiro do par
   (analito, RUN) custa O(1) por linha. O banco deixa de ter 200.000 fórmulas e
   passa a ter zero. Isso sozinho tira o recálculo de 406 s para a casa dos
   segundos, e reduz o arquivo em ~72 KB/mês.
   *Observação de método:* isso é troca de estrutura de dados, e a regra do
   projeto exige causa comprovada. A causa está comprovada acima e é
   reproduzível pelos scripts em `scripts_fase3/` desta análise.

4. **PROVISIONAMENTO.** Com `BB`/`BC` fora do caminho, estender o banco de
   15.000 para ~80.000 linhas passa a custar quase nada em bytes e nada em
   tempo. `VIEW_ROWS` de 3.000 para 12.000, e `Liberação` de 200 para 1.000.

5. **LONGO PRAZO (opção C), quando passar de ~7 anos ou se o cenário PESADO se
   confirmar:** um `QC_Bioquimica.xlsm` operacional com os últimos 24 meses e um
   `QC_Bioquimica_Historico.xlsm` só de leitura, alimentado por rotina de
   arquivamento que **preserva a trilha** (o registro migra com o log, não sem
   ele). Nunca corte por ano-calendário; corte por idade do registro.

**Não recomendo a Opção D** (base externa) para este produto. Ele é vendido para
laboratórios que precisam de um arquivo que funcione sem TI. A dependência
externa destruiria a proposta de valor, e os números acima mostram que ela não é
necessária dentro do horizonte de 5 a 7 anos.

---

## Apêndice — método e limitações

**Testes físicos realizados:** sim, todos os sete pedidos (janeiro, 3, 6, 12, 24,
36, 48 e 60 meses), sobre cópias, com dado gerado por replicação fiel do bloco
real de janeiro — mesmos 31 analitos, 2 níveis, 18 corridas, datas deslocadas de
um mês, RUNs em sequência.

**Onde o teste teve de intervir na estrutura:** acima de 15.003 linhas as
fórmulas `BA:BC` e os intervalos `r*` precisaram ser estendidos, senão o Excel
não teria o que calcular e o tempo medido seria falso — rápido por omissão.

**Dois erros de bancada cometidos e corrigidos durante esta análise**, registrados
porque afetam a leitura dos números:

1. A primeira extensão usou `Copy` + `PasteSpecial`, que replicou o bloco de
   15.000 linhas **uma única vez** e parou na linha 30.003, sem erro. Os pontos
   de 36 e 60 meses mediram ~26 s — rápido porque dois terços das linhas não
   tinham fórmula. Refeito com `FormulaR1C1` sobre a faixa inteira, com
   conferência da última linha.
2. Os intervalos nomeados foram reapontados com `.RefersTo` recebendo texto A1;
   o Excel gravou `DB_Resultados!L4C5:L26643C5` e os nomes deixaram de resolver.
   O gráfico foi a zero pontos e **eu quase reportei "o Painel morre aos 24
   meses" como defeito do produto**. Era defeito do instrumento: reescritos por
   `RefersToR1C1`, os 180 pontos voltaram. As medições finais foram refeitas
   com os nomes conferidos.

**Limitações declaradas:**

- Só existe **um mês** de dado real. Não há como estimar sazonalidade; as
  projeções assumem estabilidade mensal e dizem isso onde importa.
- Os tempos são de **uma máquina** (i7-1255U, 16 GB). Uma máquina de
  laboratório mais modesta — 8 GB, i5 de duas gerações atrás — deve piorar os
  números da tabela D.1 por um fator de 1,5 a 3, e o ponto de 60 meses (2 GB de
  RAM) pode virar paginação em disco, que degrada de forma abrupta e não
  proporcional.
- Power Query não foi medido porque **não existe** no arquivo.
- A projeção de "5 anos viáveis após as correções" é ESTIMADA por remoção
  analítica do termo n². Só uma implementação e uma nova rodada de medição
  confirmam.

**Scripts desta análise** (no scratchpad da sessão, reproduzíveis):
`inventario.py`, `db_detalhe.py`, `nomes_e_dados.py`, `teste_escala.py`,
`teste_operacoes2.py`, `diag_painel.py`, `medir_final.py`.
