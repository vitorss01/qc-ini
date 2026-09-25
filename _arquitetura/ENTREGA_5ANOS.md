# Entrega — QC Bioquímica e QC Hematologia prontos para 5 anos e vários lotes

**Data:** 19–20/09/2026 · **Decisões:** ADR-049 (multi-lote), ADR-050 (desempenho),
ADR-051 (experiência de aplicativo), ADR-052 (Hematologia) e ADR-053 (filtro por
trimestre e nenhuma tela some debaixo do usuário) e ADR-054 (o Painel dos dois
produtos no mesmo desenho) · **Fontes, instaladores e testes:**
`_arquitetura/etapa1_multilote/`

## 1. O que o usuário faz agora

| Tarefa | Onde | Como |
|---|---|---|
| Chegou lote novo | Início → **Novo lote de controle** (ou Lotes → **+ Novo lote**) | Pede o código e a validade, pergunta se ele passa a ser o lote dos lançamentos e abre a digitação de média/DP. O lote novo nasce **vazio**, sem herdar nada. |
| Média e DP do lote | aba **Analitos** (Média/DP) | A tela mostra o lote em análise ("PARÂMETROS DO LOTE X"). A gravação é imediata e fica no Audit_Log. O atalho **Trocar lote ▸** leva ao seletor. |
| Ver outro lote | **Painel**, lista de lote (**H3** nos dois produtos) | Média, DP, limites, gráfico, Westgard e Estatística passam a usar **só** esse lote. Lote antigo continua consultável. |
| Trocar de analito | **Painel**: spinner ▲▼ **ou** a lista com o nome (C3) | Menos de 0,2 s por troca. |
| Ver um período | **Painel**: `Ano` (K3) + `Período` (K4), ou as datas De/Até (G3/G4) | Tudo (sem filtro) · Ano inteiro · T1..T4 · S1/S2 · Jan..Dez. |
| Lançar o dia | **+ Lançar** | Bioquímica: aba Importar. Hematologia: formulário da corrida. |
| Saber como está | **Início → Situação agora** | Lote em uso e validade (⛔ vencido / ⚠ vence em N dias), lote em análise, última corrida, ocupação do banco e analitos sem média/DP. |

A faixa de status do Painel muda de cor: **verde** quando o lote tem alvo, **âmbar**
quando não há corrida do lote no período e **vermelho** quando o lote está sem média/DP.

## 2. Tempos medidos com cinco anos de dados

Premissas do gestor:
- **Bioquímica:** 4 lotes/ano = **20 lotes**. São 31 analitos × 2 níveis, 1 corrida por
  dia útil e uma 2ª corrida em 40% dos dias: **115.196 resultados**.
- **Hematologia:** 8 lotes/ano = **40 lotes**. São 28 analitos × 3 níveis e 1 corrida
  **todo dia**: **153.468 resultados**.

Teste: `testes/t_cinco_anos.py`. O histórico foi montado pelo caminho do usuário
(parâmetros digitados na aba Analitos, lote a lote) e o arquivo foi reaberto antes de
medir.

| Operação | Bioquímica antes* | **Bioquímica agora** | **Hematologia agora** |
|---|---|---|---|
| Clique no spinner (troca de analito) | 1,9 s (26 s antes da correção do spinner) | **0,06–0,15 s** | **0,06–0,13 s** |
| Trocar o lote em análise | 3,7–4,8 s | **0,7–1,1 s** | **0,2–0,3 s** |
| Trocar o lote em uso | ~15 s | **1,4 s** | **1,6 s** |
| Lançar a corrida do dia | 25–33 s | **4,0 s** (62 resultados) | **5,2 s** (84 resultados) |
| Abrir o arquivo | 14 s | **2,7 s** | **2,6 s** |
| Recálculo completo | 15,6 s | **0,7 s** | **0,1 s** |
| Tamanho do arquivo | 11,3 MB | 12,3 MB | 11,2 MB |
| Capacidade do banco | 120.000 | **150.000** (77% usados em 5 anos no cenário pesado) | **200.000** (77% usados) |

\* A coluna "antes" é a Bioquímica com a Etapa 1 e cinco anos de banco. A Hematologia
não tinha "antes" comparável: o Painel dela estava congelado (§4).

Com os dados **reais** (10.342 resultados, lotes 8973/8974), os tempos são:
- **Clique no spinner:** 0,06 s.
- **Troca de lote:** 0,4–0,9 s.
- **Importar uma corrida:** 0,95 s.

## 3. Provas

| Prova | Resultado |
|---|---|
| Suíte T0–T9 da Etapa 1 (multi-lote), com o arquivo instalado final | **760/760** |
| Cinco anos, Bioquímica (20 lotes) | **34/34** |
| Cinco anos, Hematologia (40 lotes) | **34/34** |
| Navegação e experiência (`prova_ux.py`), nos dois produtos | ver §6 |
| Equivalência do desempenho no arquivo real (`prova_v2.py`) | idêntico |

A prova de equivalência fotografa, antes e depois, o **Calc inteiro, o Painel, a
Eng_Saida, os Eventos_Westgard (8.057 eventos), a Estatística e a Liberação**, em 6
combinações de lote × analito × período. Todos os números ficaram iguais. A única
diferença é o carimbo de hora do motor.

## 4. Achado grave na Hematologia (corrigido)

O arquivo de produção da Hematologia rodava o **motor antigo**, que gravava direto nas
telas. Os efeitos:
- As fórmulas do **Calc (12.614)**, do **Painel (45)** e da **Estatística (1.320)**
  tinham virado valores fixos.
- **O gráfico e os indicadores estavam congelados desde agosto.** Nenhum resultado novo
  apareceria.
- A seção de Sigma do Painel apontava para linhas que não existiam.
- A estrutura da pasta estava protegida, e isso impedia o login de mostrar e ocultar telas.

Tudo foi reconstruído a partir da última versão íntegra, já no motor único dos dois
produtos (ADR-052).

## 4b. Correção do filtro por trimestre (relato de uso) — ADR-053

Você relatou: *"cliquei na caixinha do trimestre, deu erro e fui parar na aba de eventos
de Westgard"*. O clique simulado não reproduziu o erro, mas a investigação achou o que o
tornava possível — e corrigiu:

- **O filtro chamava o motor inteiro**, inclusive os eventos de Westgard, que são do lote
  e não mudam com o período. Agora ele faz só a janela de corridas e os indicadores.
- **A tela é guardada e devolvida**, inclusive no erro: nenhuma operação deixa você em
  outra aba (que, com as guias ocultas, é onde o usuário se perde).
- **Falha vira mensagem** com o número e a descrição do erro.
- **Estado preso é zerado** no início de toda operação de clique: se uma operação anterior
  morreu no meio (por exemplo, "Finalizar" numa caixa de erro do VBA), o Excel ficava em
  cálculo manual com a tela congelada, e o clique seguinte parecia dar erro.

Uma hipótese foi descartada por medição: com os dois arquivos abertos, o clique **não**
dispara a macro do outro — o Excel amarra o controle à pasta que o contém.

Na mesma passada, quatro defeitos que a revisão adversarial levantou:
a aba **Resultados** mostrava as 3.000 linhas mais **antigas** do lote e escondia as novas
(agora mostra as mais recentes); a barreira de **capacidade** rodava depois de gravar;
a **exclusão lógica** era o único caminho de gravação fora da casca de operação (escrevia
célula a célula); e a lista da **Liberação** não avisava quando o lote passava das 200
corridas.

## 4c. O Painel dos dois produtos no mesmo desenho — ADR-054

Você comparou as duas telas e apontou que a da Hematologia estava mais limpa. As duas
tinham divergido; agora o Painel dos dois é construído pelo **mesmo código**, e o que
muda é só o que precisa mudar (número de níveis, nomes das regras de Westgard).

| | Antes | Agora (os dois) |
|---|---|---|
| Lote em análise | Bio em E3, Hema em H3 | **H3** |
| Ano / Período | só na Bioquímica | **J3:K3 e J4:K4 nos dois** |
| Trimestre | 4 caixas de seleção | **saiu**: virou item da lista Período |
| Indicadores por nível | Bio tinha 5 colunas | **10 colunas**: + ETp %, Bias %, ET %, Sigma |
| Última violação | só na Hematologia | **S6:U6 nos dois**, em cima, como você pediu |
| Desempenho Sigma | coluna R (Bio) / J (Hema) | **coluna A**, abaixo dos gráficos |
| Faixa de status | O3:X4 (Bio) | **A39:U40**, abaixo do gráfico do nível 2 |
| Controle externo (EQA) | K3/L3, por baixo das caixas | **M3:N4**, sem sobreposição |
| Menu de navegação | encostado na direita, fora da tela | **começa em I1** |
| Largura das colunas | diferentes | **A..U iguais nos dois** |
| Gráfico | largura fixa | **acompanha a largura visível, em qualquer zoom** |

Duas observações sobre o que você pediu:

- **O Controle Externo ficou em M3:N4, não M3:N5.** A linha 5 é o título da tabela de
  Westgard (uma célula mesclada de L5 a U5); usar M5 quebraria esse título. O bloco está
  dentro da faixa que você indicou e não encosta mais em nada.
- **O filtro de trimestre deixou de ser caixa de seleção.** Ele existia em três lugares
  ao mesmo tempo (as caixas, a lista Ano/Período e uma cláusula escondida dentro da
  fórmula do Calc), e os três podiam se contradizer — com "2º Tri" marcado e "T1"
  escolhido na lista, o Painel ficava vazio sem dizer por quê. Agora a lista **Período**
  faz tudo o que as caixas faziam, mais os meses e os semestres, e "Tudo (sem filtro)"
  mostra os cinco anos de uma vez.

### Três defeitos que a mudança revelou

1. **Apagar os nomes das caixas quebrava o filtro do Calc em silêncio.** A fórmula que
   decide quais corridas entram no gráfico citava esses nomes; sem eles virava `#NOME?`,
   e `#NOME?` ali significa **zero corridas no gráfico, sem nenhum aviso na tela**. O
   instalador passou a conferir isso e a recusar salvar se o filtro não deixar passar
   nenhuma corrida.
2. **O erro que você viu** ("variável não definida"). A Hematologia não tinha o módulo de
   períodos da Bioquímica, e o VBA só compila uma rotina quando alguém a chama — por isso
   a instalação terminou sem acusar nada e o erro só apareceu no primeiro clique. A conta
   de trimestre virou um módulo próprio, instalado nos dois, e os instaladores agora
   **chamam** as rotinas novas antes de salvar, que é a única forma de forçar a
   compilação delas.
3. **Ano vazio fazia o filtro não responder.** Em VBA, célula vazia passa no teste de
   "é número"; o ano virava 0 e a rotina saía calada. Agora, sem ano, o sistema assume o
   ano da última corrida e escreve na tela qual assumiu.

**Provas:** `testes/prova_ux.py` — **61/61 nos dois produtos** · `testes/prova_painel.py` (desenho) — **33/33** na Bioquímica e **30/30** na Hematologia. Inclui o filtro de
trimestre (Bioquímica 110 → 52 corridas em T1; Hematologia 25 → 13), o gráfico
acompanhando o zoom (1617 pt em 70%, 868 pt em 130%) e a troca de analito depois de tudo
isso, sem erro.

## 4d. Legenda, linha 1 e paridade entre os dois — ADR-055

Três pedidos do gestor depois de usar as telas do ADR-054.

**A legenda cortava os rótulos** ("Viol…", "Rep…"). Eram **14 entradas numa caixa de
80 pt**: seis são as linhas de limite (−3s..+3s), que se leem pela posição no gráfico, e
"Repetição" aparecia três vezes (as três réplicas). Ficam seis, embaixo e em linha — e o
gráfico recupera 80 pt do lado direito, onde estão os pontos mais recentes.

**A barra de navegação cobria o título** em 61,5 pt na Bioquímica. A causa: a largura do
título era *estimada* por `len × tamanho × 0,62`, conta que erra mais de 15% em fonte
proporcional. Agora é medida. O título ficou Segoe UI 16 nos dois — era Calibri 22 na
Bioquímica e 18 na Hematologia, o mesmo título em dois tamanhos — e a barra começa na
coluna seguinte à que o título termina, pulando uma coluna: **K1 nos dois**.

**A auditoria de paridade achou um defeito meu.** O realce "esta regra é a recomendada
pelo Sigma" existia só na Bioquímica e lá estava **morto** desde o ADR-054: lia
`Painel!V10:V11`, células que o Painel novo esvaziou. Consequência visível: nenhuma regra
iluminada e as duas notas do bloco de desempenho em branco. Corrigido e portado — agora
existe nos dois.

Continuam diferentes: a importação em massa (aba `Importar` na Bioquímica × formulário
`frmMassa` na Hematologia — duas implementações do mesmo recurso), o período da
Estatística (só a Bioquímica configura) e a fonte do ETp (diferença científica
deliberada). As três estão na lista de decisões do gestor, §5.

**Design:** um primeiro passo (gráfico sem moldura, cabeçalho fixo) foi escrito e
**retirado**. Formatar elementos de gráfico faz o Excel abrir um modal — *"a formatação
complexa aplicada ao gráfico… deseja continuar?"* — que `DisplayAlerts = False` não
suprime, com 14 séries × 180 pontos. E o gestor pediu que a parte visual venha depois de
um commit, para poder reverter. Fica para um commit próprio.

**Provas:** `testes/prova_painel.py` — **51/51 nos dois** (eram 33 e 30).

## 5. Riscos que permanecem

1. **Versionamento de parâmetro.** Mudar a média/DP de um lote reinterpreta todo o
   histórico **daquele** lote. A mudança fica registrada no Audit_Log, mas os resultados
   antigos não guardam o parâmetro da época. Isso é decisão do gestor (ADR-049).
2. **Volume maior que o previsto.** Duas corridas por dia na Hematologia dariam ~306 mil
   linhas em 5 anos, acima do teto de 200 mil. O sistema **recusa com mensagem**, sem
   perder dado em silêncio. O caminho seria arquivar por ano.
3. **Power Query.** Os dados continuam entrando pelo caminho do motor
   (`UpsertResultados`). Um alimentador futuro deve gravar numa área de staging e passar
   pelo mesmo caminho (ADR-017), nunca direto no `DB_Resultados`.
4. **Camada BI** fora do escopo, por decisão do gestor. Com 100 mil linhas, o
   `AtualizarBIData` não termina em 20 min. O botão existe, mas não é chamado por nenhuma
   operação do dia.
5. **Reimportar um resultado excluído o reativa em silêncio.** `UpsertResultados` marca
   `Ativo` ao atualizar uma chave que existia, mesmo que ela estivesse `Excluído`, e sem
   registro nominal. Precisa da sua decisão: bloquear o reenvio (como a linhagem de build
   fazia) ou manter e registrar.
6. **As assinaturas da Liberação estão presas à posição da linha**, não ao RUN: excluir
   uma corrida desloca as assinaturas das seguintes. Vale chavear por RUN.
7. **Imunologia** não foi tocada.
8. **Lote em uso vencido.** A tela Início já mostra: a Bioquímica está no lote 8974,
   com validade em 30/06/2026.

## 6. Instalação

`python instalar_producao.py` faz três coisas:
1. Copia os dois arquivos para `_backup_pre_5anos_<data>/` e confere a cópia byte a byte.
2. Roda os **mesmos** instaladores validados:
   - `aplicar_etapa1.py`, para a Bioquímica;
   - `portar_hema.py`, com a referência `ref/QC_Hematologia_build_h1.xlsm`.
3. Devolve os arquivos originais do backup se algo falhar.

A validação completa é `bash testes/validar_tudo.sh`.
