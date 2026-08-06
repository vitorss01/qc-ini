# QUALITY GATE — Sistema QC (CQI laboratorial)

Documento oficial de liberação do projeto. Enquanto houver item ⏳ ou ❌ **não se avança
para a fase seguinte**. Cada ✅ exige evidência verificável — marcação sem evidência
não vale.

**Última atualização:** 05/08/2026
**Sprint corrente:** HARDENING (nenhuma funcionalidade nova)

---

## 1. Motor estatístico

| # | Item | Status | Evidência |
|---|---|---|---|
| 1.1 | Projeto compila | ✅ | Execução de procedimento força compilação; 3 erros de sintaxe (`aS`) corrigidos. **Este item é a causa raiz do 2.1** — ver nota abaixo |
| 1.2 | Nenhum travamento | ✅ | `RegistrarEventosWestgard` 0,098 s · `AtualizarEstatistica` 0,341 s |
| 1.3 | Média / DP / CV% | ✅ | Idêntico a referência Python em 6 casas: 14,000000 · 3,162278 · 22,587698 |
| 1.4 | Bias / ET / Sigma / Z | ✅ | 40,000000 · 21,500000 · 5,000000 · 1,000000 |
| 1.5 | Westgard 12s/13s/22s/R4s/41s/10x | ✅ | 9/9 asserções, incluindo níveis opostos → R4s e **não** 22s |
| 1.6 | Elegibilidade CLSI EP05/C24 | ✅ | 6/6 estados; desconhecido nunca entra no cálculo |
| 1.7 | Robustez (n=0, n=1, DP=0, alvo=0) | ✅ | Todos devolvem 0, sem erro |
| 1.8 | Sem vazamento de estado entre grupos | ✅ | `grupos_com_0_violacoes_reportando_regra = 0` |

> **Nota — por que as fórmulas foram destruídas (liga 1.1 a 2.1).**
> `aS` não é identificador válido: VBA é insensível a maiúsculas, então `aS` é o
> mesmo token que a palavra reservada `As`, e o módulo não compila. Como o VBA
> compila sob demanda, procedimento a procedimento, o defeito ficava invisível —
> `AtualizarCalc` rodava normalmente enquanto `AtualizarPainelEng`,
> `AtualizarEstatisticaAba` e `RegistrarEventosWestgard` morriam com "não foi
> possível executar a macro". **Na produção esses três nunca rodaram, e é por
> isso que as fórmulas do `Painel` e da `Estatística` sobreviveram lá.** Ao
> corrigir o `aS`, a Fase 3A deu vida aos três e eles passaram por cima das
> fórmulas. O gate registrava a correção em 1.1 e o estrago em 2.1 sem ligar um
> ao outro. `gerar_mEstatistica.ps1` agora renomeia `aS`/`aM` para
> `alvoS`/`alvoM` e tem lint que barra identificador colidindo com reservada.

## 2. Integridade dos dados

| # | Item | Status | O que fecha |
|---|---|---|---|
| 2.1 | Fórmulas preservadas | ✅ | Marcos 1–5. Motor escreve só em `Eng_Saida`. Diff célula a célula contra a produção: `AUSENTE 0 · ALTERADA 4362 · VALOR 0 · EXTRA 9`, 61.425 fórmulas. As 4362 são lista fechada e versionada (`src_hardening1/marco*_celulas_redirecionadas.csv`): 3240 Calc + 42 Painel + 1080 Estatística. As 9 EXTRA são `Painel!S:U`, onde o motor gravava valor direto |
| 2.1b | ADR-019 verificado | ✅ | `varredura_adr019.ps1`: produção tinha **120** fórmulas de interface calculando parâmetro estatístico sobre o banco (`Estatística!D`, `AVERAGEIFS`); build tem **0**. As 900 restantes no `Calc` são seleção/filtro, permitidas. Estatística agregada que sobra na interface é só a cadeia do bias EQC (840, mantida por decisão) e 8 células de limite de eixo |
| 2.2 | Sem sobrescrita silenciosa | ✅ | Suíte 2.2: altera um resultado real pela camada de dados e exige que `AU_RESANT` contenha o valor original **numérico e correto** — `OK\|92,0028`. Fechou junto com a correção de um defeito grave: `Rotulo()` devolvia String, e `"92,0028"` gravado em célula Geral era relido pelo Excel como **920028** (vírgula tratada como separador de milhar). O `delta` ficava certo, o que escondia o erro, e o hash da cadeia certificava o registro corrompido |
| 2.3 | Exclusão lógica não é revertida sozinha | ✅ | `UpsertResultados` força `Status = Ativo`, ressuscitando excluído. Reverter exclusão deve ser ação própria, com justificativa |
| 2.4 | RUN representa a corrida analítica | ✅ | Aba `Corridas` + contador persistido `proxRUN` com auto-reparo só para cima. Testado: corrida existente não consome número; datas distintas geram RUNs distintos; turno separa duas corridas no mesmo dia; contador rebaixado à mão volta acima do maior usado. Migração aditiva — nenhum RUN existente mudou de valor |
| 2.5 | `Cfg_Status` não altera histórico retroativamente | ✅ | Uma célula redefine o que entra na estatística de todo o histórico. Exige versionamento da tabela + log |


> **Por que 2.1 e 2.1b voltaram para ⏳ (03/08/2026).**
> A evidência que os fechou era verdadeira sobre as *planilhas* e falsa sobre o
> *sistema*: as 4.362 células foram mesmo redirecionadas para `Eng_Saida`, mas o
> **motor corrigido nunca entrou no `.xlsm`** — nem aqui, nem na máquina do
> trabalho. Verificado lendo o VBA de dentro dos três artefatos: `mEstatistica`
> com 1.028 linhas, 8 ocorrências de `aS` e zero referência a `Eng_Saida`,
> quando o módulo versionado tem 1.049 linhas e nenhuma `aS`. As fórmulas
> apontavam para uma camada que ninguém abastecia.
>
> Quatro defeitos no pipeline, nenhum em VBA nem em lógica de CQI:
> 1. `Select-Object -First` **encerra o script produtor** (`StopUpstreamCommandsException`).
>    As etapas morriam depois das primeiras mensagens e **antes** do `Save()`.
>    `criar_eng_saida` usava `-Last`, que bufferiza — por isso só ela persistia.
> 2. `$xl.Quit()` não mata o processo enquanto o PowerShell segura referências COM.
>    O Excel órfão travava o arquivo e a etapa seguinte o abria em **somente leitura**,
>    sem aviso, porque `DisplayAlerts = $false`.
> 3. Nenhum dos 12 scripts tinha `$ErrorActionPreference = 'Stop'`: erro de COM é
>    não-terminante, então falha virava mensagem de sucesso.
> 4. `aplicar_vba.ps1` imprimia "VBA aplicado" sem conferir nada.
>
> **Lição estrutural:** havia lint de VBA, diff célula a célula e varredura de
> arquitetura — e nenhum controle verificava se a etapa anterior tinha acontecido.
> O pipeline reportava intenção, não resultado. É o mesmo princípio que o gestor
> exigiu para as fórmulas: não basta restaurar, precisa comparar.
>
> Reabrir estes dois itens é o certo. Gate que aceita evidência sem lastro dá
> confiança falsa, que é pior do que não ter gate.

## 3. Auditoria e segurança

| # | Item | Status | O que fecha |
|---|---|---|---|
| 3.1 | Trilha de auditoria | ✅ | `RegistrarLog` é stub vazio. Aba `Audit_Log` append-only: data/hora, usuário, papel, ação, chave, valor ANTES, valor DEPOIS, justificativa |
| 3.2 | Log em todo caminho de gravação | ✅ | Chamar de **dentro do `mDados`**, não dos formulários — senão algum caminho escapa |
| 3.3 | Proteção persistida no arquivo salvo | ✅ | Nenhuma das 18 abas tem `<sheetProtection>`. `UserInterfaceOnly:=True` **não persiste** entre sessões |
| 3.4 | Projeto VBA com senha | ⏳ | DPB decodifica para payload vazio — sem senha |
| 3.5 | Arquivo distribuído em estado bloqueado | ✅ | Salvo em sessão autenticada: abrir com macros desabilitadas dá acesso total a `DB_Resultados` |


> **Reabertos e agora fechados com evidência que se sustenta (03/08/2026).**
> Os itens 2.1 e 2.1b tinham sido rebaixados porque o motor validado nunca
> estivera dentro do `.xlsm`. Agora fecham porque `verificar_tudo.ps1` prova, na
> **mesma execução**, que (a) o VBA dentro do arquivo é idêntico byte a byte à
> fonte versionada — `mEstatistica`, `mDados` e `mAuditoria`, por hash — e (b) o
> diff das fórmulas dá `AUSENTE 0 · ALTERADA 4362 · VALOR 0 · EXTRA 9`. Sem (a),
> o (b) não significava nada.
>
> **3.1 e 3.2** fecham com o teste que dá sentido à trilha: a suíte grava no log,
> **adultera uma linha fora do sistema** e exige que a verificação acuse a linha
> certa — `QUEBRADO|5|conteudo alterado` — e que a cadeia volte a fechar quando o
> valor é restaurado. "Append-only" deixou de ser promessa e virou propriedade
> demonstrável diante de um auditor.
>
> **2.3 fecha com teste, não com leitura de código.** Os itens 3.6 e 3.7 da suíte
> percorrem o vetor de fraude inteiro: marcam um registro real como `Excluído`,
> **reenviam a mesma chave pela camada de dados** — como um formulário faria — e
> exigem que o status continue `Excluído` e que a tentativa apareça no log como
> `UPSERT_BLOQUEADO`. Depois restauram o estado original.
>
> **Como reproduzir tudo isto:** `.\_arquitetura\scripts_fase3erificar_tudo.ps1`
> — um comando, 21 verificações, relatório em `VERIFICACAO.md` e código de saída
> diferente de zero se algo falhar. Roda sem agente.


> **3.3 e 3.5 fecham; 3.4 continua ⏳ de propósito (03/08/2026).**
> A proteção era aplicada só em tempo de execução: `Workbook_Open` chama
> `LockApp`, que usa `UserInterfaceOnly:=True`. Duas falhas somadas —
> **`Workbook_Open` não roda com macros desabilitadas**, e `UserInterfaceOnly`
> é atributo de sessão, não persiste. Como o arquivo era salvo em sessão de ADM
> (que roda `UnprotectAll`), ele saía **distribuído destravado**.
>
> Agora `blindar_artefato.ps1` produz o entregável já travado: 21 abas
> protegidas com senha **gravada no arquivo**, 20 em `veryHidden` (só `Login`
> visível) e estrutura da pasta protegida. Com macros habilitadas nada muda para
> o usuário — o login revela o que o papel permite.
>
> **A verificação lê o `.xlsm` como zip e inspeciona o XML**, em vez de perguntar
> ao Excel. Se a pergunta é "o que o auditor vê com macros desligadas",
> perguntar ao Excel com macros ligadas não responde.
>
> **3.4 é passo manual e vai continuar sendo.** Pôr senha no projeto VBA exige
> escrever os campos `DPB`/`DPx` do `vbaProject.bin`, e errar corrompe o projeto.
> Faça no VBE: *Ferramentas → Propriedades do VBAProject → Proteção*. A suíte
> detecta e reporta enquanto não for feito.
>
> **Limite declarado:** proteção de planilha é barreira, não cofre — a senha está
> em texto no VBA e a marcação sai descompactando o arquivo. Ela detém o acidente
> e o curioso, não a fraude decidida. A garantia de integridade está na trilha
> encadeada por hash.


> **2.5 fecha (03/08/2026) — versionada e registrada, não bloqueada.**
> `Cfg_Status` decide o que entra em média, DP, CV, Bias, Sigma e Westgard.
> Trocar uma célula redefinia **retroativamente** o que compôs a estatística de
> todo o histórico, sem erro e sem rastro. Numa auditoria, *"com que critério
> este Sigma foi calculado em março?"* não tinha resposta.
>
> `mConfig.bas` acrescenta três coisas: **versão** em `Cfg_Status!F1` (inteiro que
> só sobe), **sombra** nas colunas ocultas H/I com o par Status/Elegível da
> versão anterior — sem ela não há como saber o valor antigo, porque
> `Worksheet_Change` dispara *depois* da alteração — e **log** de cada diferença
> no `Audit_Log`, distinguindo `CFG_ELEGIBILIDADE_ALTERADA`,
> `CFG_ESTADO_INCLUIDO` e `CFG_ESTADO_REMOVIDO`. Ao final chama `InvalidarCache`,
> senão a sessão corrente seguiria calculando com a regra antiga.
>
> **Não impede a mudança, de propósito.** A tabela existe para ser extensível sem
> tocar em código (ADR-006). O que se exige é registro e data. O usuário recebe
> aviso explícito de que a alteração é retroativa.
>
> Testes 3.8 e 3.9 da suíte: alteram a elegibilidade de verdade, exigem que a
> versão suba (1 → 2) e que a ação apareça na trilha; depois restauram.


> **Audit_Log v2 — de arquivo técnico a base de auditoria (03/08/2026).**
> A trilha passou de 21 para **33 colunas**, formatada como **tabela do Excel**
> (`tblAuditoria`). O objetivo mudou: o auditor abre a aba e responde tudo com os
> filtros nativos — só o WBC, só um RUN, só um lote, só o que fulano alterou, só
> entre duas datas, só entre duas horas, só exclusões — **sem nenhuma tela em VBA**
> e sem conhecer a estrutura interna. Power Query e Power BI consomem pelo nome.
>
> Acrescentados: `Data` e `Hora` em colunas próprias (o filtro sobre data-hora não
> resolve "entre 8h e 12h em qualquer dia"); `ResultadoAnterior`/`ResultadoNovo`/
> `Delta`/`Delta%` — o valor antigo saiu de dentro do texto do parecer e virou dado;
> `ChaveRegistro` e `SeqAlteracao`, que isolam toda a vida de um resultado num
> filtro só (o valor original é sempre `Seq = 1`); `Categoria` + `Ação` + `Módulo`;
> `Motivo` codificado ao lado do `Parecer` livre — motivo em lista fechada **conta**,
> texto livre não. `RESULTADO_APAGADO` virou ação própria, com `<VAZIO>` explícito.
> Mais a aba **`Audit_Legenda`**, que explica cada ação e cada coluna.
>
> **O arquivo distribuído deixa a trilha visível**, protegida e filtrável
> (`AllowFiltering`). Um auditor cauteloso abre com macros desligadas — se nesse
> estado ele não enxerga a trilha, a trilha não serve. A integridade não vem de
> esconder: vem da cadeia de hash.
>
> **Correção de princípio no encadeamento:** o hash passou a cobrir **o valor
> gravado**, não o valor pretendido. Calcular sobre o array em memória e verificar
> lendo as células compara coisas diferentes — o Excel converte na viagem (hora
> vira fração de dia, número vira texto, ponto flutuante perde bit). Isso produzia
> falso positivo, e um verificador que dá alarme falso é pior que nenhum: ninguém
> confia nele no dia em que importa.


> **Sprint NC entregue (04/08/2026) — 42 de 43 verificações.**
> `frmResultadoNaoConforme` e `frmExcluirRegistroNC`, com botões nas abas
> `Resultados` e `Registros`. O analista informa o RUN e o sistema preenche data,
> lote e os resultados da corrida com o valor de cada analito; ele marca os
> errados, escolhe o tipo (lido do `Cfg_Status`) e escreve o parecer técnico.
>
> **Zero fórmulas alteradas.** A especificação pedia trocar `Rep 1/2/3` por uma
> coluna única, o que quebraria 3.240 fórmulas do `Calc` e o marcador do gráfico.
> A aba `Registros` já tinha as colunas certas — `E` era a corrida e passa a
> receber o RUN; `F` era `Rep 1` e é exatamente a coluna que o gráfico plota.
> `G` e `H` ficam livres para as repetições posteriores.
>
> Testes 3.12 a 3.15: parecer curto **impede** marcar como não conforme mesmo
> chamando a camada de dados direto — a exigência é propriedade do sistema, não
> cortesia da interface; o resultado sai dos cálculos e o **valor original fica**;
> a ocorrência aparece em `Registros`; as três camadas registram com o mesmo
> `ID_Auditoria`.
>
> **Desempenho na máquina limpa:** motor completo 0,46 s · gravação em camada
> dupla 165 ms/evento · verificação da cadeia (59 eventos) 0,31 s.
>
> **Única falha da suíte: 5.4**, a senha do projeto VBA — passo manual do gestor,
> adiado por decisão dele para o fim do projeto.

## 4. Qualidade e testes

| # | Item | Status | O que fecha |
|---|---|---|---|
| 4.1 | Testes unitários do motor | ✅ | Suíte em `scripts_fase1/` e `scripts_fase2/` |
| 4.2 | QA integrado | ⏳ | **Nunca executou** — agente interrompido por limite de gastos |
| 4.3 | Regressão Fases 1 e 2 pós-Fase 3 | ⏳ | RUN, exclusão lógica, troca de lote, login, 3 UserForms, upsert sem duplicar |
| 4.4 | Casos extremos | ⏳ | Banco vazio · 1 resultado · 2 resultados · analito sem média/DP · lote sem resultado · status inválido · RUN duplicado · data futura |
| 4.5 | Teste de estresse | ✅ | `teste_estresse.ps1`, curva ate o teto de 15.000 linhas (375 corridas, 20 analitos x 2 niveis): recalculo 1,15s → 4,61s e motor 0,63s → 2,28s. Total 6,88s no teto. Expoente 1,1–1,46 — **superlinear leve, nao quadratico**. As 30.000 `COUNTIFS` de intervalo expansivo do banco NAO sao gargalo: a suspeita levantada na inspecao foi refutada pela medicao. Achado real: buffer de eventos de Westgard estourava a partir de ~5.200 linhas (14.317 eventos a 15.000). Corrigido com dimensionamento dinamico |
| 4.6 | Zero erro de fórmula | ✅ | Suíte 4.6: varre todas as abas por `#DIV/0!`, `#VALOR!`, `#REF!`, `#NOME?`, `#NÚM!` e `#NULO!`. Zero achados. `#N/A` fica fora de propósito — é a lacuna deliberada das séries do gráfico |
| 4.7 | Estado global restaurado | ✅ | Suíte 4.7: após `AtualizarEstatistica`, `ScreenUpdating = True` e `Calculation = xlCalculationAutomatic`. Rotina que sai deixando cálculo manual faz o painel mentir sem dar erro |
| 4.8 | Idempotência | ✅ | Suíte 4.8: hash de `Eng_Saida` (corridas + estatística + parâmetros) após 1 execução e após 11. Idênticos. Pega coleção que não zera, contador que soma e linha duplicada |
| 4.9 | Persistência entre sessões | ✅ | Suíte 4.9: salva, **fecha o Excel de verdade**, reabre noutro processo e recalcula. Saída idêntica. Expõe dependência de variável em memória, cache `Static` e `UserInterfaceOnly` que não persiste |

## 5. Cobertura por produto

| Produto | Fase 1 | Fase 2 | Fase 3 | Especificações | Observação |
|---|---|---|---|---|---|
| Hematologia | ✅ | ✅ | ✅ | ❌ | Motor instalado e verificado pela suíte. Sem `mEspecificacoes` — o subsistema é da Bioquímica |
| Bioquímica | ✅ | ✅ | ✅ | ✅ | Produto de referência. Aba `Importar` com 31 analitos, motor de especificações (ADR-022/023) |
| Imunologia | ✅ | ✅ | ❌ | ❌ | Fase 3 nunca aplicada. Adiado por decisão do gestor |

> Esta tabela dizia "Fase 3 nunca aplicada" para a Bioquímica até 06/08/2026,
> quando ela já passava a suíte inteira. A escrituração ficou para trás do
> trabalho — anotado aqui porque um gate que descreve errado o próprio estado
> é pior do que um gate sem tabela.

## 6. UX

| # | Item | Status | Observação |
|---|---|---|---|
| 6.1 | Fluxo diário validado com usuário | ⏳ | Simulação feita por agente; falta validação real |
| 6.2 | Padronização "Corrida" → "RUN" | ✅ | Bioquímica e Hematologia. O identificador não se chama mais "Seq": `DB_Resultados!A2`, `Calc!B2`, `RegistrosStore!D1`. **"Corrida" permanece onde designa o evento** — `Lançar corrida`, `Registro de corridas`, `DataCorrida` estão corretos. Prova 6.2 |
| 6.2a | `Audit_Legenda!B22` preservada | ✅ | Ali "Seq" é a sequência de **auditoria** (1ª, 2ª, 3ª alteração do resultado), outro conceito. A prova 6.2 falha nos dois sentidos: reprova se o Seq do identificador sobrar **e** se o Seq da auditoria for apagado por um replace global |
| 6.3 | Eixo X alinhado ao RUN | ✅ | 2 gráficos na Bioquímica, 3 na Hematologia: `Corrida (RUN)` → `RUN`. Prova 6.3 |
| 6.4 | 52 achados de UX/arquitetura triados | ⏳ | `FASE3A_achados_estaticos.md` |
| 6.5 | Importação por aba substitui o `frmMassa` | ✅ | Bioquímica. Analitos na horizontal, colagem direta, botão **Registrar** migra para `DB_Resultados` e limpa a aba. Provas 1.7–1.10 (montagem), 3.16–3.17 (execução) e 3.18 (rastro) |

## 7. Motor de Especificações (ADR-022 / ADR-023)

Entrou em 05–06/08/2026 com **575 linhas que decidem CONFORME / NÃO CONFORME**, e
`Painel` e `Estatística` passaram a ler dele. Chegou sem nenhuma prova na suíte —
era a camada mais nova, a de maior consequência e a única de cálculo sem cobertura.

| # | Item | Status | Observação |
|---|---|---|---|
| 7.1 | Meta resolvida pelo ano do resultado | ✅ | Regra de **vigência**: maior `Ano` ≤ ano do resultado. Prova cadastra 2020 e 2026 e exige que 2023 resolva pela de 2020, 2026 pela de 2026 e 2019 devolva `NAO_CADASTRADA` |
| 7.2 | Os três modelos derivam o que o ADR manda | ✅ | `ETP_DIRETO` (CVtp = ETp/3), `VB` (CVtp = CVi·fi, BIAStp = √(CVi²+CVg²)·fb), `CV_BIAS_DIRETO` (ETp = BIAStp + 1,65·CVtp) |
| 7.3 | Quatro estados, e ausência nunca é aprovação | ✅ | `CONFORME`, `NAO CONFORME`, `SEM ESPECIFICACAO`, `SEM DADOS`. A prova exige explicitamente que nenhum dos dois de ausência saia `CONFORME` |
| 7.4 | Especificação desativada sai de vigência | ✅ | **Regressão de defeito real.** A comparação era contra `"NAO"` cru e o gestor desativa digitando `"Não"`. `UCase$("Não")` = `"NÃO"` ≠ `"NAO"`, então a especificação recém-desativada seguia valendo e mudando veredito, sem aviso. A prova usa de propósito a grafia acentuada — com `"NAO"` cru ela passaria mesmo com o defeito de volta |
| 7.5 | Protocolo invariante de localidade | ✅ | `Str$`/`Val`, nunca `CStr`. Mesma família do defeito que gravava `92,0028` e o Excel relia `920028` (item 2.2) |
| 7.6 | `mEspecificacoes` tem instalador versionado | ✅ | `instalar_especificacoes_producao.ps1`. O módulo tinha entrado pela VBE — `criar_form_especificacoes.ps1` apenas **exigia** que ele já estivesse lá. Sem instalador, corrigir a fonte versionada não mudava nada em lugar nenhum. Coberto também pela prova 1.1 (idêntico à fonte) |

## 8. Defeitos do pipeline multi-produto, achados em 06/08/2026

Os três vieram da mesma causa: passos criados para a Bioquímica passaram a rodar
para **todos** os produtos com valores da Bioquímica embutidos. Nenhum deles
aparecia na Bioquímica — só na Hematologia, que ninguém tinha reexecutado.

| # | Defeito | Status | Como se manifestava |
|---|---|---|---|
| 8.1 | Analito de referência fixo em `Glicose` | ✅ corrigido | `build_all.ps1` passava `-Analito 'Glicose'` para todo produto. O build **inteiro** da Hematologia morria: *"analito de referencia 'Glicose' nao esta cadastrado"*. Agora vem de `$REFERENCIA_POR_PRODUTO` (Hematologia = `WBC`, verificado com 75 resultados no banco) |
| 8.2 | Aba `Importar` montada com os analitos da Bioquímica | ✅ corrigido | `criar_aba_importar.ps1` tem `analitos_bioquimica.csv` como padrão e o build nunca passava `-Csv`. A Hematologia recebia 31 siglas de outro setor. **A aba montava, ficava bonita e passava em todas as provas estruturais** — só a prova de ponta a ponta pegava, com `ERRO\|6`. Criado `analitos_hematologia.csv` (28 analitos, extraídos da própria aba `Analitos`) e o build passa o CSV do produto |
| 8.3 | A prova 1.7 confirmava o defeito em vez de pegá-lo | ✅ corrigido | Comparava a aba contra `analitos_bioquimica.csv` **para qualquer produto** — a mesma lista errada dos dois lados, então batia. E a 1.11 fazia `continue` no de/para vazio: relatou **"0 órfãs"** com a aba inteira desmapeada. Coluna sem de/para agora é defeito, não caso previsto |

> A lição das três é a mesma, e é a do ADR-021 noutra forma: **uma prova que lê a
> mesma fonte que o passo escreveu não é uma prova.** As de estrutura (1.7–1.12)
> conferiam a aba contra ela mesma; a única que discriminou foi a que exercita o
> comportamento (3.16).

### Pendência que não é minha de decidir

**Linha de base de fórmulas da Hematologia está velha.** Foi tirada em 03/08
(`9ad7e65`); a da Bioquímica foi regerada em 05/08 (`2b11f3b`) porque o pipeline
mudou. A suíte acusa `VALOR 4965` na Hematologia, e a leitura das divergências
mostra que **não são fórmulas destruídas**: são dados de teste que o próprio
build semeia (`DB_Resultados!BA1940` vazio → `999999`) e valores do `Calc` que
mudaram por fixar a referência em `WBC`.

Regerar a linha de base faria o item passar — e é exatamente assim que um gate
morre, redefinindo a verdade para caber no resultado. Fica registrado para
decisão sua, junto da retomada da Hematologia.

---

## Resumo

> **Regra de merge:** nada entra na `main` enquanto houver ⏳ ou ❌.
> A branch `fase3a-motor-cqi` é de trabalho; se um PR for aberto, deve ser **draft**.

**40 ✅ · 6 ⏳ · 2 ❌**  (contagem inclui a tabela de cobertura por produto)

**Suíte, 06/08/2026:** Bioquímica **68 de 69** · Hematologia **61 de 63**.
Na Bioquímica a única falha é a `5.4` (senha do VBA, passo manual seu). Na
Hematologia somam-se a `5.4` e a `4.2`, que é a linha de base velha descrita
na seção 8 — não é defeito de produto.

O **motor** está pronto e validado. O **sistema** ainda não é auditável nem confiável.

---

# Anexo — Redesenho proposto do RUN (item 2.4)

Modelo atual: `RUN = f(Data, Lote)` — determinístico, bom para deduplicação, mas
**incapaz de representar duas corridas no mesmo dia** (turno, pós-calibração,
pós-manutenção), que a CLSI C24 trata como eventos distintos.

Modelo proposto: **RUN = inteiro sequencial**, com os atributos da corrida em colunas
próprias:

```
RUN | Data | Hora | Turno | Lote | Equipamento | Usuário
```

Assim `RUN 501 (10/08 08:00)` e `RUN 502 (10/08 13:20)` coexistem naturalmente.

### Pré-requisito antes de escrever código
O custo aqui **não é o código, é a migração**. Os três produtos já têm dados gravados
com o RUN atual (Hematologia 1.575 linhas, Bioquímica e Imunologia 1.000 cada).
Trocar a semântica da chave exige:

1. Script de migração que preserve a correspondência RUN antigo → RUN novo
2. Backup obrigatório antes (a migração é irreversível)
3. Revalidar tudo que consome RUN: `Calc`, `Liberação`, `Eventos_Westgard`, os 3 UserForms
4. Decidir o que fazer com corridas históricas sem hora registrada
   (sugestão: hora nula = turno único, preservando o comportamento atual)

Planejar a migração **antes** de executar, não durante.

---

# Ordem de execução da Sprint HARDENING

1. **2.1 Fórmulas** — muda onde o motor escreve e afeta o teste de todo o resto. Primeiro.
2. **2.4 RUN** — mexe no modelo de dados. Depois das fórmulas, antes da auditoria.
3. **3.1–3.2 Auditoria** — precisa da chave final do RUN para logar corretamente.
4. **2.2–2.3 Versionamento e exclusão** — apoiam-se na trilha de auditoria.
5. **3.3–3.5 Proteção** — por último nas correções: muda o estado de distribuição.
6. **4.2–4.7 QA completo** — só faz sentido com tudo acima fechado.
7. **Propagar para Bioquímica e Imunologia** — o pipeline é paramétrico; propagar apenas
   depois de Hematologia passar em todo o gate.

---

# Roadmap até o Release Candidate

**Arquitetura congelada.** Nenhuma funcionalidade nova — nem dashboard, nem indicador,
nem UserForm, nem gráfico — até RC1. Só correção, teste e validação.

### Sprint HARDENING 1 — Fórmulas
- Criar `Eng_Saida`; motor deixa de escrever em `Painel` e `Estatística`
- Restaurar as 1.365 fórmulas destruídas (fonte: `_backup_pre_fase3/`)
- Painel e Estatística voltam a consumir por fórmula, a partir de `Eng_Saida` (ADR-019)
- **Validar a restauração, não confiar nela.** Diff automatizado célula a célula contra o
  backup: endereço, fórmula em R1C1 e resultado. Aceite = 100% idêntico; qualquer
  divergência é listada, não resumida em contagem
- **Varredura ADR-019:** nenhuma fórmula de interface referencia `DB_Resultados`. Esperado: zero
- **Aceite:** contagem **e conteúdo** idênticos · zero referência direta ao banco · regressão completa

> **Histórico não é reprocessado.** A unificação do Westgard e a correção de
> ET/Sigma mudam veredictos passados. Decisão do gestor: **não** reprocessar o
> histórico — o efeito cascata é maior que o benefício. A partir da RC1, as
> regras Westgard e o cálculo de Erro Total e Sigma passam a ser determinados
> **exclusivamente pelo motor**. Evidência do impacto medido fica em
> `src_hardening1/marco2_impacto_westgard.csv`.

### Sprint HARDENING 2 — Modelo do RUN
- RUN sequencial + `Data | Hora | Turno | Lote | Equipamento | Usuário` (ADR-011)
- Script de migração dos 3.575 registros existentes, com backup obrigatório
- **Aceite:** unicidade garantida · 3 UserForms validados · nada órfão

### Sprint HARDENING 3 — Auditoria (ISO 15189)
- `Audit_Log` append-only: usuário, timestamp, ação, chave, valor antes, valor depois, motivo
- Log chamado de dentro de `mDados` — nenhum caminho de gravação escapa
- **Aceite:** toda gravação e exclusão rastreável · log não editável pela interface

### Sprint HARDENING 4 — Validação
- Estresse · regressão · usabilidade · desempenho
- **Idempotência (4.8):** 1× · 10× · 100× execuções consecutivas → saída idêntica
- **Persistência (4.9):** gerar → fechar → reabrir → gerar → resultado idêntico
- **Teste com macros desabilitadas** (o cenário do auditor)
- Teste de recuperação após falha
- **Aceite:** Quality Gate 100% verde

### Release Candidate (RC1)
Somente quando **todos** forem verdadeiros:
- Quality Gate 100% verde
- Nenhum bug crítico aberto
- Regressão 100%
- Auditoria completa e funcional
- Desempenho validado sob carga
- **Hematologia, Bioquímica e Imunologia executando arquitetura idêntica**

A partir do RC1: implantação piloto, e toda mudança passa a ser evolução funcional —
não mais correção estrutural.
