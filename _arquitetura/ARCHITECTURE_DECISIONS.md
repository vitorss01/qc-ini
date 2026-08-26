# Architecture Decision Records — Sistema QC (CQI laboratorial)

Registro das decisões arquiteturais do projeto: o que foi decidido, **por quê**, e o que
se abriu mão em troca. Serve para não rediscutir o já decidido, para orientar manutenção
futura e para sustentar a defesa técnica em banca ou auditoria.

**Formato:** contexto → decisão → consequências → status.
**Status possíveis:** ✅ vigente · ⏳ decidida, não implementada · ↩️ revertida · 🔄 em revisão

---

## ADR-001 · `DB_Resultados` como única fonte operacional
**Status:** ✅ vigente

**Contexto.** Os cálculos liam de uma aba que também era usada para digitação e exibição.
Não havia separação entre dado bruto, cálculo e apresentação.

**Decisão.** A aba `Resultados` foi **renomeada** para `DB_Resultados` e passou a ser a
única fonte operacional. Uma aba nova `Resultados` foi criada como camada de
visualização somente leitura.

**Por que renomear e não recriar.** Ao renomear, o Excel reaponta automaticamente todos
os nomes definidos e fórmulas que referenciam a aba. Recriar exigiria reescrever ~45
referências à mão, com risco alto de erro silencioso.

**Consequências.** Nenhum formulário escreve em célula diretamente; toda gravação passa
pela camada `mDados`. Em troca, a view precisa ser reconstruída explicitamente após cada
alteração — não é automática como fórmula.

---

## ADR-002 · Armazenamento vertical, entrada horizontal
**Status:** ✅ vigente

**Contexto.** O analista pensa em "uma corrida com N analitos" (horizontal). O banco
precisa de uma linha por observação (vertical) para permitir filtro, agregação e auditoria.

**Decisão.** O banco é vertical e normalizado. Os formulários trabalham em formato
horizontal e a conversão acontece na camada de dados, não na interface.

**Consequências.** Consultas estatísticas ficam triviais e escaláveis. O custo é uma
etapa de transformação a cada gravação — resolvida com arrays em memória, não célula a
célula.

---

## ADR-003 · Exclusão lógica, nunca física
**Status:** ✅ vigente

**Contexto.** Apagar fisicamente um resultado destrói a rastreabilidade e é incompatível
com ISO 15189 §8.4.2, que exige que o valor original permaneça recuperável após emenda.

**Decisão.** Coluna `Status`. Nenhum registro é removido do banco.

**Consequências.** O banco cresce monotonicamente. Todos os consumidores precisam filtrar
por elegibilidade — o que levou ao ADR-006.

---

## ADR-004 · `RUN` como chave lógica da corrida
**Status:** 🔄 em revisão (ver ADR-011)

**Contexto.** O identificador anterior (`Seq`) era sequencial **por lote** — existia
`Seq = 1` em vários lotes. Não era chave.

**Decisão.** `RUN`, inteiro único global, atribuído por par (Data + lote de 6 dígitos).
Todos os níveis e analitos da mesma corrida compartilham o RUN.

**Consequências.** Deduplicação natural: relançar a mesma data reaproveita o RUN.
**Efeito colateral descoberto depois:** impede representar duas corridas no mesmo dia
(turno, pós-calibração), que a CLSI C24 trata como eventos distintos. Ver ADR-011.

---

## ADR-005 · Chave de unicidade = `RUN + Nível + Analito`
**Status:** ✅ vigente

**Contexto.** O requisito original dizia "mesmo RUN + Analito → atualizar".

**Decisão.** O **Nível entra na chave**.

**Por quê.** O RUN identifica a *corrida*, que abrange todos os níveis. Com RUN+Analito
apenas, gravar o Nível 2 sobrescreveria o resultado do Nível 1 do mesmo analito na mesma
corrida — perda silenciosa de dado. Com o Nível na chave, o comportamento pedido
(não duplicar, atualizar o existente) é preservado sem esse efeito.

---

## ADR-006 · Elegibilidade como tabela de configuração (CLSI EP05 / C24)
**Status:** ✅ vigente

**Contexto.** A norma exige que resultados inválidos, de calibração, manutenção, troca de
lote ou treinamento **não componham** média, DP, CV, Bias, Sigma nem Westgard. E que
novos estados possam surgir.

**Decisão.** Aba `Cfg_Status` com pares Status → Elegível. Uma função única,
`EhElegivel()`, é o ponto de verdade. Novos estados entram na tabela **sem tocar em código**.

**Falha segura.** Estado desconhecido devolve *não elegível* — na dúvida, o dado fica fora
do cálculo, nunca dentro.

**Consequências.** Evita `If Status = "Excluído"` espalhado. **Risco identificado depois:**
a tabela é editável e redefine retroativamente o histórico inteiro — pendente de
versionamento e log (item 2.5 do Quality Gate).

---

## ADR-007 · Westgard: regras clássicas apenas; 12s é alerta
**Status:** ✅ vigente

**Contexto.** A implementação anterior usava variantes por setor (6X/8X, 3-1s, 2of3-2s),
inconsistentes entre produtos.

**Decisão.** Conjunto clássico: 12s, 13s, 22s, R4s, 41s, 10x. **12s é alerta, não
rejeição** — cerca de 5% dos resultados excedem 2 DP por variação aleatória esperada.
Status passa a ter três estados: OK / ALERTA / REJEITADO.

**Classificação normativa.** Alerta: 12s. Erro aleatório: 13s, R4s. Erro sistemático:
22s, 41s, 10x.

**Consequências.** Comportamento uniforme entre os três produtos e alinhado à literatura
(Westgard et al., Clin Chem 1981;27:493-501).

---

## ADR-008 · Trend, Slope, Deriva e Shift fora do escopo do motor
**Status:** ↩️ revertida (implementada e depois removida)

**Contexto.** Foram implementadas como indicadores independentes, com regressão linear.

**Decisão.** **Removidas.** Pertencem a um módulo futuro de *Analytics*, não ao motor
operacional de CQI.

**Por quê.** O comportamento sistemático já é capturado por 22s, 41s e 10x. Manter dois
mecanismos concorrentes para detectar a mesma coisa aumenta a superfície de erro sem
ganho clínico. O esforço foi redirecionado para explicabilidade das regras existentes.

**Registro honesto.** Foram implementadas antes de a decisão ser revista — o código
existiu e foi retirado. Custou uma rodada de retrabalho.

---

## ADR-009 · Conhecimento das regras centralizado em `mWestgardKnowledge`
**Status:** ✅ vigente

**Contexto.** Interpretação, causas prováveis e sugestões de investigação tendem a ser
duplicadas em mensagens, relatórios e painéis.

**Decisão.** Um módulo único com `Enum eWestgard`, `Type WestgardRuleInfo` e funções de
consulta. **Nenhum outro módulo duplica esses textos.**

**Por que Enum e não string.** Comparação por string espalhada pelo código é fonte
clássica de bug silencioso (erro de digitação não falha, apenas não casa).

**Consequências.** Atualizar a interpretação de uma regra é edição em um lugar só.

---

## ADR-010 · Eventos separados de parâmetros
**Status:** ✅ vigente

**Contexto.** As violações de Westgard iam ser gravadas na aba `Estatística`.

**Decisão.** Aba própria `Eventos_Westgard` (Data · RUN · Analito · Nível · Regra ·
Classificação · Resultado · Z-Score). A `Estatística` permanece dedicada a **parâmetros**.

**Por quê.** São naturezas de dado diferentes: parâmetro é estado atual, evento é
histórico. Misturar impede auditoria, filtro, pivot e análise temporal.

**Ganho colateral de performance.** A separação permitiu gerar todos os eventos em **uma
passagem** pelo banco, em vez de 120 varreduras (40 analitos × 3 níveis).

---

## ADR-011 · Redesenho do `RUN` para representar a corrida analítica
**Status:** ⏳ decidida, não implementada

**Contexto.** `RUN = f(Data, Lote)` colapsa múltiplas corridas do mesmo dia. A CLSI C24
trata corrida analítica como unidade de avaliação — turno, pós-calibração e pós-manutenção
são eventos distintos.

**Decisão.** `RUN` passa a ser **inteiro sequencial**, com os atributos da corrida em
colunas próprias: `RUN | Data | Hora | Turno | Lote | Equipamento | Usuário`.

**Pré-requisito.** O custo não é o código — é migrar 3.575 linhas já gravadas nos três
produtos. Exige script de migração com correspondência preservada, backup obrigatório
(irreversível), revalidação de tudo que consome RUN, e decisão sobre corridas históricas
sem hora (sugestão: hora nula = turno único).

---

## ADR-012 · Motor escreve em área intermediária, não sobre fórmulas
**Status:** ⏳ decidida, não implementada

**Contexto.** O requisito era "mover os cálculos para VBA, mantendo no Excel apenas o
resultado". Implementado literalmente, o motor grava valores sobre células que continham
fórmulas — **1.365 fórmulas destruídas de forma irreversível** (Painel 58→13,
Estatística 2160→840).

**Decisão.** O motor passa a escrever em aba própria (`Eng_Saida`); `Painel` e
`Estatística` voltam a ter fórmulas referenciando-a.

**Por quê.** Com macros desabilitadas, valores literais não têm procedência. Um auditor
precisa rastrear o cálculo célula a célula. Valores sem fórmula não são defensáveis em
ISO 15189 nem em banca.

**Trade-off aceito.** Duplica a lógica de apresentação (o motor calcula, a fórmula exibe).
Em troca, a planilha permanece auditável sem depender de macro.

**Registro honesto.** A destruição de fórmulas era consequência do desenho que implementei,
e não foi explicitada ao gestor antes da execução.

---

## ADR-013 · Multi-lote por armazém oculto, não por duplicação de abas
**Status:** ✅ vigente

**Contexto.** Requisito de até 50 lotes com especificações, registros e liberação
independentes por lote.

**Decisão.** Abas ocultas (`LotesStore`, `LiberStore`, `RegistrosStore`) guardam um bloco
por lote; as abas visíveis são *views* do lote ativo, trocadas por macro.

**Alternativa rejeitada.** Duplicar fisicamente as abas por lote geraria ~300 abas por
arquivo — inviável em tamanho, desempenho e manutenção.

**Herança.** Ao criar um lote novo, as especificações analíticas são herdadas do anterior
(evita redigitar 40 analitos); assinaturas de liberação **nunca** são herdadas.

---

## ADR-014 · Não adotar orientação a objetos (`clsAnalito`)
**Status:** ✅ vigente (revisar se o sistema crescer)

**Contexto.** Modelar analitos como classes seria elegante.

**Decisão.** Manter módulos + arrays + `Dictionary`.

**Por quê.** Em VBA, OO adiciona instanciação, gerenciamento de objetos, serialização e
depuração mais difícil, com ganho prático pequeno nesta escala. Módulos com arrays são
mais simples, mais rápidos e mais fáceis de manter.

---

## ADR-015 · `frmMassa` usa área de colagem, não grade editável
**Status:** ✅ vigente

**Contexto.** O requisito pedia "janela semelhante a uma planilha" para colar até 50 linhas.

**Decisão.** Área de colagem TSV + lista de erros, em vez de grade célula a célula.

**Por quê.** Uma grade real de 50 linhas × 43 colunas exigiria ~2.000 controles, o que
trava o VBA. A colagem direta do Excel é o caminho que o analista já usa.

**Mesma lógica em `frmExcluir`:** `ListBox` em modo opção (checkbox por item) em vez de 40
CheckBoxes dinâmicos — mesmo efeito visual, rola bem, carrega instantaneamente.

---

## ADR-016 · Validação acumulativa, nunca interrompida no primeiro erro
**Status:** ✅ vigente

**Contexto.** Importação em massa com uma inconsistência por vez força o usuário a
descobrir os erros um a um.

**Decisão.** A validação percorre tudo e reporta **todas** as inconsistências, com
linha · coluna · motivo. Nada é gravado enquanto houver erro.

---

# Decisões pendentes de registro

Serão convertidas em ADR quando implementadas:

- **Trilha de auditoria** — `Audit_Log` append-only, chamada de dentro de `mDados`
  (nunca dos formulários), com valor antes/depois e justificativa
- **Versionamento de resultado** — substituir sobrescrita silenciosa por nova versão
  com a anterior marcada como superada
- **Proteção persistente** — `<sheetProtection>` no arquivo salvo, senha no projeto VBA,
  distribuição em estado bloqueado

---

## ADR-017 · Power Query como alimentador do banco
**Status:** ⏳ decidida, implementação pelo gestor (Claude auxilia)

**Contexto.** A entrada de dados hoje é por UserForm ou pela área de dados interfaceados
em `DB_Resultados`. Alimentar o banco a partir de exportações do equipamento/LIS reduz
digitação e erro de transcrição.

**Decisão.** Power Query passará a alimentar o banco. **Implementação a cargo do gestor**;
o assistente atua em apoio (revisão, integração com o motor, testes).

**Pontos de integração a resolver no desenho — cada um toca um ADR vigente:**

1. **Atribuição do `RUN` (ADR-004/011).** **O ETL nunca gera RUN.** RUN pertence ao
   domínio da aplicação. Power Query importa exatamente o que veio da origem; a atribuição
   do RUN acontece no motor VBA, depois da validação. Se cada origem (CSV, XML, API, LIS)
   trouxesse seu próprio RUN, cada uma traria também sua regra de numeração — e a unicidade
   global se perderia. Com a geração centralizada, qualquer origem futura entra pelo mesmo
   caminho sem alterar o modelo.

2. **Trilha de auditoria (pendente).** A decisão de logar **dentro de `mDados`** existe
   justamente para nenhum caminho de gravação escapar. Power Query escreve direto na
   planilha, contornando essa camada. Sem tratamento, a carga externa fica invisível ao
   log — exatamente a lacuna que ISO 15189 §8.4.1 cobra. Opções: rotina de pós-carga que
   registra o lote de importação, ou área de *staging* separada que só entra no banco via
   `mDados`.

3. **Proteção persistente (pendente).** Uma consulta que atualiza precisa da área
   gravável. Isso conflita com distribuir o arquivo em estado bloqueado. Provável
   solução: PQ escreve numa aba de staging desprotegida; `DB_Resultados` permanece
   protegido e só é alimentado por código.

**Recomendação de desenho (não vinculante).** `Power Query → aba de staging →
validação/atribuição de RUN → mDados → DB_Resultados`. Preserva os três ADRs acima sem
abrir exceção para nenhum deles.

---

## ADR-018 · Pipeline de importação
**Status:** ⏳ decidida, não implementada

**Contexto.** Sem um pipeline explícito, cada nova origem de dados tenderia a criar seu
próprio caminho até o banco — multiplicando regras de validação, de numeração e de log.

**Decisão.** Toda entrada de dados, de qualquer origem, percorre **um único pipeline**:

```
Origem (LIS · CSV · XML · API · equipamento)
        ↓
Power Query                      ← importa exatamente o que veio; não transforma, não numera
        ↓
STG_Importacao (aba oculta)      ← única aba gravável; staging
        ↓
Validação estrutural             ← colunas presentes, tipos, cardinalidade
        ↓
Validação semântica              ← analito existe, lote cadastrado, nível válido, data plausível
        ↓
Normalização                     ← decimal, data, código de lote, nome de analito
        ↓
Atribuição de RUN                ← SOMENTE aqui; domínio da aplicação
        ↓
mDados.Upsert()                  ← ponto único de gravação; dispara a trilha de auditoria
        ↓
DB_Resultados                    ← única fonte operacional, protegida
        ↓
Motor estatístico
        ↓
Eng_Saida
        ↓
Estatística · Painel · Liberação  ← consomem por fórmula
```

**Invariantes que o pipeline garante:**

- Nenhuma origem escreve em `DB_Resultados` diretamente
- Nenhuma origem gera RUN
- Nenhuma gravação escapa da trilha de auditoria — o log vive em `mDados`
- `STG_Importacao` é a **única** aba gravável; todas as demais permanecem protegidas
- Trocar ou acrescentar origem não altera nada a jusante do staging

**Consequência.** Uma integração futura com LIS ou API não exige rediscutir arquitetura:
basta fazê-la depositar em `STG_Importacao`.

---

## ADR-019 · O motor é a única camada autorizada a calcular
**Status:** ⏳ decidida, não implementada (fecha com o Sprint HARDENING 1)

**Contexto.** A separação `DB_Resultados → Motor → Eng_Saida → Interface` já estava
implícita nas decisões anteriores, mas nunca foi escrita como regra. O que não está
escrito não é padrão — é hábito, e hábito se perde na primeira manutenção feita por
outra pessoa.

**Decisão.** **Nenhuma planilha executa cálculo estatístico sobre `DB_Resultados`.**
Todo cálculo — média, DP, CV%, Bias, ET, Sigma, Z, Westgard, elegibilidade — é executado
exclusivamente pelo motor (`mEstatistica`), que grava em `Eng_Saida`. As abas `Painel`,
`Estatística`, `Liberação` e qualquer interface futura **apenas consomem** `Eng_Saida`,
por fórmula ou consulta.

**Fronteira, na prática:**

| Camada | Pode | Não pode |
|---|---|---|
| `DB_Resultados` | armazenar | conter fórmula de cálculo |
| `mEstatistica` | ler o banco, calcular, gravar em `Eng_Saida` | escrever em aba de interface |
| `Eng_Saida` | receber do motor | ser editada manualmente |
| `Painel` · `Estatística` · `Liberação` | referenciar `Eng_Saida`, formatar, filtrar | `AGGREGATE`/`MÉDIA`/`DESVPAD`/`CONT.SES` sobre `DB_Resultados` |

**Por quê.** Dois motivos concretos, ambos já observados neste projeto:

1. **Divergência silenciosa.** Uma fórmula de aba e o motor podem calcular a mesma
   grandeza com critérios de elegibilidade diferentes. Os dois números aparecem, nenhum
   acusa erro, e o laboratório libera com o que estiver na tela.
2. **Erosão do padrão.** Uma fórmula nova colocada direto na `Estatística` daqui a seis
   meses não quebra nada visivelmente — só recria, aos poucos, o acoplamento que a
   Fase 3 desfez.

**Consequência.** Toda grandeza nova nasce no motor, não na planilha. Isso é mais lento
para prototipar e é intencional: é o preço de ter um número só, rastreável até a linha
do banco que o gerou.

**Como se verifica.** Item 2.1 do Quality Gate: varredura de fórmulas nas abas de
interface procurando referência direta a `DB_Resultados`. Esperado: zero.

---

## ADR-020 — A importação em massa é uma ABA, não um formulário

**Data:** 05/08/2026
**Status:** aceito
**Substitui:** o `frmMassa` (F2-5)

**Contexto.** O `frmMassa` pedia que o analista colasse dados numa caixa de texto.
Isso inverte a forma como ele pensa: na bancada o dado nasce numa tabela com os
analitos lado a lado, uma corrida por linha — que é exatamente o formato em que
o equipamento exporta e a planilha de origem já vive. Colar isso num TextBox
obrigava o usuário a conferir no escuro, sem ver colunas nem cabeçalho.

**Decisão.** A entrada em massa passa a ser a aba `Importar`:

- cabeçalho `Data | Nível | Lote | <analitos do produto>`, gerado **da aba
  Analitos** do próprio produto (Bioquímica 20, Hematologia 28 — paramétrico por
  construção, não por lista fixa);
- só a área de colagem é destravada; cabeçalho e resto da aba ficam protegidos;
- o botão **Registrar** migra os dados para o `DB_Resultados` e limpa a aba;
- o botão da aba `Resultados` deixa de abrir o `frmMassa` e passa a navegar
  para cá.

**O que NÃO mudou.** A origem é outra; o pipeline é o mesmo do ADR-018. A aba
não escreve no banco: ela alimenta a mesma sequência
`ParseData → ParseNum → CodigoLote → ObterOuCriarRUN → UpsertResultados`.
Nenhuma origem gera RUN, e nenhuma escrita escapa da trilha de auditoria.
Reaproveitar essas rotinas em vez de reescrevê-las é o que garante que a
importação por aba e o `frmCorrida` não possam divergir em regra.

**Tudo-ou-nada.** Se qualquer linha tiver problema, **nada** é gravado: as
inconsistências aparecem numa coluna à direita e o usuário corrige. O banco
nunca recebe importação pela metade. Duas consequências que valeram correção
depois do primeiro teste:

1. Os valores dos analitos são validados **mesmo quando o cabeçalho da linha
   está inválido**. Se o laço só rodasse com data/nível/lote corretos, um valor
   não numérico ficaria escondido até o usuário consertar a data e clicar de
   novo — e o tudo-ou-nada deixaria de mostrar tudo de uma vez.
2. Duas linhas para a **mesma corrida e nível** são recusadas. O `Upsert`
   obedeceria e a segunda sobrescreveria a primeira em silêncio; numa colagem
   isso é erro de digitação, não intenção, e perda de dado sem aviso contradiz
   a prioridade de integridade.

**Interface separada da regra.** `RegistrarImportacao` (o botão) só traduz
resultado em `MsgBox`; toda a regra vive em `ExecutarImportacao(silencioso)`.
Assim o teste automatizado exercita o **mesmo** código que o botão executa, e
não uma cópia que poderia divergir dele.

**Como se verifica.** Itens 1.7–1.10 do Quality Gate (cabeçalho bate com os
analitos do produto, botões ligados, só a área de colagem editável) e 3.16–3.17
(colagem válida migra e some da aba; colagem com erro não grava nada, nem a
linha válida).

---

## ADR-021 — Produção é a fonte; o build é o produto

**Data:** 05/08/2026
**Status:** aceito

**Contexto.** O sistema passou a existir em dois arquivos que ninguém tinha
declarado como linhas distintas:

| | `QC_Bioquimica.xlsm` (raiz) | `_entregas/QC_Bioquimica_hardening1.xlsm` |
|---|---|---|
| Aba `Importar`, 31 analitos | ✅ | ❌ |
| Trilha de auditoria, motor, trava | ❌ | ✅ |

O gestor perguntou "você já fez a trilha de auditoria?" e a resposta honesta era
"sim, mas não no arquivo que você usa". Pior: `RegistrarLog` na produção é um
**stub vazio**, então a importação recém-entregue gravava no banco sem rastro —
lacuna direta em ISO 15189 §8.4.

**Decisão.** As duas não são linhas rivais. `build_all.ps1` **começa copiando a
produção**, logo:

- **produção = fonte.** É onde o laboratório trabalha e onde nascem dados de
  configuração: analitos, lotes, especificações, usuários.
- **`_entregas/` = produto.** Recebe o que só o build sabe aplicar: motor,
  `Eng_Saida`, trilha de auditoria encadeada, trava de estrutura, Sprint NC.

Tudo que for **configuração** entra pela produção e sobe no próximo build de
graça. Tudo que for **engenharia** entra por script versionado e nunca é
aplicado à mão na produção.

**Consequência prática.** O que era conflito virou pipeline. Três ajustes
bastaram, e nenhum foi reescrita:

1. `criar_aba_importar.ps1` deixou de gerar o cabeçalho da aba `Analitos` e
   passou a ler o mesmo CSV que a produção usa — os dois arquivos produzem a
   aba **idêntica**.
2. `mImportar` existe em duas versões, `src_producao/` e `src_hardening1/`, com
   layout igual e só a API de RUN diferente (`NovoRUN` × `PreverRUN` +
   `ObterOuCriarRUN`). Não é duplicação preguiçosa: VBA não tem compilação
   condicional, e um módulo chamando rotina inexistente **derruba o projeto
   inteiro**, não só a si mesmo.
3. `patch_forms_run.ps1` passou a tratar a ausência do `frmMassa` como estado
   esperado — ele foi substituído, não perdido.

**O que isto impede de repetir.** O defeito de julho (motor validado que nunca
entrou no `.xlsm`) e o de hoje (aba entregue no build e ausente na produção) são
o mesmo erro: **confundir onde a mudança foi feita com onde ela precisa estar.**
A regra de ouro passa a ser: se o gestor abre o arquivo e não vê, não está
pronto — independentemente de quantos testes passaram noutro artefato.

---

## ADR-022 — Especificações de Qualidade como módulo, com histórico por ano

**Data:** 05/08/2026
**Status:** aceito
**Relacionado:** ADR-018 (pipeline), ADR-019 (motor é a única camada de cálculo)

**Contexto.** A meta analítica define *contra o que* o laboratório é julgado. Hoje
ela existe na aba `Analitos`, colunas `K`–`W`, e a matemática está correta:
`S = TEa/3`, `T = CVi × fi`, `R` resolvido por cascata, e a tabela de rigor
`U:W` com ÓTI 0,25/0,125 · DES 0,5/0,25 · MIN 0,75/0,375. Três defeitos de
**modelagem**, porém, a tornam inadequada para uso real:

1. **Não há dimensão ANO.** O CLIA mudou em 2024; a variação biológica é
   revisada periodicamente. Sem histórico, reabrir uma corrida de 2025 em 2030 a
   julga pela meta de 2030 — o laudo deixa de ser reproduzível, o que é
   inaceitável sob ISO 15189 §7.6.
2. **A spec mora no armazém POR LOTE** (`aInput` = `E4:P43`, que inclui `K..P`).
   TEa CLIA é constante regulatória e CVi/CVg são constantes biológicas: **não
   são propriedades do lote**. Trocar o lote hoje troca o CLIA junto.
3. **Não há escolha de fonte** — é cascata fixa `Config > CLIA > VB`. O
   Fabricante sequer participa: existe `L = CV Fab %` que não alimenta nada, e
   não existe BIAStp em lugar nenhum.

**Decisão — camadas, iguais às do resto do sistema.**

```
DB_Especificacoes  (banco tabular, append)
        ↓
mEspecificacoes    (motor: resolve Ano+Analito+Fonte → CVtp, BIAStp, ETp)
        ↓
Eng_Saida / Painel / Estatística   (só leem)
```

Nenhuma aba de interface volta a calcular meta. Vale o ADR-019 sem exceção.

**FONTE é dado; MODELO é um conjunto fechado.**

Aqui divergimos de "cadastre qualquer fonte e nunca mexa no código". Cada fonte
**informa grandezas diferentes**, e é isso que decide a conta:

| Modelo | A fonte informa | O motor deriva |
|---|---|---|
| `ETP_DIRETO` | ETp | `CVtp = ETp/3`; BIAStp indefinido |
| `VB` | CVi, CVg, rigor | `CVtp = CVi·fi` · `BIAStp = √(CVi²+CVg²)·fb` · `ETp = BIAStp + 1,65·CVtp` |
| `CV_BIAS_DIRETO` | CVtp, BIAStp | `ETp = BIAStp + 1,65·CVtp` |

CLIA usa `ETP_DIRETO`; Variação Biológica usa `VB`; Fabricante usa
`CV_BIAS_DIRETO`. **Ricós, EFLM, CAP, RCPA, Rilibak entram como LINHA, escolhendo
um dos três modelos — sem tocar em código.** Uma fonte com matemática realmente
nova exige um modelo novo, e dizer isso é mais honesto do que prometer
extensibilidade infinita.

**A meta é resolvida pelo ANO DO RESULTADO, nunca pelo ano corrente.**

Este é o requisito central. A regra é de vigência: vale a especificação de maior
`Ano` que seja **≤ ano do resultado**. Assim um resultado de 2025 continua sendo
julgado pela meta de 2025 em 2035, e cadastrar a meta de 2027 não reescreve o
passado. Casar por ano exato seria pior: abriria buracos em todo ano sem
cadastro.

**Rastreabilidade.** O motor devolve, além dos três números, a `Fonte`, o `Ano`
vigente, o `Rigor` e o **ID da especificação** aplicada. Esse ID é o que permite
ao `Audit_Log` registrar *contra qual meta* uma corrida foi julgada — sem ele, a
trilha diz o que foi decidido mas não com base em quê.

**Migração, não convivência.** As colunas `K`–`W` da `Analitos` são a
implementação anterior. Manter as duas seria criar duas fontes de verdade para o
ETp — exatamente a divergência que o ADR-019 existe para impedir. Os valores
atuais são migrados para `DB_Especificacoes` com o ano vigente, e a `Analitos`
passa a **exibir** o que o motor resolveu.

---

## ADR-024 — O núcleo do lote é derivado por comprimento, não por posição fixa

**Data:** 06/08/2026
**Status:** aceito
**Relaciona-se a:** ADR-004 e ADR-011 (RUN), ADR-009 e ADR-019 (conhecimento em um lugar só)

**Contexto.** O laboratório colou 555 resultados pela aba `Importar`, o banco
gravou tudo corretamente — e o `Calc`, o `Painel` e a aba `Resultados` ficaram
**em branco**. Nenhum erro, nenhuma mensagem. O dado entrava e não saía.

A causa era uma conta de duas linhas:

```
CodigoLote("8974", 1)  ->  "QC-897401"      (monta)
Mid(codigo, 4, 6)      ->  "897401"          (desmonta)
loteAtivo                  "8974"
```

`CodigoLote` monta `"QC-" & núcleo & nível(2 dígitos)`. O inverso correto é
portanto `Mid(codigo, 4, Len(codigo) - 5)` — tira os 3 do prefixo e os 2 do
nível, **seja qual for o tamanho do núcleo**. O sistema usava `6` fixo, que só
acerta quando o núcleo tem exatamente 6 dígitos. O lote real do laboratório tem
4, então a extração engolia os dígitos do nível e o filtro do `Calc`
(`(""&rLote)=(""&loteAtivo)`) nunca casava.

**Por que sobreviveu a tantas revisões.** `mEntrada.NucleoLote` já existia e já
fazia exatamente essa conta. **Nenhum chamador a usava:** os 13 pontos que
precisavam do núcleo reescreveram o `Mid` inline. Conhecimento duplicado em 13
cópias é a falha que o ADR-009 e o ADR-019 tratam noutro domínio — e aqui ela
teve a consequência clássica: corrigir um lugar não corrigia nada, e o rótulo
"Lote (6 díg.)" na `Configuração` fazia a suposição parecer razoável.

**Decisão.**

1. A derivação do núcleo vive em **uma** função, `NucleoLote`, e é
   `Mid(codigo, 4, Len(codigo) - 5)`. Todo consumidor chama a função.
2. A mesma conta na camada de fórmula: `DB_Resultados!BA` passa a
   `MID($D4,4,LEN($D4)-5)`.
3. O sistema **não exige** núcleo de 6 dígitos. Aceitar um lote de qualquer
   comprimento e depois não encontrá-lo é pior do que recusá-lo na entrada.

**Consequência sobre o RUN, que é o achado maior.** Como o *código* do lote
embute o nível, montar a chave da corrida com ele fazia `QC-897401` e
`QC-897402` — Nível 1 e Nível 2 da **mesma** corrida — virarem chaves
diferentes. Cada nível ganhava o seu próprio RUN, e 18 corridas viravam 36, com
os níveis plotados em faixas separadas do eixo X. Isso contradiz o ADR-011
diretamente. Com o núcleo (que não tem nível), a chave volta a ser
`data + lote` e os níveis compartilham o RUN, como a chave lógica
`RUN + Analito + Nível` sempre pressupôs.

**Nota de método.** O defeito do RUN só apareceu porque a conferência contou
**36 corridas para 18 datas** — um número que não fechava. A correção do lote
já tinha feito os gráficos plotarem, e parar ali teria deixado o modelo de RUN
quebrado sob uma tela que parecia certa.

---

## ADR-023 — Conformidade tem quatro estados, e ausência nunca é aprovação

**Data:** 06/08/2026
**Status:** aceito
**Refina:** ADR-022

**Contexto.** `AvaliarConformidade` tinha três estados: `CONFORME`,
`NAO CONFORME` e `SEM META`. O terceiro fundia duas situações que só parecem a
mesma:

- **não existe meta cadastrada** para (analito, ano, fonte);
- **existe meta, faltam medições** — nenhum resultado no período, ou n < 2.

São causas opostas, com donos opostos. A primeira é falha de gestão da
qualidade e quem resolve é quem cadastra a especificação; a segunda é lacuna
operacional e quem resolve é a bancada. Numa auditoria ISO 15189 são achados
diferentes, e fundir os dois esconde qual é.

**Decisão.** Quatro estados:

| Estado | Significa | Quem age |
|---|---|---|
| `CONFORME` | mediu e está dentro | ninguém |
| `NAO CONFORME` | mediu e está fora | analista investiga a corrida |
| `SEM ESPECIFICACAO` | não há meta cadastrada | gestor da qualidade cadastra |
| `SEM DADOS` | há meta, faltam medições | bancada roda controles |

Nomeados pelo **que falta**, não pelo que não aconteceu (`NAO_CADASTRADA` /
`NAO_AVALIADA`): o nome sozinho diz qual é a ação.

**Nenhum dos dois últimos é aprovação.** Tratar ausência — de meta ou de dado —
como conformidade é o jeito mais silencioso de um laboratório se achar em dia.

**Um quinto caso, resolvido noutra camada.** Linha cadastrada mas **inválida**
(ETp textual, modelo inexistente) não é o mesmo que ausência de cadastro:
alguém *tentou* e o dado está ruim, o que é problema de qualidade de dado.
Para a *avaliação* não há critério utilizável, então cai em
`SEM ESPECIFICACAO`; o diagnóstico não se perde:

- `ResolverEspec` devolve `ESPEC_INVALIDA|<ID>|<motivo>`;
- `SituacaoEspec` devolve `INVALIDA: <motivo>`;
- `Eng_Especificacoes` publica a coluna **Situação** por analito.

Assim o veredito fica com quatro estados limpos e o motor continua sabendo
mais do que o veredito mostra — que é a divisão certa entre as duas perguntas:
*"o laboratório está dentro da meta"* e *"existe meta utilizável, e se não, por
quê"*.

**Correção de contrato junto.** `ResolverEspec` só devolve prefixo `OK` quando
há **pelo menos uma grandeza derivada**. Antes, achar a linha bastava — mesmo
com as três metas vazias. Nenhum consumidor se enganava, porque todos conferem
se o valor existe antes de usar; mas contrato que só funciona porque quem chama
é cuidadoso é contrato quebrado esperando a hora.

**Nota sobre a origem desta mudança.** A revisão externa que a motivou descrevia
o sistema aprovando analito sem especificação (`sem especificação → OK`). Não
era o caso: `AvaliarConformidade` já devolvia `SEM META`, e havia teste para
isso. O `OK` estava no *protocolo interno* de `ResolverEspec`, não no veredito.
A recomendação estava certa; o diagnóstico, não — e a diferença importa, porque
uma leva a corrigir um contrato e a outra levaria a procurar um defeito que não
existe.

---

## ADR-025 — Capacidade de 60 meses: as flags do banco saem da fórmula e vão para o VBA

**Contexto.** A análise de escalabilidade de 12/08/2026 mediu o arquivo real e
achou dois problemas na mesma estrutura, o `DB_Resultados`.

O primeiro é de **integridade**, e era o urgente. As colunas `BA`, `BB` e `BC`
eram fórmula provisionada até a linha 15.003, e os intervalos nomeados `r*`
tinham essa mesma altura fixa. Janeiro consome 1.110 registros; o teto chegava
em **13,5 meses**. Passado ele, `UltimaLinhaBanco` continuava gravando — usa
`End(xlUp)`, que não tem limite — mas as linhas novas nasciam sem as fórmulas
derivadas e fora dos intervalos. O dado entrava, ficava salvo, e desaparecia do
Painel, do Calc, dos gráficos e da Estatística. Sem erro. É a falha que este
projeto mais teme: não o erro visível, e sim o número plausível e errado.

O segundo é de **custo**. `BB` e `BC` usavam `COUNTIFS` de faixa expansiva
(`$E$4:$E4`): na linha 4 varrem uma célula, na linha 66.603 varrem 66.600. A
soma é ~2,5n². Medido: recálculo completo de 5,7 s com um ano de dados e 221 s
com cinco.

**Decisão.** `BA`, `BB` e `BC` deixam de ser fórmula e passam a ser **valor
gravado** por `AtualizarFlagsBanco`, em `mBanco.bas`: uma varredura, dois
dicionários, uma escrita em bloco — O(n). E os intervalos nomeados deixam de ter
altura fixa: `RedimensionarNomes` os ajusta à última linha real a cada gravação.

**Por que redimensionar em vez de só provisionar mais alto.** A aba `Calc`
avalia `AGGREGATE` sobre esses intervalos 180 vezes, e o custo é O(180 × altura)
**sem curto-circuito em célula vazia**. Provisionar 120.000 linhas de intervalo
com 1.110 preenchidas custaria cem vezes mais do que precisa. Com o nome
acompanhando o dado, o custo acompanha o dado — e provisionamento alto deixa de
ter preço.

**A semântica que não podia mudar.** `BB` vale 1 quando a linha é a primeira
**Ativa** do par (analito, RUN); `BC`, quando é a primeira Ativa do RUN. O
detalhe que decide o desenho: *primeira entre as ativas*. A flag **não é estável
no momento da inserção** — excluir logicamente uma linha promove a próxima
duplicata a "primeira". Por isso `AtualizarFlagsBanco` recalcula o banco inteiro
e é chamada tanto por `UpsertResultados` quanto por `ExcluirLogico`. Calcular a
flag só na inserção seria mais rápido e estaria errado.

**Capacidade explícita.** `CAP_LINHAS = 120000`, dimensionada pelo pior caso
plausível (40 analitos × 2 níveis × 23 dias úteis = 1.840/mês × 60 meses =
110.400, mais 8,7% de margem). `ExigirCapacidade` recusa a gravação **antes** de
escrever, com mensagem. A regra: é preferível bloquear com mensagem clara a
aceitar o dado e deixá-lo invisível para os cálculos.

**Rastreabilidade preservada.** A conta continua explicável — é a mesma regra da
fórmula, transcrita no cabeçalho de `mBanco.bas` junto com a fórmula de origem.
`ConferirFlagsBanco` recalcula por um caminho independente (contagem direta, sem
dicionário, deliberadamente O(n²)) e devolve o número de divergências. Serve de
prova sob demanda, inclusive em auditoria.

**Efeitos colaterais medidos.** O banco perdeu 45.000 fórmulas e o arquivo caiu
de 1,71 MB para 1,15 MB. A suíte vai acusar variação grande no diff de fórmulas
do `DB_Resultados` — é esperado e é o objetivo, não regressão.

---

## ADR-026 — Camada de dados para BI: uma tabela fato, não uma cópia da planilha

**Contexto.** O QC_INI precisa alimentar um Power BI acessível por web e celular,
sem que o Excel deixe de ser o núcleo operacional. O risco óbvio era exportar
"a planilha" e criar uma segunda verdade sobre o mesmo controle.

**Decisão.** A interface é a tabela estruturada `tblBI_Fato`, na aba oculta
`BI_Data`, com **granularidade (Lote, RUN, Nível, Analito)** — a mesma chave
natural que o `UpsertResultados` usa. Chaves textuais estáveis (`ID_Result`,
`ID_Corrida`, `ID_Analito`, `ID_Lote`), nenhuma dependente de posição de célula.
`ListObject` e não faixa: o Power Query referencia pelo nome e o banco cresce sem
tocar na consulta.

**Alvo por lote, não o da tela.** A aba `Analitos` mostra as metas do lote ativo;
o histórico vive no `LotesStore`. Ler o alvo da tela para um resultado de outro
lote daria um Z errado, plausível e silencioso. `mBI` lê sempre o bloco do lote a
que o registro pertence.

**O que não entrou.** `Equipamento` e `Setor`: a produção não tem as abas que os
guardariam. `4_1s` e `10x`: `Calc!N3` e `Calc!O3` são `IF(OR(FALSE;FALSE);1;0)` —
o motor não implementa essas regras. Coluna sem origem é coluna que alguém filtra
e conclui errado; regra fabricada no painel do gestor é pior do que ausência.

**Reconciliação como portão de build.** `mBI.ReconciliarComCalc` compara Z e
veredito linha a linha contra o `Calc`, e o build **falha** com qualquer
divergência. Não é cerimônia: na primeira execução acusou **45 divergências em
50**, com o Z idêntico e o veredito grudado — `Dim` dentro de laço no VBA não
cria variável nova a cada volta, e as flags de Westgard acumulavam da primeira
violação em diante. O painel reprovaria corrida boa, com número plausível.

**Sobre o `.pbix`.** Power BI Desktop está instalado, mas sem `pbi-tools`,
Tabular Editor ou `MicrosoftPowerBIMgmt` não há caminho suportado para gerá-lo
por script. **Nenhum `.pbix` foi fabricado.** Entregam-se as queries M, as
medidas DAX, o desenho do modelo e o roteiro de montagem.

---

## ADR-027 — Analitos!R/S/T como fonte única das especificações ativas

**Contexto.** A auditoria encontrou **três fontes concorrentes de ETp** no mesmo
arquivo:

| Consumidor | De onde vinha o ETp | Alimentava |
|---|---|---|
| `Estatística!O` | `LimEspec(analito, ano, "CLIA", "ETP")` — **CLIA cravado no argumento** | Sigma da Estatística (`R`) |
| `Painel!F7/F8` | `INDEX(engETp, …)` — motor `Eng_Especificacoes` (ADR-022) | Sigma do Painel (`I`) |
| `Analitos!S` | seleção conforme `Analitos!R` | **ninguém** |

Consequência concreta: Painel e Estatística podiam exibir **Sigmas diferentes
para o mesmo analito**, e trocar a fonte em `Analitos!R` não mudava nada em lugar
nenhum. O mesmo valia para o CV: `Estatística!G` chamava `LimEspec(…,"CLIA","CV")`
enquanto `Analitos!T` já resolvia o CVTp conforme a fonte.

**Decisão.** `Analitos` passa a ser a *single source of truth* das especificações
**ativas**:

```
Analitos!R (fonte: CLIA | VB | FAB)
   ↓
Analitos!S = ETp%   ·   Analitos!T = CVTp%
   ↓
Estatística (O, G) → Sigma → Painel → BI_Data
```

Publicados como nomes — `espFonte`, `etpOficial`, `cvtpOficial` — para que os
consumidores refiram o conceito e não a coordenada.

**As fórmulas da `Analitos` não foram tocadas.** Foram criadas e testadas pelo
gestor; esta mudança só liga quem as consome.

**Sigma só calcula com ETp numérico.** `Analitos!S` devolve o texto
`"DEFINIR FONTE QUE CONTENHAM DADOS"` quando a fonte escolhida não tem dado. O
guarda `ISNUMBER` não mascara erro: diz *"não há meta utilizável"*, que é a
verdade, e é o mesmo critério de quatro estados do ADR-023. Zero seria lido como
"desempenho péssimo"; vazio é lido como "sem meta".

**Sobre `Cfg_/DB_/Eng_Especificacoes` — mantidas (Cenário C).** Não são
redundantes e não podem ser removidas:

- `DB_Especificacoes` guarda o **histórico por ano** com campos de auditoria
  (`ES_ANO`, `ES_USUARIO`, `ES_DATACAD`, `ES_ATIVO`) que a `Analitos` **não tem** —
  ela só representa o estado ativo, sem dimensão temporal.
- `Cfg_Especificacoes` guarda o catálogo de fontes e seus modelos
  (`ETP_DIRETO`, `VB`, `CV_BIAS_DIRETO`), que alimenta o cadastro.
- `Eng_Especificacoes` continua como saída do motor por ano, consumida por
  `LimEspec` — que segue servindo às **comparações informativas** da Estatística
  (colunas `L`, `M`, `P`), onde o ponto é justamente mostrar a meta de *outra*
  fonte ao lado da escolhida.

A separação passa a ser: **`Analitos` = o que está em uso agora**;
**`DB_Especificacoes` = o que valia em cada ano, com rastro**. Nenhuma das duas
compete pela mesma pergunta.

---

## ADR-028 — O ano em `Analitos!S2` passa a mandar, e a camada antiga sai

**Contexto.** A arquitetura pretendida era `S2 (ano) → histórico nas linhas 46–89
→ A4:W43`. A auditoria mediu: das 23 colunas do bloco vigente, **zero fórmulas
citavam `S2` e zero citavam as linhas 46–89**. As colunas K, M, N, O, P, Q e S
eram literais digitados. O histórico estava lá; ninguém o lia. Trocar o ano não
mudava nada em lugar nenhum.

**Decisão.** Construir a ligação que faltava e retirar a camada antiga inteira.

```
Analitos!S2 (ano)
   ↓  INDEX no bloco 46..89, por ano e por analito
Analitos!K, M, N, O, P          entradas do ano
   ↓  seleção do gestor em S (CLIA | VB | FAB)
Analitos!T = ETp%   ·   Analitos!U = CVTp%
   ↓
Estatística · Painel · Sigma · BI_Data
```

**O que continua digitado, de propósito.** `Q` (Desemp.) e `S` (fonte) são
**decisões** do laboratório, não dados do ano. Puxá-las do histórico
transformaria uma escolha em consequência.

**Classificação das abas, depois de rastrear dependências:**

| Aba | Veredito | Por quê |
|---|---|---|
| `Cfg_Especificacoes` | **C — obsoleta** | catálogo de fontes; hoje é a validação `"CLIA,VB,FAB"` em `Analitos!S` |
| `DB_Especificacoes` | **C — obsoleta** | era o histórico por ano; o histórico agora vive na `Analitos`. E estava **vazia**: `LimEspec` devolvia branco para os 40 analitos |
| `Eng_Especificacoes` | **C — obsoleta** | saída do motor; sem consumidor desde o ADR-027, e o `mBI` foi reapontado |

**Mudança de veredito registrada.** O ADR-027 classificou `DB_Especificacoes`
como *indispensável* porque a `Analitos` não tinha dimensão de ano. Isso deixou
de ser verdade. Registrar que o veredito mudou importa tanto quanto o veredito.

**Dois defeitos achados no caminho:**

1. **76 `#REF!` na `Estatística`** (colunas H, I). A causa não era a coluna: era
   o mapeamento **posicional** `H14 = Analitos!N4`. Com 40 analitos × 2 níveis,
   as linhas do nível 2 apontavam para `Analitos!N44:N83`, fora do bloco.
   Corrigido para busca por **chave de analito**.
2. **O `mBI` publicava o campo errado.** Lia `Analitos` coluna 18 como ETp e 17
   como Fonte — no layout atual, *ETp VB* e *Desemp.* O painel do gestor exibia
   o nível de desempenho (`OTI`/`DES`/`MIN`) como se fosse a fonte da
   especificação. Reapontado para S/T/U.

**Consequência que o gestor precisa saber.** Saiu junto a tela de cadastro
(`frmEspecificacoes` + botão). O cadastro passa a ser feito digitando no próprio
bloco histórico da `Analitos`. Não houve perda de informação — o `DB` estava vazio.

**Lição de método.** A primeira tentativa de remover `LimEspec` usou
`ProcStartLine`/`ProcCountLines` e apagou 46 linhas em vez de 23, quebrando a
compilação de `mEstatPeriodo`. Como **todas** as UDFs vivem nesse módulo, 642
células viraram `#NOME?` de uma vez. Recortar pelo texto exato da função é
determinístico; e nenhuma remoção de VBA deve ser dada por boa sem **chamar uma
UDF depois** — função quebrada não levanta erro, devolve `#NOME?` em silêncio.

---

## ADR-030 — O bias de ET e Sigma passa a vir do ensaio de proficiência

**Contexto.** `Estatística!K` — a célula que alimenta Erro Total e Sigma — chamava
`EstatPeriodo(…,"BIAS")`, que resolve em
`mEstatistica.CalcularBias(mediaObs, alvoDoLote)`:

```
mediaObs = média do CONTROLE INTERNO no período
alvo     = média atribuída AO LOTE do controle interno
```

Isso mede a **deriva do CQI contra o alvo do próprio lote**. É uma medida útil, e
não é erro sistemático: erro sistemático se mede contra um valor **externo e
independente** — o consenso do grupo no ensaio de proficiência. O próprio
cabeçalho dizia *"Bias % (alvo do lote)"*. Um laboratório pode estar perfeitamente
centrado no alvo do fabricante e ainda assim 8% acima do grupo, e era esse 8% que
sumia do Sigma.

**A aritmética sempre esteve certa.** `CalcularErroTotal = |bias| + 1,65·CV` e
`CalcularSigma = (ETp − |bias|)/CV` já estavam corretas, e o Sigma já usava o CV
**observado**, não o CVTp. O que estava errado era o `ref` do bias.

**A fonte já existia.** `EQC_Dados` guarda o EP linha a linha, com
`N = (X_lab − X_ref)/X_ref × 100` (assinado) e `O = |N|`. Ninguém as consumia
para ET e Sigma. `mCEQ` **consome** N e O — não recalcula. A conta acontece uma
vez, na célula, onde é visível e auditável.

**Consolidação de múltiplas rodadas: média das magnitudes.**

```
|Bias|consolidado = Σ|Bias_i| / n
```

Nunca a média dos assinados. Rodadas de +5% e −5% descrevem um método que oscila
5% em torno do grupo; a média assinada daria 0% e afirmaria exatidão perfeita. O
bias assinado continua disponível em `Estatística!M`, para leitura da direção.

**Vigência: a rodada mais recente que não ultrapasse o ano em análise** — a mesma
regra do ADR-022 para especificação. Exigir coincidência exata de ano fazia o
bias sumir sempre que o CQI passasse na frente do último ciclo publicado (EP de
2025, análise de 2026).

**Ausência de dado devolve o texto `"SEM EP"`, nunca `Empty`.** Na primeira versão
deste módulo as 80 linhas exibiram bias `0,00` — porque **o Excel renderiza o
`Empty` de uma UDF como zero** — e esse zero entrou em ET e Sigma produzindo
números de aparência perfeita. É o mesmo defeito que `AlvoDoLote` tem em
`mEstatPeriodo`. Texto não é número: `ISNUMBER` reprova, e ET e Sigma ficam
vazios.

**Dois defeitos menores corrigidos junto, no Painel:**

1. `H7 = E7*1,65 + IF(ISNUMBER(G7);G7;0)` somava o bias **com sinal** — um bias
   negativo *reduzia* o erro total.
2. Tanto `H7` quanto `I7` somavam **zero** quando faltava bias, afirmando exatidão
   que ninguém mediu. Agora os dois exigem bias numérico.

---

## ADR-031 — Arquivo de trabalho sem proteção durante o desenvolvimento

**Contexto.** Proteção de estrutura, senha de aba e células travadas custaram
tempo real nesta fase: vários scripts gastaram tentativas com *"aba protegida"*,
`AllowFormattingCells=False` e `RPC_E_CALL_REJECTED` escrevendo em célula
travada. O produto ainda está sendo construído; a trava protege contra um
usuário que ainda não existe.

**Decisão.** O `.xlsm` de trabalho fica **sem senha, sem proteção de aba e sem
célula travada**. `veryHidden` vira `hidden`, para que as abas de estrutura
apareçam em *Reexibir* quando alguém precisar olhar.

Isso **não** afeta o produto entregue: `travar_estrutura.ps1` e
`blindar_artefato.ps1` continuam no build e reaplicam tudo no artefato.

**Nota sobre um fantasma.** `Cfg_/DB_/Eng_Especificacoes` foram removidas no
ADR-028 — o script reportou `aba removida` para as três e `SALVO` — e
reapareceram. O blob commitado em `fe4d372` **já as continha**, ou seja, a
remoção nunca chegou ao commit, enquanto as mudanças de VBA e de fórmula do
mesmo ADR chegaram. A pasta de trabalho fica dentro do OneDrive; o padrão é
compatível com o arquivo ter sido restaurado por sincronização entre a gravação
e o commit.

Consequência prática: **`SALVO` impresso por um script não é prova de que o byte
sobreviveu.** Depois de remover estrutura de um `.xlsm` que mora em pasta
sincronizada, vale reabrir e conferir — foi o que se fez aqui.

---

## ADR-032 — A aba de controle externo vira operável

**Contexto.** O laboratório participa de **mais de um programa** — Controllab com
4 rodadas anuais, CAP com 3 — e qual usar para julgar o desempenho é decisão
técnica dele, não do sistema. A aba `EQC_Dados` guardava um provedor só, rodadas
numéricas e nenhuma validação; a `Estatística` consumia tudo sem poder escolher.

**Decisão.**

`EQC_Dados`:

| Coluna | O quê |
|---|---|
| `C` Rodada | vira **letra** (A–D), com validação. O laboratório fala *"rodada A"*, e o número se confundia com contagem de amostra |
| `E` Provedor | validação `Controllab` / `CAP` |
| `G` `H` `I` `K` `L` | **digitados**: resultado do lab, média do grupo, DP do grupo, limite inferior e superior |
| `J` SDI | `(X_lab − média grupo) / DP grupo` |
| `M` Status limites | dentro da faixa do provedor? |
| `P` Status SDI | **novo** — `\|SDI\| ≤ 2` aprova |

`Estatística` ganha o bloco de filtros em `K3:P5` — **provedor, ano e rodada** —
publicados como `eqProvedor`, `eqAnoEP` e `eqRodada`. Os três alimentam **todas**
as colunas de EP, para que não exista um card lendo um filtro e outro lendo
outro. `K5` mostra em texto o que está em uso. Duas colunas novas: `T` Status SDI
e `U` Status limites, consolidados por analito.

**Status por pior amostra, não pela média.** Uma rodada com SDI +3 e outra com −3
dão média zero e descrevem um desempenho que ninguém aprovaria. É a pior que
reprova o conjunto.

**Limite ausente não é aprovação.** A amostra entra como *não avaliada* e aparece
na contagem — `OK (3 dentro; 1 sem limite)` — em vez de passar por conforme
alguém que ninguém conferiu.

**Sobre o separador da validação.** No XML e no COM a lista vai com **vírgula**
(`"Controllab,CAP"`); o Excel exibe e aceita ponto‑e‑vírgula na caixa de diálogo,
conforme a configuração regional. Gravar `;` criaria um item único chamado
`Controllab;CAP`. As validações que já existiam na aba usam vírgula.

**Rodadas A–D para os dois programas.** O CAP simplesmente não usa a D.
Restringir a lista por provedor exigiria validação dependente, que quebra ao
copiar linha — e o custo não paga: rodada D de CAP não existe nos dados e não
entra em média nenhuma.

---

## ADR-033 — Bias, especificação, Sigma e orçamento de erro deixam de estar embaralhados

**Contexto.** A `Estatística` chegou a esta fase com quatro famílias de indicador
misturadas nas mesmas colunas, e o gestor reestruturou a aba deixando `G`, `I`,
`J`, `K` e `M` vazias para serem preenchidas. As duas fórmulas que já existiam
tinham um defeito de modelo:

| Coluna | Fórmula encontrada | Problema |
|---|---|---|
| `H` ET | `(F*1,65)+G` | usa o bias **com sinal** |
| `L` Sigma | `(K−G)/F` | usa o bias **com sinal** |

Com um bias de −8%, a primeira **encolhia** o erro total e a segunda **inflava** o
Sigma — os dois sentidos invertidos. Westgard *et al.* (2018), p. 3:
*"SM = (TEa% – bias%) / CV. […] the bias will be an absolute percentage (the
presence of any bias always shrinks the allowable error, never enlarges it)."*

**Decisão — layout da `Estatística`.**

| Col | O quê | Origem |
|---|---|---|
| `G` | Bias EQC (abs) % | `mCEQ.BiasEQ(…,"ABS",…)` |
| `H` | ET % | `1,65×CV + \|Bias\|` |
| `I` `J` `K` | Fonte, CVTp %, ETp % | `Analitos` `S` / `U` / `T` (ADR-028) |
| `L` | Sigma | `(ETp − \|Bias\|) / CV` |
| `M` | Status sigma | `mQualidade.ClassificarSigma` |
| `N` `O` | Margem ETp em p.p. e em % | `ETp − ET` e `(ETp−ET)/ETp×100` |
| `P` | Status margem | `mQualidade.ClassificarMargem` |
| `Q` `R` `S` | Status CV, SDI (EP), limites (EP) | |
| `T` | Bias EQC **com sinal** % | direção do desvio, para leitura |
| `U` | ordem do crítico (oculta) | apoio da lista, evita fórmula matricial |

Abaixo da tabela: resumo de conformidade (linha 96) e a lista dos analitos em
margem crítica ou com ETp excedido (linha 106).

**Decisão — `Painel` em blocos separados.** O bloco descritivo fica só com
`Nível / n / Média / DP / CV%`. O que saiu dele virou dois blocos próprios —
`SIX SIGMA` em `J10` e `ERRO TOTAL vs ETp` em `J16` — ao lado do bloco Westgard,
que ganhou a coluna `Status`. Contadores globais de margem crítica em `J21`.

**A escada de classificação existe uma vez.** Enquanto estava escrita como `IF`
encadeado na `Estatística`, no `Painel` e no BI, mexer numa faixa exigia lembrar
dos três; o primeiro esquecido divergiria em silêncio, e só apareceria quando o
gestor comparasse a planilha com o relatório. Agora é `mQualidade`, versionada em
`src_producao` e importada pelo build. Custo medido: **1,1 s** de recálculo com
164 chamadas de UDF.

**Sigma baixo não reprova corrida.** A classificação qualifica o **método**; quem
reprova **corrida** é Westgard. O artigo, p. 8–9, trata Sigma baixo como exigência
de *mais* regras, limites mais estreitos e CQ mais frequente — *"For Three Sigma
methods and lower, however, QC frequency must be greatly increased"* —, não como
motivo de rejeição. A prova 6 confere que nenhum status de corrida lê `L` ou `M`.

**Cinco faixas, e não as seis do artigo.** A p. 6 descreve seis zonas: *World
Class* (≥6), *Excellent* (5–6), *Good* (4–5), *Marginal* (3–4), *Poor* (2–3) e
*unacceptable* (<2). O produto une as duas últimas sob **Inadequado**, porque
abaixo de 3 Sigma a conduta operacional é a mesma. É decisão de produto, não do
artigo, e está anotada em `mQualidade`: separá-las é trocar uma linha.

**`IsNumeric(Empty)` devolve `True`.** E `CDbl(Empty)` devolve `0`. Sem guarda de
`IsEmpty`/`IsNull`, célula vazia viraria Sigma 0 → *"Inadequado"*, e margem
ausente → *"ETp excedido"*: reprovação inventada em cima de nada. É a terceira
aparição da mesma armadilha no projeto, depois de `AlvoDoLote` e `BiasEQ`.

**Sem EP não vira zero.** `G` mostra `SEM EP` (texto) e `H`, `L`, `M`, `N`, `O`,
`P` ficam **vazias**. Zero seria exatidão inventada.

**Margem zero é crítica, não excedida.** A faixa é `< 0` → excedido; `0 ≤ m ≤ 10`
→ crítica. Um método que consome exatamente todo o orçamento ainda não estourou.

**BI: 60 → 65 colunas.** `Bias_Observado_abs_pct`, `Classificacao_Sigma`,
`Margem_ETp_pp`, `Margem_ETp_pct`, `Status_Margem_ETp`. A coluna 51 continua
publicando o bias **assinado**; a magnitude ganhou coluna própria porque uma
medida que faça `AVERAGE` sobre bias assinado cancela desvios opostos e devolve
um viés falso perto de zero — o mesmo motivo que levou `mCEQ` a consolidar
rodadas por magnitude.

**O que NÃO mudou.** Lógica de Westgard, banco histórico, estrutura append-only e
os registros históricos de EP.

**Correção de uma leitura errada minha.** Eu havia registrado que `Calc!N` e
`Calc!O` (4-1s e 8x) eram `IF(OR(FALSE;FALSE);1;0)` — sempre zero — e cheguei a
gravar isso como nota no `Painel`. É falso, e a origem do erro vale mais do que o
erro: eu conferi **a linha 3**, que é a primeira do bloco. Numa regra sequencial,
a primeira linha é o único lugar onde ela **não pode** disparar, porque não existe
ponto anterior; os termos degeneram para `FALSE`. Da linha 4 em diante as fórmulas
crescem e as regras funcionam: nas 180 linhas do `Calc`, **4-1s dispara 45 vezes e
8x dispara 52**. As cinco regras estão implementadas. A nota do `Painel` foi
corrigida para dizer o que é verdade: as primeiras linhas de cada lote não têm
histórico suficiente para 4-1s e 8x.

**O eco do filtro de EP passa a contar linhas.** O arquivo estava com o filtro em
`Controllab`, provedor sem uma única linha na `EQC_Dados` (que tem 90 linhas, todas
`CAP`/2025, 6 analitos). O resultado eram 80 células `SEM EP` sem explicação
visível. `Estatística!K5` agora termina com *"N linha(s) no banco de EP"* — zero
para `Controllab`, 90 para `CAP`.

---

## ADR-034 — Controle externo: interface separada, motor consolidado

**Contexto.** A `EQC_Dados` era uma aba só, com CAP e Controllab dividindo as
mesmas colunas, uma linha por analito-rodada e nenhuma rastreabilidade até o
laudo. Não respondia às perguntas que importam: *o desvio foi isolado ou as
cinco amostras andaram juntas? de que página do PDF saiu esse valor? o CAP e o
Controllab estão dizendo a mesma coisa?*

**Decisão — sete abas.**

| Aba | Papel |
|---|---|
| `EQA.CAP_Dados` | digitação do CAP, tabela `tblEQA_CAP_Dados` |
| `EQA.CAP_Resumo` | indicadores + matriz analito × survey, **dinâmica** |
| `EQA.CAP_Nao_Aceitaveis` | derivada por fórmula, só `Unacceptable` |
| `EQA.Controllab_Dados` | digitação do Controllab, `tblEQA_Controllab_Dados` |
| `EQA.Controllab_Resumo` | os mesmos indicadores, no vocabulário do provedor |
| `EQA.Controllab_Desvios` | derivada, critério do próprio Controllab |
| `EQA_Base` | consolidada, **veryHidden**, `tblEQA_Base` |

As seis visíveis ficam agrupadas e nessa ordem, com guia azul (`#263B4D`) para o
CAP e verde (`#1F6F5C`) para o Controllab. As duas abas de digitação são
**gêmeas**: mesmas 18 colunas, mesmas larguras, mesmos formatos — o QA compara
coluna a coluna e não achou diferença. A seção 5 da missão proíbe um provedor
sofisticado e o outro simplificado.

**Granularidade por amostra, não por analito.** Amostra 1 com +8% e amostra 2
com −8% dão média zero e descrevem um método que ninguém aprovaria. Cada
resultado, cada alvo, cada SDI e cada avaliação ficam guardados; a consolidação
acontece depois, conforme a finalidade.

**Um só ponto lê a aba de EP.** `mCEQ` era o único consumidor direto da
`EQC_Dados`. Repontá-lo para a `EQA_Base` fez as **403 células** da `Estatística`
e do `Painel` e a coluna de bias do BI acompanharem sem uma edição de fórmula.
Foi a razão de a migração caber num commit.

**Analito do provedor ≠ analito canônico.** O CAP reporta `Urea Nitrogen`; a
pasta chama `Ureia`. Guardar só um dos dois destrói alguma coisa: só o do
provedor impede o cruzamento com a `Analitos`; só o canônico apaga a
rastreabilidade até o PDF. A base guarda **os dois**, e o mapa entre eles vive em
`EQA_Base!W:Y`, editável.

`Ferritin` e `Thyroid Stim Hormone` não existem na Bioquímica desta pasta: entram
com canônico vazio — visíveis, contados (20 linhas), rastreáveis e fora da
`Estatística`, porque não há especificação de qualidade para cruzar. Inventar
correspondente seria pior do que a lacuna.

**BUN e ureia não são a mesma escala** (fator 2,14), mas o Bias em **porcentagem**
não muda: resultado e alvo estão na mesma unidade e o fator cancela na razão. SD,
SDI e limites ficam na escala do provedor, que é onde são usados.

**`Uso_Analitico`: preservar não é misturar.** Os 90 registros que vinham da
`EQC_Dados` **não são resultado real de EP** — glicose entre 250 e 262 mg/dL nas
quinze amostras, com quatro casas decimais, e **nenhum valor em comum** com o
relatório do CAP. São simulação. Apagar violaria *"não apagar histórico"*;
misturar contaminaria todo Sigma e todo ET da pasta. A coluna existe para não ser
preciso escolher: a linha fica visível e contada, com `NAO`, e não entra em conta
nenhuma. A prova 12 mede os dois números — 2,733% só com o dado real contra
1,843% se misturasse.

**`EQC_Dados` continua na pasta**, intacta, oculta, marcada como legado. As
dependências foram levantadas antes (403 fórmulas + `mBI`), e todas passam por
`mCEQ`.

**Chave duplicada é contada, não descartada.** `Provedor|Ano|Rodada|Analito|Amostra`.
Sumir com duplicata em silêncio esconderia digitação dobrada — exatamente o que a
chave existe para revelar. O carimbo em `Z1` relata o número.

**Status padronizado não aprova por omissão.** Termo desconhecido vira
`NAO AVALIADO`, nunca `ACEITO`.

**EQA não é Westgard.** `Unacceptable` no CAP não reprova corrida de CQI. São
dimensões diferentes: Westgard controla a corrida, o EQA avalia o desempenho
analítico contra pares.

### O que a construção custou aprender

**Argumento nomeado não vinculou neste dispatch.** `Worksheets.Add(After=x)`
ignorou a âncora, e pior: `Move(After=x)` mandou a aba para uma **pasta nova** —
que é o que `Move` faz sem `Before` nem `After`. A aba sumiu do arquivo sem erro
nenhum. Só o primeiro parâmetro posicional chega intacto, então a ordenação usa
`Move(Before)`.

**`Dim ws` sombreia a função `Ws()`.** VBA não distingue maiúsculas: dentro de
`CarregarMapa`, `Set ws = Ws(nome)` chamava o membro padrão da própria variável
vazia — **erro 91**, em caixa modal, numa instância sem tela. A função virou
`AbaPorNome`.

**UDF dentro de `SUMPRODUCT` recebe array.** `PadronizarStatus(faixa)` não é
chamada célula a célula: recebe a faixa inteira, `CStr(array)` dá erro de tipo e
o resumo vira `#VALOR!`. O status virou **coluna** (`W`), calculada uma vez por
linha — e os resumos passaram a usar `COUNTIFS`, que ainda é mais rápido.

**Ordinal por `SUMPRODUCT` é O(n³).** `SUMPRODUCT((COUNTIF(faixa;faixa)=1)*1)` faz
n×n comparações por linha. Trocado por `MAX` da própria coluna até a linha
anterior, mais um.

**`NumberFormat` comportou-se como `NumberFormatLocal`.** Num Excel pt-BR o ponto
é separador de **milhar**, então `"0.00"` foi lido como *milhar, sem decimais*.
Oito colunas nasceram com casas erradas: `-2,12` exibia `-002` e o SDI `-0,9`
exibia `-01`. Só apareceu porque o QA visual mede o que a célula **exibe**, não o
código guardado. Corrigido via `NumberFormatLocal` com vírgula decimal.

**Dezesseis abas estavam protegidas**, apesar de o ADR-031 declarar a pasta
liberada. Descoberto porque `Validation.Add` e `Font.Bold` foram recusados. A
etapa mede o estado, registra quais estavam protegidas e desprotege — a fase de
desenvolvimento pede a pasta aberta.

**`mCEQ` não estava na lista de módulos do build.** O artefato sairia sem ele e as
403 células virariam `#NOME?` — a mesma falha que o `mBanco` pagou no ADR-025.
`mEQA` e `mCEQ` entraram na lista.

---

## ADR-035 — Do Sigma até a decisão operacional: DPM, rendimento e plano de CQ

**Contexto.** O `Painel` mostrava `Sigma = 5,4` e parava aí. Um número isolado
não decide nada: o que decide é a consequência — quantas regras rodar, quantos
controles medir, quantos pacientes podem passar entre um evento de CQ e o
próximo.

**Decisão — a cadeia completa, calculada uma vez.**

```
CV + |Bias| + ETp → Sigma → Classificação → DPM teórico → Rendimento teórico
                                          → Regras → N → Run Size
```

`mPlanoQC` (novo, versionado) é o único lugar onde essa tradução acontece.
`Estatística`, `Painel` e Power BI chamam as mesmas funções.

**DPM é contínuo, não lookup.**

```
DPM = [1 − Φ(Sigma − 1,5)] × 1.000.000
```

Sigma real é 4,27, não 4. Arredondar antes de converter jogaria fora a
resolução do indicador: `DPM(4,27) = 2.802,8`, entre `DPM(4,5) = 1.349,9` e
`DPM(4,0) = 6.209,7`. As nove provas conferem contra a tabela publicada em
Westgard *et al.*, 2018 — 3,4 / 32 / 233 / 1.350 / 6.210 / 22.750 / 66.807 /
158.655 / 308.538.

**DPM é benchmark teórico, não contagem de erro.** A conversão usa short-term
Sigma com deslocamento de 1,5 SD. A nota aparece na `Estatística`, no `Painel` e
na `Cfg_PlanoQC`, e não deve ser removida: sem ela, "66.807 DPM" lê-se como
"o laboratório errou 66.807 resultados", que é falso.

**O plano vem de uma tabela.** `tblPlanoQC_Sigma` (aba `Cfg_PlanoQC`):

| Sigma | Classificação | Regras | N | Run Size |
|---|---|---|---|---|
| ≥ 6 | Classe mundial | `1_3s` | 2 | 1000 |
| 5 – 6 | Excelente | `1_3s / 2_2s / R_4s` | 2 | 450 |
| 4 – 5 | Bom | `1_3s / 2_2s / R_4s / 4_1s` | 4 | 200 |
| 3 – 4 | Marginal | `1_3s / 2_2s / R_4s / 4_1s / 8x` | 6 | 45 |
| < 3 | Desempenho inadequado | — | — | — |

Mudar uma faixa é editar uma linha. `IF`s encadeados em três abas divergiriam no
primeiro esquecimento — e divergiram: ver abaixo.

**A regra de sequência do produto é 8x — decisão do laboratório.** A família
6x / 8x / 10x responde à mesma pergunta: quantos resultados consecutivos do
mesmo lado da média denunciam desvio sistemático. Tabelas publicadas de Sigma
rules trazem ora 6x, ora 8x. O QC_INI opera com 8x, que é o que o motor do
`Calc` avalia, e a tabela recomenda 8x. Por isso `Cobertura_Motor_Westgard`
fecha **TOTAL**: recomendação e motor falam da mesma regra.

**Abaixo de 3 Sigma não se atribui N nem run size.** Preencher ali sugeriria
existir plano de CQ estatístico capaz de sustentar o método. O campo traz a
orientação: investigar e melhorar o desempenho analítico, ou reavaliar o método.

**N não é nível de controle.** É o número **total** de medições no evento. Como
distribuir entre níveis, materiais e replicatas depende da configuração real do
laboratório, e o sistema não presume.

**Run Size não é `R_4s`.** `R_4s` é regra; run size é quantos pacientes passam
entre eventos de CQ. A interface nunca abrevia os dois junto.

**O Painel não recalcula nada.** Lê a `Estatística` por `INDEX/MATCH` na chave
`analito|nível`. É o que faz o card e a célula coincidirem — provado campo a
campo. A consequência está dita na tela: o Sigma do Painel responde ao período e
ao filtro de EQA definidos na `Estatística`; o filtro de datas do Painel manda no
gráfico e nos descritivos.

**Consolidação de |Bias| em duas etapas.** Antes era média simples de todas as
amostras juntas, o que dava peso maior à rodada que por acaso teve mais
amostras. Agora: média dos `|bias|` **por rodada**, depois média das médias.
Com C-A de 2 amostras (`+4%`, `+6%`) e C-B de 1 (`+20%`), a média simples daria
10,00 e as duas etapas dão **12,50**. O modo `DETALHE` publica a memória de
cálculo rodada a rodada.

### O que as provas encontraram

**`Empty` de UDF renderiza como 0 — quarta vez.** `PlanoQC` devolvia o conteúdo
da célula vazia das faixas abaixo de 3 Sigma; a tela mostrava **`N = 0` e
`Run Size = 0`**, que não é "não há plano automático" e sim "rode zero
controles". A conversão para `""` agora mora na saída da função, e a prova 6c
olha o **texto exibido**, não o valor devolvido ao Python — foi assim que o
defeito escapou da prova anterior.

**Duas etiquetas para a mesma faixa.** `mQualidade.ClassificarSigma` dizia
`Inadequado` e `tblPlanoQC_Sigma` dizia `Desempenho inadequado`. O `COUNTIF` do
resumo contava zero, e o painel diria que nenhum analito está inadequado
enquanto vários estão. `ClassificarSigma` passou a **ler a tabela**; a escada de
`IF`s continua como reserva para o caso de a `Cfg_PlanoQC` faltar, e a prova 6b
confere que as duas concordam em todas as fronteiras.

**Benchmark publicado vem arredondado.** Sigma 5,5 dá 31,671 pela conta exata e
a tabela imprime 32 — 1,03% de diferença. A tolerância de 1% reprovava a
matemática correta por causa da precisão com que o artigo imprimiu o número.

**`Norm_S_Dist` em vez de fórmula própria.** A Φ vem da mesma função que o Excel
expõe ao usuário: assim o número da célula e o número do VBA não podem divergir.
`NormSDist` fica como reserva para instalações anteriores a 2010.

---

## ADR-036 — O layout do Painel é do gestor; a automação encosta só onde foi chamada

**Contexto.** O gestor reorganizou a aba `Painel` à mão: moveu os blocos da
coluna `J` para a `R`, redistribuiu, mudou cores. Essa organização está aprovada
e passa a ser a **base oficial**. No meio da reorganização, as fórmulas de status
das regras de Westgard se perderam, e a nova posição definida para elas é `M7`
(QC nível 1) e `M8` (QC nível 2).

**Decisão.** Intervenção cirúrgica: escrever `M7` e `M8`, criar duas colunas de
rastreabilidade na `Estatística`, e **provar** que nada mais mudou.

### A lógica restaurada foi localizada, não inventada

No commit `07f3ff1` o status vivia em `Painel!S7`/`S8`:

```
=IF($B7="","",IF($R7>0,"REPROVA — "&$R7&" violação(ões)","Sem violação"))
```

`$B7` = n de resultados do nível; `$R7` = total das cinco regras. No layout novo
o bloco Westgard ocupa `F6:M8` e o `Total` desceu de `R` para `L`. A fórmula é a
mesma; muda a referência da coluna.

**Defeito encontrado ao auditar a lógica original.** O guarda `$B7=""` nunca
dispara: `B7` é `AGGREGATE(2;6;…)`, uma **contagem**, que devolve **zero** quando
não há resultado no período — nunca texto vazio. Com o período vazio a célula
exibia *"Sem violação"*, que se lê como "está tudo certo" onde não há dado
nenhum. O guarda passou a testar o zero e a resposta virou *"sem dados no
período"*.

### De onde M7/M8 bebem — e de onde não

```
Calc!K..O   × Calc!D  →  violações do NÍVEL 1 no filtro de período
Calc!AG..AK × Calc!D  →  violações do NÍVEL 2
Painel!L7, L8         →  soma das cinco regras
Painel!B7, B8         →  n de resultados do nível
```

Nenhuma dessas fontes lê Sigma, classificação, DPM, rendimento, run size ou
margem de ETp. Provado por leitura da fórmula **e** por cenário: com Sigma
`−1,80`, DPM `999.517` e margem *"ETp excedido"*, `M7` e `M8` continuaram
*"Sem violação"*.

### Rastreabilidade por analito

`Estatística!AC` e `AD` — `N EQA Resultados` e `N EQA Rodadas`, por analito e por
recorte. O eco em `K5` continua informando o total global do filtro (455), que é
outra pergunta: quantos registros existem no recorte, não quantos entraram no
`|Bias|` daquele analito.

### Como se prova que o layout não foi tocado

Comparar o índice de estilo (`s="123"`) **não funciona**: ao salvar, o Excel
renumera a tabela `cellXfs`, remove estilos sem uso e recompacta o resto. Duas
gravações do mesmo arquivo saem com índices diferentes apontando para o mesmo
visual — a primeira tentativa acusou milhares de mudanças inexistentes.

`comparar_layout_painel.py` resolve cada índice até o **conteúdo** — fonte,
preenchimento, borda, formato numérico, alinhamento — e compara as tuplas. Mais
largura de coluna, altura de linha, células mescladas, formatação condicional e
posição de cada objeto. Contra o checkpoint do gestor: **1794 células, zero
diferenças**.

### Testes que leem por nome, não por coordenada

Localizar bloco por coordenada quebrou três vezes nesta sessão: quando o bloco
desceu duas linhas, quando foi para baixo dos gráficos, e quando o gestor o moveu
de `J` para `R`. A forma que sobrevive: achar o título em qualquer coluna, achar
a linha de cabeçalho abaixo dele, montar `{texto do cabeçalho: coluna}` e ler por
nome. Coluna que o gestor tirar do bloco vira *"não publicado neste bloco"*, e
não falha.

Uma âncora precisa ser **exata**, não "contém": `"DPM teórico"` também é
cabeçalho do bloco Six Sigma, e a busca por trecho casava com ele antes de chegar
à tabela de referência. `"Rendimento %"` só existe num lugar.

---

## ADR-037 — O Sigma acende as regras que ele exige

**Contexto.** O bloco Westgard mostrava contagens de violação. Não dizia *quais
regras deveriam estar rodando* para o desempenho daquele analito.

**Decisão.** As cinco células de regra (`Painel!G6:K6`) ganharam formatação
condicional dirigida pelo Sigma. Nada foi movido, redesenhado ou recolorido: as
condições existentes em `G7:K9` continuam intactas.

### Duas semânticas, duas cores, uma prioridade

| | |
|---|---|
| verde escuro `#146C43`, fonte branca, negrito | regra **recomendada** pelo Sigma |
| vermelho (codificação já existente) | regra **violada** na corrida |

Uma regra pode ser as duas coisas. A condição de violação entra com prioridade
**menor** (7 contra 8, 9 contra 10, …), então vence: regra violada não fica
verde só porque está no plano.

### A matriz é uma leitura do texto, não uma cópia dele

`tblPlanoQC_Sigma` ganhou sete colunas — uma por regra da família
`1-3S · 2-2S · R4S · 4-1S · 8X` — com a bandeira

```
=IF(ISNUMBER(SEARCH("1_3s";IF($D4="";$D$1;$D4)));1;0)
```

Cada bandeira **lê a coluna `Regras` da própria linha**. Recomendação textual e
realce visual não podem divergir: são a mesma informação lida de dois jeitos.
Mudar uma faixa muda os dois.

### `SUMPRODUCT`, e não `INDEX/MATCH`

Sondado célula a célula: a formatação condicional **aceita** `SUMPRODUCT` sobre
nomes que apontam para outra aba e **recusa** `INDEX/MATCH` sobre os mesmos
nomes — inclusive dentro de `IFERROR`. São funções que devolvem *referência*, e
a CF não atravessa planilha com elas. A fórmula final é

```
=SUMPRODUCT((regrasRotulos=G6)*regrasAtivas)=1
```

### O plano cobre o nível pior

O bloco tem uma linha por nível, mas a célula da regra é uma só. O realce usa o
**menor Sigma dos dois níveis**.

Isso produziu uma incoerência de leitura que o QA pegou: Lactato exibia
*"Classe mundial"* no card (nível 1) com as cinco regras acesas (nível 2, Sigma
1,82). Cada metade estava certa; o conjunto, contraditório. O texto de apoio
passou a nomear o número e o nível — *"Base: menor Sigma entre os dois níveis =
1,82 (nível 2). O plano precisa cobrir o nível que está pior"*.

### Abaixo de 3 Sigma

A faixa não tem regras na tabela — de propósito, porque não existe plano
estatístico que sustente o método ali. Para o realce, a bandeira cai no conjunto
que o **motor** suporta (estratégia intensiva), a tabela devolve
`Status_Plano_QC = REAVALIAR MÉTODO`, o Painel exibe o alerta, e **nenhum run
size é atribuído**. Não é prescrição de que aquelas regras resolvem; é o máximo
que o CQ estatístico alcança enquanto o método não melhora.

### O conflito 6x / 8x, declarado

A especificação desta missão pede `6x` na faixa 3–4 Sigma. O gestor havia
determinado antes: *"devemos decidir entre 6x, 8x ou 10x e nós já decidimos —
irá ser 8x"*. A mesma missão exige que o realce saia da **mesma fonte** da
recomendação textual. As duas coisas só fecham com 8x, que é o que a tabela diz
e o que o motor executa — e por isso `Cobertura_Motor_Westgard` fecha `TOTAL`.
*(Fechado no ADR-038: `6x` e `10x` saíram do projeto; a matriz tem exatamente as cinco regras.)*

### Achado colateral, não corrigido

`Painel!I7:I9` carrega condições `cellIs` com limiares `4` e `3`, e `J7:J9`
compara com `"OK"`/`"REJEITADO"` — sobras de um layout em que `I` era Sigma e
`J` era Status. Hoje `I` é a contagem de R4S e `J` a de 4-1S, então essas
condições podem colorir a contagem por um critério que não é dela. Não foram
removidas porque remover é mudança visual, e o layout é do gestor.

---

## ADR-038 — 8x é a regra sequencial; o pior nível governa o plano

**Três decisões fechadas pelo gestor**, sem margem: a família sequencial é `8x` e
só ela; o plano de CQ é governado pelo **pior** nível; área visual sem função
pode ser removida ou reconectada.

### `6x` e `10x` saíram do projeto

| Onde estavam | O que foi feito |
|---|---|
| `Cfg_PlanoQC` colunas `O`/`P` da matriz | removidas — a matriz tem exatamente cinco colunas |
| `mPlanoQC.bas` (comentários) | reescritos: `8x` é a definitiva, sem discussão de família |
| `montar_plano_qc.py` (comentários) | idem |
| `realce_regras_westgard.py` (`MATRIZ`) | cinco pares |
| `testar_realce_regras.py` | expectativas ajustadas |
| `ARCHITECTURE_DECISIONS.md` (ADR-037) | afirmações que ficaram falsas corrigidas |

Os ADRs anteriores **não** foram reescritos: eles registram por que a decisão foi
tomada, e apagar isso destruiria a auditabilidade que este projeto usa como
método. Só as frases que passaram a ser factualmente falsas foram corrigidas.

### O defeito que a governança pelo pior nível corrige

**Lactato** tem Sigma `6,99` no nível 1 e `1,83` no nível 2. O bloco de plano
exibia a linha do nível 1:

```
1_3s   ·   N = 2   ·   Run Size máx = 1000 pacientes
```

O CQ mais leve que existe, num analito cujo nível 2 não sustenta nem 3 Sigma. Um
plano lido assim autoriza mil pacientes entre eventos de controle.

```
Sigma_Plano = MIN(Sigma_N1; Sigma_N2)     entre os níveis válidos
```

Com um só nível válido, usa o disponível. Sem nenhum, `SEM DADOS`. Isso alimenta
regras, N, run size, frequência **e** o realce — uma decisão, um número.

O bloco passou a mostrar **um** plano:

| | |
|---|---|
| `R22` `Plano` | Sigma do pior nível + regras + N + run size + frequência + cobertura |
| `R23` `Base` | *"Pior nível governa: nível 2 (Sigma 1,82). N1 = 6,99 · N2 = 1,83"* |

### Destino dos blocos auditados

**`R15:R18` — ERRO TOTAL vs ETp:** estava **conectado e correto**. Aparecia vazio
porque o analito selecionado no momento não tinha Sigma. Com Lactato mostra
`N1 ET 5,65 / ETp 15,21 / margem 62,83% / Dentro do orçamento` e `N2 ET 14,01 /
margem 7,87% / Margem crítica` — e a diferença entre os dois níveis é exatamente
a informação que o bloco existe para dar. **Mantido.**

**`R21:X24` — PLANO DE CQ:** tinha propósito e estava conectado, mas ao nível
errado. **Reconectado** ao pior nível.

### Realce

Verde escuro, fonte branca, **negrito e itálico**. Prioridade inalterada:
violação vence recomendação.

### Coerência que faltava

Abaixo de 3 Sigma o realce acendia as cinco regras enquanto a coluna `Regras`
ficava vazia — cor e texto discordando na mesma tela. A faixa passou a declarar
`1_3s / 2_2s / R_4s / 4_1s / 8x`. `N` e run size **continuam vazios**: não existe
run size tabelado abaixo de 3 Sigma, e inventar um seria autorizar intervalo de
CQ que o método não sustenta.

### Duas armadilhas da construção

**Escrever em `Cfg!A3`/`B3` destrói a tabela.** São os cabeçalhos `Sigma_Min` e
`Sigma_Max` de `tblPlanoQC_Sigma`. O Painel passou a exibir *"Governa o PIOR
nível: 0"* — o Excel devolveu zero no lugar do texto. A célula do nível que
governa foi para `C2`/`D2`, acima da tabela.

**A proteção volta sozinha.** Nesta etapa, **25 das 25 abas** estavam protegidas
de novo, e `Range.Clear` devolveu `0x800A03EC` no meio da alteração. O estado
passa a ser medido e registrado antes de qualquer escrita.

### Layout

Contra o checkpoint do gestor: **1798 células, zero diferenças** de estilo,
largura, altura, mesclagem, formatação condicional e posição de objeto. Só o
conteúdo de `R22`, `R23`, `S22` e `F10` mudou — a reconexão pedida.

---

## ADR-039 — A formatação condicional fala português, e aceitava o inglês calada

**Data:** 2026-08-23 · **Status:** aceito · **Origem:** validação final integrada

### O defeito

`FormatConditions.Add(..., Formula1)` nesta instalação se comporta como
`FormulaLocal`: espera os nomes de função no idioma do Excel. Escrita em inglês,
a condição é **aceita sem erro e nunca avalia verdadeira**.

Consequência: o realce inteiro de `G6:K6` estava **morto desde o ADR-037**. Nem
a recomendação (`SUMPRODUCT`) nem a violação (`SUM`) pintavam. O que a tela
mostrava era só o preenchimento base.

### Por que passou pelos testes anteriores

Porque os testes anteriores **liam a fórmula da condição e a prioridade dela** e
concluíam por raciocínio qual venceria. Nenhum media o que o Excel realmente
pintava. Uma condição sintaticamente presente, com a prioridade certa, no
`AppliesTo` certo, apontando para nomes que resolvem corretamente — e inerte.

`Range.DisplayFormat` devolve o formato **efetivo**, já resolvida a CF. É a única
leitura honesta. Passou a ser obrigatória em qualquer prova de cor.

### Como foi isolado

Uma condição por vez na mesma célula, medindo `DisplayFormat` depois de cada uma:

| Fórmula | Resultado |
|---|---|
| `=TRUE` | não pintou |
| `=$G$6="1-3S"` | **pintou** |
| `=SUM($G$7:$G$8)=0` | não pintou |
| `=sigmaDoPlano>=5` | **pintou** |
| `=SUMPRODUCT((regrasRotulos=G6)*regrasAtivas)=1` | não pintou |
| `=INDEX(regrasAtivas,1)=1` | recusada |

O que pintava não tinha nome de função. Refeito em pt-BR — `VERDADEIRO()`,
`SOMA`, `ÍNDICE`, `SOMARPRODUTO`, separador `;` — **os oito construtos pintaram**,
inclusive o `ÍNDICE` que em inglês era recusado.

### A teoria anterior estava errada

O ADR-037 registrou que *"a CF não atravessa planilha com funções que devolvem
referência"*, porque `INDEX` era recusado e `SUMPRODUCT` era aceito. Não era isso.
Era idioma. O comentário no script foi corrigido para não induzir o mesmo
diagnóstico errado de novo.

### Correção

`_arquitetura/scripts_fase3/corrigir_cf_idioma.py`:

- `=SOMA($G$7:$G$8)>0` — violação
- `=SOMARPRODUTO((regrasRotulos=G6)*regrasAtivas)=1` — recomendação

Os construtores `padronizar_8x_pior_nivel.py` e `realce_regras_westgard.py` foram
corrigidos junto, para que uma re-execução não reintroduza o inglês.

### Um segundo caso, mesma causa

`Estatística!Z3` — o aviso *"BASE DESATUALIZADA"* — tinha sido **recusado** no
ADR-034 usando `LEFT`, e ficou registrado como "sem cor, o texto continua". Era o
mesmo defeito. Com `ESQUERDA` foi aceito e pinta.

### O padrão

É a terceira vez que a fronteira COM↔Excel troca o protocolo pelo idioma local:

| Propriedade | Comportamento real | Sintoma |
|---|---|---|
| `NumberFormat` | `NumberFormatLocal` | `2,70` exibido como `003` |
| `FormatConditions.Formula1` | `FormulaLocal` | condição aceita e inerte |

Nenhum dos dois deu erro. Ambos falharam **em silêncio, na renderização**. A
regra que fica: nesta fronteira, **escrever em português e provar pelo que a tela
mostra**, nunca pelo que o objeto aceitou.

### Por que o `openpyxl` nunca acusaria isso

O arquivo guarda a fórmula **sempre em inglês**. Depois da correção escrita em
português via COM, o `.xlsm` contém:

```
G6  prio=57  SUM($G$7:$G$8)>0
G6  prio=58  SUMPRODUCT((regrasRotulos=G6)*regrasAtivas)=1
```

A tradução acontece **na entrada do COM**, não no formato de arquivo. Ler o
`.xlsm` mostra inglês tanto para a versão morta quanto para a viva — são
idênticas em disco. Só `DisplayFormat`, com o Excel aberto, separa uma da outra.

### Um susto que não era

Comparando contra a base oficial, `Painel!O3` sumiu da lista de formatação
condicional. Não sumiu: o comparador trunca a amostra em 8 itens. Verificado
pelo `openpyxl` nas três versões — a condição `OR($G$3<>"";$G$4<>"")` do aviso
de filtro de data está intacta. O comparador passou a declarar o total ao lado
da amostra.

### Um defeito no teste, não no sistema

O bloco C da validação esperava *"verde = toda regra recomendada pelo Sigma"*.
Isso só vale sem violação — a condição que o bloco D cria de propósito. Com dado
real, violação tem prioridade sobre recomendação, então uma regra recomendada **e
violada** aparece vermelha.

Medido célula a célula em cinco analitos reais (`sonda_bloco_c.py`), **25 de 25
células** se comportam como projetado. Exemplos:

| Analito | Sigma_Plano | Recomendadas | Verdes | Vermelhas |
|---|---|---|---|---|
| Ácido úrico | 6,12 | `1-3S` | `1-3S` | `2-2S` `4-1S` `8X` |
| Cálcio | 4,12 | `1-3S` `2-2S` `R4S` `4-1S` | `1-3S` `R4S` | `2-2S` `4-1S` `8X` |
| Lactato | 1,82 | as cinco | `R4S` | `1-3S` `2-2S` `4-1S` `8X` |

O `R4S` fica verde em quase todos porque é a única que raramente dispara: é
regra de amplitude entre níveis na mesma corrida.

A expectativa do bloco C foi corrigida para *"verde onde recomendada **e sem
violação**"*, mais uma asserção nova: *"toda regra violada tem de estar
vermelha, recomendada ou não"*.

### Matriz final

| Bloco | Testes | Resultado |
|---|---|---|
| A. Fronteiras exatas do Sigma | 64/64 | PASS |
| B. Pior nível governa | 18/18 | PASS |
| C. Analitos reais (cadeia completa) | 15/15 | PASS |
| D. Realce efetivo de `G6:K6` | 25/25 | PASS |
| E. Violação vence recomendação | 5/5 | PASS |
| F. `M7`/`M8` independentes do plano Sigma | 8/8 | PASS |
| G. `8x` é a única regra sequencial | — | PASS (varredura) |
| H. ET × ETp | 10/10 | PASS |
| I. Plano de CQ | 10/10 | PASS |
| **TOTAL** | **155/155** | **PASS** |

Layout do Painel contra a base oficial do gestor (`3a21910`): zero diferença de
estilo, largura, altura, mesclagem e posição de objeto. As únicas mudanças são as
autorizadas — `F10:F13` (apoio, alerta, legenda, cobertura) e a formatação
condicional de `G6:K6`.

### Fora do escopo, registrado

`qclab/` — o app Streamlit, com banco SQLite próprio, que **não lê o `.xlsm`** —
mantém `6X` e `3-1S` como regras de rejeição em `qc_engine.py`, com testes
próprios em `test_westgard.py`. Diverge do conjunto de cinco regras do ADR-038.
Não foi alterado: esta etapa era validação, e mudar o motor dele exigiria decidir
o que fazer com os testes que hoje cobrem `6X`. Fica para decisão do gestor.

---

## ADR-040 — O pior nível governa o plano de CQ também no BI

**Status:** aceito · **Sucede:** ADR-038

### O defeito

O ADR-038 estabeleceu `Sigma_Plano = MIN` entre os níveis **na planilha**. O
contrato do BI ficou para trás: `mBI` preenchia regras, N e run size a partir do
Sigma da **própria linha** — ou seja, do nível daquela linha.

Lactato tem Sigma 7,14 no nível 1 e 1,86 no nível 2. As linhas do nível 1
publicavam *"1_3s, N=2, run size 1000"*: o CQ mais leve que existe, num analito
cujo nível 2 não sustenta nem 3 Sigma. Mil pacientes entre eventos de controle.
Excel e Power BI diriam coisas diferentes sobre o mesmo analito.

### Por que no motor e não em DAX

A regra é decisão de negócio e o motor é a única camada de cálculo (ADR-019). O
`MIN` entre níveis reimplementado em DAX seria a segunda cópia, que diverge no
primeiro ajuste de faixa — o que o ADR-027 passou uma sessão eliminando.

### Contrato 76 → 84

`Sigma_Plano`, `Nivel_Governante`, `Classificacao_Sigma_Plano`, `Usar_1_3s`,
`Usar_2_2s`, `Usar_R_4s`, `Usar_4_1s`, `Usar_8x`.

As booleanas existem para o BI **não procurar substring** dentro de
`"1_3s / 2_2s / R_4s / 4_1s / 8x"`. Busca ingênua por `1_3s` casaria dentro de
`11_3s`, e no dia em que alguém escrevesse `4-1s` a flag viraria `FALSE` em
silêncio — apagando a regra da tela. `mPlanoQC.RegraNoPlano` compara por token
normalizado, no mesmo módulo que produz a cadeia.

`SigmaValidoBI` **rejeita zero**: Sigma exatamente zero não existe num método que
produziu resultado — ele nasce de ETp, bias e CV. Zero é ausência de dado
disfarçada, e aceitá-lo jogaria o analito em "reavaliar método" por falta de
informação, conclusão diferente de "o método é ruim".

### Recomendada e violada são dimensões separadas

A primeira versão somava `W_*` sobre o histórico inteiro e marcava **todas** as
regras como violadas — com 110 resultados por analito em seis meses qualquer
regra dispara em algum momento, e a tela ficava vermelha sem informar nada.
`[Violada X]` passou a ser da **última corrida**; a contagem do período vive em
`[Violações X no período]`.

E colapsava as duas dimensões num rótulo só: uma regra violada **fora** do plano
aparecia como "VIOLADA", indistinguível de uma regra do plano. Passaram a ser
quatro estados — `RECOMENDADA`, `RECOMENDADA - VIOLADA`, `FORA DO PLANO -
VIOLADA`, `fora do plano` — com cor derivada do mesmo lugar, para cor e texto
nunca discordarem.

### Validação

| Produto | Resultado |
|---|---|
| Bioquímica | 304 PASS, 0 FAIL |
| Hematologia | 147 PASS, 0 FAIL |

| Analito | N1 | N2 | Plano | Governa | Regras |
|---|---|---|---|---|---|
| Lactato | 7,136 | 1,863 | 1,863 | nível 2 | cinco; N e run size **vazios** |
| Ácido úrico | 6,445 | 10,377 | 6,445 | nível 1 | `1_3s`, N=2, Run=1000 |
| Fixação do ferro | 2,656 | 2,209 | 2,209 | nível 2 | cinco |

Fronteiras corretas nos dois sentidos: 2,99 inadequado / 3,00 Marginal; 5,99
Excelente / 6,00 Classe mundial.

### Dois achados alheios à mudança

`montar_plano_qc.py` tinha `Regras = ''` na faixa abaixo de 3 Sigma enquanto a
planilha já declara as cinco — **re-executá-lo reverteria o ADR-038**.

`padronizar_run.ps1` derrubava o build exigindo título de eixo `RUN` nos gráficos
de `EQA.*`, que são indexados por **rodada** do provedor. A exceção é nomeada e
contada no relatório, para não confundir "não se aplica" com "passou".

---

## ADR-041 — Duas matrizes de Westgard, uma fonte de verdade

### Regra de negócio

A seleção das regras é específica da configuração do controle interno.

| Produto | Níveis | Matriz |
|---|---|---|
| Bioquímica | 2 | `1_3s / 2_2s / R_4s / 4_1s / 8x` |
| Hematologia | 3 | `1_3s / 2of3_2s / R_4s / 3_1s / 6x` |

As matrizes **não são intercambiáveis**, e regra da outra família na matriz ativa
é erro de configuração.

### O defeito

Os dois produtos rodavam o **mesmo** código de Westgard — só `NLV` diferia. A
Hematologia, que mede três níveis, era avaliada pelas regras de dois. As faixas
existem para casar com a probabilidade de detecção daquele desenho de controle; a
matriz errada muda sensibilidade e taxa de falsa rejeição **sem produzir nenhum
sintoma visível**.

### 10x corrigido para 8x na Bioquímica

`Cfg_PlanoQC` e `mPlanoQC` declaravam `8x`; o motor contava 10. Nomes batendo,
limiares não — e `CoberturaWestgard` respondendo TOTAL porque só comparava
rótulos.

Efeito medido no banco real: as violações de tendência passaram de **1.965 para
2.374** — 409 sequências de oito resultados do mesmo lado da média que passavam
sem ser denunciadas.

### Semântica, não renomeação

| Regra | O que conta |
|---|---|
| `2of3_2s` | quantos dos três níveis da corrida passam o mesmo 2s |
| `3_1s` | duas leituras: três níveis na corrida, ou três corridas no nível |
| `6x` | seis corridas consecutivas do mesmo lado, no mesmo nível |

A variante consecutiva da `2_2s` foi **desligada** para `NLV=3`: somá-la a
`2of3_2s` faria a Hematologia rodar seis regras em vez de cinco, aumentando a
rejeição falsa.

### Uma fonte

`MatrizWestgard()` decide a partir de `NLV` e publica os nomes. `mPlanoQC` consome
dela **por vínculo tardio** — chamada direta criaria dependência de *compilação*,
e numa pasta sem a função o projeto VBA inteiro deixaria de compilar, o que
`On Error` não captura. A lista fixa `"1_3s;2_2s;R_4s;4_1s;8x"` saiu do `mPlanoQC`.

### Fail-fast

`ValidarMatrizWestgard` confere os **dois** sentidos: regra alheia presente e
regra própria faltando. `CoberturaWestgard` deixou de comparar só nomes — o motor
publica `LimiarSequencialWestgard` e `LimiarUmSigmaWestgard`, e divergência entre
o número que o nome promete e o que o motor conta devolve **erro de cobertura**.

`gerar_mEstatistica.ps1` recortava por **índice fixo** (299, 512, 679, 750, 799).
Acrescentar linhas ao motor deslocava tudo e o build parava. Passou a localizar
por marcador.

### Testes funcionais

Séries de z-score injetadas no motor por uma ponte VBA. Provam o disparo **e o
não-disparo**, que é o que evidencia o limiar.

| Produto | Resultado | Evidência |
|---|---|---|
| Bioquímica | 22 PASS, 0 FAIL | 7 corridas não acendem `8x`; 8 acendem |
| Hematologia | 23 PASS, 0 FAIL | 5 não acendem `6x`; 6 acendem; `2of3_2s` +2,3/+2,4/+0,5 dispara; `3_1s` +1,2/+1,4/+1,1 dispara |

---

## ADR-042 — Motor Westgard por detector e escopo

### O defeito central

O ADR-041 acertou os **nomes** das regras e errou as **janelas**. `8x` era
avaliado como oito *corridas* consecutivas do mesmo lado num único nível — o
detector longitudinal. A matriz Sigma da Bioquímica pede N=2, R=4: dois níveis,
quatro corridas, oito **observações**. Duas regras diferentes dividindo o mesmo
nome. Idem `6x`, que na matriz de três níveis é N=3, R=2.

### Escopos explícitos

`WITHIN_RUN`, `ACROSS_RUN`, `WITHIN_MATERIAL`, `ACROSS_MATERIALS`. Cada violação
declara o seu.

`R_4s` ficou rígida: só dentro da corrida, todos os pares (N1×N2, N1×N3, N2×N3),
com o par normalizado para N1×N3 e N3×N1 não virarem duas violações. Amplitude de
4 DP entre a corrida de ontem e a de hoje **não é `R_4s`** — é deriva, e aplicar
`R_4s` ali inventa rejeição que o método não cometeu. Teste negativo obrigatório
incluído.

### Oficial × complementar

A tabela `DETECTORES` é a única fonte, e `DetectorAtivo()` é a única porta. Os
longitudinais (`SAME_LEVEL_R8/R6/R4/R3`) são calculados e registrados no trace, e
**não consolidados**. Somar longitudinal com N/R multiplicaria as oportunidades de
rejeição e subiria a probabilidade de falsa rejeição sem ninguém decidir.

### `NAO_AVALIAVEL` × `FALSE`

`FALSE` quer dizer *"avaliei e não violou"*. Janela sem os dados que a regra exige
passou a ser contada em `NaoAvaliaveisWestgard`, por detector. Um `6x` N3/R2 com
um nível faltando numa das duas corridas **não é aprovação**.

### Contagens antes × depois

| Regra | Produto | Antes | Depois | Motivo |
|---|---|---|---|---|
| `4_1s` | Bioq. | 1.474 | 1.148 | longitudinal → N2/R2 |
| `8x` | Bioq. | 2.374 | 2.088 | longitudinal → N2/R4 |
| `3_1s` | Hema. | 20 | 6 | longitudinal saiu da decisão |
| `R_4s` | Hema. | 24 | 16 | marca só o par que violou |

A queda contraria a intuição de que janela mais curta dispara mais: N2/R4 é mais
curta **e** mais exigente, porque pede que os **dois níveis concordem** durante as
quatro corridas. Manter um nível do mesmo lado por oito corridas é mais fácil do
que fazer dois concordarem por quatro.

`1_3s`, `2_2s`, `R_4s` (Bioq.), `6x` e alertas: inalterados.

### Testes

Bioquímica 26 PASS · Hematologia 29 PASS, 0 FAIL. Cobrem: `R_4s` nunca cruza
corridas; `2of3_2s` falso e `R_4s` verdadeiro com lados opostos; `6x` N3/R2 com
sequência interrompida não dispara; corrida sem N3 devolve `NAO_AVALIAVEL` e não
`FALSE`.

### Declarações de módulo

`gTrace` e `gNaoAval` tinham ficado no **meio** do arquivo depois da substituição
do motor. Em VBA declaração de módulo vem antes da primeira procedure; o resultado
foi *"variável não definida"* em `gNaoAval`. Compilação conferida com
**Debug > Compile**, não só por leitura.

---

## ADR-043 — Uma escada de classificação Sigma para todo o QC_INI

### Três implementações da mesma coisa

| Onde | O quê |
|---|---|
| `mQualidade.ClassificarSigma` | canônica, lê as faixas de `Cfg_PlanoQC` |
| `mEstatistica.ClassificarSigma` | duplicata, com a escada **errada** |
| célula da Estatística (Hema.) | terceira escada, fixa na fórmula |

A duplicata do `mEstatistica` **vencia** a canônica: `AtualizarEstatisticaAba`
chamava `ClassificarSigma` **sem qualificar**, e chamada não qualificada resolve
para a função do próprio módulo. A coluna de classificação dos **dois** produtos
vinha dali, não de `mQualidade`.

### Três defeitos na escada errada

- `>=6` devolvia *"Excelente"*; a faixa é *"Classe mundial"*
- a faixa 5 a <6 **não existia**: Sigma 5,5 caía em *"Bom"*
- abaixo de 3 dizia *"Inaceitável"*; o projeto diz *"Desempenho inadequado"*

Um método de Sigma 5,5 — excelente — aparecia como apenas bom, e um de 6,7
aparecia como excelente em vez de classe mundial.

### Conflito de arquitetura resolvido

Classificação de **desempenho** e faixa do **plano** são perguntas diferentes e
estavam na mesma tabela. A matriz de três níveis não muda de regra entre 3 e 4
Sigma, então `PLANO_HEMATOLOGIA` tinha quatro faixas — e `ClassificarSigma`, que
lê a classificação dessa tabela, chamaria um método de Sigma 3,5 de *"Desempenho
inadequado"*, divergindo da Bioquímica **para o mesmo número**.

A Hematologia passou a ter cinco faixas: 3 a <4 *"Marginal"* e <3 *"Desempenho
inadequado"*, **ambas com as mesmas cinco regras**. O conjunto de regras não muda;
muda o rótulo e o tratamento de N/R abaixo de 3, onde entra reavaliar método.
Assim continua havendo **uma** escada de classificação, sem tabela paralela.

### Fronteiras provadas nos dois produtos

2,99 Desempenho inadequado · 3,00 e 3,99 Marginal · 4,00 e 4,99 Bom · 5,00 e 5,99
Excelente · 6,00 e 6,01 Classe mundial.

Hematologia: 40 fórmulas trocadas, 9 fronteiras PASS. Bioquímica: 0 trocadas, 80
já centralizadas, 9 fronteiras PASS.

O script **recusa salvar** se qualquer fronteira falhar — e recusou na primeira
execução, quando a tabela de faixas ainda tinha quatro linhas.

---

## ADR-044 — Cobertura confere contrato, e os slots deixam de mentir

### Cobertura estrutural × QA funcional

`CoberturaWestgard` não tem como saber se a suíte de testes passou, e fingir que
sabe faria a função parecer mais forte do que é. As duas garantias ficam separadas
e declaradas:

| Garantia | O que prova |
|---|---|
| `CoberturaWestgard` | **contrato**: regra, detector oficial, ativo, N, R, escopo |
| `testar_westgard_*` | **comportamento**: dispara, não dispara, fronteira, ausência |

### O plano não declara detector

`Cfg_PlanoQC` continua dizendo **quais** regras são necessárias — `8x`, não
`8x / N2_R4`. Fazer o plano conhecer detectores criaria a segunda fonte que esta
série de ADRs existe para eliminar.

A resolução é derivada: o plano pede a regra, `CONTRATO` diz qual é a
implementação oficial dela naquela área, e `DETECTORES` diz o que o motor declara
implementar. Cobertura compara os dois. Isso **não é duplicar a verdade**:
`CONTRATO` é especificação e `DETECTORES` é declaração — o mesmo papel de um teste
que afirma o valor esperado. Se fossem o mesmo texto lido do mesmo lugar, a
conferência não provaria nada.

### Testes negativos reais

`ConferirContrato(regras, tabela)` recebe a tabela **por parâmetro** para o QA
poder passar uma versão mutada. Sem isso os testes negativos seriam encenação —
não há como alterar uma `Const` em execução.

| Produto | Cenários que devem reprovar |
|---|---|
| Bioquímica (5 PASS) | `8x` oficial trocado pelo longitudinal · `8x` com R=8 · `4_1s` sem detector oficial ativo · `R_4s` com escopo `ACROSS_RUN` |
| Hematologia (5 PASS) | `6x` e `3_1s` trocados pelos longitudinais · `6x` com N=2 |

As mensagens são específicas: *"8x: detector oficial é SAME_LEVEL_R8, esperado
N2_R4"*.

### Slots legados renomeados

`r22 → r2sMulti`, `r41 → r1sMulti`, `r10 → rSeq` (e os `q*` correspondentes). 92
ocorrências. São nomes locais e as chamadas são posicionais, então nada muda de
comportamento — o que muda é o código **deixar de mentir**: `r10` se lê como 10x,
e 10x não existe mais em lugar nenhum operacional.

Regressão: Bioquímica 26 PASS · Hematologia 29 PASS, contagens **idênticas**.

`CONTRATO` nasceu no meio do `mPlanoQC` e a pasta parou de compilar — mesma
armadilha de `gNaoAval`.

---

## ADR-045 — `EVENTOS_WESTGARD` é um fato de evento

### O defeito

A aba tinha granularidade de (corrida × analito × nível) com as regras
concatenadas numa célula (`"13s+22s"`), e era produzida por `AvaliarWestgard1N`:
uma **segunda implementação** das regras, escrita para uma série de um nível.

Consequências, todas medidas:

- `R_4s`, `2of3_2s`, `3_1s` e `6x` eram **estruturalmente invisíveis** (exigem
  visão multi-nível);
- a aba materializava *"10 corridas seguidas do mesmo lado"* — o `10x` que o
  ADR-041 aposentou;
- a Hematologia mostrava **9 eventos** onde a camada de BI via **80 marcações**.

`AvaliarWestgard1N` foi **removida**. A aba passou a vir do trace de
`AvaliarWestgard`. Hematologia foi de 9 para **116 eventos**; Bioquímica registra
**7.003** — acima do teto antigo de 5.000, que teria abortado o registro.

`Evid` ganhou `niveis=` e `runIni=`, para o trace ser **fato** e não texto livre.
Sem eles o consumidor teria de adivinhar por arqueologia de string quais
resultados o evento marcou.

### Evento e marca são números diferentes

Um `6x` N3/R2 é **um** evento e marca **seis** resultados. O Painel dizia "3x",
que não distinguia as duas coisas. `MetricasWestgard` devolve os três:

| Medida | O que conta |
|---|---|
| `N_Eventos_Violacao` | quantas vezes a **regra** disparou |
| `N_Resultados_Marcados` | quantos **resultados** ficaram marcados |
| `N_Corridas_Envolvidas` | união das janelas dos eventos oficiais |

Exemplo real da fixture: BASO# N2 = 3 eventos / 2 marcados / 3 corridas.

### Buffer sem teto

A premissa de `buffer_dinamico.ps1` (*"eventos nunca excedem as linhas do
banco"*) morreu com granularidade de evento: dez detectores podem emitir
evidência sobre a mesma corrida. O buffer cresce por `ReDim Preserve`, com a
matriz orientada em **coluna × evento** — VBA só preserva a última dimensão.

### Achado registrado e não remendado

`mEstatPeriodo.AlvoDoLote` varria até a linha 200 do `LotesStore`. O bloco tem 40
linhas por lote (`mBI.LS_CAP`), então do **5º lote em diante devolvia vazio** —
sem Z, sem bias, sem erro. Num sistema dimensionado para 60 meses, chega lá dentro
da vida útil. O teto passou a vir do dado.

A duplicação em si **não foi unificada**: `mBI.AlvoDoLote` lê o mesmo alvo por
aritmética de bloco, `mEstatPeriodo` por varredura. Concordam hoje por
coincidência de duas leituras corretas. Unificar exige mexer na Estatística por
período, que não tem prova automatizada — e `instalar_estat_periodo.ps1` **nem é
chamado pelo `build_all`**, então esse módulo hoje é ingerenciado pelo build.

---

## ADR-046 — Escrita em aba protegida fecha a própria janela

### A causa é uma só, e não é a proteção estar errada

Dois erros 1004 relatados em produção. `ReprotectAll` aplica
`UserInterfaceOnly:=True`, que é a configuração certa, mas **o Excel não persiste
essa flag ao salvar**. Reaberto o arquivo, a aba volta protegida também para o VBA
até `Workbook_Open` rodar `LockApp` — e `Workbook_Open` **não roda em automação**,
porque todo script de build e de QA abre com `EnableEvents = False`.

Logo, "aplicar `UserInterfaceOnly`" não corrige (já é aplicado) e "desproteger"
reduziria a proteção. Cada rotina trata a própria janela, com restauro garantido —
o que `mBI`, `mBanco` e `mImportar` já faziam.

### A auditoria estática achou oito pontos sem guarda, não dois

| Módulo | Rotinas |
|---|---|
| `mEstatistica` | `AtualizarCalc`, `AtualizarPainelEng`, `AtualizarEstatisticaAba`, `RegistrarEventosWestgard` |
| `mAuditoria` | `Auditar` |
| `mConfig` | `SincronizarSombraCfg`, `AuditarMudancaCfg` |
| `mLogDB` | `RegistrarLogDB` |
| `mRegistros` | `MarcarNaoConforme`, `ExcluirRegistroNC` |
| `mImportar` | `MostrarErros`, `LimparAreaImport` |

O par `LiberarEscrita`/`RestaurarProtecao` mora em `mSeguranca`, **não no motor**:
a Imunologia ainda não tem `mEstatistica`, e um primitivo compartilhado ali a
deixaria sem ele. `mEstatistica` ficou com **zero** ocorrências da senha.

`Auditar` e `RegistrarLogDB` mantêm a reproteção própria
(`AllowFiltering`/`AllowSorting`): trocar pelo genérico reduziria permissão sem
aviso.

### Onde a correção quase se perdeu

Três vezes, todas na mesma família — tratar o gerador como se preservasse o que se
escreve:

1. helpers colocados **depois** do banner `MOTOR: MONTAR Calc`, que é a fronteira
   inicial do splice: sumiam do módulo gerado (*"Sub ou Function não definida"*);
2. o comentário de aviso **citava** o banner, e a busca por substring o elegeu
   como marcador — apagou os helpers de novo;
3. correções feitas em `src_hardening1/<Produto>/`, que é **saída** de build:
   `build_all` copia `src_hardening1/*.bas` por cima a cada execução. Compilava,
   passava no teste, e desapareceria no build seguinte **sem sinal**.

### A prova de que a auditoria enxerga

`auditar_vba.py` cobre 10 categorias sobre o projeto **inteiro** exportado do
artefato (65 módulos), porque duplicata entre módulo gerado e módulo legado da
produção só aparece olhando o conjunto.

`testar_auditar_vba.py` injeta um defeito de cada categoria e exige que a
categoria apareça: **11 PASS**. Sem isso *"nada encontrado"* seria indistinguível
de cegueira — e foi cegueira duas vezes:

- `RE_DECL` exigia `Dim`/`Const` e não via `Private gTrace As Object`, que é a
  forma exata do defeito que já quebrou este projeto três vezes;
- `exportar_vba.py` gravava `\r\r\n` (o VBE já devolve `\r\n`) e a releitura em
  modo texto contava cada `\r` solto como quebra: os módulos apareciam com o
  **dobro** de linhas e fronteiras de procedure deslocadas. O *"0 graves"*
  anterior foi calculado sobre lixo.

A verificação de identificador não declarado só entrou depois de sair de **959
para 0** falsos positivos (continuação `_`, intrínsecos, chamadas com `(`,
argumentos nomeados `:=`, `Declare` de API, hexadecimal `&H`, controles de
UserForm que vivem no designer).

### Regressão (artefatos de build limpo)

| Verificação | Bioquímica | Hematologia |
|---|---|---|
| compilação (11 sondas) | 0 FAIL | 0 FAIL |
| auditoria estática | 0 graves | 0 graves |
| Westgard por módulo | 26 PASS | 29 PASS |
| cobertura × contrato | 5 PASS | 5 PASS |
| eventos (grão de evento) | 0 FAIL | 0 FAIL |
| proteção × escrita | 0 FAIL | 0 FAIL |
| Painel exibe o do motor | 0 FAIL | 0 FAIL |
| `Cfg_Westgard_Escopo` | sincronizada | sincronizada |
| pior nível com 3 níveis | — | 8 PASS |

Proteção preservada: `ProtectContents` True, células `Locked`,
`UserInterfaceOnly` reaplicado, estrutura travada.

**O teste de teclado (usuário digitando) continua sendo humano** — escrita por COM
é programática e `UserInterfaceOnly` a autoriza por definição, então usá-la como
prova de bloqueio reprovava código correto.

---

## ADR-047 — O último módulo fora do build, e a correção que nunca embarcou

**Status:** aceito · **Sucede:** ADR-045 · **Mesma família:** ADR-025, ADR-034

### O achado

O ADR-045 registrou a correção do teto do `LotesStore` em
`mEstatPeriodo.AlvoDoLote` e deixou dito, de passagem, que
`instalar_estat_periodo.ps1` *"nem é chamado pelo `build_all`"*. A consequência
não tinha sido medida. Ela é esta: **a correção nunca chegou ao arquivo.**

Comparando `src_producao/mEstatPeriodo.bas` com a cópia dentro de
`QC_Bioquimica.xlsm`:

| | fonte | dentro da pasta |
|---|---|---|
| `AlvoDoLote` | teto de `End(xlUp)` | `For i = 2 To 200` |
| `LimEspec` | presente | **ausente** |

A produção seguiu varrendo até a linha 200. Com 40 linhas por lote
(`mBI.LS_CAP`), isso cobre quatro blocos e meio: **do 5º lote em diante o alvo
cai fora da varredura** e a função devolve vazio — sem Z, sem bias, sem erro.

### Por que passou despercebido por uma série inteira de ADRs

Porque `build_all` **parte da produção e só adiciona módulos**. A cópia velha
que já vivia dentro do arquivo sobrevivia intacta, e não havia `#NOME?` para
denunciar a falta. Um módulo ausente grita; um módulo **desatualizado** é mudo.

É a terceira vez nesta família — `mBanco` (ADR-025) e `mCEQ` (ADR-034) —, e a
única das três que não deu sintoma nenhum.

### O que impedia a importação

`LimEspec` chama `EspecCVtp`/`EspecBIAStp`/`EspecETp`, que vivem no
`mEspecificacoes` — módulo que **não existe em nenhum dos dois produtos**.
Chamada direta cria dependência de **compilação**: numa pasta sem o módulo, o
projeto VBA inteiro deixa de compilar, e `On Error` não captura erro de
compilação. O sintoma é um diálogo modal invisível que trava qualquer automação.

Resolvido pelo vínculo tardio do ADR-041: `Application.Run`. A ausência vira
erro de **execução**, capturado, e `LimEspec` degrada para vazio. `StatusCV` e
`StatusETP` já tratam ausência de limite como **não-aprovação** (ADR-023), então
ninguém se vê aprovado por um critério que não foi avaliado.

### O portão do instalador estava velho em quatro pontos

Ele existia e era bom; tinha envelhecido junto com a planilha.

| Defeito | Sintoma | Correção |
|---|---|---|
| janela `linha 7 a 86` fixa | o cabeçalho desceu para a 13; `[double]"n"` estourava | cabeçalho localizado pelo rótulo `Analito` |
| fim pela última linha da coluna A | invadia o bloco de margem crítica; `[double]"Margem critica"` estourava | tabela termina na **primeira linha vazia** |
| assinatura de 8 argumentos | exclusões viraram um intervalo N×2; COM recusava | 7 argumentos, exclusões como bloco |
| comparação circular | comparava `EstatPeriodo` com células que **chamam** `EstatPeriodo` | referência passa a ser o estado **anterior ao import** |

O quarto é o mais importante. O cabeçalho do script prometia *"conferência
contra o caminho antigo"*, e era verdade na migração: `C..F` guardavam as
fórmulas velhas. Hoje `C14` é `=EstatPeriodo(...)`. O portão parecia provar, e
não provava.

Somou-se um piso: **conferir zero linha não é aprovar**. Sem ele, uma mudança de
layout que fizesse o laço não casar nada passaria como sucesso.

### Falhou fechado, três vezes

Todas as três recusas aconteceram **depois** do import e **antes** do `Save`. O
arquivo ficou intacto em todas — verificado por hash contra o backup. A ordem
importa: importar é barato de desfazer, salvar não é.

### A prova de alcance, e as duas versões dela que mentiam

`testar_alvo_do_lote.py` escreve um lote além da linha 200 e exige que
`AlvoDoLote` o encontre. As duas primeiras versões deram falso resultado:

1. **analito inventado** — `AlvoDoLote` não procura pelo nome no `LotesStore`:
   resolve o nome para um **índice** na aba `Analitos` e procura pelo par
   lote+idx. Nome que não existe ali faz a função sair antes de varrer. A versão
   1 ainda assim marcou OK, porque aceitava "qualquer valor" — e o valor que
   veio era o do lote ativo lá em cima, sem ter tocado na linha 240.
2. **linha isolada em 240** — a varredura sai no primeiro branco (`Exit For`), e
   o vão de 42 a 239 a fazia parar em 42. Reprovava código correto por simular
   um layout que não acontece: no dado real os lotes são blocos contíguos.

A versão que prova: neutraliza o casamento antigo mudando **só a coluna `idx`**
daquela linha — a coluna 1 continua preenchida, então não abre buraco —,
preenche o bloco de forma contígua até 240 e põe o par lote+idx no fim.

```
linha 2 responde hoje          : 4,767727272727272
bloco contiguo de 42 ate 240
varredura alcanca a linha 240  esp 4,321  obt 4,321      OK
desfeito: volta a responder o de antes    obt 4,767727…  OK
```

**6 OK, 0 FALHA.** Nada é salvo: o teste fecha o arquivo sem gravar.

### O VBE reescreve a caixa dos identificadores

A asserção do vínculo tardio reprovou na primeira execução procurando
`Application.Run("EspecCVtp"`. Dentro do arquivo está `Application.run`. É a
mesma normalização que já aparecia no diff como `Str$` → `stR$` e `ws.Rows` →
`ws.rows`: identificador VBA é caso-insensível e o editor reescreve a caixa ao
importar. **Toda busca estrutural sobre código devolvido pelo VBE tem de ser
sem caixa** — a versão sensível reprova código correto.

### A etapa no build

`build_all.ps1`, etapa **7a**, depois do motor e só para a Bioquímica:

- **depois do motor** porque o portão lê `n`, média, DP e CV da Estatística
  antes de importar; com a aba vazia não haveria o que conferir;
- **só a Bioquímica** porque é o único produto cujas células chamam
  `EstatPeriodo` — 323 delas. Na Hematologia, nenhuma.

### O que continua registrado e não remendado

`mBI.AlvoDoLote` lê o **mesmo** alvo por aritmética de bloco; `mEstatPeriodo`,
por varredura. Duas implementações do mesmo contrato de layout, que concordam
hoje por coincidência de duas leituras corretas. Unificar exige mexer na
Estatística por período — que agora tem um portão de regressão, mas ainda não
tem suíte própria. Fica para uma decisão do gestor, não para um remendo aqui.

E `mEspecificacoes` não existe em nenhum dos dois produtos. Enquanto for assim,
`LimEspec` devolve vazio e o status diz `SEM LIMITE` — correto pelo ADR-023, e
declarado pelo portão em vez de silencioso. Ligar o banco de especificações da
Bioquímica é decisão de escopo, não consequência desta correção.
