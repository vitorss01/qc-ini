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
