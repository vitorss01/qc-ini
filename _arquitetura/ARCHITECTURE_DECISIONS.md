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
