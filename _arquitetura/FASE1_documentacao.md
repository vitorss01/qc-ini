# Fase 1 — Refatoração da Arquitetura

**Escopo:** reorganização estrutural sem novas funcionalidades operacionais.
**Arquivos:** `QC_Hematologia.xlsm`, `QC_Bioquimica.xlsm`, `QC_Imunologia.xlsm`
**Backup pré-refatoração:** `_backup_pre_fase1/`

---

## 1. Nova arquitetura adotada

Arquitetura em camadas, com a aba **`Resultados`** como *Single Source of Truth*:

```
Interface        frmCorrida · frmAssinar · frmDev · abas visíveis · Shape btnDev
     ↓
Banco de dados   Resultados  (vertical, normalizado, chave RUN, exclusão lógica)
     ↓
Motor            Calc  (descoberta de corridas, Westgard, limites, Z-score)
     ↓
Apresentação     Painel · Estatística · Liberação · Registros
     ↓
Segurança        login SHA-256 · papéis · assinatura · proteção · Modo Desenvolvedor
```

> **Nota de nomenclatura.** O pedido citava "aba Registros" como fonte oficial. Nesta
> pasta de trabalho a aba de resultados chama-se **`Resultados`**; `Registros` é o log de
> repetições/calibração. A SSOT implementada é `Resultados` — renomear abas quebraria
> fórmulas, nomes definidos e VBA sem ganho algum.

### Camadas do VBA

O monolito `mSystem` (439 linhas, tudo misturado) foi dividido por responsabilidade:

| Módulo | Responsabilidade | Procedures |
|---|---|---|
| `mDados` | Acesso ao banco (SSOT), RUN, gravação, exclusão lógica | `UltimaLinhaBanco`, `LoteAtivoCore`, `NovoRUN`, `GravarCorrida`, `MarcarStatusRUN`, `AtualizarBanco`, `RegistrarLog` |
| `mEstatistica` *(ver nota)* | Consolidado dentro de `mUI` nesta fase | `AtualizarEstatistica` |
| `mUI` | Interface, eixos, gráficos, orquestração | `SystemLook`, `AbrirFormCorrida`, `AtualizarEixos`, `HookCharts`, `AtualizarResultados`, `AtualizarEstatistica`, `AtualizarGraficos`, `AtualizarPainel`, `AtualizarTudo` |
| `mLotes` | Troca de lote e persistência das views por lote | `BlocoDoLote`, `SalvarViewNoBloco`, `CarregarBlocoNaView`, `TrocarLote`, `FlushLoteAtual` |
| `mSeguranca` | SHA-256, login, papéis, assinatura, proteção, Modo Desenvolvedor | `SHA256Hex`, `DoLogin`, `CadastrarUsuario`, `AssinarCom`, `Assinar`, `LockApp`, `UnlockApp`, `UnprotectAll`, `ReprotectAll`, `Logout`, `ModoDesenvolvedor` |

Cadeia de atualização de responsabilidade única:
`AtualizarTudo` → `AtualizarBanco` → `AtualizarResultados` → `AtualizarEstatistica` → `AtualizarPainel` → `AtualizarGraficos`

---

## 2. Alterações realizadas

### Etapa 3-4 — Schema e identificador RUN

Schema anterior → novo (aba `Resultados`):

| Antes | Depois |
|---|---|
| A=Data, B=Nível, C=Seq, D=Lote, E=Analito, F=Valor, G=NC | **A=RUN, B=Data, C=Nível, D=Lote, E=Analito, F=Resultado, G=Status, H=NC** |

- **RUN** é chave lógica **única global**, atribuída por par *(Data + lote de 6 dígitos)*.
  Todos os níveis e analitos da mesma corrida compartilham o mesmo RUN.
- O `Seq` anterior era sequencial **por lote** — não era chave única (havia `Seq=1` em vários lotes).
- Migração: 25 corridas → RUN 1..25, sem perda de dados (1.575 linhas em Hematologia, 1.000 nas demais).
- Formato numérico de `A` fixado como inteiro (herdava formato de data da coluna antiga).

### Etapa 5 — Exclusão lógica

- Coluna **`Status`** (`Ativo` / `Excluído`), com validação por lista e formatação condicional
  (linha excluída fica cinza e itálica).
- **Nenhum registro é apagado fisicamente.** Todos os consumidores passaram a filtrar `Status="Ativo"`:
  `Calc` (descoberta de RUN, data e valores), `Estatística` (n/média/DP), `Liberação`,
  e as colunas auxiliares `1ªOc` e `RunUnico`.
- Verificado: excluir logicamente as linhas de um analito derruba `n` de 25→24 no Painel e na
  Estatística, e reverter restaura 25.

### Etapa 6 — Modo Desenvolvedor

- **Shape `btnDev`** no Painel (canto superior direito, célula `R1`), 14×14 px,
  **sem preenchimento e sem borda** — imperceptível ao usuário comum. Sem atalho de teclado.
- Ao clicar: abre `frmDev` pedindo senha **mascarada**; valida por hash SHA-256.
- Se correta: desprotege todas as abas, exibe as ocultas e desliga o visual de sistema.
- **Senha:** `QCDEV@2026` — o hash fica em `mSeguranca.DEV_HASH`. Para trocar, gere o
  SHA-256 da nova senha e substitua a constante.

### Etapa 7 — Refatoração do VBA

- Monolito dividido em 4 módulos por responsabilidade (tabela acima).
- Gravação de corrida centralizada em `mDados.GravarCorrida` — o formulário não escreve mais
  célula a célula; delega para a camada de dados.
- `NovoRUN` usa **array em memória** (leitura de bloco único) em vez de varrer célula a célula.
- Constantes de schema (`COL_RUN`, `COL_DATA`, …) eliminam índices mágicos espalhados pelo código.
- `RegistrarLog` criado como **stub documentado** (a trilha de auditoria em si é Fase 2).

---

## 3. Problemas estruturais encontrados

| # | Achado | Situação |
|---|---|---|
| 1 | `Seq` não era chave única (sequencial por lote) | **Corrigido** (RUN global) |
| 2 | Sem exclusão lógica — remoção seria física | **Corrigido** (coluna Status) |
| 3 | `mSystem` monolítico misturando 6 responsabilidades | **Corrigido** (4 camadas) |
| 4 | `Calc` usa referências absolutas fixas (`$BT$1`/`$BU$1` p/ média e DP) | **Pendente** — frágil a inserção de colunas |
| 5 | `LiberStore` gravado de forma preguiçosa (só na troca de lote / `BeforeSave`) | **Pendente** — risco em fechamento por crash |
| 6 | 15 módulos de classe de planilha, só 5 com código | **Pendente** — ruído (inofensivo) |
| 7 | Fonte VBA recuperado vinha com `\r\r\n`, quebrando continuações `_` | **Corrigido** (normalização) |
| 8 | Coluna `RUN` herdava o formato de **data** da coluna antiga, quebrando o casamento de critérios de `COUNTIFS`/`MAXIFS` | **Corrigido** (formato numérico explícito) |

### Duas armadilhas registradas para referência futura

**Formato numérico herdado.** Ao reescrever uma coluna com significado novo, o formato da
coluna antiga permanece. A coluna `RUN` ficou com formato de data (`1` exibido como
`31/12/1899`) e, pior, `COUNTIFS`/`MAXIFS` deixaram de casar o critério — sem gerar erro de
fórmula, apenas devolvendo vazio. O sintoma foi um painel zerado com dados íntegros.
Sempre fixar `NumberFormat` explicitamente após reescrever colunas.

**Exportação de VBA em modo texto.** Ao extrair o código com `olevba` e gravar com
`open(...,"w")`, o conteúdo (que já continha `\r\n`) virou `\r\r\n`, inserindo uma linha em
branco após cada linha — inclusive após as continuações `_` do `Array(...)`. Resultado: erro
de compilação que só se manifestava ao chamar a primeira função, travando o Excel em um
diálogo modal invisível. Gravar sempre em modo binário ou com `newline=""`.

---

## 4. Recomendações para a Fase 2

Em ordem de prioridade:

1. **Trilha de auditoria** — implementar `RegistrarLog` gravando em aba dedicada
   (usuário, data/hora, ação, RUN afetado, valor antes/depois). O gancho já existe em todos
   os pontos de escrita.
2. **Formulário de exclusão** — UI para `MarcarStatusRUN`, exigindo justificativa e
   assinatura (a camada de dados já suporta; falta só a interface).
3. **Backup/autosave automático** — versionamento antes de operações destrutivas.
   *Observação:* durante esta fase o AutoSave do OneDrive persistiu um estado parcial após
   uma falha de COM; convém avaliar desativá-lo para estes arquivos.
4. **Eliminar as referências absolutas do `Calc`** (achado #4) trocando por nomes definidos.
5. **Gravação imediata no `LiberStore`** (achado #5) em vez de só no `BeforeSave`.
6. **Índice de RUN** — com o crescimento do banco (limite atual 15.000 linhas), `NovoRUN`
   passa a varrer todo o range; considerar cache em `Dictionary`.
7. **Proteger o projeto VBA com senha** — hoje o código está acessível pelo editor.

---

## 5. Compatibilidade verificada

Testado nos 3 arquivos após a refatoração:

- Compilação do VBA (`SHA256Hex` confere com `hashlib`)
- Login e papéis (`DoLogin` → ANALISTA)
- `AtualizarTudo` (cadeia completa) → n = 25
- `NovoRUN`: data existente devolve RUN 1; data nova devolve 26
- Troca de lote: lote novo n=0, retorno ao lote original n=25
- `frmCorrida` carrega (66 controles)
- Shape `btnDev` presente, invisível, ligado a `ModoDesenvolvedor`
- Varredura de erros de fórmula: **NENHUM**
