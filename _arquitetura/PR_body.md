## O que muda

Reestrutura os três produtos de Controle Interno da Qualidade (Hematologia, Bioquímica, Imunologia) de planilha para **sistema em camadas**, com banco normalizado, operação por UserForms e motor estatístico validado matematicamente.

```
Interface   frmCorrida · frmMassa · frmExcluir   (nenhum escreve em célula)
    ↓
mDados      RUN · upsert em lote · exclusão lógica
    ↓
DB_Resultados   ← única fonte operacional (vertical, normalizado)
    ↓
mEstatistica    motor: média/DP/CV/Bias/ET/Sigma/Z + Westgard
    ↓
Painel · Estatística · Eventos_Westgard · Resultados (view)
```

## Por quê

O sistema tinha três defeitos estruturais que impediam uso auditável:

| Defeito | Correção |
|---|---|
| `Seq` era sequencial **por lote** — havia `Seq=1` em vários lotes, logo não era chave | **RUN**, único por (Data + lote de 6 dígitos) |
| Sem coluna de status — remoção seria física | **Exclusão lógica**; nada é apagado |
| Estatística espalhada em fórmulas por várias abas | **`mEstatistica`** centraliza; o Excel guarda só o resultado |

## Destaques

**Elegibilidade CLSI EP05/C24** — tabela `Cfg_Status` com 8 estados (Ativo, Excluído, Inválido, Calibração, Manutenção, Troca de Lote, Treinamento, Rejeitado-Operacional). Novos estados entram sem tocar em código. Estado desconhecido **nunca** entra no cálculo (falha segura).

**Westgard** — 12s/13s/22s/R4s/41s/10x, com **12s como alerta, não rejeição** (comportamento correto). `mWestgardKnowledge` centraliza classificação, interpretação, causas e sugestões — sem duplicação de texto.

**Chave de upsert = RUN + Nível + Analito.** O Nível entra na chave porque um RUN abrange todos os níveis da corrida; sem ele, o Nível 2 sobrescreveria o Nível 1 do mesmo analito.

## Dois bugs de VBA que valem registro

**1. Travamento (Fase 3A).** `RegistrarEventosWestgard` congelava sem lançar erro. Causa: o identificador **`aS` é a palavra reservada `As`** (VBA é case-insensitive) — erro de sintaxe. Ficou latente porque `CodeModule.AddFromString` não valida sintaxe na inserção e o VBA compila **sob demanda**: só os procedimentos exercitados compilavam. Sob `/automation` o diálogo de compilação fica invisível e o Excel bloqueia com 0% de CPU — e `On Error GoTo` não intercepta erro de compilação.

Provado por: trace em arquivo (o marcador da primeira instrução nunca gravou), amostragem de CPU por 300 s, enumeração da janela modal `#32770` com texto "Erro de sintaxe", e repro mínimo isolado (`aS` congela, `aSd` roda em 0,01 s).

`AtualizarEstatistica`: de travamento para **0,172 s**.

**2. Erro 9.** Array local `val()` sombreava a função nativa `Val()`, transformando conversão em indexação de array bidimensional.

Ambos são a mesma classe: identificador injetado por script colidindo com nome reservado do VBA.

## Validação

| Métrica | Motor VBA | Referência Python |
|---|---|---|
| Média | 14,000000 | 14.000000 |
| DP (amostral, n−1) | 3,162278 | 3.162278 |
| CV% | 22,587698 | 22.587698 |
| Bias% | 40,000000 | 40.000000 |
| Erro Total | 21,500000 | 21.500000 |
| Sigma | 5,000000 | 5.000000 |
| Z-score | 1,000000 | 1.000000 |

Westgard **9/9 asserções**, incluindo o caso sutil: níveis em lados opostos disparam R4s e **não** 22s. Elegibilidade 6/6. Zero erros de fórmula nas três pastas.

## Para o revisor

⚠️ **O arquivo vivo `QC_Hematologia.xlsm` ainda contém o travamento.** A correção validada está em `_arquitetura/fase3a_corrigido/QC_Hematologia.xlsm` e **não foi promovida** — a regra da fase era não tocar no arquivo de produção. Promoção é decisão do gestor:

```
Copy-Item "_arquitetura\fase3a_corrigido\QC_Hematologia.xlsm" "QC_Hematologia.xlsm" -Force
```

Bioquímica e Imunologia estão funcionais (Fase 2) e não foram tocadas na Fase 3.

**Dois pontos que precisam de decisão antes de promover:**

1. `priR`/`priRun`/`ultR`/`ultRun` não são reinicializados por grupo em `mEstatistica`. Em VBA o escopo de `Dim` é o procedimento, não o laço — a partir do primeiro grupo com violação, os seguintes herdam a regra/RUN do anterior. Defeito provado por construção, hoje mascarado porque o procedimento nunca executou. Correção: zerar as quatro variáveis junto com `nv` e `maxZ`.

2. A primeira execução real **destrói fórmulas permanentemente**: `AtualizarPainelEng` sobrescreve `Painel!B7:J9` e `AtualizarEstatisticaAba` sobrescreve `Estatística!C7:M126` com valores literais. É o desenho da Fase 3 (motor único em VBA), mas é irreversível.

`_arquitetura/FASE3A_achados_estaticos.md` traz 64 achados de análise estática (12 críticos) de código, auditoria ISO 15189 e UX — ainda não triados. Backups locais ficaram fora do versionamento via `.gitignore`.

🤖 Generated with [Claude Code](https://claude.com/claude-code)
