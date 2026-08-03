# Snapshot da produção — ponto de restauração da RC1

Extraído de `QC_Hematologia.xlsm` (raiz do repositório) em 03/08/2026, **antes**
de qualquer alteração do Sprint HARDENING 1.

Gerado por [`scripts_fase3/snapshot_projeto.ps1`](../scripts_fase3/snapshot_projeto.ps1)
e [`scripts_fase3/snapshot_formulas.ps1`](../scripts_fase3/snapshot_formulas.ps1).

## Conteúdo

| Arquivo | O que é |
|---|---|
| `vba/` | 33 componentes VBA — 8 módulos `.bas`, 18 classes `.cls`, 4 UserForms `.frm` + `.frx` |
| `formulas.csv` | 61.416 fórmulas: aba, endereço, R1C1 e valor |
| `nomes.csv` | 55 nomes definidos, com `RefersTo` e visibilidade |
| `abas.csv` | 18 abas: nome, `CodeName`, índice, visibilidade, proteção, `UsedRange` |
| `objetos.csv` | 21 objetos: gráficos, formas e tabelas, por aba |

## Regra de uso

**O `.xlsm` é artefato de build; a fonte oficial do código é `vba/`.**
Toda alteração de código nasce aqui e é aplicada à pasta de trabalho — nunca o
contrário. Isso resolve o problema de `.xlsm` não fazer merge no Git: duas
máquinas podem editar `.bas` diferentes sem conflito binário.

A produção permanece **intacta** até o Quality Gate fechar. O hardening trabalha
sobre cópia.

## Verificar que uma build não regrediu

```powershell
$root = (Get-Item "...\CLONE_PLANILHA_QC__*\qc-ini").FullName
& "$root\_arquitetura\scripts_fase3\snapshot_formulas.ps1" `
    -Workbook "<build_candidata.xlsm>" -OutCsv "$env:TEMP\candidata.csv"
& "$root\_arquitetura\scripts_fase3\diff_formulas.ps1" `
    -Referencia "$root\_arquitetura\snapshot_producao\Hematologia\formulas.csv" `
    -Candidato "$env:TEMP\candidata.csv" -OutCsv "$env:TEMP\divergencias.csv"
```

Aceite do item 2.1 do Quality Gate = `ACEITE: OK`.

## Pendências conhecidas

- Snapshot equivalente de **Bioquímica** e **Imunologia** ainda não gerado.
- `_arquitetura/vba_recuperado/` é de outra build: contém `mSystem.bas`, que **não
  existe** na produção, e não contém `mDados`, `mEstatistica`, `mOperacao`,
  `mSeguranca`, `mLotes`, `mUI`, `mEntrada` nem `mWestgardKnowledge`. Não usar
  como referência — está superado por esta pasta.
