# Investigação do CV alto do CIQ (06 meses) — 13/10/2026

**Somente leitura.** Nenhum resultado foi inativado. A decisão é do laboratório, pela aba **Inativar** (ID + analito + motivo).

**Método** (`cv_investigacao.py`): mesma população do Sigma (ADR-071).

- Escopo: `PARTICIPA_ESTATISTICA = SIM`, janela de 6 meses até o último resultado, por equipamento × analito × nível.
- CV pooled dos lotes com n ≥ 20.
- Um resultado é **suspeito** quando:
  - tem valor ≤ 0; ou
  - está a mais de 5 DP robustos da mediana do próprio lote (MAD; quando MAD = 0, desvio médio absoluto).
- Para cada par, o CV pooled é recalculado **sem** os suspeitos.

| Produto | Pares | CV > meta | Voltam à meta sem os suspeitos | Suspeitos |
|---|---|---|---|---|
| Bioquímica | 158 | 64 | 45 | 253 (232 muito distante do lote, 21 valor <= 0) |
| Hematologia | 93 | 15 | 1 | 35 (35 muito distante do lote) |

## Bioquímica: os 15 pares de maior CV acima da meta

| Equip. | Analito | Nível | CV pooled % | CV sem suspeitos % | Meta CVTp % | Suspeitos |
|---|---|---|---|---|---|---|
| DIMENSION 1 | Amilase | 1 | 60.2 | 1.6 | 10.00 | 3 |
| DIMENSION 2 | Amilase | 1 | 49.0 | 1.3 | 10.00 | 1 |
| DIMENSION 2 | Proteína C-reativa | 1 | 38.1 | 5.9 | 10.00 | 7 |
| DIMENSION 1 | AST (TGO) | 1 | 31.5 | 4.4 | 6.67 | 1 |
| DIMENSION 2 | AST (TGO) | 1 | 28.1 | 4.1 | 6.67 | 2 |
| DIMENSION 1 | ALT (TGP) | 1 | 27.2 | 5.4 | 6.67 | 4 |
| DIMENSION 2 | Fosfatase alcalina | 1 | 25.8 | 3.3 | 10.00 | 1 |
| DIMENSION 1 | Fosfatase alcalina | 1 | 25.3 | 5.9 | 10.00 | 1 |
| DIMENSION 2 | Glicose | 1 | 22.1 | 1.5 | 3.33 | 1 |
| DIMENSION 1 | Glicose | 1 | 20.7 | 2.5 | 3.33 | 2 |
| DIMENSION 2 | Ferro | 2 | 20.4 | 1.4 | 5.00 | 1 |
| DIMENSION 2 | ALT (TGP) | 1 | 19.8 | 5.0 | 6.67 | 1 |
| DIMENSION 1 | Creatina fosfoquinase | 1 | 19.0 | 4.1 | 6.67 | 1 |
| DIMENSION 2 | Creatina fosfoquinase | 1 | 17.6 | 3.5 | 6.67 | 1 |
| DIMENSION 1 | Ureia | 1 | 16.6 | 5.9 | 3.00 | 3 |

## Hematologia: os 15 pares de maior CV acima da meta

| Equip. | Analito | Nível | CV pooled % | CV sem suspeitos % | Meta CVTp % | Suspeitos |
|---|---|---|---|---|---|---|
| XN1000 | MONO# | 1 | 12.3 | 10.0 | 9.98 | 8 |
| XN1000 | MONO% | 1 | 12.2 | 9.0 | 9.98 | 9 |
| XN1000 | MPV | 1 | 3.1 | 3.1 | 1.65 | 0 |
| XN1000 | MPV | 2 | 1.9 | 1.9 | 1.65 | 0 |
| XN1000 | MCHC | 1 | 1.5 | 1.5 | 0.68 | 0 |
| XN1000 | HCT | 1 | 1.4 | 1.4 | 1.33 | 0 |
| XN1000 | RDW-SD | 3 | 1.3 | 1.3 | 1.20 | 0 |
| XN1000 | MCHC | 2 | 1.2 | 1.2 | 0.68 | 0 |
| XN1000 | MCHC | 3 | 1.2 | 1.2 | 0.68 | 0 |
| XN1000 | MCH | 1 | 1.2 | 1.2 | 0.60 | 1 |
| XN1000 | MCV | 1 | 1.0 | 1.0 | 0.60 | 0 |
| XN1000 | MCV | 2 | 1.0 | 1.0 | 0.60 | 0 |
| XN1000 | MCV | 3 | 0.9 | 0.9 | 0.60 | 0 |
| XN1000 | MCH | 2 | 0.8 | 0.8 | 0.60 | 0 |
| XN1000 | MCH | 3 | 0.8 | 0.8 | 0.60 | 0 |

## Leitura

- **Bioquímica:** o CV alto é **dado anômalo, não imprecisão do método**.
  - Poucos resultados absurdos sobem o CV de valores próximos de 2% para dezenas de %. Exemplos: Amilase N1, 482 U/L num controle de 67 U/L; resultados ≤ 0.
  - Revisados e, se procedente, inativados com motivo, levam a maioria dos pares de volta à meta.
- **Hematologia:** os pares acima da meta não vêm de valores anômalos. São imprecisão real diante de metas apertadas da variação biológica (MCHC, MCV, MCH, MPV, com CVTp de 0,6 a 1,7%).
  - É uma questão de especificação: o RT pode avaliar a meta de desempenho (Q: ÓTI / DES / MIN) ou outra fonte.
  - MONO% e MONO# N1 têm alguns resultados distantes.

## Como usar

1. Abra `cv_suspeitos_<produto>.csv` (ID, equipamento, analito, nível, lote, data/hora, resultado, motivo).
2. Para cada ID que o RT considerar não conforme: na aba **Inativar**, digite o ID (só o número), escolha o analito e escreva o motivo.
3. Clique em ⟳ **ATUALIZAR DADOS**.
4. O CV pooled e o Sigma se recalculam sozinhos.

Arquivos: `cv_pares_<produto>.csv` (todos os pares) e `cv_suspeitos_<produto>.csv` (os IDs).
