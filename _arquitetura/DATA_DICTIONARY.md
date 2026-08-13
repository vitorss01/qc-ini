# QC_INI — Dicionário de dados da camada BI

Tabela: **`tblBI_Fato`** na aba `BI_Data` (oculta — é camada de dados, não tela).
Granularidade: **um resultado de controle = (Lote, RUN, Nível, Analito)**.
Reconstruída inteira por `mBI.AtualizarBIData`. Não editar à mão.

---

| # | Campo | Tipo | Origem | Definição | Regra |
|--:|---|---|---|---|---|
| 1 | `ID_Result` | texto | derivado | `ID_Lote\|RUN\|Nivel\|ANALITO` | **chave primária**; unicidade validada no build |
| 2 | `ID_Corrida` | texto | derivado | `ID_Lote\|RUN` | agrupa os resultados de uma corrida |
| 3 | `Data` | data | `DB_Resultados!B` | dia da corrida | truncada para o dia (`Int` do serial) |
| 4 | `Ano` | inteiro | derivado | ano da `Data` | |
| 5 | `Mes` | inteiro | derivado | 1–12 | |
| 6 | `Trimestre` | inteiro | derivado | 1–4 | `Int((Mes−1)/3)+1` |
| 7 | `Competencia` | texto | derivado | `AAAA-MM` | ordenável como texto |
| 8 | `ID_Analito` | texto | derivado | nome em maiúsculas | junta com `Dim_Analito` |
| 9 | `Analito` | texto | `DB_Resultados!E` | nome de exibição | |
| 10 | `Area` | texto | `Analitos!B` | área/setor analítico do analito | vazio se não cadastrado |
| 11 | `Unidade` | texto | `Analitos!C` | unidade de medida | |
| 12 | `ID_Lote` | texto | derivado | núcleo do lote (`NucleoLote`) | junta com `Dim_Lote` |
| 13 | `Lote` | texto | `DB_Resultados!D` | código completo `QC-nnnnnn-NN` | formato texto (preserva zero à esquerda) |
| 14 | `Nivel` | inteiro | `DB_Resultados!C` | nível do controle (1–3) | |
| 15 | `RUN` | inteiro | `DB_Resultados!A` | identificador da corrida | |
| 16 | `Resultado` | número | `DB_Resultados!F` | valor medido | vazio quando não numérico |
| 17 | `Status` | texto | `DB_Resultados!G` | `Ativo` / `Excluído` | exclusão lógica, com trilha |
| 18 | `Ativo` | inteiro | derivado | 1 se `Status = "Ativo"` | o Power Query filtra por este campo |
| 19 | `Media_Alvo` | número | `LotesStore` | média-alvo **do lote do registro** | vazio se o lote não tem alvo |
| 20 | `DP_Alvo` | número | `LotesStore` | DP-alvo do lote | precisa ser `> 0` |
| 21 | `Z` | número | calculado | `(Resultado − Media_Alvo) / DP_Alvo` | idêntico a `Calc!G`/`Calc!AC` |
| 22 | `Lim_m3s` | número | calculado | `Media_Alvo − 3·DP_Alvo` | Levey-Jennings |
| 23 | `Lim_m2s` | número | calculado | `− 2·DP` | |
| 24 | `Lim_m1s` | número | calculado | `− 1·DP` | |
| 25 | `Lim_p1s` | número | calculado | `+ 1·DP` | |
| 26 | `Lim_p2s` | número | calculado | `+ 2·DP` | |
| 27 | `Lim_p3s` | número | calculado | `+ 3·DP` | |
| 28 | `CVtp_pct` | número | `Eng_Especificacoes!D` | CV total permitido (%) | meta vigente no ano |
| 29 | `BIAStp_pct` | número | `Eng_Especificacoes!E` | bias permitido (%) | |
| 30 | `ETp_pct` | número | `Eng_Especificacoes!F` | erro total permitido (%) | vazio = **sem meta**, não zero |
| 31 | `W_1_3s` | 0/1 | calculado | `ABS(Z) > 3` | idêntico a `Calc!K` |
| 32 | `W_2_2s` | 0/1 | calculado | ambos os níveis da corrida `> 2` ou `< −2` | idêntico a `Calc!L` |
| 33 | `W_R_4s` | 0/1 | calculado | níveis em lados opostos, ambos `> 2` em módulo | idêntico a `Calc!M` |
| 34 | `Veredito` | texto | calculado | `REJEITADO` se alguma regra = 1, senão `OK` | idêntico a `Calc!P`/`Calc!AL` |

---

## Campos ausentes, e por quê

| Campo pedido | Situação |
|---|---|
| `Equipamento` | **não existe na produção.** A aba `Corridas` (que teria `CC_EQUIP`) só existe no artefato do motor da Hematologia, não no `QC_Bioquimica.xlsm`. |
| `Setor` | **não existe.** O mais próximo é `Area` do analito (campo 10), que é área analítica, não setor organizacional. |
| `Responsável` / `Assinatura` | existem na aba `Liberação`, e **não** são exportados: são dado pessoal e não servem a indicador de qualidade. |
| `W_4_1s`, `W_10x` | **não implementados no motor** — `Calc!N3`/`O3` são `IF(OR(FALSE;FALSE);1;0)`. Não foram criados como coluna sempre-zero para não sugerir um monitoramento que não acontece. |

Quando `Equipamento` passar a existir no QC_INI, o gancho é: acrescentar a coluna
em `mBI.Cab()`, ler da origem no passo 2 de `AtualizarBIData`, e criar
`Dim_Equipamento` por `Table.Distinct` no Power Query. Nada mais muda.

---

## Invariantes que o build verifica

1. `ID_Result` é único em todas as linhas.
2. A tabela tem exatamente 34 colunas.
3. `tblBI_Fato` existe como `ListObject` (não como faixa).
4. `Z` e `Veredito` coincidem com o `Calc` para o analito e lote em tela,
   tolerância `1e-6` no `Z` e igualdade exata no veredito.

Falhando qualquer uma, o build **para** e o artefato não é gerado.
