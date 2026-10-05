# Anexo do QUALITY_GATE — mapa real da camada de dados (ADR-057, revisto pelo ADR-066)

> Mapa levantado no código para avaliar a proposta QA-ETL-001 (04/10/2026). As seções 1–9 descrevem o comportamento
> **antes** das correções do ADR-066; a seção 10 diz o que mudou. As linhas citadas são da versão anterior.

## 1. Fluxo real × nomes da proposta
| Proposta | Objeto real | Consulta / aba | Papel |
|---|---|---|---|
| TB1 Recebimento | `tblDB_Recebimento` | `pq/DB_RECEBIMENTO.m` (aba DB) | Acumula o DB_SEAC e nunca apaga. Anti-junção por ID_REGISTRO (`:58-62`); RECEBIDO_EM carimba a 1ª chegada (`:63-64`) |
| TB2 Organização | `tblDB_Organizado` | `pq/DB_ORGANIZADO.m` | **Visão horizontal só de conferência.** NÃO alimenta a final (a final lê `tblDB_Recebimento` em `DB_CQ_FINAL.m:114`) |
| TB3 Exclusão | `tblInativacao_NaoConformes` | aba Inativar | **Soft-delete por ID**, coluna `REGISTRAR - LJ` |
| — | `tblComentariosTecnicos` | COMENTARIOS_TECNICOS | Justificativa; sem ela, E01 |
| TB4 Manual | `tblResultados_Manuais` | Digitar Resultados | MAN_####, RUN calculado pela mesma regra |
| TB5 Final | `tblCQ_Final` | `pq/DB_CQ_FINAL.m` (Principal - Resultados) | UNION ALL interface + manuais; uma linha = um resultado |
| — | `tblQA_Integracao` | `pq/QA_INTEGRACAO.m` | Só confere o que a final decidiu |
| Estatística/LJ | `mDados.CarregarDB` → `mEstatistica` / `mEstatPeriodo` (Bio) / `mIncerteza` | VBA | Leem **só** a tblCQ_Final |

Fluxo: `DB_SEAC (ou vazio em HISTORICO) → SEAC_ORIGEM → tblDB_Recebimento → [Interface + Manuais] → join Inativação/Comentários → classificação → RUN → tblCQ_Final → QA_INTEGRACAO / VBA (LJ, Painel, Estatística, Westgard, u(Rw))`.

## 2. Chave
- **Interface**: `ID_REGISTRO = PREFIXO_ID & "-" & ID_ORIGEM` (Int64) (`SEAC_ORIGEM.m:47,51`). Não é índice de linha.
- **ID_ORIGEM nulo** dá ID nulo: a linha é regravada a cada refresh e descartada na final (`DB_CQ_FINAL.m:118`), sem achado de QA (**P-02**).
- **ID_ORIGEM repetido na mesma carga**: as duas linhas entram com o mesmo ID; Canon não marca DUP (`:150,155`); o join com RunPorId multiplica (`:364-368`).
- **Retransmissão** (ID diferente, mesma amostra/item/instante/valor) vira DUPLICIDADE_ORIGEM (`:146-156`) e A02.
- **ID digitado** (Inativar, Comentários, Manual) passa pela normalização idêntica em VBA e PQ: Trim, Upper, sem espaço ASCII; número puro → `PREFIXO-n`; `MAN_n` → `MAN_` com 4 dígitos (`mIntegracao.bas:306-319`; `DB_CQ_FINAL.m:79-85`). Não remove zeros à esquerda nem NBSP.
- **Manual**: ID vazio recebe `ProximoIdManual` = maior MAN_ presente + 1 (`mIntegracao.bas:429-448`). Fora do padrão vira `MAN_INVALIDO_L<n>` + MANUAL_INCOMPLETO (`DB_CQ_FINAL.m:185,191-192`). Repetido: a 1ª linha física vale e as demais viram `ID~L<n>` + CONFLITO_MANUAL (`:213-216,240`).
- **Inativação repetida**: deduplicada, vale a 1ª (`:265`) e gera E02. Inativação de ID inexistente: E03 sem efeito.

## 3. Exclusão (inativação) e checkbox
- Join **LEFT OUTER** por ID normalizado (`:278`): a linha **continua** na final.
- Precedência (`:286-291`): MANUAL_INCOMPLETO > DUPLICIDADE_ORIGEM > CONFLITO_MANUAL > SEM_VALOR > INATIVADO > ATIVO. Inativar algo já excluído por regra gera A04.

| STATUS | PARTICIPA | REGISTRAR - LJ | TIPO_PLOTAGEM_LJ |
|---|---|---|---|
| ATIVO | SIM | — | NORMAL |
| INATIVADO | NÃO | SIM (TRUE, vazio, "SIM/S/V/X/1", número ≠ 0) | X_VERMELHO |
| INATIVADO | NÃO | qualquer outro valor | NAO_PLOTAR |
| demais | NÃO | — | NAO_PLOTAR |
(`:86-91,267,292,300,302`)
- Na criação o VBA grava TRUE (`mIntegracao.bas:383`); depois disso vale o que o usuário marcar.
- **Desmarcar não reabilita.** Só apagar a linha da inativação reabilita (T14).
- O INATIVADO **mantém o RUN**, porque entra em `Reais` (`:306,321`).

## 4. Inserção manual
- Linha totalmente vazia: ignorada (`:163-164`).
- Incompleta: ID, DATA, HORA, LOTE, NIVEL fora de 1..NIVEIS, ANALITO, RESULTADO, ou data posterior a REGISTRADO_EM. Resultado: MANUAL_INCOMPLETO, E04, fora do RUN (`:184-189,321`).
- RESULTADO em texto é lido em pt-BR (`:176`).
- **Manual × manual**: mesma chave + instante ao segundo dá CONFLITO (`:220-225,241`).
- **Manual × interface**: mesmo EQUIP, MATRIZ, LOTE, NIVEL, ANALITO e **DATA**, |Δ| ≤ TOLERANCIA_CONFLITO_MIN. A interface prevalece, **inclusive se INATIVADA** (`:229-243`), e o manual vira CONFLITO_MANUAL + A01. Fora da tolerância no mesmo dia: A06, e os dois participam.
- Não há inserção física: a posição é derivada (seções 5 e 6).

## 5. Ordenação e RUN (regra formal)
1. Entram na corrida ATIVO, INATIVADO e SEM_VALOR. Ficam fora incompleto, DUP e conflito (`:321`).
2. `_GK` = EQUIP|MATRIZ|LOTE|ANALITO|yyyyMMdd. Ordenação: _GK, DATA_HORA, NIVEL, ITEM_ID, ID_REGISTRO. Abre bloco novo quando muda o _GK ou quando Δ > GAP_CORRIDA_MIN (`:334-341`).
3. `_POS`: dentro de (bloco, nível), o 2º resultado do mesmo nível abre posição nova (`:344-351`).
4. CORRIDA_NO_DIA = ordem de (bloco, posição) no _GK; RUN = yyMMdd·100 + corrida, ou nulo se > 99, que gera E05 (`:354-366`).
5. Desempate final por ID_REGISTRO, comparado como texto.
6. **Consequências**: o RUN é estável com refresh e com inativação. Um manual que abre corrida antes de outras do mesmo _GK/dia **renumera as posteriores (+1)**, preservando a ordem relativa; outros _GK não mudam. Um resultado tardio também renumera.
7. Ordem da tabela: DATA_HORA, ANALITO, NIVEL, ID (`:400-401`); no recebimento, DATA_HORA, ID_ORIGEM.

## 6. Final × Estatística × LJ
- **Fonte única**: `mDados.CarregarDB` lê a tblCQ_Final; PARTICIPA=SIM vira "Ativo", qualquer outro valor vira "Inativo" (`src/hema/mDados.bas:60-87`).
- **Índices**: `mIdx` = Ativo (valores, z, Westgard, Painel). `mIdxX` = não-Ativo com X_VERMELHO, usado só para o marcador. NAO_PLOTAR não entra em nenhum (`mEstatistica.bas:127-163`).
- **X**: vai para a linha do RUN, no nível COL_NIVEL, até 3 slots, no bloco AD:AL de Eng_Saida, fora das colunas de valor Y.. que o Painel soma (`:1257-1271,1358`). Uma corrida só com X entra no eixo, mas não na sequência de Westgard (`:1277-1281`). **Ela conta na janela NK=180** (`:1219`).

| Consumidor | População | Observação |
|---|---|---|
| Painel (n, média, DP, CV, bias, ET, Sigma) | Ativo da janela de 180 corridas | Pode divergir da aba Estatística em períodos longos |
| Estatística Hema | EstatBasica: Ativo, ano e lote; lote vazio junta todos os lotes | (`:365-393`) |
| Estatística Bio | mEstatPeriodo: Ativo, período, equipamento, lote | |
| Eventos_Westgard | Ativo, lote do Painel | |
| u(Rw) | Ativo numérico por lote, n ≥ 20 | **PEXCL usa INATIVADO por desenho** |
| Fórmulas de cabeçalho | Só contam SIM, X_VERMELHO, INATIVADO e MANUAL | |

- **Hema**: `selEquipamento=""`, então todos os equipamentos entram numa série só (`instalar_integracao.py:155`).

## 7. Refresh e modos
- **Auto-leitura** (`DB_RECEBIMENTO.m:40`): é idempotente para ID não nulo.
- **HISTORICO**: origem vazia com as colunas do recebido; erro se o recebimento estiver ausente; A07 em todo refresh; A03 inoperante (`SEAC_ORIGEM.m:26-33`; `QA_INTEGRACAO.m:151-157`).
- **Falha de fonte em SEAC**: o refresh para e não declara "concluída" (T20).

## 8. Identidades de reconciliação (oráculo do gate)
1. `IDs(Final) = IDs(Recebimento, ID≠nulo) ∪ IDs_manuais_derivados`, com `|Final| = |Recebimento| + |linhas manuais não vazias|` e IDs únicos.
2. `Final = ATIVO ⊔ INATIVADO ⊔ SEM_VALOR ⊔ DUPLICIDADE_ORIGEM ⊔ CONFLITO_MANUAL ⊔ MANUAL_INCOMPLETO`.
3. `{PARTICIPA=SIM} = {ATIVO} = {TIPO=NORMAL}`; `{X_VERMELHO} = {INATIVADO ∧ LJ}`.
4. `{INATIVACAO_REGISTRADA=SIM} = NormId(Inativar) ∩ IDs(Final)`, sendo `{INATIVADO}` esse conjunto menos as precedências (A04).
5. Para cada (analito, lote, nível, filtro), a Estatística é igual ao cálculo independente sobre {ATIVO}. Pontos do LJ = {ATIVO} da janela; X do LJ = {X_VERMELHO} da janela.
6. Toda linha da origem sem ID, ou com ID repetido, gera achado no QA (**hoje não gera**: P-02 e D01).

## 9. QA_INTEGRACAO
- **Detecta**: E01–E05, A01–A07, I01–I03 (`QA_INTEGRACAO.m:44-159`).
- **Não detecta**: ID nulo, ID duplicado no recebimento ou na final, tabela de entrada ausente, troca de PREFIXO, mudança de analito/lote/instante na origem, conflito que atravessa a meia-noite.

## 10. O que mudou com o ADR-066 (regras vigentes)

| Tema | Regra vigente |
|---|---|
| ID nulo ou inválido na origem | Não entra no recebimento; **E06** no QA, com amostra, analito, nível e data/hora |
| ID repetido na mesma carga | Entra uma vez (vale o menor ITEM_ID; empate, a 1ª DATA_HORA); **E07**; a final nunca multiplica linhas |
| Tabela de entrada ausente (inativação, manual, comentário, de/para) | A atualização **para** com erro citando a tabela |
| Recebimento ausente com a final preenchida | A atualização **para**; carga inicial só com CFG `CARGA_INICIAL = SIM` |
| Prefixo | Sempre em maiúsculas; prefixo recebido diferente do CFG gera **E08** |
| Interface sem DATA_HORA | Status **SEM_DATA_HORA** (não participa, não plota, sem RUN); **E09** |
| Normalização de ID (PQ, QA e VBA, idênticas) | Também remove NBSP e zeros à esquerda (`00123` e `HEM-00123` = `HEM-123`) |
| Próximo ID manual | Nunca reaproveita: máximo de manual ∪ inativação ∪ comentários ∪ final ∪ marca `qcUltimoIdManual` |
| "REGISTRAR - LJ" | Vazio = SIM só na linha nova; vazio após o carimbo = NÃO; valor não reconhecido = NÃO e **A08** |
| RESULTADO manual em texto com ponto e sem vírgula | MANUAL_INCOMPLETO + E04 "RESULTADO ambíguo" |
| MAN_ repetido entre linha incompleta e completa | Vale a completa, qualquer que seja a ordem física |
| Conflito manual × interface | Tolerância no DATA_HORA, atravessando a meia-noite (dia, véspera e dia seguinte) |
| Manual que coincide com interface INATIVADO | Continua CONFLITO_MANUAL (fora da estatística) + **A09** — decisão do RT pendente |
| Mudança na origem de um ID já recebido | **A03** compara VALOR, FLAG, ANALITO, LOTE, NÍVEL e DATA_HORA (ao segundo) |
| Janela de 180 corridas do Painel | Conta **só corridas elegíveis**; corrida só com X não tira lugar de resultado |
| Hematologia (sem seletor de equipamento) | Painel, Estatística, Westgard e u(Rw) do `EQUIPAMENTO_PADRAO`; outro equipamento gera **A10** |
| Gravar (Ctrl+S) | Sempre trancado e sem a identidade da sessão (ADR-065) |

Identidades do gate (seção 8) continuam valendo, com o status SEM_DATA_HORA na partição e com a linha sem ID
contada como **rejeitada e reportada** (E06), nunca como perdida.
