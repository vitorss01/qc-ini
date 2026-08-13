# QC_INI — Medidas DAX

Criar numa tabela dedicada: **Inserir → Inserir dados → tabela vazia chamada
`_Medidas`** → depois ocultar a coluna. Mantém as medidas fora das tabelas de
dados e o painel de campos legível.

Regra que atravessa tudo: **onde o QC_INI já calcula, o DAX reproduz a mesma
conta.** Diferença numérica entre Excel e Power BI não é arredondamento, é
defeito — ver [`BI_TEST_PLAN.md`](BI_TEST_PLAN.md).

---

## Contagens

```dax
n Resultados = COUNTROWS ( Fato_QC )
```

```dax
n Corridas = DISTINCTCOUNT ( Fato_QC[ID_Corrida] )
```

```dax
n Analitos = DISTINCTCOUNT ( Fato_QC[ID_Analito] )
```

---

## Estatística descritiva

`DP Observado` usa **STDEV.S** (amostral, denominador n−1), que é o que
`DESVPAD.A` do Excel faz. Trocar por `STDEV.P` mudaria o CV% e, por consequência,
o Sigma — silenciosamente, e sempre para melhor. É o erro clássico neste domínio.

```dax
Media Observada = AVERAGE ( Fato_QC[Resultado] )
```

```dax
DP Observado =
VAR n = COUNTROWS ( Fato_QC )
RETURN IF ( n >= 2, STDEV.S ( Fato_QC[Resultado] ) )
```

```dax
CV% =
VAR m = [Media Observada]
VAR s = [DP Observado]
RETURN IF ( NOT ISBLANK ( s ) && m <> 0, DIVIDE ( s, m ) * 100 )
```

**Bias** compara a média observada com a **média-alvo do lote** — não com a média
de outro período. Como o alvo é o mesmo para todas as linhas de um par
(analito, nível, lote), `AVERAGE` sobre a coluna devolve o próprio alvo; se o
filtro misturar lotes, devolve a média dos alvos, que é o comportamento correto
para um agregado.

```dax
Media Alvo = AVERAGE ( Fato_QC[Media_Alvo] )
```

```dax
Bias% =
VAR obs = [Media Observada]
VAR alvo = [Media Alvo]
RETURN IF ( NOT ISBLANK ( alvo ) && alvo <> 0, DIVIDE ( obs - alvo, alvo ) * 100 )
```

```dax
ETp% = AVERAGE ( Fato_QC[ETp_pct] )
```

**Sigma métrico** — `(ETa − |Bias|) / CV`. O valor absoluto no bias não é
detalhe: sem ele, um bias negativo *aumentaria* o Sigma e faria um método ruim
parecer excelente.

```dax
Sigma =
VAR eta = [ETp%]
VAR bias = [Bias%]
VAR cv = [CV%]
RETURN
    IF (
        NOT ISBLANK ( eta ) && NOT ISBLANK ( cv ) && cv > 0,
        DIVIDE ( eta - ABS ( bias ), cv )
    )
```

Sem meta cadastrada, `ETp%` é vazio e `Sigma` fica vazio — **não zero**. Zero
seria lido como "péssimo desempenho"; vazio é lido como "não há meta", que é a
verdade. Mesma distinção que o ADR-023 estabeleceu para a conformidade.

---

## Controle e Westgard

```dax
n Rejeitados = CALCULATE ( COUNTROWS ( Fato_QC ), Fato_QC[Veredito] = "REJEITADO" )
```

```dax
% Aceitáveis =
VAR n = [n Resultados]
RETURN IF ( n > 0, DIVIDE ( n - [n Rejeitados], n ) * 100 )
```

```dax
% Fora de Controle = IF ( [n Resultados] > 0, 100 - [% Aceitáveis] )
```

```dax
Viol 1_3s = SUM ( Fato_QC[W_1_3s] )
Viol 2_2s = SUM ( Fato_QC[W_2_2s] )
Viol R_4s = SUM ( Fato_QC[W_R_4s] )
Viol Total = [Viol 1_3s] + [Viol 2_2s] + [Viol R_4s]
```

> `4_1s` e `10x` **não existem** como medida porque não existem no motor
> (`Calc!N3` e `Calc!O3` são `IF(OR(FALSE;FALSE);1;0)`). Criar a medida daria ao
> gestor um indicador que sempre marca zero e que ele leria como "nunca
> violamos". Ver [`BI_ARCHITECTURE.md` §5](BI_ARCHITECTURE.md).

---

## Levey-Jennings

As linhas de controle vêm **materializadas na fato**, calculadas com o alvo do
lote correto. Em visual, use-as como linhas constantes por série:

```dax
LJ Média  = AVERAGE ( Fato_QC[Media_Alvo] )
LJ +1s = AVERAGE ( Fato_QC[Lim_p1s] )
LJ +2s = AVERAGE ( Fato_QC[Lim_p2s] )
LJ +3s = AVERAGE ( Fato_QC[Lim_p3s] )
LJ -1s = AVERAGE ( Fato_QC[Lim_m1s] )
LJ -2s = AVERAGE ( Fato_QC[Lim_m2s] )
LJ -3s = AVERAGE ( Fato_QC[Lim_m3s] )
```

**Só faz sentido com um analito, um nível e um lote selecionados.** Com o filtro
aberto, a média dos limites de analitos diferentes não significa nada. Guarda:

```dax
LJ Válido =
IF (
    HASONEVALUE ( Dim_Analito[ID_Analito] )
        && HASONEVALUE ( Dim_Nivel[Nivel] )
        && HASONEVALUE ( Dim_Lote[ID_Lote] ),
    1, 0
)
```

```dax
LJ Aviso =
IF ( [LJ Válido] = 0,
     "Selecione um analito, um nível e um lote para ver o Levey-Jennings." )
```

Coloque `LJ Aviso` num cartão sobre o gráfico. É preferível dizer "não dá para
mostrar" a desenhar uma curva sem sentido.

---

## Semáforo executivo

Critérios **derivados das metas cadastradas**, não arbitrários:

```dax
Status Global =
VAR s = [Sigma]
VAR fora = [% Fora de Controle]
RETURN
    SWITCH (
        TRUE (),
        ISBLANK ( [n Resultados] ), "SEM DADOS",
        ISBLANK ( s ), "SEM META",
        s < 3 || fora > 5, "🔴 FORA DE CONTROLE",
        s < 4 || fora > 2, "🟡 ATENÇÃO",
        "🟢 SOB CONTROLE"
    )
```

Os cortes seguem a prática consolidada da métrica Sigma em laboratório clínico:
σ ≥ 4 desempenho bom, 3 ≤ σ < 4 aceitável com monitoramento reforçado, σ < 3
inaceitável. Estão aqui em um lugar só, para que mudar a política seja uma
edição e não uma caça.

---

## Tendência

```dax
Media Movel 7 Corridas =
VAR corridaAtual = SELECTEDVALUE ( Fato_QC[RUN] )
RETURN
    CALCULATE (
        [Media Observada],
        FILTER (
            ALLSELECTED ( Fato_QC ),
            Fato_QC[RUN] <= corridaAtual && Fato_QC[RUN] > corridaAtual - 7
        )
    )
```

```dax
Deslocamento vs Alvo (SD) =
VAR d = [Media Observada] - [Media Alvo]
VAR sd = AVERAGE ( Fato_QC[DP_Alvo] )
RETURN IF ( sd > 0, DIVIDE ( d, sd ) )
```

Deslocamento em **unidades de SD** e não em unidade do analito: é o que permite
comparar glicose com cálcio na mesma tela.

---

## Resumo do dia

```dax
Data Mais Recente = MAX ( Fato_QC[Data] )
```

```dax
n Hoje =
CALCULATE ( [n Resultados], Fato_QC[Data] = [Data Mais Recente] )
```

```dax
Rejeitados Hoje =
CALCULATE ( [n Rejeitados], Fato_QC[Data] = [Data Mais Recente] )
```

```dax
Analitos Críticos =
CALCULATE (
    DISTINCTCOUNT ( Fato_QC[ID_Analito] ),
    Fato_QC[Veredito] = "REJEITADO"
)
```
