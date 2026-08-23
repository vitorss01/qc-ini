# -*- coding: utf-8 -*-
"""
gerar_pbip.py -- gera o projeto Power BI (PBIP) do QC_INI a partir da
tabela tblBI_Fato do artefato Excel.

Por que um gerador e nao um .pbix editado a mao: ADR-021. O PBIX e binario
(o DataModel e um stream do Analysis Services) e nao versiona, nao revisa e
nao reproduz. O PBIP e texto -- TMDL para o modelo, JSON para o relatorio --
entao o dashboard entra no repositorio pela mesma porta que todo o resto:
script versionado, nunca edicao manual no produto final.

Uso:
    python gerar_pbip.py --xlsm "C:\\...\\QC_Bioquimica.xlsm" --saida "C:\\...\\PowerBI"

O caminho do .xlsm vira o parametro pCaminhoQC dentro do modelo; trocar de
maquina e editar um parametro, nao reescrever consultas.
"""
import argparse
import json
import os
import shutil
import uuid

NOME = "QC_INI_Bioquimica"

# ---------------------------------------------------------------- utilidades

def guid(semente):
    """GUID estavel por nome: reexecutar o gerador nao troca todos os
    lineageTag do modelo, o que geraria um diff inteiro a cada build."""
    return str(uuid.uuid5(uuid.NAMESPACE_URL, "qcini/pbip/" + semente))


def escrever(caminho, texto):
    os.makedirs(os.path.dirname(caminho), exist_ok=True)
    with open(caminho, "w", encoding="utf-8", newline="\r\n") as f:
        f.write(texto)


def escrever_json(caminho, obj):
    os.makedirs(os.path.dirname(caminho), exist_ok=True)
    with open(caminho, "w", encoding="utf-8", newline="\r\n") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


def plataforma(caminho, tipo, nome):
    escrever_json(caminho, {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": tipo, "displayName": nome},
        "config": {"version": "2.0", "logicalId": guid(tipo + "/" + nome)},
    })


# ------------------------------------------------------- esquema da tabela fato
# As 76 colunas de tblBI_Fato. O tipo vem do que o motor grava, nao de
# adivinhacao: ver mBI.bas.
TEXTO = "string"
INT = "int64"
REAL = "double"
DATA = "dateTime"

COLUNAS_FATO = [
    ("ID_Result", TEXTO, True), ("ID_Corrida", TEXTO, True),
    ("Data", DATA, False), ("Ano", INT, False), ("Mes", INT, False),
    ("Trimestre", INT, False), ("Competencia", TEXTO, False),
    ("ID_Analito", TEXTO, True), ("Analito", TEXTO, False),
    ("Area", TEXTO, False), ("Unidade", TEXTO, False),
    ("ID_Lote", TEXTO, True), ("Lote", TEXTO, False),
    ("Nivel", INT, False), ("RUN", INT, False),
    ("Resultado", REAL, False), ("Status", TEXTO, True), ("Ativo", INT, True),
    ("Media_Alvo", REAL, False), ("DP_Alvo", REAL, False), ("Z", REAL, False),
    ("Lim_m3s", REAL, False), ("Lim_m2s", REAL, False), ("Lim_m1s", REAL, False),
    ("Lim_p1s", REAL, False), ("Lim_p2s", REAL, False), ("Lim_p3s", REAL, False),
    ("CVtp_pct", REAL, False), ("BIAStp_pct", REAL, False), ("ETp_pct", REAL, False),
    ("W_1_3s", INT, False), ("W_2_2s", INT, False), ("W_R_4s", INT, False),
    ("W_4_1s", INT, False), ("W_10x", INT, False), ("A_1_2s", INT, False),
    ("Veredito", TEXTO, False),
    ("Produto", TEXTO, True), ("WorkbookID", TEXTO, True),
    ("VersaoContrato", TEXTO, True), ("AtualizadoEmUTC", TEXTO, True),
    ("FonteArquivo", TEXTO, True),
    ("ID_Result_Global", TEXTO, True), ("ID_Corrida_Global", TEXTO, True),
    ("ID_Lote_Global", TEXTO, True), ("ID_Analito_Global", TEXTO, True),
    ("N_Observado", INT, False), ("Media_Observada", REAL, False),
    ("DP_Observado", REAL, False), ("CV_Observado_pct", REAL, False),
    ("Bias_Observado_pct", REAL, False), ("ET_Observado_pct", REAL, False),
    ("Sigma_Obs", REAL, True),   # exibida pela medida [Sigma]; ver RENOMEAR
    ("Fonte_Especificacao", TEXTO, False), ("ID_Especificacao", TEXTO, True),
    ("Vigencia_Inicio", TEXTO, True), ("Vigencia_Fim", TEXTO, True),
    ("Situacao_Especificacao", TEXTO, False),
    ("Usuario_Atualizacao", TEXTO, True), ("Tipo_Evento", TEXTO, True),
    # ADR-033. Bias_Observado_pct (col. 51) continua ASSINADO; a magnitude tem
    # coluna propria porque AVERAGE sobre bias assinado cancela desvios opostos
    # e devolve um vies falso perto de zero.
    ("Bias_Observado_abs_pct", REAL, False),
    ("Classificacao_Sigma", TEXTO, False),
    ("Margem_ETp_pp", REAL, False),
    ("Margem_ETp_pct", REAL, False),
    ("Status_Margem_ETp", TEXTO, False),
    # ADR-035: a cadeia de decisao inteira, calculada no motor e nao em DAX.
    ("Provedor_EQA", TEXTO, False), ("Ano_EQA", TEXTO, False),
    ("Rodada_EQA", TEXTO, False),
    ("DPM_Teorico", REAL, False), ("Yield_Teorico", REAL, False),
    ("Regra_Westgard_Recomendada", TEXTO, False),
    ("N_Controle_Recomendado", INT, False),
    ("RunSize_Max_Recomendado", INT, False),
    ("Frequencia_QC_Descricao", TEXTO, False),
    ("Cobertura_Motor_Westgard", TEXTO, False),
    ("Referencia_Plano_QC", TEXTO, False),
]

# colunas acrescentadas pelo Power Query (nao existem no Excel)
COLUNAS_DERIVADAS = [
    ("Nivel_Rotulo", TEXTO, False),
    ("Especificacao_Efetiva", TEXTO, False),
]

# Colunas cujo nome no modelo difere do nome que o Power Query entrega.
# A chave e o nome exibido; o valor e o sourceColumn real.
RENOMEAR = {"Sigma_Obs": "Sigma"}

# M nao aceita todos os tipos direto; mapa para o Table.TransformColumnTypes
M_TIPO = {TEXTO: "type text", INT: "Int64.Type", REAL: "type number", DATA: "type date"}


def m_transform_fato():
    """Lista de tipagem do Power Query. Vigencia_* fica fora de proposito:
    vem vazia do motor na maioria das linhas e uma conversao para data
    quebraria a consulta inteira por causa de um registro mal formado."""
    fora = {"Vigencia_Inicio", "Vigencia_Fim", "AtualizadoEmUTC"}
    partes = []
    for nome, tipo, _ in COLUNAS_FATO:
        if nome in fora:
            continue
        partes.append('{{"{0}", {1}}}'.format(RENOMEAR.get(nome, nome), M_TIPO[tipo]))
    linhas, atual = [], "        "
    for i, p in enumerate(partes):
        virgula = "," if i < len(partes) - 1 else ""
        if len(atual) + len(p) > 76:
            linhas.append(atual.rstrip())
            atual = "        "
        atual += p + virgula + " "
    linhas.append(atual.rstrip())
    return "\n".join(linhas)


# ------------------------------------------------------------ consultas (M)

def m_fato():
    return '''let
    Fonte = Excel.Workbook(File.Contents(pCaminhoQC), null, true),
    // Le o ListObject PELO NOME, nunca por faixa de celulas: e o que permite
    // o banco crescer sem tocar na consulta. A aba BI_Data e xlVeryHidden --
    // isso nao esconde a tabela do Excel.Workbook, que enxerga o arquivo em
    // disco e nao a janela do Excel.
    Tabela = Fonte{[Item="tblBI_Fato", Kind="Table"]}[Data],
    Tipada = Table.TransformColumnTypes(Tabela, {
%s
    }),
    // Exclusao logica e decisao do laboratorio, com trilha de auditoria.
    // O BI a respeita; nao a revisita.
    Ativos = Table.SelectRows(Tipada, each [Ativo] = 1),
    // O motor grava Area em duas grafias ("Bioquimica" e "Bioquimica" com
    // acento), o que produziria dois itens identicos no filtro. Normaliza na
    // camada de apresentacao -- o motor nao se altera por causa de visual.
    AreaNorm = Table.TransformColumns(Ativos, {{"Area", each
        if _ = null then null
        else if Text.Contains(_, "Bioquimica") then "Bioquímica"
        else _, type text}}),
    // Nivel como rotulo: em visual, o numero 1 vira medida e o Power BI
    // tenta soma-lo.
    ComRotulo = Table.AddColumn(AreaNorm, "Nivel_Rotulo",
        each "N" & Text.From([Nivel]), type text),
    // Fonte_Especificacao vem preenchida mesmo onde Situacao e NAO
    // CADASTRADA. Exibi-la assim afirmaria um criterio de qualidade que nao
    // existe -- ausencia de meta nunca e aprovacao (ADR-023).
    ComEspec = Table.AddColumn(ComRotulo, "Especificacao_Efetiva", each
        // A meta existe quando ha ETp e fonte, nao quando a Situacao tem um
        // texto especifico: o ADR-028 trocou "CADASTRADA" por
        // "ANALITOS_VIGENTE" e a comparacao literal rotulou TODOS os
        // analitos como "Sem meta". Mesmo criterio de [Grupos com Meta].
        if [ETp_pct] <> null
           and [Fonte_Especificacao] <> null
           and Text.Trim(Text.From([Fonte_Especificacao])) <> ""
        then [Fonte_Especificacao] else "Sem meta", type text)
in
    ComEspec''' % m_transform_fato()


M_DIM_DATA = '''let
    // Cobre exatamente o intervalo dos dados. Um calendario 1900-2100 infla
    // o modelo sem servir a nada.
    Min = List.Min(Fato_QC[Data]),
    Max = List.Max(Fato_QC[Data]),
    FimMes = Date.EndOfMonth(Max),
    N = Duration.Days(FimMes - Min) + 1,
    Dias = List.Dates(Min, N, #duration(1,0,0,0)),
    Tab = Table.FromList(Dias, Splitter.SplitByNothing(), {"Data"}),
    Tipada = Table.TransformColumnTypes(Tab, {{"Data", type date}}),
    C1 = Table.AddColumn(Tipada, "Ano", each Date.Year([Data]), Int64.Type),
    C2 = Table.AddColumn(C1, "Mes", each Date.Month([Data]), Int64.Type),
    C3 = Table.AddColumn(C2, "Trimestre", each Date.QuarterOfYear([Data]), Int64.Type),
    C4 = Table.AddColumn(C3, "Competencia", each Date.ToText([Data], "yyyy-MM"), type text),
    C5 = Table.AddColumn(C4, "MesNome", each Date.ToText([Data], "MMM", "pt-BR"), type text),
    C6 = Table.AddColumn(C5, "TrimestreRotulo", each "T" & Text.From(Date.QuarterOfYear([Data])), type text),
    C7 = Table.AddColumn(C6, "AnoMes", each Date.Year([Data]) * 100 + Date.Month([Data]), Int64.Type)
in
    C7'''

M_DIM_ANALITO = '''let
    Base = Table.SelectColumns(Fato_QC, {"ID_Analito", "Analito", "Area", "Unidade"}),
    Unicos = Table.Distinct(Base, {"ID_Analito"}),
    Ord = Table.Sort(Unicos, {{"Analito", Order.Ascending}})
in
    Ord'''

M_DIM_LOTE = '''let
    Base = Table.SelectColumns(Fato_QC, {"ID_Lote"}),
    Unicos = Table.Distinct(Base, {"ID_Lote"}),
    // QC-897401 e QC-897402 sao o MESMO lote 8974 em dois niveis: o codigo
    // embute o nivel (ADR-024). Rotular a dimensao com o codigo por nivel
    // faria o filtro de lote se comportar como filtro de nivel.
    Rot = Table.AddColumn(Unicos, "Lote", each "Lote " & [ID_Lote], type text)
in
    Rot'''

M_DIM_NIVEL = '''let
    Base = Table.SelectColumns(Fato_QC, {"Nivel", "Nivel_Rotulo"}),
    Unicos = Table.Distinct(Base, {"Nivel"}),
    Ord = Table.Sort(Unicos, {{"Nivel", Order.Ascending}})
in
    Ord'''


# --------------------------------------------------------------- medidas DAX
# (nome, DAX, formatString, pasta)
# Regra que atravessa tudo: onde tblBI_Fato ja traz o valor calculado pelo
# motor, a medida LE esse valor -- nao recalcula. Recalcular criaria um
# segundo numero para a mesma pergunta, que e exatamente o defeito que o
# ADR-019 existe para impedir.
#
# AVERAGEX sobre SUMMARIZE(analito, nivel, lote) e o padrao usado nas metricas
# materializadas: como o motor repete o mesmo valor em todas as linhas do
# grupo, a media simples ponderaria pelo numero de resultados. Com um grupo
# selecionado o resultado e identico ao do Excel; com varios, e a media entre
# analitos, que e a leitura correta de um agregado.
GRUPO = "SUMMARIZE ( Fato_QC, Fato_QC[ID_Analito], Fato_QC[Nivel], Fato_QC[ID_Lote] )"


def por_grupo(coluna, agregado="AVERAGE"):
    return ("AVERAGEX (\n    %s,\n    CALCULATE ( %s ( Fato_QC[%s] ) )\n)"
            % (GRUPO, agregado, coluna))


MEDIDAS = [
    # ---- contagens
    ("n Resultados", "COUNTROWS ( Fato_QC )", "#,0", "01 Contagens"),
    ("n Corridas", "DISTINCTCOUNT ( Fato_QC[ID_Corrida] )", "#,0", "01 Contagens"),
    ("n Analitos", "DISTINCTCOUNT ( Fato_QC[ID_Analito] )", "#,0", "01 Contagens"),
    ("n Lotes", "DISTINCTCOUNT ( Fato_QC[ID_Lote] )", "#,0", "01 Contagens"),
    ("N Observado", "SUMX (\n    %s,\n    CALCULATE ( MAX ( Fato_QC[N_Observado] ) )\n)" % GRUPO,
     "#,0", "01 Contagens"),

    # ---- alvo do lote e ultimo resultado
    ("Media Alvo", "AVERAGE ( Fato_QC[Media_Alvo] )", "#,0.000", "02 Descritiva"),
    ("DP Alvo", "AVERAGE ( Fato_QC[DP_Alvo] )", "#,0.000", "02 Descritiva"),
    ("Ultimo Resultado",
     "VAR ultRun = MAX ( Fato_QC[RUN] )\n"
     "RETURN CALCULATE ( AVERAGE ( Fato_QC[Resultado] ), Fato_QC[RUN] = ultRun )",
     "#,0.000", "02 Descritiva"),
    ("Z Score Atual",
     "VAR ultRun = MAX ( Fato_QC[RUN] )\n"
     "RETURN CALCULATE ( AVERAGE ( Fato_QC[Z] ), Fato_QC[RUN] = ultRun )",
     "#,0.00", "02 Descritiva"),

    # ---- metricas observadas (materializadas pelo motor)
    ("Media Observada", por_grupo("Media_Observada"), "#,0.000", "02 Descritiva"),
    ("DP Observado", por_grupo("DP_Observado"), "#,0.0000", "02 Descritiva"),
    ("CV Observado %", por_grupo("CV_Observado_pct"), "#,0.00", "02 Descritiva"),
    ("Bias Observado %", por_grupo("Bias_Observado_pct"), "#,0.00", "02 Descritiva"),
    ("ET Observado %", por_grupo("ET_Observado_pct"), "#,0.00", "02 Descritiva"),
    ("Sigma", por_grupo("Sigma_Obs"), "#,0.00", "02 Descritiva"),

    # ---- limites da especificacao vigente (respeita a escolha do Excel)
    ("CVtp %", por_grupo("CVtp_pct"), "#,0.00", "03 Especificacao"),
    ("BIAStp %", por_grupo("BIAStp_pct"), "#,0.00", "03 Especificacao"),
    ("ETp %", por_grupo("ETp_pct"), "#,0.00", "03 Especificacao"),
]


MEDIDAS += [
    # ---- conformidade: observado x especificacao vigente
    # Ausencia de meta NUNCA e aprovacao. "Sem meta" e um terceiro estado,
    # nao um "Conforme" por omissao (ADR-023).
    ("Status CV",
     'VAR lim = [CVtp %]\nVAR obs = [CV Observado %]\n'
     'RETURN SWITCH ( TRUE (), ISBLANK ( obs ), "Sem dado", ISBLANK ( lim ), "Sem meta", '
     'obs <= lim * 0.8, "Conforme", obs <= lim, "Proximo ao limite", "Nao conforme" )',
     None, "04 Conformidade"),
    ("Status Bias",
     'VAR lim = [BIAStp %]\nVAR obs = ABS ( [Bias Observado %] )\n'
     'RETURN SWITCH ( TRUE (), ISBLANK ( obs ), "Sem dado", ISBLANK ( lim ), "Sem meta", '
     'obs <= lim * 0.8, "Conforme", obs <= lim, "Proximo ao limite", "Nao conforme" )',
     None, "04 Conformidade"),
    ("Status ET",
     'VAR lim = [ETp %]\nVAR obs = [ET Observado %]\n'
     'RETURN SWITCH ( TRUE (), ISBLANK ( obs ), "Sem dado", ISBLANK ( lim ), "Sem meta", '
     'obs <= lim * 0.8, "Conforme", obs <= lim, "Proximo ao limite", "Nao conforme" )',
     None, "04 Conformidade"),
    ("Folga ET %",
     "VAR lim = [ETp %]\nVAR obs = [ET Observado %]\n"
     "RETURN IF ( NOT ISBLANK ( lim ) && NOT ISBLANK ( obs ), lim - obs )",
     "#,0.00", "04 Conformidade"),
    # conta pares (analito,nivel,lote) e nao linhas: 110 resultados do mesmo
    # analito nao sao 110 nao conformidades.
    ("Grupos com Meta",
     "COUNTROWS ( FILTER ( %s, NOT ISBLANK ( CALCULATE ( AVERAGE ( Fato_QC[ETp_pct] ) ) ) ) )" % GRUPO,
     "#,0", "04 Conformidade"),
    ("Grupos ET Nao Conforme",
     "COUNTROWS ( FILTER ( %s,\n"
     "    VAR l = CALCULATE ( AVERAGE ( Fato_QC[ETp_pct] ) )\n"
     "    VAR o = CALCULATE ( AVERAGE ( Fato_QC[ET_Observado_pct] ) )\n"
     "    RETURN NOT ISBLANK ( l ) && NOT ISBLANK ( o ) && o > l ) )" % GRUPO,
     "#,0", "04 Conformidade"),
    ("% Grupos ET Conforme",
     "VAR com = [Grupos com Meta]\nVAR nc = [Grupos ET Nao Conforme]\n"
     "RETURN IF ( com > 0, DIVIDE ( com - nc, com ) * 100 )",
     "#,0.0", "04 Conformidade"),

    # ---- Westgard e veredito
    ("n Rejeitados", 'CALCULATE ( COUNTROWS ( Fato_QC ), Fato_QC[Veredito] = "REJEITADO" )',
     "#,0", "05 Westgard"),
    ("n Aceitos", 'CALCULATE ( COUNTROWS ( Fato_QC ), Fato_QC[Veredito] = "OK" )',
     "#,0", "05 Westgard"),
    ("% Aceitaveis",
     "VAR n = [n Resultados]\nRETURN IF ( n > 0, DIVIDE ( n - [n Rejeitados], n ) * 100 )",
     "#,0.0", "05 Westgard"),
    ("% Fora de Controle", "IF ( [n Resultados] > 0, 100 - [% Aceitaveis] )",
     "#,0.0", "05 Westgard"),
    ("Viol 1_3s", "SUM ( Fato_QC[W_1_3s] )", "#,0", "05 Westgard"),
    ("Viol 2_2s", "SUM ( Fato_QC[W_2_2s] )", "#,0", "05 Westgard"),
    ("Viol R_4s", "SUM ( Fato_QC[W_R_4s] )", "#,0", "05 Westgard"),
    ("Viol 4_1s", "SUM ( Fato_QC[W_4_1s] )", "#,0", "05 Westgard"),
    ("Viol 10x", "SUM ( Fato_QC[W_10x] )", "#,0", "05 Westgard"),
    ("Viol Total",
     "[Viol 1_3s] + [Viol 2_2s] + [Viol R_4s] + [Viol 4_1s] + [Viol 10x]",
     "#,0", "05 Westgard"),
    ("Alertas 1_2s", "SUM ( Fato_QC[A_1_2s] )", "#,0", "05 Westgard"),
    # 4_1s e 10x existem como coluna mas o motor ainda nao as avalia -- ver
    # BI_ARCHITECTURE.md. A medida diz isso em vez de exibir um zero que o
    # gestor leria como "nunca violamos".
    ("Analitos Criticos",
     'CALCULATE ( DISTINCTCOUNT ( Fato_QC[ID_Analito] ), Fato_QC[Veredito] = "REJEITADO" )',
     "#,0", "05 Westgard"),

    # ---- Levey-Jennings
    ("LJ Media", "AVERAGE ( Fato_QC[Media_Alvo] )", "#,0.000", "06 Levey-Jennings"),
    ("LJ +1s", "AVERAGE ( Fato_QC[Lim_p1s] )", "#,0.000", "06 Levey-Jennings"),
    ("LJ +2s", "AVERAGE ( Fato_QC[Lim_p2s] )", "#,0.000", "06 Levey-Jennings"),
    ("LJ +3s", "AVERAGE ( Fato_QC[Lim_p3s] )", "#,0.000", "06 Levey-Jennings"),
    ("LJ -1s", "AVERAGE ( Fato_QC[Lim_m1s] )", "#,0.000", "06 Levey-Jennings"),
    ("LJ -2s", "AVERAGE ( Fato_QC[Lim_m2s] )", "#,0.000", "06 Levey-Jennings"),
    ("LJ -3s", "AVERAGE ( Fato_QC[Lim_m3s] )", "#,0.000", "06 Levey-Jennings"),
    ("LJ Resultado", "AVERAGE ( Fato_QC[Resultado] )", "#,0.000", "06 Levey-Jennings"),
    # Com o filtro aberto, a media dos limites de analitos diferentes nao
    # significa nada. E preferivel dizer "nao da para mostrar" a desenhar uma
    # curva sem sentido.
    ("LJ Valido",
     "IF ( HASONEVALUE ( Dim_Analito[ID_Analito] ) && HASONEVALUE ( Dim_Nivel[Nivel] ) "
     "&& HASONEVALUE ( Dim_Lote[ID_Lote] ), 1, 0 )", "0", "06 Levey-Jennings"),
    ("LJ Aviso",
     'IF ( [LJ Valido] = 0, "Selecione UM analito, UM nivel e UM lote para ver o '
     'Levey-Jennings." )', None, "06 Levey-Jennings"),

    # ---- tendencia
    ("Deslocamento vs Alvo (SD)",
     "VAR d = [Media Observada] - [Media Alvo]\nVAR sd = AVERAGE ( Fato_QC[DP_Alvo] )\n"
     "RETURN IF ( sd > 0, DIVIDE ( d, sd ) )", "#,0.00", "07 Tendencia"),
    ("Media Movel 7 Corridas",
     "VAR corridaAtual = SELECTEDVALUE ( Fato_QC[RUN] )\n"
     "RETURN CALCULATE ( AVERAGE ( Fato_QC[Resultado] ), FILTER ( ALLSELECTED ( Fato_QC ), "
     "Fato_QC[RUN] <= corridaAtual && Fato_QC[RUN] > corridaAtual - 7 ) )",
     "#,0.000", "07 Tendencia"),

    # ---- resumo executivo
    ("Data Mais Recente", "MAX ( Fato_QC[Data] )", None, "08 Executivo"),
    ("n Hoje", "CALCULATE ( [n Resultados], Fato_QC[Data] = [Data Mais Recente] )",
     "#,0", "08 Executivo"),
    ("Rejeitados Hoje", "CALCULATE ( [n Rejeitados], Fato_QC[Data] = [Data Mais Recente] )",
     "#,0", "08 Executivo"),
    ("Grupos Sem Meta",
     "COUNTROWS ( FILTER ( %s, ISBLANK ( CALCULATE ( AVERAGE ( Fato_QC[ETp_pct] ) ) ) ) )" % GRUPO,
     "#,0", "08 Executivo"),
    # Cortes da pratica consolidada da metrica Sigma: >= 4 bom, 3 a <4
    # aceitavel com monitoramento reforcado, < 3 inaceitavel. Ficam em UM
    # lugar so, para que mudar a politica seja uma edicao e nao uma cacada.
    ("Status Global",
     "VAR s = [Sigma]\nVAR fora = [% Fora de Controle]\n"
     "RETURN SWITCH ( TRUE (),\n"
     '    ISBLANK ( [n Resultados] ), "SEM DADOS",\n'
     '    ISBLANK ( s ), "SEM META",\n'
     '    s < 3 || fora > 5, "FORA DE CONTROLE",\n'
     '    s < 4 || fora > 2, "ATENCAO",\n'
     '    "SOB CONTROLE" )', None, "08 Executivo"),
    ("Cor Status Global",
     'SWITCH ( [Status Global], "FORA DE CONTROLE", "#C0392B", "ATENCAO", "#D68910", '
     '"SOB CONTROLE", "#1E8449", "#7F8C8D" )', None, "08 Executivo"),
    ("Periodo Analisado",
     'VAR a = MIN ( Fato_QC[Data] )\nVAR b = MAX ( Fato_QC[Data] )\n'
     'RETURN IF ( NOT ISBLANK ( a ), FORMAT ( a, "dd/mm/yyyy" ) & "  ate  " & '
     'FORMAT ( b, "dd/mm/yyyy" ) )', None, "08 Executivo"),
    ("Atualizado Em",
     'VAR t = MAX ( Fato_QC[AtualizadoEmUTC] )\n'
     'RETURN IF ( NOT ISBLANK ( t ), "Dados do QC_INI de " & t & " (UTC)" )',
     None, "08 Executivo"),
]

# coluna calculada de faixa Sigma: classificacao para exibicao, nao
# recalculo do motor. O Sigma ja vem materializado por (analito, nivel, lote).
COLUNAS_CALCULADAS = [
    ("Faixa_Sigma",
     'SWITCH ( TRUE (), ISBLANK ( Fato_QC[Sigma_Obs] ), "Sem meta", Fato_QC[Sigma_Obs] < 3, "< 3", '
     'Fato_QC[Sigma_Obs] < 4, "3 a <4", Fato_QC[Sigma_Obs] < 5, "4 a <5", Fato_QC[Sigma_Obs] < 6, '
     '"5 a <6", ">= 6" )', TEXTO),
    ("Faixa_Sigma_Ordem",
     "SWITCH ( TRUE (), ISBLANK ( Fato_QC[Sigma_Obs] ), 9, Fato_QC[Sigma_Obs] < 3, 1, "
     "Fato_QC[Sigma_Obs] < 4, 2, Fato_QC[Sigma_Obs] < 5, 3, Fato_QC[Sigma_Obs] < 6, 4, 5 )", INT),
]



# ------------------------------------------------------------- emissao TMDL
# TMDL e sensivel a indentacao e usa TAB. Todo bloco aninhado desce um TAB.

def bloco(texto, niveis):
    """Reindenta um trecho multilinha para dentro de um bloco TMDL."""
    tab = "\t" * niveis
    return "\n".join(tab + l if l.strip() else "" for l in texto.split("\n"))


def tmdl_coluna(tabela, nome, tipo, oculta, coluna_fonte=None, dax=None,
                ordenar_por=None):
    # Coluna calculada: em TMDL o DAX mora no CABECALHO ("column X = expr"),
    # nunca numa propriedade "expression:". A forma de propriedade existe na
    # gramatica mas exige valor inline e recusa bloco multilinha.
    if dax and "\n" not in dax:
        # DAX de uma linha vai INLINE no cabecalho. Num bloco recuado ele
        # ficaria no mesmo nivel das propriedades e o analisador do TMDL
        # engoliria "dataType:", "lineageTag:" e "sortByColumn:" para dentro
        # da formula -- o modelo ate carrega, mas a coluna calculada nasce em
        # erro e todo visual que a usa aparece quebrado.
        ln = ["\tcolumn %s = %s" % (nome, dax)]
    elif dax:
        # multilinha exige bloco delimitado por crase tripla
        ln = ["\tcolumn %s = ```" % nome, bloco(dax, 2), "\t\t```"]
    else:
        ln = ["\tcolumn %s" % nome]
    ln.append("\t\tdataType: %s" % tipo)
    if oculta:
        ln.append("\t\tisHidden")
    ln.append("\t\tlineageTag: %s" % guid("col/%s/%s" % (tabela, nome)))
    ln.append("\t\tsummarizeBy: none")
    if not dax:
        ln.append("\t\tsourceColumn: %s" % (coluna_fonte or nome))
    if ordenar_por:
        ln.append("\t\tsortByColumn: %s" % ordenar_por)
    if tipo == DATA:
        ln.append('\t\tformatString: dd/mm/yyyy')
    ln.append("")
    ln.append("\t\tannotation SummarizationSetBy = Automatic")
    ln.append("")
    return "\n".join(ln)


def tmdl_medida(nome, dax, formato, pasta):
    seguro = "'%s'" % nome if (" " in nome or "%" in nome or "(" in nome) else nome
    ln = ["\tmeasure %s =" % seguro]
    ln.append(bloco(dax, 3))
    ln.append("")
    if formato:
        ln.append("\t\tformatString: %s" % formato)
    ln.append("\t\tlineageTag: %s" % guid("med/%s" % nome))
    if pasta:
        ln.append("\t\tdisplayFolder: %s" % pasta)
    ln.append("")
    ln.append("\t\tannotation PBI_FormatHint = {\"isGeneralNumber\":true}"
              if not formato else "\t\tannotation PBI_FormatHint = {}")
    ln.append("")
    return "\n".join(ln)


def tmdl_particao_m(tabela, codigo_m):
    ln = ["\tpartition %s = m" % tabela]
    ln.append("\t\tmode: import")
    ln.append("\t\tsource =")
    ln.append(bloco(codigo_m, 4))
    ln.append("")
    return "\n".join(ln)


def tmdl_fato():
    p = ["table Fato_QC", "\tlineageTag: %s" % guid("tab/Fato_QC"), ""]
    for nome, tipo, oculta in COLUNAS_FATO + COLUNAS_DERIVADAS:
        p.append(tmdl_coluna("Fato_QC", nome, tipo, oculta,
                             coluna_fonte=RENOMEAR.get(nome, nome)))
    for nome, dax, tipo in COLUNAS_CALCULADAS:
        ordem = "Faixa_Sigma_Ordem" if nome == "Faixa_Sigma" else None
        p.append(tmdl_coluna("Fato_QC", nome, tipo,
                             nome.endswith("_Ordem"), dax=dax, ordenar_por=ordem))
    for nome, dax, fmt, pasta in MEDIDAS:
        p.append(tmdl_medida(nome, dax, fmt, pasta))
    p.append(tmdl_particao_m("Fato_QC", m_fato()))
    p.append("\tannotation PBI_ResultType = Table")
    p.append("")
    return "\n".join(p)


def tmdl_dim(nome, colunas, codigo_m, chave=None):
    p = ["table %s" % nome, "\tlineageTag: %s" % guid("tab/%s" % nome), ""]
    for c_nome, c_tipo, c_oculta, c_ordem in colunas:
        p.append(tmdl_coluna(nome, c_nome, c_tipo, c_oculta, ordenar_por=c_ordem))
    p.append(tmdl_particao_m(nome, codigo_m))
    p.append("\tannotation PBI_ResultType = Table")
    p.append("")
    return "\n".join(p)


def tmdl_medidas_vazia():
    """Tabela so para hospedar nada -- as medidas moram em Fato_QC, onde o
    contexto de filtro delas de fato vive. Uma tabela '_Medidas' separada e
    cosmetica e some do painel de campos quando o usuario oculta a coluna."""
    return None


def tmdl_modelo():
    return """model Model
\tculture: pt-BR
\tdefaultPowerBIDataSourceVersion: powerBI_V3
\tsourceQueryCulture: pt-BR
\tdataAccessOptions
\t\tlegacyRedirects
\t\treturnErrorValuesAsNull

annotation PBI_QueryOrder = ["pCaminhoQC","Fato_QC","Dim_Data","Dim_Analito","Dim_Lote","Dim_Nivel"]

annotation PBI_ProTooling = ["DevMode"]

ref table Fato_QC
ref table Dim_Data
ref table Dim_Analito
ref table Dim_Lote
ref table Dim_Nivel

ref expression pCaminhoQC

ref relationship Fato_Data
ref relationship Fato_Analito
ref relationship Fato_Lote
ref relationship Fato_Nivel
"""


def tmdl_relacoes():
    # Todas simples, nenhuma bidirecional. Filtro cruzado bidirecional numa
    # estrela cria caminhos ambiguos e medidas que mudam de valor conforme o
    # visual -- num painel de CQ isso significa dois numeros diferentes para
    # a mesma pergunta.
    rels = [
        ("Fato_Data", "Fato_QC", "Data", "Dim_Data", "Data"),
        ("Fato_Analito", "Fato_QC", "ID_Analito", "Dim_Analito", "ID_Analito"),
        ("Fato_Lote", "Fato_QC", "ID_Lote", "Dim_Lote", "ID_Lote"),
        ("Fato_Nivel", "Fato_QC", "Nivel", "Dim_Nivel", "Nivel"),
    ]
    p = []
    for nome, tf, cf, tt, ct in rels:
        p.append("relationship %s" % nome)
        p.append("\tfromColumn: %s.%s" % (tf, cf))
        p.append("\ttoColumn: %s.%s" % (tt, ct))
        p.append("")
    return "\n".join(p)


def tmdl_expressao_parametro(caminho_xlsm):
    literal = caminho_xlsm.replace('"', '""')
    return ('expression pCaminhoQC = "%s" meta [IsParameterQuery=true, Type="Text", '
            'IsParameterQueryRequired=true]\n'
            '\tlineageTag: %s\n'
            '\n'
            '\tannotation PBI_NavigationStepName = Navegação\n'
            '\n'
            '\tannotation PBI_ResultType = Text\n' % (literal, guid("expr/pCaminhoQC")))


# ------------------------------------------------------------ paginas (JSON)
# paleta institucional: sobria, alto contraste, sem decoracao
AZUL = "#1F3864"
AZUL_CLARO = "#2E75B6"
CINZA = "#595959"
CINZA_CLARO = "#F2F2F2"
VERDE = "#1E8449"
AMBAR = "#D68910"
VERMELHO = "#C0392B"

# Series do Levey-Jennings: o resultado em destaque, os limites em tons que
# recuam. Se os limites competirem visualmente com o ponto medido, o grafico
# vira um emaranhado e deixa de responder a unica pergunta que importa.
COR_LJ = {
    "LJ Resultado": AZUL_CLARO,
    "LJ Media": "#7F7F7F",
    "LJ +1s": "#BFBFBF", "LJ -1s": "#BFBFBF",
    "LJ +2s": AMBAR, "LJ -2s": AMBAR,
    "LJ +3s": VERMELHO, "LJ -3s": VERMELHO,
}

INCLUIR = {"texto", "cartao", "slicer", "grafico", "tabela", "cores"}

TIPO_CLASSE = {"textbox": "texto", "card": "cartao", "slicer": "slicer",
               "tableEx": "tabela"}


def _quer(classe):
    return classe in INCLUIR


def _lit(v):
    """Literal do formato de expressao do Power BI."""
    if isinstance(v, bool):
        return {"expr": {"Literal": {"Value": "true" if v else "false"}}}
    if isinstance(v, (int, float)):
        return {"expr": {"Literal": {"Value": "%sD" % v}}}
    return {"expr": {"Literal": {"Value": "'%s'" % v}}}


def _cor(hexa):
    return {"solid": {"color": _lit(hexa)}}


class Campo(object):
    """Referencia a uma coluna ou medida, na forma que a prototypeQuery pede."""

    def __init__(self, entidade, propriedade, medida=False, rotulo=None):
        self.entidade = entidade
        self.propriedade = propriedade
        self.medida = medida
        self.rotulo = rotulo or propriedade

    @property
    def ref(self):
        return "%s.%s" % (self.entidade, self.propriedade)

    def select(self, alias):
        chave = "Measure" if self.medida else "Column"
        return {
            chave: {"Expression": {"SourceRef": {"Source": alias}},
                    "Property": self.propriedade},
            "Name": self.ref,
            "NativeReferenceName": self.rotulo,
        }


def C(ent, prop, rotulo=None):
    return Campo(ent, prop, False, rotulo)


def M(prop, rotulo=None):
    return Campo("Fato_QC", prop, True, rotulo)


def _consulta(campos, ordenar=None):
    """Monta From/Select a partir dos campos, dando um alias por entidade."""
    alias, frm = {}, []
    for c in campos:
        if c.entidade not in alias:
            a = c.entidade[0].lower() + str(len(alias))
            alias[c.entidade] = a
            frm.append({"Name": a, "Entity": c.entidade, "Type": 0})
    q = {"Version": 2, "From": frm,
         "Select": [c.select(alias[c.entidade]) for c in campos]}
    if ordenar:
        campo, desc = ordenar
        chave = "Measure" if campo.medida else "Column"
        q["OrderBy"] = [{
            "Direction": 2 if desc else 1,
            "Expression": {chave: {
                "Expression": {"SourceRef": {"Source": alias[campo.entidade]}},
                "Property": campo.propriedade}},
        }]
    return q


# ===========================================================================
# EMISSAO DO RELATORIO -- formato PBIR
#
# O Power BI Desktop 2.156 IGNORA em silencio o "report.json" legado de
# pagina unica: abre o projeto, nao acusa erro e mostra um relatorio sem
# nenhuma pagina. O formato aceito e o PBIR, uma arvore de arquivos:
#
#   definition/report.json
#   definition/pages/pages.json
#   definition/pages/<pagina>/page.json
#   definition/pages/<pagina>/visuals/<visual>/visual.json
#
# A consulta tambem muda de forma: em vez de prototypeQuery + projections,
# cada visual declara query.queryState.<papel>.projections[].field.
#
# Limitacao conhecida do PBIR 1.0.0: nao ha layout de telefone no formato.
# Ele existe so na interface do Desktop ("Exibicao para telefone").
# ===========================================================================

ESQ = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition"
LARG, ALT = 1280, 720


def _id(semente):
    """Nome estavel e curto: o schema limita a 50 caracteres e o nome precisa
    ser unico dentro do relatorio."""
    return guid(semente).replace("-", "")[:20]


def _campo_pbir(c):
    """QueryExpressionContainer. No PBIR nao ha clausula From, entao o
    SourceRef aponta direto para a entidade."""
    chave = "Measure" if c.medida else "Column"
    return {chave: {"Expression": {"SourceRef": {"Entity": c.entidade}},
                    "Property": c.propriedade}}


def visual(nome, tipo, x, y, w, h, papeis, titulo=None, objetos=None,
           z=0, celular=None, sem_titulo=False, filtros=None):
    """Um visual.json do PBIR. `celular` e aceito e ignorado: o formato nao
    carrega layout de telefone (ver cabecalho)."""
    if not _quer(TIPO_CLASSE.get(tipo, "grafico")):
        return None
    if not _quer("cores"):
        objetos = {k: v for k, v in (objetos or {}).items() if k != "dataPoint"}

    cfg = {"visualType": tipo}

    estado = {}
    for papel, lista in papeis.items():
        if not lista:
            continue
        estado[papel] = {"projections": [
            {"field": _campo_pbir(c), "queryRef": c.ref,
             "nativeQueryRef": c.rotulo} for c in lista]}
    if estado:
        cfg["query"] = {"queryState": estado}
    if objetos:
        cfg["objects"] = objetos

    vco = {}
    if titulo:
        vco["title"] = [{"properties": {
            "show": _lit(True), "text": _lit(titulo),
            "fontColor": _cor(AZUL), "fontSize": _lit(11),
            "bold": _lit(True), "alignment": _lit("left")}}]
    elif sem_titulo:
        vco["title"] = [{"properties": {"show": _lit(False)}}]
    vco["background"] = [{"properties": {
        "show": _lit(True), "color": _cor("#FFFFFF"), "transparency": _lit(0)}}]
    vco["border"] = [{"properties": {
        "show": _lit(True), "color": _cor("#E1E5EA"), "radius": _lit(6)}}]
    cfg["visualContainerObjects"] = vco

    return {
        "$schema": ESQ + "/visualContainer/1.0.0/schema.json",
        "name": _id("vis/" + nome),
        "position": {"x": x, "y": y, "z": z, "width": w, "height": h,
                     "tabOrder": z * 1000 + x},
        "visual": cfg,
    }


def texto(nome, x, y, w, h, corridas, z=0, celular=None, fundo=None):
    """Caixa de texto. corridas: lista de (texto, tamanho_pt, negrito, cor)."""
    if not _quer("texto"):
        return None
    runs = [{"value": t, "textStyle": {
        "fontSize": "%dpt" % pt, "fontWeight": "bold" if neg else "normal",
        "color": cor, "fontFamily": "Segoe UI"}} for t, pt, neg, cor in corridas]
    return {
        "$schema": ESQ + "/visualContainer/1.0.0/schema.json",
        "name": _id("vis/" + nome),
        "position": {"x": x, "y": y, "z": z, "width": w, "height": h,
                     "tabOrder": z * 1000 + x},
        "visual": {
            "visualType": "textbox",
            "objects": {"general": [{"properties": {
                "paragraphs": [{"textRuns": runs}]}}]},
            "visualContainerObjects": {
                "title": [{"properties": {"show": _lit(False)}}],
                "background": [{"properties": {
                    "show": _lit(bool(fundo)), "color": _cor(fundo or "#FFFFFF"),
                    "transparency": _lit(0)}}]},
        },
    }


def secao(ordem, titulo, visuais):
    return {"ordem": ordem, "titulo": titulo,
            "nome": _id("sec/" + titulo),
            "visuais": [v for v in visuais if v is not None]}


def escrever_relatorio(relat, secoes):
    plataforma(os.path.join(relat, ".platform"), "Report", NOME)
    escrever_json(os.path.join(relat, "definition.pbir"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/"
                   "item/report/definitionProperties/1.0.0/schema.json",
        "version": "4.0",
        "datasetReference": {"byPath": {"path": "../" + NOME + ".SemanticModel"}},
    })
    defs = os.path.join(relat, "definition")
    # version.json e OBRIGATORIO no PBIR e nao aparece na lista de schemas do
    # relatorio: sem ele o Power BI Desktop abre e MORRE com
    # "Cannot find file 'version.json'". O formato e major.minor.patch com
    # patch sempre 0.
    escrever_json(os.path.join(defs, "version.json"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/"
                   "item/report/definition/versionMetadata/1.0.0/schema.json",
        "version": "2.0.0",
    })
    escrever_json(os.path.join(defs, "report.json"), {
        "$schema": ESQ + "/report/2.0.0/schema.json",
        "themeCollection": {"baseTheme": {
            "name": "CY24SU10", "reportVersionAtImport": "5.55",
            "type": "SharedResources"}},
        "settings": {"useStylableVisualContainerHeader": True,
                     "defaultFilterActionIsDataFilter": True},
    })
    pastas = os.path.join(defs, "pages")
    escrever_json(os.path.join(pastas, "pages.json"), {
        "$schema": ESQ + "/pagesMetadata/1.0.0/schema.json",
        "pageOrder": [s["nome"] for s in secoes],
        "activePageName": secoes[0]["nome"],
    })
    for s in secoes:
        pp = os.path.join(pastas, s["nome"])
        escrever_json(os.path.join(pp, "page.json"), {
            "$schema": ESQ + "/page/1.0.0/schema.json",
            "name": s["nome"],
            "displayName": s["titulo"],
            "displayOption": "FitToPage",
            "height": ALT, "width": LARG,
        })
        for v in s["visuais"]:
            escrever_json(
                os.path.join(pp, "visuals", v["name"], "visual.json"), v)


def cartao(nome, medida, rotulo, x, y, w, h, z=0, celular=None, pt=22):
    """Cartao de KPI. O rotulo vai no titulo do visual e nao na categoria:
    o cartao com categoria repete o nome da medida e come metade da altura."""
    obj = {
        "labels": [{"properties": {"color": _cor(AZUL), "fontSize": _lit(pt),
                                   "bold": _lit(True),
                                   "fontFamily": _lit("Segoe UI")}}],
        "categoryLabels": [{"properties": {"show": _lit(False)}}],
    }
    return visual(nome, "card", x, y, w, h, {"Values": [medida]},
                  titulo=rotulo, objetos=obj, z=z, celular=celular)


def filtro_slicer(nome, campo, x, y, w, h, modo="Dropdown", z=0, celular=None,
                  titulo=None, unico=False):
    obj = {
        "general": [{"properties": {"mode": _lit(modo)}}],
        "header": [{"properties": {"show": _lit(False)}}],
        "items": [{"properties": {"fontSize": _lit(9)}}],
    }
    if unico:
        obj["selection"] = [{"properties": {"singleSelect": _lit(True)}}]
    return visual(nome, "slicer", x, y, w, h, {"Values": [campo]},
                  titulo=titulo or campo.rotulo, objetos=obj, z=z,
                  celular=celular)

# --------------------------------------------------------- campos do modelo
F_ANALITO = C("Dim_Analito", "Analito")
F_NIVEL = C("Dim_Nivel", "Nivel_Rotulo", "Nível")
F_LOTE = C("Dim_Lote", "Lote")
F_ANO = C("Dim_Data", "Ano")
F_MES = C("Dim_Data", "MesNome", "Mês")
F_TRI = C("Dim_Data", "TrimestreRotulo", "Trimestre")
F_COMP = C("Dim_Data", "Competencia", "Competência")
F_DATA = C("Dim_Data", "Data")
F_RUN = C("Fato_QC", "RUN")
F_VEREDITO = C("Fato_QC", "Veredito")
F_FAIXA = C("Fato_QC", "Faixa_Sigma", "Faixa Sigma")
F_ESPEC = C("Fato_QC", "Especificacao_Efetiva", "Especificação")
F_SITUACAO = C("Fato_QC", "Situacao_Especificacao", "Situação")

SERIES_LJ = ["LJ Resultado", "LJ Media", "LJ +3s", "LJ +2s", "LJ +1s",
             "LJ -1s", "LJ -2s", "LJ -3s"]


def cores_series(nomes):
    """Cor por serie de medida. selector.metadata aponta para o queryRef --
    e assim que o formato legado amarra uma cor a uma medida especifica."""
    return {"dataPoint": [
        {"properties": {"fill": _cor(COR_LJ.get(n, AZUL_CLARO))},
         "selector": {"metadata": "Fato_QC." + n}} for n in nomes]}


# ---------------------------------------------------------------------------
# Medidas de contexto/apresentacao acrescentadas para o dashboard.
# Nenhuma delas recalcula estatistica: elas selecionam, contam ou rotulam o
# que o motor ja materializou em tblBI_Fato.
# ---------------------------------------------------------------------------

def _grupo_filtrado(condicao):
    return ("COUNTROWS ( FILTER ( %s,\n    VAR s = CALCULATE ( AVERAGE ( Fato_QC[Sigma_Obs] ) )\n"
            "    RETURN %s ) )" % (GRUPO, condicao))


def _pior(coluna, desc=True):
    """Nome do analito/nivel no extremo da metrica. TOPN sobre os grupos, nao
    sobre as linhas: um analito com 110 resultados nao deve pesar mais."""
    # O filtro de vazio nao e detalhe: TOPN devolve TODOS os empates, e os
    # grupos sem meta empatam todos em BLANK -- sem ele, "pior Sigma" lista
    # 22 analitos em vez de um.
    return (
        "VAR t = FILTER (\n"
        "    SUMMARIZE ( Fato_QC, Fato_QC[Analito], Fato_QC[Nivel] ),\n"
        "    NOT ISBLANK ( CALCULATE ( AVERAGE ( Fato_QC[%s] ) ) ) )\n"
        "VAR alvo = TOPN ( 1, t, CALCULATE ( AVERAGE ( Fato_QC[%s] ) ), %s )\n"
        'RETURN CONCATENATEX ( alvo, Fato_QC[Analito] & "  N" & Fato_QC[Nivel], ", " )'
        % (coluna, coluna, "DESC" if desc else "ASC"))


MEDIDAS += [
    # ---- rotulos de contexto (cartoes do Painel)
    ("Veredito Atual",
     "VAR ultRun = MAX ( Fato_QC[RUN] )\n"
     "RETURN CALCULATE ( CONCATENATEX ( VALUES ( Fato_QC[Veredito] ), "
     'Fato_QC[Veredito], " / " ), Fato_QC[RUN] = ultRun )',
     None, "02 Descritiva"),
    ("Situacao da Especificacao",
     "VAR n = DISTINCTCOUNT ( Fato_QC[Situacao_Especificacao] )\n"
     "RETURN IF ( n = 1, SELECTEDVALUE ( Fato_QC[Situacao_Especificacao] ), "
     'FORMAT ( n, "0" ) & " situações" )', None, "03 Especificacao"),
    ("Especificacao em Uso",
     "VAR n = DISTINCTCOUNT ( Fato_QC[Especificacao_Efetiva] )\n"
     "RETURN IF ( n = 1, SELECTEDVALUE ( Fato_QC[Especificacao_Efetiva] ), "
     'FORMAT ( n, "0" ) & " fontes" )', None, "03 Especificacao"),

    # ---- contagens por faixa Sigma (para a leitura executiva)
    ("Grupos Sigma < 3", _grupo_filtrado("NOT ISBLANK ( s ) && s < 3"),
     "#,0", "08 Executivo"),
    ("Grupos Sigma 3 a 4", _grupo_filtrado("NOT ISBLANK ( s ) && s >= 3 && s < 4"),
     "#,0", "08 Executivo"),
    ("Grupos Sigma >= 6", _grupo_filtrado("NOT ISBLANK ( s ) && s >= 6"),
     "#,0", "08 Executivo"),

    # ---- extremos: quem investigar primeiro
    ("Pior CV", _pior("CV_Observado_pct"), None, "08 Executivo"),
    ("Pior Bias", _pior("Bias_Observado_pct"), None, "08 Executivo"),
    ("Pior ET", _pior("ET_Observado_pct"), None, "08 Executivo"),
    ("Pior Sigma", _pior("Sigma_Obs", desc=False), None, "08 Executivo"),

    # ---- cores para formatacao condicional da matriz
    # Uma medida por status, e nao um SWITCH espalhado por cada visual: mudar
    # a politica de cor vira uma edicao unica.
    ("Cor Status CV",
     'SWITCH ( [Status CV], "Conforme", "%s", "Proximo ao limite", "%s", '
     '"Nao conforme", "%s", "%s" )' % (VERDE, AMBAR, VERMELHO, CINZA),
     None, "09 Cores"),
    ("Cor Status Bias",
     'SWITCH ( [Status Bias], "Conforme", "%s", "Proximo ao limite", "%s", '
     '"Nao conforme", "%s", "%s" )' % (VERDE, AMBAR, VERMELHO, CINZA),
     None, "09 Cores"),
    ("Cor Status ET",
     'SWITCH ( [Status ET], "Conforme", "%s", "Proximo ao limite", "%s", '
     '"Nao conforme", "%s", "%s" )' % (VERDE, AMBAR, VERMELHO, CINZA),
     None, "09 Cores"),
    ("Cor Sigma",
     "VAR s = [Sigma]\n"
     'RETURN SWITCH ( TRUE (), ISBLANK ( s ), "%s", s < 3, "%s", s < 4, "%s", "%s" )'
     % (CINZA, VERMELHO, AMBAR, VERDE), None, "09 Cores"),
    ("Cor Veredito",
     'IF ( [n Rejeitados] > 0, "%s", "%s" )' % (VERMELHO, VERDE),
     None, "09 Cores"),
]


def cor_por_medida(pares):
    """Formatacao condicional de tabela: pinta a fonte de uma coluna com o
    valor devolvido por uma medida de cor. pares: [(coluna_exibida, medida_cor)]"""
    return {"values": [
        {"properties": {"fontColor": {"solid": {"color": {"expr": {"Measure": {
            "Expression": {"SourceRef": {"Entity": "Fato_QC"}},
            "Property": medida}}}}}},
         "selector": {"metadata": "Fato_QC." + alvo}}
        for alvo, medida in pares]}


# ===========================================================================
# PAGINAS
# ===========================================================================

def painel_filtros(prefixo, celulares=None):
    """Coluna de filtros identica nas tres paginas: trocar de aba nao deveria
    obrigar o usuario a reaprender onde ficam os filtros."""
    celulares = celulares or {}
    v = [texto(prefixo + "/lbl_filtros", 8, 8, 198, 22,
               [("FILTROS", 10, True, CINZA)], z=1)]
    # Analito, Nivel e Lote em selecao unica: o Levey-Jennings so tem
    # significado com um de cada, e um filtro que ja nasce util evita que o
    # grafico principal abra vazio.
    esp = [
        ("analito", F_ANALITO, 34, 58, "Dropdown", True),
        ("nivel", F_NIVEL, 96, 54, "Dropdown", True),
        ("lote", F_LOTE, 154, 54, "Dropdown", True),
        ("ano", F_ANO, 212, 54, "Dropdown", False),
        ("tri", F_TRI, 270, 54, "Dropdown", False),
        ("mes", F_MES, 328, 54, "Dropdown", False),
        ("comp", F_COMP, 386, 54, "Dropdown", False),
        ("data", F_DATA, 444, 76, "Between", False),
    ]
    for nome, campo, y, h, modo, so_um in esp:
        s = filtro_slicer(prefixo + "/f_" + nome, campo, 8, y, 198, h,
                          modo=modo, z=1, celular=celulares.get(nome),
                          unico=so_um)
        if s is not None:
            v.append(s)
    v.append(cartao(prefixo + "/c_periodo", M("Periodo Analisado"), "Período",
                    8, 526, 198, 66, z=1, pt=10))
    v.append(cartao(prefixo + "/c_espec", M("Especificacao em Uso"),
                    "Especificação em uso", 8, 596, 198, 58, z=1, pt=11))
    v.append(cartao(prefixo + "/c_atualizado", M("Atualizado Em"), "Origem",
                    8, 658, 198, 54, z=1, pt=8))
    return v


def cabecalho(prefixo, titulo, subtitulo, larg=1058):
    return texto(prefixo + "/titulo", 214, 8, larg, 44, [
        (titulo + "   ", 16, True, AZUL),
        (subtitulo, 10, False, CINZA),
    ], z=1)


def linha_kpi(prefixo, itens, y=58, h=92, x0=214, larg_total=1058,
              celulares=None):
    """Cartoes de KPI distribuidos por igual na largura util.
    celulares: lista paralela de posicoes no canvas de telefone (ou None)."""
    n = len(itens)
    gap = 5
    w = (larg_total - gap * (n - 1)) // n
    fora = []
    for i, (medida, rotulo, pt) in enumerate(itens):
        cel = celulares[i] if celulares else None
        c = cartao(prefixo + "/kpi%d_%d" % (y, i), M(medida), rotulo,
                   x0 + i * (w + gap), y, w, h, z=1, celular=cel, pt=pt)
        if c is not None:
            fora.append(c)
    return fora


def _par_obs_lim(obs, lim):
    """Observado em azul, limite em cinza: o limite e referencia, nao
    concorrente visual do valor medido."""
    return {"dataPoint": [
        {"properties": {"fill": _cor(AZUL_CLARO)},
         "selector": {"metadata": "Fato_QC." + obs}},
        {"properties": {"fill": _cor("#B7B7B7")},
         "selector": {"metadata": "Fato_QC." + lim}},
    ]}


# ----------------------------------------------------------------- Painel

def pagina_painel():
    p = "painel"
    v = painel_filtros(p, celulares={"analito": (8, 300, 304, 58),
                                     "nivel": (8, 362, 304, 54)})
    v.append(cabecalho(p, "Painel de Controle Interno da Qualidade",
                       "Bioquímica · Levey-Jennings, Westgard e situação do controle"))

    v += linha_kpi(p, [
        ("Ultimo Resultado", "Último resultado", 19),
        ("Media Alvo", "Média alvo", 19),
        ("DP Alvo", "DP alvo", 19),
        ("Z Score Atual", "Z da última corrida", 19),
        ("Veredito Atual", "Veredito da última corrida", 13),
        ("Situacao da Especificacao", "Situação da especificação", 11),
    ], y=58, h=74, celulares=[(8, 8, 304, 76), None, None, None,
                              (8, 92, 304, 76), None])

    v += linha_kpi(p, [
        ("CV Observado %", "CV observado (%)", 19),
        ("Bias Observado %", "Bias observado (%)", 19),
        ("ET Observado %", "ET observado (%)", 19),
        ("Sigma", "Sigma", 19),
        ("% Aceitaveis", "% conformes", 19),
        ("Alertas 1_2s", "Alertas 1-2s", 19),
    ], y=138, h=74, celulares=[(8, 176, 148, 76), (164, 176, 148, 76),
                               (8, 260, 148, 76), (164, 260, 148, 76),
                               None, None])

    # e preferivel dizer "nao da para mostrar" a desenhar um Levey-Jennings
    # com limites de analitos diferentes misturados
    v.append(cartao(p + "/lj_aviso", M("LJ Aviso"), None, 214, 218, 700, 24,
                    z=2, pt=9))
    v.append(visual(
        p + "/lj", "lineChart", 214, 246, 700, 250,
        {"Category": [F_RUN], "Y": [M(n) for n in SERIES_LJ]},
        titulo="Levey-Jennings — resultado, média e limites ±1s / ±2s / ±3s",
        objetos=cores_series(SERIES_LJ), z=1))

    v.append(visual(
        p + "/westgard", "clusteredColumnChart", 920, 218, 352, 278,
        {"Y": [M("Alertas 1_2s"), M("Viol 1_3s"), M("Viol 2_2s"),
               M("Viol R_4s"), M("Viol 4_1s"), M("Viol 10x")]},
        titulo="Alertas e violações por regra", z=1,
        celular=(8, 424, 304, 208)))

    v.append(visual(
        p + "/matriz_west", "tableEx", 214, 502, 1058, 210,
        {"Values": [F_ANALITO, F_NIVEL, M("n Resultados"), M("n Rejeitados"),
                    M("% Aceitaveis"), M("Alertas 1_2s"), M("Viol 1_3s"),
                    M("Viol 2_2s"), M("Viol R_4s"), M("Viol 4_1s"),
                    M("Viol 10x"), M("Viol Total")]},
        titulo="Westgard por analito e nível — alertas, violações e veredito",
        objetos=cor_por_medida([("% Aceitaveis", "Cor Veredito")]), z=1))
    return v


# ------------------------------------------------------------ Estatistica

def pagina_estatistica():
    p = "estat"
    v = painel_filtros(p, celulares={"analito": (8, 496, 304, 58)})
    v.append(cabecalho(p, "Estatística e Desempenho Analítico",
                       "observado × especificação vigente, por analito e nível"))

    v += linha_kpi(p, [
        ("N Observado", "N", 19),
        ("Media Observada", "Média observada", 19),
        ("DP Observado", "DP observado", 19),
        ("CV Observado %", "CV (%)", 19),
        ("Bias Observado %", "Bias (%)", 19),
        ("ET Observado %", "ET (%)", 19),
        ("Sigma", "Sigma", 19),
        ("% Grupos ET Conforme", "% conformes (ET)", 19),
    ], y=58, h=72, celulares=[None, None, None, (8, 8, 148, 76),
                              (164, 8, 148, 76), (8, 92, 148, 76),
                              (164, 92, 148, 76), (8, 176, 304, 76)])

    larg = 346
    for i, (chave, obs, lim, rot) in enumerate([
        ("cv", "CV Observado %", "CVtp %", "CV observado × CV permitido (%)"),
        ("bias", "Bias Observado %", "BIAStp %", "Bias observado × Bias permitido (%)"),
        ("et", "ET Observado %", "ETp %", "ET observado × ETp permitido (%)"),
    ]):
        v.append(visual(
            p + "/" + chave, "clusteredBarChart",
            214 + i * (larg + 10), 138, larg, 200,
            {"Category": [F_ANALITO], "Y": [M(obs), M(lim)]},
            titulo=rot, z=1, objetos=_par_obs_lim(obs, lim)))

    v.append(visual(
        p + "/sigma", "clusteredColumnChart", 214, 346, 400, 170,
        {"Category": [F_FAIXA], "Y": [M("n Analitos")]},
        titulo="Analitos por faixa Sigma", z=1,
        celular=(8, 260, 304, 190)))

    v.append(visual(
        p + "/piores", "clusteredBarChart", 624, 346, 648, 170,
        {"Category": [F_ANALITO], "Y": [M("Folga ET %")]},
        titulo="Folga até o ETp (%) — negativo significa fora da especificação",
        z=1))

    # A matriz reproduz a aba Estatistica inteira. As colunas de status
    # ganham cor pelo valor da medida correspondente.
    v.append(visual(
        p + "/matriz", "tableEx", 214, 524, 1058, 188,
        {"Values": [F_ANALITO, F_NIVEL, M("N Observado"), M("Media Observada"),
                    M("DP Observado"), M("CV Observado %"), M("CVtp %"),
                    M("Status CV"), M("Bias Observado %"), M("BIAStp %"),
                    M("Status Bias"), M("ET Observado %"), M("ETp %"),
                    M("Status ET"), M("Sigma"), F_ESPEC, F_SITUACAO]},
        titulo="Matriz estatística — observado, limite e situação por analito e nível",
        objetos=cor_por_medida([("Status CV", "Cor Status CV"),
                                ("Status Bias", "Cor Status Bias"),
                                ("Status ET", "Cor Status ET"),
                                ("Sigma", "Cor Sigma")]), z=1))
    return v


# -------------------------------------------------------- Visao Gerencial

def pagina_gerencial():
    p = "gestao"
    v = painel_filtros(p, celulares={"analito": (8, 560, 304, 58)})
    v.append(cabecalho(p, "Visão Gerencial — Controle Interno da Qualidade",
                       "leitura executiva do período selecionado"))

    v += linha_kpi(p, [
        ("Status Global", "Situação geral", 14),
        ("% Aceitaveis", "% conformes", 21),
        ("Analitos Criticos", "Analitos com rejeição", 21),
        ("Grupos ET Nao Conforme", "Fora da especificação", 21),
        ("Grupos Sem Meta", "Sem meta cadastrada", 21),
    ], y=58, h=80, celulares=[(8, 8, 304, 88), (8, 104, 148, 76),
                              (164, 104, 148, 76), (8, 188, 148, 76),
                              (164, 188, 148, 76)])

    v += linha_kpi(p, [
        ("Grupos Sigma < 3", "Sigma < 3", 21),
        ("Grupos Sigma 3 a 4", "Sigma 3 a <4", 21),
        ("Grupos Sigma >= 6", "Sigma ≥ 6", 21),
        ("Pior CV", "Maior CV", 11),
        ("Pior Bias", "Maior Bias", 11),
        ("Pior ET", "Maior ET", 11),
    ], y=144, h=72, celulares=[(8, 272, 148, 76), None, (164, 272, 148, 76),
                               None, None, (8, 356, 304, 76)])

    v.append(visual(
        p + "/tendencia", "lineChart", 214, 224, 640, 244,
        {"Category": [F_COMP], "Y": [M("% Aceitaveis")]},
        titulo="Tendência de conformidade por competência", z=1,
        celular=(8, 440, 304, 190)))

    v.append(visual(
        p + "/faixa", "donutChart", 860, 224, 412, 244,
        {"Category": [F_FAIXA], "Y": [M("n Analitos")]},
        titulo="Analitos por faixa Sigma", z=1))

    v.append(visual(
        p + "/ranking", "tableEx", 214, 474, 640, 238,
        {"Values": [F_ANALITO, F_NIVEL, M("CV Observado %"),
                    M("Bias Observado %"), M("ET Observado %"), M("Sigma"),
                    M("Status ET")]},
        titulo="Desempenho por analito — ordenar por qualquer coluna",
        objetos=cor_por_medida([("Status ET", "Cor Status ET"),
                                ("Sigma", "Cor Sigma")]), z=1))

    v.append(visual(
        p + "/espec", "tableEx", 860, 474, 412, 238,
        {"Values": [F_ESPEC, M("n Analitos"), M("n Resultados")]},
        titulo="Especificação de qualidade em uso", z=1))
    return v


def construir_secoes():
    return [
        secao(0, "Painel", pagina_painel()),
        secao(1, "Estatística", pagina_estatistica()),
        secao(2, "Visão Gerencial", pagina_gerencial()),
    ]


COLS_DIM_DATA = [
    ("Data", DATA, False, None),
    ("Ano", INT, False, None),
    ("Mes", INT, True, None),
    ("Trimestre", INT, True, None),
    ("Competencia", TEXTO, False, None),
    ("MesNome", TEXTO, False, "Mes"),
    ("TrimestreRotulo", TEXTO, False, "Trimestre"),
    ("AnoMes", INT, True, None),
]
COLS_DIM_ANALITO = [
    ("ID_Analito", TEXTO, True, None),
    ("Analito", TEXTO, False, None),
    ("Area", TEXTO, False, None),
    ("Unidade", TEXTO, False, None),
]
COLS_DIM_LOTE = [
    ("ID_Lote", TEXTO, True, None),
    ("Lote", TEXTO, False, None),
]
COLS_DIM_NIVEL = [
    ("Nivel", INT, True, None),
    ("Nivel_Rotulo", TEXTO, False, "Nivel"),
]


def tmdl_dim_data():
    """dataCategory: Time + isKey na coluna de data e como 'Marcar como
    tabela de datas' se materializa. Sem isso, time intelligence do DAX
    silenciosamente usa a tabela de datas automatica e ignora o calendario."""
    p = ["table Dim_Data", "\tdataCategory: Time",
         "\tlineageTag: %s" % guid("tab/Dim_Data"), ""]
    for nome, tipo, oculta, ordem in COLS_DIM_DATA:
        col = tmdl_coluna("Dim_Data", nome, tipo, oculta, ordenar_por=ordem)
        if nome == "Data":
            col = col.replace("\t\tsummarizeBy: none", "\t\tisKey\n\t\tsummarizeBy: none")
        p.append(col)
    p.append(tmdl_particao_m("Dim_Data", M_DIM_DATA))
    p.append("\tannotation PBI_ResultType = Table")
    p.append("")
    return "\n".join(p)


# ------------------------------------------------------------- esqueleto PBIP

def escrever_projeto(raiz, caminho_xlsm, secoes):
    modelo = os.path.join(raiz, NOME + ".SemanticModel")
    relat = os.path.join(raiz, NOME + ".Report")
    defs = os.path.join(modelo, "definition")

    # ---- arquivo raiz do projeto
    escrever_json(os.path.join(raiz, NOME + ".pbip"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        "version": "1.0",
        "artifacts": [{"report": {"path": NOME + ".Report"}}],
        "settings": {"enableAutoRecovery": True},
    })

    # ---- modelo semantico
    plataforma(os.path.join(modelo, ".platform"), "SemanticModel", NOME)
    escrever_json(os.path.join(modelo, "definition.pbism"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
        "version": "4.2", "settings": {}})
    escrever_json(os.path.join(modelo, "diagramLayout.json"), {
        "version": "1.1.0",
        "diagrams": [{
            "ordinal": 0, "scrollPosition": {"x": 0, "y": 0}, "nodes": [
                {"location": {"x": 420, "y": 40}, "nodeIndex": "Fato_QC",
                 "nodeLineageTag": guid("tab/Fato_QC"), "size": {"height": 300, "width": 230}, "zIndex": 0},
                {"location": {"x": 60, "y": 40}, "nodeIndex": "Dim_Data",
                 "nodeLineageTag": guid("tab/Dim_Data"), "size": {"height": 200, "width": 200}, "zIndex": 1},
                {"location": {"x": 60, "y": 300}, "nodeIndex": "Dim_Analito",
                 "nodeLineageTag": guid("tab/Dim_Analito"), "size": {"height": 160, "width": 200}, "zIndex": 2},
                {"location": {"x": 760, "y": 40}, "nodeIndex": "Dim_Lote",
                 "nodeLineageTag": guid("tab/Dim_Lote"), "size": {"height": 140, "width": 200}, "zIndex": 3},
                {"location": {"x": 760, "y": 240}, "nodeIndex": "Dim_Nivel",
                 "nodeLineageTag": guid("tab/Dim_Nivel"), "size": {"height": 140, "width": 200}, "zIndex": 4},
            ],
            "name": "Modelo QC_INI", "zoomValue": 100, "pinKeyFieldsToTop": False,
            "showExtraHeaderInfo": False, "hideKeyFieldsWhenCollapsed": False,
            "tablesLocked": False,
        }],
        "selectedDiagram": "Modelo QC_INI", "defaultDiagram": "Modelo QC_INI",
    })

    escrever(os.path.join(defs, "database.tmdl"), "database\n\tcompatibilityLevel: 1567\n")
    escrever(os.path.join(defs, "model.tmdl"), tmdl_modelo())
    escrever(os.path.join(defs, "expressions.tmdl"), tmdl_expressao_parametro(caminho_xlsm))
    escrever(os.path.join(defs, "relationships.tmdl"), tmdl_relacoes())
    escrever(os.path.join(defs, "tables", "Fato_QC.tmdl"), tmdl_fato())
    escrever(os.path.join(defs, "tables", "Dim_Data.tmdl"), tmdl_dim_data())
    escrever(os.path.join(defs, "tables", "Dim_Analito.tmdl"),
             tmdl_dim("Dim_Analito", COLS_DIM_ANALITO, M_DIM_ANALITO))
    escrever(os.path.join(defs, "tables", "Dim_Lote.tmdl"),
             tmdl_dim("Dim_Lote", COLS_DIM_LOTE, M_DIM_LOTE))
    escrever(os.path.join(defs, "tables", "Dim_Nivel.tmdl"),
             tmdl_dim("Dim_Nivel", COLS_DIM_NIVEL, M_DIM_NIVEL))

    # ---- relatorio (PBIR)
    escrever_relatorio(relat, secoes)



def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsm", required=True)
    ap.add_argument("--saida", required=True)
    ap.add_argument("--forcar", action="store_true",
                    help="regenera mesmo que o relatorio em disco tenha mais paginas")
    ap.add_argument("--incluir", default="texto,cartao,slicer,grafico,tabela,cores",
                    help="classes de visual a emitir (bissecao de diagnostico)")
    a = ap.parse_args()

    global INCLUIR
    INCLUIR = set(x.strip() for x in a.incluir.split(",") if x.strip())

    xlsm = os.path.abspath(a.xlsm)
    if not os.path.isfile(xlsm):
        raise SystemExit("nao encontrei o .xlsm: " + xlsm)

    raiz = os.path.abspath(a.saida)

    # Trava de seguranca: este gerador emite as TRES paginas originais. Depois
    # do ADR-029 o relatorio em disco tem CINCO (dois produtos), montadas por
    # montar_relatorio_4paginas.py em cima da saida daqui. Regenerar sem pensar
    # apagaria esse trabalho -- a ordem correta e gerar, depois montar.
    relat = os.path.join(raiz, NOME + ".Report")
    pags = os.path.join(relat, "definition", "pages")
    if os.path.isdir(pags) and not a.forcar:
        quantas = len([d for d in os.listdir(pags)
                       if os.path.isdir(os.path.join(pags, d))])
        if quantas > len(construir_secoes()):
            raise SystemExit(
                "Recusado: o relatorio em disco tem %d paginas e este gerador\n"
                "emite %d. Regenerar apagaria as paginas montadas depois.\n"
                "Se e mesmo isso que voce quer, repita com --forcar; em seguida\n"
                "rode montar_relatorio_4paginas.py para remontar os dois produtos."
                % (quantas, len(construir_secoes())))

    for sub in (NOME + ".SemanticModel", NOME + ".Report"):
        alvo = os.path.join(raiz, sub)
        if os.path.isdir(alvo):
            shutil.rmtree(alvo)

    escrever_projeto(raiz, xlsm, construir_secoes())
    print("PBIP gerado em: " + raiz)
    print("  modelo   : %s.SemanticModel" % NOME)
    print("  relatorio: %s.Report" % NOME)
    print("  fonte    : %s" % xlsm)
    print("  medidas  : %d" % len(MEDIDAS))
    print("  incluido : %s" % ",".join(sorted(INCLUIR)))
    print("  colunas na fato: %d" % (len(COLUNAS_FATO) + len(COLUNAS_DERIVADAS)
                                     + len(COLUNAS_CALCULADAS)))


if __name__ == "__main__":
    main()
