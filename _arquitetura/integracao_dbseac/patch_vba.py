# -*- coding: utf-8 -*-
"""patch_vba.py -- ADR-057: gera os modulos VBA da nova camada de dados.

Entrada : o PROPRIO .xlsm de producao (o VBA e extraido com oletools na hora,
          entao o patch sempre parte do codigo que esta no arquivo).
Saida   : integracao_dbseac/src/<bio|hema>/  (um arquivo por modulo).

Cada troca e feita por ANCORA EXATA e conferida: se a ancora nao aparece o
numero esperado de vezes, o script para (ADR-047: patch que "passa" sem ter
casado e a correcao que nunca embarca).

Uso: python patch_vba.py <Bioquimica|Hematologia> <arquivo.xlsm>
"""
import os
import re
import sys

from oletools.olevba import VBA_Parser

AQUI = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(AQUI, 'src')
COMUM = os.path.join(SRC, 'comum')

# modulos que SAEM do projeto (a funcao deles foi substituida pela camada de dados)
REMOVER = ['mBanco', 'mImportar', 'mOperacao', 'frmCorrida', 'frmMassa', 'frmExcluir', 'frmConfigEstatistica']


def extrair(xlsm):
    vp = VBA_Parser(xlsm)
    out = {}
    for (_f, _s, vname, code) in vp.extract_macros():
        nome = re.sub(r'\.(bas|cls|frm)$', '', vname)
        out[nome] = code.replace('\r\n', '\n')
    vp.close()
    return out


def troca(txt, velho, novo, vezes=1, rotulo=''):
    n = txt.count(velho)
    if n != vezes:
        raise SystemExit(f'ANCORA [{rotulo}] casou {n}x, esperado {vezes}x:\n{velho[:300]}')
    return txt.replace(velho, novo)


def troca_bloco(txt, inicio, fim, novo, rotulo=''):
    """Substitui do inicio (inclusive) ate fim (exclusive). Inicio unico."""
    a = txt.count(inicio)
    if a != 1:
        raise SystemExit(f'ANCORA-INICIO [{rotulo}] casou {a}x')
    i = txt.index(inicio)
    j = txt.index(fim, i + len(inicio))
    return txt[:i] + novo + txt[j:]


def remover_funcao(txt, cabecalho, rotulo=''):
    i = txt.index(cabecalho)
    j = txt.index('\nEnd Function\n', i) + len('\nEnd Function\n')
    return txt[:i] + txt[j:]


def ler(caminho):
    with open(caminho, encoding='utf-8') as f:
        return f.read().replace('\r\n', '\n')


# =============================================================================
def patch_mestatistica(t):
    t = troca(t, """'  Elegibilidade (CLSI EP05 / C24): apenas resultados ELEGIVEIS compoem
'  media, DP, CV, Bias, Sigma e as regras de Westgard. Os estados vivem na aba
'  Cfg_Status; acrescentar um novo estado NAO exige alterar nenhuma rotina.
""", """'  Elegibilidade (CLSI EP05 / C24): apenas resultados ELEGIVEIS compoem
'  media, DP, CV, Bias, Sigma e as regras de Westgard. Quem decide e o Power
'  Query (tblCQ_Final.PARTICIPA_ESTATISTICA, ADR-057); o motor so le. O
'  resultado inativado com REGISTRAR - LJ marcado vira X vermelho no grafico e
'  nunca entra em calculo.
""", rotulo='cabecalho')
    t = troca(t, "Private Const COL_DATA_ENG As Long = 29 ' Eng_Saida: data da corrida (AC) -- o Calc le daqui (ADR-050)\n",
              "Private Const COL_DATA_ENG As Long = 29 ' Eng_Saida: data da corrida (AC) -- o Calc le daqui (ADR-050)\n"
              "Private Const COL_X0 As Long = 30      ' Eng_Saida: 1o X vermelho (AD) -- ADR-057\n"
              "Private Const NXN As Long = 3          ' X vermelhos por corrida e nivel\n"
              "Private Const NX As Long = 9           ' colunas do bloco X (3 niveis x 3), iguais nos dois produtos\n",
              rotulo='const X')
    t = troca(t, "Private mElig As Object                ' status -> True/False\n",
              "Private mIdxX As Object                ' \"ANALITO|LOTE\" -> linhas X_VERMELHO de mDB (ADR-057)\n"
              "Private mEquipIdx As String            ' equipamento para o qual mIdx/mIdxX foram montados (\"\" = todos)\n",
              rotulo='decl mElig')
    t = troca(t, "Private mReg As Variant                ' snapshot de Registros (repeticoes/calibracao)\n", '',
              rotulo='decl mReg')
    t = troca_bloco(t, "' ============================ ELEGIBILIDADE ============================\n",
                    "' ============================ CACHE ============================\n", """' ============================ ELEGIBILIDADE (ADR-057) ============================
' Quem decide se um resultado participa da estatistica e o Power Query
' (tblCQ_Final.PARTICIPA_ESTATISTICA); mDados.CarregarDB traduz SIM para "Ativo".
' A aba Cfg_Status saiu: uma segunda regra de elegibilidade aqui seria uma
' populacao analitica paralela, exatamente o que a arquitetura proibe.

' Equipamento em analise. Bioquimica: o seletor do Painel (selEquipamento) --
' DIMENSION 1 e DIMENSION 2 sao series proprias de Levey-Jennings e Westgard.
' Hematologia: vazio = todos (um equipamento so).
Public Function EquipFiltro() As String
    On Error Resume Next
    EquipFiltro = UCase$(Trim$(CStr(ThisWorkbook.Names("selEquipamento").RefersToRange.Value)))
End Function

""", rotulo='elegibilidade')
    t = troca(t, "    Set mElig = Nothing\n", "    Set mIdxX = Nothing\n    mEquipIdx = \"\"\n", rotulo='inval mElig')
    t = troca(t, "    mReg = Empty\n", '', rotulo='inval mReg')
    t = troca_bloco(t, "Private Sub GarantirIndice()\n",
                    "' Linhas elegiveis de (analito, lote). Lote vazio = todos os lotes do analito.\n", """Private Sub GarantirIndice()
    Dim i As Long, k As String, an As String, cod As String, eq As String
    GarantirDB
    eq = EquipFiltro()
    If Not mIdx Is Nothing Then
        If eq = mEquipIdx Then Exit Sub
        ' ADR-057: o equipamento do Painel mudou -- o indice, o cache de
        ' estatistica e os eventos de Westgard eram do outro equipamento
        Set mIdx = Nothing: Set mIdxX = Nothing: Set mAgg = Nothing: mEvLote = ""
        Set mCache = CreateObject("Scripting.Dictionary")
        mCache.CompareMode = 1
    End If
    mEquipIdx = eq
    Set mIdx = CreateObject("Scripting.Dictionary")
    mIdx.CompareMode = 0                     ' lote exato; o analito vai em maiusculas
    Set mIdxX = CreateObject("Scripting.Dictionary")
    mIdxX.CompareMode = 0
    If IsEmpty(mDB) Then Exit Sub
    For i = 1 To UBound(mDB, 1)
        an = Trim$(CStr(mDB(i, COL_ANALITO)))
        If Len(an) > 0 Then
            If Len(eq) = 0 Or UCase$(Trim$(CStr(mDB(i, COL_EQUIP)))) = eq Then
                ' ADR-057: o lote ja vem como NUCLEO da tblCQ_Final
                cod = Trim$(CStr(mDB(i, COL_LOTE)))
                k = UCase$(an) & "|" & cod
                If mDB(i, COL_STATUS) = ST_ATIVO Then
                    If Not mIdx.Exists(k) Then mIdx.Add k, New Collection
                    mIdx(k).Add i
                ElseIf CStr(mDB(i, COL_PLOT)) = PLOT_X Then
                    ' inativado com REGISTRAR - LJ: so o X do grafico, nunca o calculo
                    If Not mIdxX.Exists(k) Then mIdxX.Add k, New Collection
                    mIdxX(k).Add i
                End If
            End If
        End If
    Next i
End Sub

""", rotulo='GarantirIndice')
    t = troca(t, "' Lote de uma chave do indice (\"ANALITO|LOTE\").\n", """' Linhas X VERMELHO (inativadas com REGISTRAR - LJ) de (analito, lote) -- ADR-057.
' Nome com sufixo De, como LinhasDe: o VBA nao diferencia maiusculas, e uma
' variavel local "linhasX" esconderia uma funcao "LinhasX" (erro 450).
Private Function LinhasXDe(ByVal analito As String, ByVal lote As String) As Collection
    Dim k As String
    GarantirIndice
    k = UCase$(Trim$(analito)) & "|" & lote
    If mIdxX.Exists(k) Then Set LinhasXDe = mIdxX(k) Else Set LinhasXDe = New Collection
End Function

' Lote de uma chave do indice ("ANALITO|LOTE").
""", rotulo='LinhasX')
    t = troca(t, """        Set wsG = ThisWorkbook.Sheets("Registros")
        mReg = wsG.Range(wsG.Cells(4, 2), wsG.Cells(203, 9)).Value      ' Data..Calibrado
""", '', rotulo='GarantirDB mReg')
    t = troca_bloco(t, "Public Sub AtualizarCalc()\n", "' Filtro de data/trimestre do Painel.\n",
                    ler(os.path.join(COMUM, 'AtualizarCalc.fragmento.bas')), rotulo='AtualizarCalc')
    t = remover_funcao(t, "Private Function MarcadoresRegistro(")
    t = troca(t, "        If Len(mEvLote) > 0 And mEvLote = mLotes.LotePainel() Then Exit Sub\n",
              "        If Len(mEvLote) > 0 And mEvLote = mLotes.LotePainel() & \"|\" & mEquipIdx Then Exit Sub\n",
              rotulo='ev cache')
    t = troca(t, "    If porAnalito.Count = 0 Then Set mAgg = CreateObject(\"Scripting.Dictionary\"): mEvLote = lote: GoTo restaura\n",
              "    If porAnalito.Count = 0 Then Set mAgg = CreateObject(\"Scripting.Dictionary\"): mEvLote = lote & \"|\" & mEquipIdx: GoTo restaura\n",
              rotulo='ev vazio')
    t = troca(t, "    mEvLote = lote                 ' ADR-050: vale para este lote ate InvalidarCache\n",
              "    mEvLote = lote & \"|\" & mEquipIdx   ' ADR-050/057: vale para este lote e equipamento ate InvalidarCache\n",
              rotulo='ev fim')
    t = troca(t, """    ws.Range("L2").Value = lote
""", """    ws.Range("L2").Value = lote
    ws.Range("M2").Value = IIf(Len(mEquipIdx) > 0, "Equip.:", "")
    ws.Range("N2").Value = mEquipIdx
""", rotulo='ev equip')
    t = troca(t, """       Or Trim$(CStr(eng.Range("E1").Value)) <> mLotes.LotePainel() Then
""", """       Or Trim$(CStr(eng.Range("E1").Value)) <> mLotes.LotePainel() _
       Or UCase$(Trim$(CStr(eng.Range("S1").Value))) <> EquipFiltro() Then
""", rotulo='garantir motor')
    t = t.rstrip('\n') + """


' ADR-057: o equipamento do Painel mudou (Bioquimica). Cada equipamento e uma
' serie propria: o indice e refeito para ele e TODO o motor acompanha --
' grafico, Westgard, eventos e a aba Estatistica.
Public Sub EquipMudou()
    Dim volta As Object, nE As Long, sE As String
    Set volta = mApp.TelaAtual()
    mApp.InicioUsuario "Trocando o equipamento para " & EquipFiltro() & "..."
    On Error GoTo falha
    AtualizarEstatistica
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!RecalcularEstatPeriodo"
    On Error GoTo falha
    mApp.fim
    mApp.VoltarPara volta
    Exit Sub
falha:
    nE = Err.Number: sE = Err.Description
    mApp.fim
    mApp.VoltarPara volta
    MsgBox "Nao foi possivel trocar o equipamento." & vbCrLf & vbCrLf & _
           "Erro " & nE & ": " & sE, vbExclamation, "Painel"
End Sub
"""
    # nenhuma referencia residual ao que saiu
    for proibido in ('EhElegivel', 'CarregarElegibilidade', 'mElig', 'mReg', 'Cfg_Status', 'NucleoLote'):
        if re.search(r'^[^\'\n]*\b' + proibido + r'\b', t, re.M):
            raise SystemExit(f'mEstatistica ainda referencia {proibido}')
    return t


def patch_mentrada(t):
    i = t.index('Public Function NucleoLote(')
    j = t.index('\nEnd Function\n', i) + len('\nEnd Function\n')
    return t[:i] + """Public Function NucleoLote(ByVal codigo As String) As String
    ' ADR-057: a fonte oficial (tblCQ_Final) ja entrega o NUCLEO ("8974",
    ' "522611"). So o codigo antigo "QC-" + nucleo + 2 digitos de nivel e
    ' convertido; nucleo devolve ele mesmo. ADR-024: corte por COMPRIMENTO.
    Dim s As String
    s = Trim$(codigo)
    If UCase$(Left$(s, 3)) = "QC-" And Len(s) > 5 Then
        NucleoLote = Mid$(s, 4, Len(s) - 5)
    Else
        NucleoLote = s
    End If
End Function
""" + t[j:]


def patch_mbi(t):
    t = troca(t, """    ult = UltimaLinhaBanco()
    If ult < BANCO_R0 Then
        LimparCorpo ws
        Exit Sub
    End If
    dados = ThisWorkbook.Sheets(BANCO).Range( _
        ThisWorkbook.Sheets(BANCO).Cells(BANCO_R0, COL_RUN), _
        ThisWorkbook.Sheets(BANCO).Cells(ult, COL_STATUS)).Value
""", """    ' ADR-057: a fonte e a tblCQ_Final -- a mesma do motor, do Painel e da
    ' Estatistica. Ativo = PARTICIPA_ESTATISTICA = SIM (decidido no Power Query).
    dados = CarregarDB()
    If IsEmpty(dados) Then
        LimparCorpo ws
        Exit Sub
    End If
""", rotulo='bi fonte')
    t = troca(t, "        If Len(lote) < 6 Then GoTo proxima1\n", "        If Len(lote) = 0 Then GoTo proxima1\n", rotulo='bi lote1')
    t = troca(t, "        If Len(lt) >= 6 Then nuc = NucleoLote(lt)\n", "        If Len(lt) > 0 Then nuc = NucleoLote(lt)\n", rotulo='bi lote2')
    return t


def patch_mlotes(t):
    t = troca(t, "    lote = LoteAtivoCore()\n    If IsMissing(dados) Then dados = CarregarDB()\n",
              "    lote = LoteAtivoCore()\n    If IsMissing(dados) Then dados = CarregarDB()\n"
              "    Dim eqF As String\n    eqF = mEstatistica.EquipFiltro()        ' ADR-057: corridas do equipamento em analise\n",
              rotulo='liber equip')
    t = troca(t, "            If Trim$(CStr(dados(i, COL_STATUS))) = ST_ATIVO Then\n                If Len(Trim$(CStr(dados(i, COL_RUN)))) > 0 Then\n                    cod = Trim$(CStr(dados(i, COL_LOTE)))\n                    If Len(cod) > 5 Then\n",
              "            If Trim$(CStr(dados(i, COL_STATUS))) = ST_ATIVO And _\n               (Len(eqF) = 0 Or UCase$(Trim$(CStr(dados(i, COL_EQUIP)))) = eqF) Then\n                If Len(Trim$(CStr(dados(i, COL_RUN)))) > 0 Then\n                    cod = Trim$(CStr(dados(i, COL_LOTE)))\n                    If Len(cod) > 0 Then\n",
              rotulo='liber lote')
    # Registros sem as colunas de repeticao (F:H): a view tem 9 colunas, nao 12
    t = troca(t, "    Set srcR = ThisWorkbook.Names(\"regView\").RefersToRange          ' B4:M203 (200 x 12)\n"
                 "    wsReg.Range(wsReg.Cells(rr0, 1), wsReg.Cells(rr0 + NLB_LIBER - 1, 12)).Value = srcR.Value\n",
              "    Set srcR = ThisWorkbook.Names(\"regView\").RefersToRange          ' B4:J203 (200 x 9) -- ADR-057: sem Rep 1-3\n"
              "    wsReg.Range(wsReg.Cells(rr0, 1), wsReg.Cells(rr0 + NLB_LIBER - 1, 9)).Value = srcR.Value\n",
              rotulo='reg salvar')
    t = troca(t, "    dstR.Value = wsReg.Range(wsReg.Cells(rr0, 1), wsReg.Cells(rr0 + NLB_LIBER - 1, 12)).Value\n",
              "    dstR.Value = wsReg.Range(wsReg.Cells(rr0, 1), wsReg.Cells(rr0 + NLB_LIBER - 1, 9)).Value\n",
              rotulo='reg carregar')
    return t


def patch_mui(t):
    t = troca(t, "Public Sub AbrirFormCorrida()\n    frmCorrida.Show\nEnd Sub\n\n", '', rotulo='form corrida')
    t = troca(t, """Public Sub AtualizarTudo()
    Application.ScreenUpdating = False
    AtualizarBanco
    AtualizarResultados
    AtualizarEstatistica
    AtualizarPainel
    Application.ScreenUpdating = True
End Sub
""", """' ADR-057: a cadeia completa e o botao ATUALIZAR DADOS (Power Query na ordem,
' sincrono, e so depois motor, Westgard, Estatistica e graficos).
Public Sub AtualizarTudo()
    mIntegracao.AtualizarDados
End Sub
""", rotulo='atualizar tudo')
    return t


def patch_mapp(t):
    t = troca(t, "Public Sub IrResultados(): Ir \"Resultados\", \"A1\": End Sub\n", '', rotulo='ir resultados')
    t = troca_bloco(t, "' Lancar resultados: Bioquimica tem a aba Importar; Hematologia, o formulario\n",
                    "Public Sub Ir(ByVal aba As String", """' Lancar resultado MANUAL (ADR-057): aba Digitar Resultados (tblResultados_Manuais), que entra
' na DB_CQ_FINAL pelo mesmo caminho do interfaceamento. A aba Importar, o
' frmCorrida e o frmMassa sairam.
Public Sub IrLancar()
    mIntegracao.IrDigitarResultados
End Sub

""", rotulo='ir lancar')
    t = troca(t, "    mx = Application.WorksheetFunction.Max(ThisWorkbook.Names(\"rData\").RefersToRange)\n",
              "    mx = mDados.UltimaDataCQ()           ' ADR-057: a fonte oficial (tblCQ_Final)\n",
              rotulo='ano dados')
    return t


def patch_mseguranca(t):
    return troca(t, '            Case "Calc", "LotesStore", "LiberStore", "RegistrosStore"\n',
                 '            Case "Calc", "LotesStore", "LiberStore", "RegistrosStore", "Cfg_Integracao"\n',
                 rotulo='seg adm')


def patch_mestatperiodo(t):
    t = troca(t, 'Public Const EP_BANCO As String = "DB_Resultados"\nPublic Const EP_R0 As Long = 4\n\n'
                 '\' colunas do banco\nPrivate Const EP_C_RUN As Long = 1\nPrivate Const EP_C_DATA As Long = 2\n'
                 'Private Const EP_C_NIVEL As Long = 3\nPrivate Const EP_C_LOTE As Long = 4\n'
                 'Private Const EP_C_ANALITO As Long = 5\nPrivate Const EP_C_VALOR As Long = 6\n'
                 'Private Const EP_C_STATUS As Long = 7\n',
              "' ADR-057: a fonte e a tblCQ_Final (mDados.CarregarDB), a mesma do motor.\n"
              "' Elegivel = PARTICIPA_ESTATISTICA = SIM; equipamento = o do Painel.\n"
              "Private mSnap As Variant          ' snapshot da fonte oficial\n"
              "Private mVersao As Long           ' sobe a cada InvalidarCacheEstat (refresh do Power Query)\n",
              rotulo='ep const')
    i = t.index('Private Function UltimaLinhaEP() As Long')
    j = t.index('\nEnd Function\n', i) + len('\nEnd Function\n')
    t = t[:i] + t[j:]
    t = troca(t, """    ult = UltimaLinhaEP()
    c = Carimbo(dtIni, dtFim, ex, nEx, lote, ult)
""", """    Dim eq As String
    eq = mEstatistica.EquipFiltro()
    c = Carimbo(dtIni, dtFim, ex, nEx, lote & "|" & eq, mVersao)
""", rotulo='ep carimbo')
    t = troca(t, """    If ult < EP_R0 Then Exit Sub

    Set ws = ThisWorkbook.Sheets(EP_BANCO)
    dados = ws.Range(ws.Cells(EP_R0, EP_C_RUN), ws.Cells(ult, EP_C_STATUS)).Value
""", """    If IsEmpty(mSnap) Then mSnap = CarregarDB()
    If IsEmpty(mSnap) Then Exit Sub
    dados = mSnap
""", rotulo='ep dados')
    t = troca(t, """        If Trim$(CStr(dados(i, EP_C_ANALITO))) = "" Then GoTo proxima
        If Trim$(CStr(dados(i, EP_C_STATUS))) <> "Ativo" Then GoTo proxima
        If Not IsDate(dados(i, EP_C_DATA)) Then GoTo proxima
        If Not IsNumeric(dados(i, EP_C_VALOR)) Then GoTo proxima
""", """        If Trim$(CStr(dados(i, COL_ANALITO))) = "" Then GoTo proxima
        If CStr(dados(i, COL_STATUS)) <> ST_ATIVO Then GoTo proxima
        If Len(eq) > 0 Then If UCase$(Trim$(CStr(dados(i, COL_EQUIP)))) <> eq Then GoTo proxima
        If Not IsDate(dados(i, COL_DATA)) Then GoTo proxima
        If Not IsNumeric(dados(i, COL_RESULT)) Then GoTo proxima
""", rotulo='ep filtros')
    t = troca(t, "        d = Int(CDbl(CDate(dados(i, EP_C_DATA))))\n", "        d = Int(CDbl(CDate(dados(i, COL_DATA))))\n", rotulo='ep d')
    t = troca(t, """        ' lote pelo NUCLEO (NucleoLote vive em mEntrada: uma conta, um lugar)
        If Len(Trim$(lote)) > 0 Then
            If NucleoLote(CStr(dados(i, EP_C_LOTE))) <> Trim$(lote) Then GoTo proxima
        End If

        v = CDbl(dados(i, EP_C_VALOR))
        k = UCase$(Trim$(CStr(dados(i, EP_C_ANALITO)))) & "|" & Trim$(CStr(dados(i, EP_C_NIVEL)))
""", """        ' lote: a fonte oficial ja entrega o NUCLEO
        If Len(Trim$(lote)) > 0 Then
            If Trim$(CStr(dados(i, COL_LOTE))) <> Trim$(lote) Then GoTo proxima
        End If

        v = CDbl(dados(i, COL_RESULT))
        k = UCase$(Trim$(CStr(dados(i, COL_ANALITO)))) & "|" & Trim$(CStr(dados(i, COL_NIVEL)))
""", rotulo='ep lote')
    t = troca(t, """' Invalida o cache. Chamar apos importacao ou exclusao logica.
Public Sub InvalidarCacheEstat()
    Set mAgg = Nothing
    mCarimbo = ""
End Sub
""", """' Invalida o cache. Chamado pelo ATUALIZAR DADOS depois do refresh do Power
' Query (ADR-057): uma inativacao muda PARTICIPA sem mudar o numero de linhas,
' entao o carimbo leva uma VERSAO, nao o tamanho da tabela.
Public Sub InvalidarCacheEstat()
    Set mAgg = Nothing
    mCarimbo = ""
    mSnap = Empty
    mVersao = mVersao + 1
End Sub

' As 320 celulas da aba Estatistica so recalculam quando muda um ARGUMENTO.
' Depois de um refresh (ou da troca de equipamento) nenhum argumento mudou:
' marca as celulas como sujas e recalcula.
Public Sub RecalcularEstatPeriodo()
    On Error Resume Next
    InvalidarCacheEstat
    ThisWorkbook.Sheets("Estatística").Range("C14:F93").Dirty
    Application.Calculate
End Sub
""", rotulo='ep inval')
    if 'EP_C_' in t or 'EP_BANCO' in t or 'EP_R0' in t:
        raise SystemExit('mEstatPeriodo ainda referencia o banco antigo')
    return t


def patch_planilha7_bio(t):
    t = troca(t, "'   M4 ........ provedor do controle externo deste analito (so a Bioquimica)\n",
              "'   M4 ........ provedor do controle externo deste analito (so a Bioquimica)\n"
              "'   L4 ........ EQUIPAMENTO em analise (selEquipamento; so a Bioquimica) -- ADR-057\n",
              rotulo='p7 cab')
    t = troca(t, "    If Not Intersect(Target, Me.Range(\"M4\")) Is Nothing Then\n",
              "    If Not Intersect(Target, Me.Range(\"L4\")) Is Nothing Then\n"
              "        ' ADR-057: DIMENSION 1 / DIMENSION 2 -- cada equipamento e uma serie propria\n"
              "        mEstatistica.EquipMudou\n"
              "    End If\n"
              "    If Not Intersect(Target, Me.Range(\"M4\")) Is Nothing Then\n",
              rotulo='p7 L4')
    return t


# =============================================================================
def sombreamentos(modulos):
    """O VBA nao diferencia maiusculas: 'Dim linhasX' + chamada 'LinhasX(a, b)' no mesmo
    procedimento = 'numero de argumentos incorreto' na compilacao. Acha esses casos,
    considerando o ESCOPO: rotinas publicas de qualquer modulo + privadas do proprio."""
    def sem_comentario(t):
        out = []
        for linha in t.split('\n'):
            aspas = 0
            for i, ch in enumerate(linha):
                if ch == '"':
                    aspas += 1
                elif ch == "'" and aspas % 2 == 0:
                    linha = linha[:i]
                    break
            out.append(linha)
        return '\n'.join(out)
    decl = r'^\s*(Public |Private |Friend )?(?:Static )?(?:Function|Sub|Property (?:Get|Let|Set))\s+(\w+)'
    publicas, privadas = set(), {}
    limpos = {n: sem_comentario(t) for n, t in modulos.items()}
    for n, t in limpos.items():
        privadas[n] = set()
        for vis, nome in re.findall(decl, t, re.M | re.I):
            (privadas[n] if vis.strip().lower() == 'private' else publicas).add(nome.lower())
    out = []
    for n, t in limpos.items():
        visiveis = publicas | privadas[n]
        for proc in re.finditer(r'^\s*(?:Public |Private |Friend )?(?:Static )?(Function|Sub|Property \w+)\s+(\w+)(.*?)^\s*End (?:Function|Sub|Property)',
                                t, re.M | re.S | re.I):
            corpo = proc.group(3)
            locais = set()
            for d in re.findall(r'^\s*(?:Dim|Static)\s+(.+)$', corpo, re.M | re.I):
                for parte in d.split(','):
                    m = re.match(r'\s*(\w+)', parte)
                    if m:
                        locais.add(m.group(1).lower())
            for v in locais & visiveis:
                if v == proc.group(2).lower():
                    continue
                if re.search(r'(?<![\w.])' + re.escape(v) + r'\s*\(', corpo, re.I):
                    out.append(f'{n}.{proc.group(2)}: local "{v}"')
    return out


def gravar(destino, nome, txt):
    ext = '.cls' if nome.startswith(('Planilha', 'EstaPasta')) else '.bas'
    with open(os.path.join(destino, nome + ext), 'w', encoding='utf-8', newline='') as f:
        f.write(txt.replace('\r\n', '\n').replace('\n', '\r\n'))


def main(produto, xlsm):
    cod = extrair(xlsm)
    bio = produto.lower().startswith('bio')
    destino = os.path.join(SRC, 'bio' if bio else 'hema')
    os.makedirs(destino, exist_ok=True)
    for f in os.listdir(destino):
        os.remove(os.path.join(destino, f))
    out = {
        'mDados': ler(os.path.join(COMUM, 'mDados.bas')),
        'mIntegracao': ler(os.path.join(COMUM, 'mIntegracao.bas')),
        'mEstatistica': patch_mestatistica(cod['mEstatistica']),
        'mEntrada': patch_mentrada(cod['mEntrada']),
        'mBI': patch_mbi(cod['mBI']),
        'mLotes': patch_mlotes(cod['mLotes']),
        'mUI': patch_mui(cod['mUI']),
        'mApp': patch_mapp(cod['mApp']),
        'mSeguranca': patch_mseguranca(cod['mSeguranca']),
    }
    if bio:
        out['mEstatPeriodo'] = patch_mestatperiodo(cod['mEstatPeriodo'])
        out['Planilha7'] = patch_planilha7_bio(cod['Planilha7'])
    for nome, txt in out.items():
        gravar(destino, nome, txt)
    # conferencia: nenhum modulo que FICA chama o que SAI
    saem = r'\b(UltimaLinhaBanco|BANCO_R0|UpsertResultados|ExcluirLogico|NovoRUN|RunsDoLote|AtualizarFlagsBanco|' \
           r'ExigirCapacidade|AtualizarViewResultados|IrParaImportar|ExecutarImportacao|AbrirFormCorrida|' \
           r'frmCorrida|frmMassa|frmExcluir|frmConfigEstatistica|AtualizarBanco|ListaLotes|ST_EXCLUIDO|' \
           r'EhElegivel|AtualizarOperacao|AbrirFormMassa|AbrirFormExcluir)\b|DB_Resultados'
    # modulos de aba que saem junto com a aba (Resultados da Bioquimica = Planilha16)
    abas_que_saem = {'Planilha16'} if bio else set()
    restantes = {**{k: v for k, v in cod.items() if k not in REMOVER and k not in out and k not in abas_que_saem}, **out}
    achados = []
    for nome, txt in restantes.items():
        for n, linha in enumerate(txt.split('\n'), 1):
            s = linha.strip()
            if s.startswith("'") or s.startswith('Attribute'):
                continue
            if re.search(saem, linha):
                achados.append(f'{nome}:{n}: {s[:120]}')
    if achados:
        print('REFERENCIAS RESIDUAIS:\n  ' + '\n  '.join(achados))
        raise SystemExit(1)
    print(f'{produto}: {len(out)} modulos gerados em {destino}; {len(REMOVER)} a remover; nenhuma referencia residual')
    return out


if __name__ == '__main__':
    main(sys.argv[1], os.path.abspath(sys.argv[2]))
