Attribute VB_Name = "mDados"
Option Explicit
' ===== CAMADA DE DADOS (Fase 2) =====
' DB_Resultados e a UNICA fonte operacional. Formato vertical normalizado.
' Schema: A=RUN | B=Data | C=Nivel | D=Lote | E=Analito | F=Resultado | G=Status
Public Const BANCO As String = "DB_Resultados"
Public Const VIEW As String = "Resultados"
Public Const BANCO_R0 As Long = 4
Public Const COL_RUN As Long = 1
Public Const COL_DATA As Long = 2
Public Const COL_NIVEL As Long = 3
Public Const COL_LOTE As Long = 4
Public Const COL_ANALITO As Long = 5
Public Const COL_RESULT As Long = 6
Public Const COL_STATUS As Long = 7
Public Const ST_ATIVO As String = "Ativo"
Public Const ST_EXCLUIDO As String = "Excluído"

Public Function UltimaLinhaBanco() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(BANCO)
    UltimaLinhaBanco = ws.Cells(ws.rows.Count, COL_RUN).End(xlUp).Row
    If UltimaLinhaBanco < BANCO_R0 Then UltimaLinhaBanco = BANCO_R0 - 1
End Function

Public Function LoteAtivoCore() As String
    On Error Resume Next
    LoteAtivoCore = Trim$(CStr(ThisWorkbook.Names("loteAtivo").RefersToRange.Value))
End Function

' Le o banco inteiro para memoria de uma vez (performance: sem loop celula a celula).
Public Function CarregarDB() As Variant
    Dim ws As Worksheet, lastRow As Long
    Set ws = ThisWorkbook.Sheets(BANCO)
    lastRow = UltimaLinhaBanco()
    If lastRow < BANCO_R0 Then
        CarregarDB = Empty
    Else
        CarregarDB = ws.Range(ws.Cells(BANCO_R0, COL_RUN), ws.Cells(lastRow, COL_STATUS)).Value
    End If
End Function

Public Function ChaveReg(ByVal run As Long, ByVal nivel As Long, ByVal analito As String) As String
    ' Chave de unicidade do registro. O Nivel ENTRA na chave: um mesmo RUN abrange
    ' todos os niveis da corrida, entao RUN+Analito sozinho colidiria entre niveis.
    ChaveReg = CStr(run) & "|" & CStr(nivel) & "|" & UCase$(Trim$(analito))
End Function

' RUN = chave logica da corrida. Unico por (Data + lote de 6 digitos).
Public Function NovoRUN(ByVal dt As Date, ByVal loteCore As String) As Long
    Dim dados As Variant, i As Long, mx As Long
    dados = CarregarDB()
    mx = 0
    If IsEmpty(dados) Then NovoRUN = 1: Exit Function
    For i = 1 To UBound(dados, 1)
        If Len(Trim$(CStr(dados(i, COL_RUN)))) > 0 And IsNumeric(dados(i, COL_RUN)) Then
            If CLng(dados(i, COL_RUN)) > mx Then mx = CLng(dados(i, COL_RUN))
            If IsDate(dados(i, COL_DATA)) Then
                If CDate(dados(i, COL_DATA)) = dt And NucleoLote(CStr(dados(i, COL_LOTE))) = loteCore Then
                    NovoRUN = CLng(dados(i, COL_RUN)): Exit Function
                End If
            End If
        End If
    Next i
    NovoRUN = mx + 1
End Function
' ===== UPSERT EM LOTE =====
' regs: array (1..n, 1..7) ja no schema do banco. Atualiza o que existir
' (mesma chave RUN|Nivel|Analito) e acrescenta o resto - nunca duplica.
' Devolve "novos|atualizados".
'
' DESEMPENHO -- por que a versao anterior levava MINUTOS
'
' Nao era o custo de chamar o COM. Era o RECALCULO. O DB_Resultados tem ~45.000
' formulas (o bloco BA:BD desce ate a linha 15.003) e a pasta passa de 56.000.
' Com o calculo em automatico, CADA escrita de celula dispara o recalculo das
' formulas dependentes. Atualizar 1.200 resultados com quatro escritas cada da
' ~4.800 recalculos de dezenas de milhares de formulas.
'
' As tres medidas, em ordem de impacto:
'   1. Calculation = manual durante a gravacao, e UM Calculate no fim.
'   2. Atualizacao na MATRIZ em memoria, com uma unica escrita no fim.
'   3. Escrita so da FAIXA TOCADA (primeira..ultima linha alterada), nao do
'      banco inteiro. Corrigir um resultado de ontem nao pode reescrever
'      15.000 linhas.
'
' UM SO SISTEMA DE COORDENADAS NO DICIONARIO.
'
' O indice guarda SEMPRE a linha da PLANILHA. Misturar indice-de-matriz para os
' existentes e linha-de-planilha para os novos -- como parece natural quando se
' passa a trabalhar em memoria -- faz a chave repetida no mesmo lote cair no
' ramo "existe" com um numero de linha de planilha e indexar a matriz fora dos
' limites. Falha intermitente, dependente do dado, do tipo mais caro de achar.
Public Function UpsertResultados(ByRef regs As Variant) As String
    Dim ws As Worksheet, dados As Variant, idx As Object
    Dim i As Long, c As Long, lastRow As Long, novos As Long, atual As Long
    Dim k As String, lin As Long, linArr As Long
    Dim addBuf() As Variant, nAdd As Long, idxNovo As Object
    Dim minTocada As Long, maxTocada As Long
    Dim calcAntes As Long, eventosAntes As Boolean, telaAntes As Boolean

    If IsEmpty(regs) Then UpsertResultados = "0|0": Exit Function

    Set ws = ThisWorkbook.Sheets(BANCO)
    Set idx = CreateObject("Scripting.Dictionary")

    ' Estado do Excel guardado ANTES de qualquer alteracao. Restaurar o que
    ' havia, e nao um valor fixo: deixar o calculo em manual porque a rotina
    ' assumiu que estava em automatico faria a planilha inteira parar de somar
    ' sem que ninguem percebesse.
    calcAntes = Application.Calculation
    eventosAntes = Application.EnableEvents
    telaAntes = Application.ScreenUpdating

    On Error GoTo Limpeza
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    dados = CarregarDB()
    lastRow = UltimaLinhaBanco()

    If Not IsEmpty(dados) Then
        For i = 1 To UBound(dados, 1)
            If Len(Trim$(CStr(dados(i, COL_ANALITO)))) > 0 Then
                k = ChaveReg(CLng(dados(i, COL_RUN)), CLng(dados(i, COL_NIVEL)), CStr(dados(i, COL_ANALITO)))
                If Not idx.Exists(k) Then idx.Add k, BANCO_R0 + i - 1
            End If
        Next i
    End If

    ReDim addBuf(1 To UBound(regs, 1), 1 To COL_STATUS)
    Set idxNovo = CreateObject("Scripting.Dictionary")
    nAdd = 0: novos = 0: atual = 0
    minTocada = 0: maxTocada = 0

    ' DOIS indices, cada um no seu dominio. A chave repetida DENTRO do mesmo
    ' lote precisa atualizar a linha que ainda esta na fila de insercao -- ela
    ' nao existe na matriz do banco, entao procurar so em 'idx' faria o
    ' registro ser DESCARTADO EM SILENCIO. Um indice unico misturando linha de
    ' planilha e indice de matriz e a origem desse tipo de perda.
    For i = 1 To UBound(regs, 1)
        k = ChaveReg(CLng(regs(i, COL_RUN)), CLng(regs(i, COL_NIVEL)), CStr(regs(i, COL_ANALITO)))
        If idx.Exists(k) Then
            lin = idx(k)                       ' linha da PLANILHA
            linArr = lin - BANCO_R0 + 1        ' indice na matriz
            dados(linArr, COL_RESULT) = regs(i, COL_RESULT)
            dados(linArr, COL_STATUS) = ST_ATIVO
            dados(linArr, COL_DATA) = regs(i, COL_DATA)
            dados(linArr, COL_LOTE) = regs(i, COL_LOTE)
            If minTocada = 0 Or linArr < minTocada Then minTocada = linArr
            If linArr > maxTocada Then maxTocada = linArr
            atual = atual + 1
        ElseIf idxNovo.Exists(k) Then
            ' ja esta na fila de insercao: sobrescreve a entrada pendente
            Dim iB As Long: iB = idxNovo(k)
            For c = 1 To COL_STATUS
                addBuf(iB, c) = regs(i, c)
            Next c
            atual = atual + 1
        Else
            nAdd = nAdd + 1
            For c = 1 To COL_STATUS
                addBuf(nAdd, c) = regs(i, c)
            Next c
            idxNovo.Add k, nAdd
            novos = novos + 1
        End If
    Next i

    ' Barreira de capacidade (ADR-025) ANTES de qualquer escrita: a mensagem diz
    ' "nenhum dado foi gravado", e so e verdade se nada tiver sido gravado ainda.
    ExigirCapacidade nAdd

    ' --- devolve so a faixa que mudou ---
    If minTocada > 0 Then
        Dim upd() As Variant
        ReDim upd(1 To maxTocada - minTocada + 1, 1 To COL_STATUS)
        For i = minTocada To maxTocada
            For c = 1 To COL_STATUS
                upd(i - minTocada + 1, c) = dados(i, c)
            Next c
        Next i
        Dim r0 As Long
        r0 = BANCO_R0 + minTocada - 1
        ws.Range(ws.Cells(r0, COL_LOTE), ws.Cells(r0 + maxTocada - minTocada, COL_LOTE)).NumberFormat = "@"
        ws.Range(ws.Cells(r0, COL_RUN), ws.Cells(r0 + maxTocada - minTocada, COL_STATUS)).Value = upd
    End If

    ' --- acrescenta os novos em um unico bloco ---
    If nAdd > 0 Then
        Dim outp() As Variant
        ReDim outp(1 To nAdd, 1 To COL_STATUS)
        For i = 1 To nAdd
            For c = 1 To COL_STATUS
                outp(i, c) = addBuf(i, c)
            Next c
        Next i
        ws.Range(ws.Cells(lastRow + 1, COL_LOTE), ws.Cells(lastRow + nAdd, COL_LOTE)).NumberFormat = "@"
        ws.Range(ws.Cells(lastRow + 1, COL_RUN), ws.Cells(lastRow + nAdd, COL_STATUS)).Value = outp
    End If

    ' As flags BA/BB/BC sao mantidas por mBanco desde o ADR-025. Recalcular aqui,
    ' e nao so nas linhas novas: uma linha inserida ou reativada muda quem e a
    ' "primeira ativa" das linhas seguintes.
    AtualizarFlagsBanco
Limpeza:
    Dim nErr As Long, sErr As String
    nErr = Err.Number: sErr = Err.Description
    On Error Resume Next
    Application.Calculation = calcAntes
    Application.EnableEvents = eventosAntes
    Application.ScreenUpdating = telaAntes
    On Error GoTo 0
    If nErr <> 0 Then Err.Raise nErr, "UpsertResultados", sErr
    UpsertResultados = CStr(novos) & "|" & CStr(atual)
End Function

' Exclusao LOGICA por RUN + Nivel + lista de analitos (Dictionary de nomes em UCase).
' ADR-053: a exclusao logica era o unico caminho de gravacao fora da casca mApp --
' escrevia celula a celula com o calculo em automatico. Com cinco anos de banco,
' cada celula alterada disparava um recalculo.
Public Function ExcluirLogico(ByVal run As Long, ByVal nivel As Long, ByRef alvoS As Object) As Long
    mApp.Inicio "Excluindo resultado(s)..."
    On Error GoTo LimpezaEx

    Dim ws As Worksheet, dados As Variant, i As Long, n As Long
    Dim minL As Long, maxL As Long, st() As Variant
    Set ws = ThisWorkbook.Sheets(BANCO)
    dados = CarregarDB()
    If IsEmpty(dados) Then ExcluirLogico = 0: GoTo LimpezaEx
    For i = 1 To UBound(dados, 1)
        If Len(Trim$(CStr(dados(i, COL_ANALITO)))) > 0 Then
            If CLng(dados(i, COL_RUN)) = run And CLng(dados(i, COL_NIVEL)) = nivel Then
                If alvoS.Exists(UCase$(Trim$(CStr(dados(i, COL_ANALITO))))) Then
                    dados(i, COL_STATUS) = ST_EXCLUIDO
                    If minL = 0 Then minL = i
                    maxL = i
                    n = n + 1
                End If
            End If
        End If
    Next i
    ' Uma escrita em bloco na faixa tocada, e nao uma por celula (ADR-053).
    If n > 0 Then
        ReDim st(1 To maxL - minL + 1, 1 To 1)
        For i = minL To maxL
            st(i - minL + 1, 1) = dados(i, COL_STATUS)
        Next i
        ws.Range(ws.Cells(BANCO_R0 + minL - 1, COL_STATUS), _
                 ws.Cells(BANCO_R0 + maxL - 1, COL_STATUS)).Value = st
        ' As flags BA/BB/BC sao mantidas por mBanco desde o ADR-025. Recalcular aqui,
        ' e nao so nas linhas novas: uma linha inserida ou reativada muda quem e a
        ' "primeira ativa" das linhas seguintes.
        AtualizarFlagsBanco
    End If
    ExcluirLogico = n
LimpezaEx:
    Dim nEx As Long, sEx As String
    nEx = Err.Number: sEx = Err.Description
    mApp.Fim
    If nEx <> 0 Then Err.Raise nEx, "mDados.ExcluirLogico", sEx

End Function

' Lista de RUNs distintos ATIVOS do lote em uso (para os combos dos formularios).
Public Function RunsDoLote(ByVal loteCore As String) As Collection
    Dim dados As Variant, i As Long, seen As Object, c As Collection
    Set c = New Collection: Set seen = CreateObject("Scripting.Dictionary")
    dados = CarregarDB()
    If IsEmpty(dados) Then Set RunsDoLote = c: Exit Function
    For i = 1 To UBound(dados, 1)
        If Len(Trim$(CStr(dados(i, COL_RUN)))) > 0 Then
            If NucleoLote(CStr(dados(i, COL_LOTE))) = loteCore Then
                If Not seen.Exists(CStr(dados(i, COL_RUN))) Then
                    seen.Add CStr(dados(i, COL_RUN)), 1
                    c.Add CLng(dados(i, COL_RUN))
                End If
            End If
        End If
    Next i
    Set RunsDoLote = c
End Function

' Nomes dos analitos cadastrados (aba Analitos).
Public Function ListaAnalitos() As Collection
    Dim ws As Worksheet, i As Long, nm As String, c As Collection
    Set c = New Collection
    Set ws = ThisWorkbook.Sheets("Analitos")
    For i = 4 To 43
        nm = Trim$(CStr(ws.Cells(i, 1).Value))
        If nm <> "" Then c.Add nm
    Next i
    Set ListaAnalitos = c
End Function

' Codigos completos de lote disponiveis no registro (Configuracao).
Public Function ListaLotes() As Collection
    Dim rng As Range, Cel As Range, c As Collection, v As String
    Set c = New Collection
    On Error Resume Next
    Set rng = ThisWorkbook.Names("regLoteCol").RefersToRange
    If rng Is Nothing Then Set ListaLotes = c: Exit Function
    For Each Cel In rng
        v = Trim$(CStr(Cel.Value))
        If v <> "" Then c.Add v
    Next Cel
    Set ListaLotes = c
End Function

Public Sub AtualizarBanco()
    Application.Calculate
End Sub




Public Sub AtualizarListasAno()
    Dim wsCfg As Worksheet
    Set wsCfg = ThisWorkbook.Sheets("Configuração")
    Dim anosCIQ As Object, anosCEQ As Object
    Set anosCIQ = CreateObject("Scripting.Dictionary")
    Set anosCEQ = CreateObject("Scripting.Dictionary")
    Dim dados As Variant
    dados = CarregarDB()
    If Not IsEmpty(dados) Then
        Dim i As Long
        For i = 1 To UBound(dados, 1)
            If IsDate(dados(i, COL_DATA)) Then
                Dim a As Long: a = Year(CDate(dados(i, COL_DATA)))
                If Not anosCIQ.Exists(a) Then anosCIQ.Add a, True
            End If
        Next i
    End If
    Dim wsEqa As Worksheet
    Set wsEqa = ThisWorkbook.Sheets("EQA_Base")
    Dim ultEqa As Long
    ultEqa = wsEqa.Cells(wsEqa.rows.Count, 2).End(xlUp).Row
    If ultEqa >= 2 Then
        Dim r As Long, v As Variant, ae As Long
        For r = 2 To ultEqa
            v = wsEqa.Cells(r, 2).Value
            If IsNumeric(v) Then
                ae = CLng(v)
                If ae > 1900 And ae < 2200 Then
                    If Not anosCEQ.Exists(ae) Then anosCEQ.Add ae, True
                End If
            End If
        Next r
    End If
    ' ADR-050: so regrava se a lista MUDOU. Regravar igual invalidava
    ' lstAnosCEQ -> Estatistica!N4 (ano de EQA) -> ~480 funcoes de EQA: 10 s
    ' extras em TODA importacao, para escrever os mesmos anos de sempre.
    If ListaIgual(wsCfg.Range("Z2:Z50").Value, AnosOrdenados(anosCIQ)) And _
       ListaIgual(wsCfg.Range("AA2:AA50").Value, AnosOrdenados(anosCEQ)) Then Exit Sub
    Dim protEstava As Boolean
    On Error GoTo restaura
    protEstava = LiberarEscrita(wsCfg)
    wsCfg.Range("Z1:Z50").ClearContents
    wsCfg.Range("AA1:AA50").ClearContents
    wsCfg.Range("Z1").Value = "lstAnosCIQ (auto - nao editar)"
    wsCfg.Range("AA1").Value = "lstAnosCEQ (auto - nao editar)"
    Dim lst As Variant, k As Long
    lst = AnosOrdenados(anosCIQ)
    For k = 0 To UBound(lst)
        wsCfg.Cells(2 + k, 26).Value = lst(k)
    Next k
    lst = AnosOrdenados(anosCEQ)
    For k = 0 To UBound(lst)
        wsCfg.Cells(2 + k, 27).Value = lst(k)
    Next k
    RestaurarProtecao wsCfg, protEstava
    Exit Sub
restaura:
    Dim nErrP As Long, sErrP As String
    nErrP = Err.Number: sErrP = Err.Description
    RestaurarProtecao wsCfg, protEstava
    On Error GoTo 0
    If nErrP <> 0 Then Err.Raise nErrP, "mDados.AtualizarListasAno", sErrP
End Sub

' A coluna (Z2:Z50 lida da planilha) tem exatamente o que a rotina escreveria?
' (AnosOrdenados devolve um unico 0 quando nao ha ano -- e ele e escrito.)
Private Function ListaIgual(ByVal col As Variant, ByVal lst As Variant) As Boolean
    Dim i As Long, n As Long
    n = UBound(lst) + 1
    For i = 1 To UBound(col, 1)
        If i <= n Then
            If CStr(col(i, 1)) <> CStr(lst(i - 1)) Then Exit Function
        ElseIf Len(CStr(col(i, 1))) > 0 Then
            Exit Function
        End If
    Next i
    ListaIgual = True
End Function

Private Function AnosOrdenados(ByVal d As Object) As Variant
    Dim n As Long, arr() As Long, i As Long, kk As Variant
    n = d.Count
    If n = 0 Then
        Dim vazio(0 To 0) As Long
        AnosOrdenados = vazio
        Exit Function
    End If
    ReDim arr(0 To n - 1)
    i = 0
    For Each kk In d.Keys
        arr(i) = CLng(kk): i = i + 1
    Next kk
    Dim aa As Long, bb As Long, tmp As Long
    For aa = 1 To n - 1
        tmp = arr(aa): bb = aa - 1
        Do While bb >= 0
            If arr(bb) <= tmp Then Exit Do
            arr(bb + 1) = arr(bb): bb = bb - 1
        Loop
        arr(bb + 1) = tmp
    Next aa
    AnosOrdenados = arr
End Function

