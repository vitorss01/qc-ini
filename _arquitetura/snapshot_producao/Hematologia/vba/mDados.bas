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
' regs: array (1..n, 1..7) já no schema do banco. Atualiza o que existir
' (mesma chave RUN|Nivel|Analito) e acrescenta o resto — nunca duplica.
' Devolve "novos|atualizados".
Public Function UpsertResultados(ByRef regs As Variant) As String
    Dim ws As Worksheet, dados As Variant, idx As Object
    Dim i As Long, lastRow As Long, novos As Long, atual As Long, k As String, lin As Long
    Dim addBuf() As Variant, nAdd As Long
    If IsEmpty(regs) Then UpsertResultados = "0|0": Exit Function
    Set ws = ThisWorkbook.Sheets(BANCO)
    Set idx = CreateObject("Scripting.Dictionary")
    dados = CarregarDB()
    If Not IsEmpty(dados) Then
        For i = 1 To UBound(dados, 1)
            If Len(Trim$(CStr(dados(i, COL_ANALITO)))) > 0 Then
                k = ChaveReg(CLng(dados(i, COL_RUN)), CLng(dados(i, COL_NIVEL)), CStr(dados(i, COL_ANALITO)))
                If Not idx.Exists(k) Then idx.Add k, BANCO_R0 + i - 1
            End If
        Next i
    End If
    lastRow = UltimaLinhaBanco()
    ReDim addBuf(1 To UBound(regs, 1), 1 To COL_STATUS)
    nAdd = 0: novos = 0: atual = 0
    Application.ScreenUpdating = False
    For i = 1 To UBound(regs, 1)
        k = ChaveReg(CLng(regs(i, COL_RUN)), CLng(regs(i, COL_NIVEL)), CStr(regs(i, COL_ANALITO)))
        If idx.Exists(k) Then
            lin = idx(k)
            ws.Cells(lin, COL_RESULT).Value = regs(i, COL_RESULT)
            ws.Cells(lin, COL_STATUS).Value = ST_ATIVO
            ws.Cells(lin, COL_DATA).Value = regs(i, COL_DATA)
            ws.Cells(lin, COL_LOTE).Value = regs(i, COL_LOTE)
            atual = atual + 1
        Else
            nAdd = nAdd + 1
            Dim c As Long
            For c = 1 To COL_STATUS
                addBuf(nAdd, c) = regs(i, c)
            Next c
            idx.Add k, lastRow + nAdd
            novos = novos + 1
        End If
    Next i
    If nAdd > 0 Then
        ws.Range(ws.Cells(lastRow + 1, COL_LOTE), ws.Cells(lastRow + nAdd, COL_LOTE)).NumberFormat = "@"
        Dim outp() As Variant
        ReDim outp(1 To nAdd, 1 To COL_STATUS)
        For i = 1 To nAdd
            For c = 1 To COL_STATUS
                outp(i, c) = addBuf(i, c)
            Next c
        Next i
        ws.Range(ws.Cells(lastRow + 1, COL_RUN), ws.Cells(lastRow + nAdd, COL_STATUS)).Value = outp
    End If
    Application.ScreenUpdating = True
    UpsertResultados = CStr(novos) & "|" & CStr(atual)
End Function

' Exclusao LOGICA por RUN + Nivel + lista de analitos (Dictionary de nomes em UCase).
Public Function ExcluirLogico(ByVal run As Long, ByVal nivel As Long, ByRef alvoS As Object) As Long
    Dim ws As Worksheet, dados As Variant, i As Long, n As Long
    Set ws = ThisWorkbook.Sheets(BANCO)
    dados = CarregarDB()
    If IsEmpty(dados) Then ExcluirLogico = 0: Exit Function
    Application.ScreenUpdating = False
    For i = 1 To UBound(dados, 1)
        If Len(Trim$(CStr(dados(i, COL_ANALITO)))) > 0 Then
            If CLng(dados(i, COL_RUN)) = run And CLng(dados(i, COL_NIVEL)) = nivel Then
                If alvoS.Exists(UCase$(Trim$(CStr(dados(i, COL_ANALITO))))) Then
                    ws.Cells(BANCO_R0 + i - 1, COL_STATUS).Value = ST_EXCLUIDO
                    n = n + 1
                End If
            End If
        End If
    Next i
    Application.ScreenUpdating = True
    ExcluirLogico = n
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
    Dim rng As Range, cel As Range, c As Collection, v As String
    Set c = New Collection
    On Error Resume Next
    Set rng = ThisWorkbook.Names("regLoteCol").RefersToRange
    If rng Is Nothing Then Set ListaLotes = c: Exit Function
    For Each cel In rng
        v = Trim$(CStr(cel.Value))
        If v <> "" Then c.Add v
    Next cel
    Set ListaLotes = c
End Function

Public Sub AtualizarBanco()
    Application.Calculate
End Sub

' Placeholder — trilha de auditoria e da Fase 5.
Public Sub RegistrarLog(ByVal acao As String, ByVal detalhe As String)
End Sub

