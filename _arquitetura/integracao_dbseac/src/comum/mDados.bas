Attribute VB_Name = "mDados"
Option Explicit
' ===== CAMADA DE DADOS (ADR-057) =====
' A fonte analitica UNICA e a tabela tblCQ_Final (aba Principal - Resultados), gerada so
' pelo Power Query a partir do DB_SEAC, dos resultados manuais, da inativacao e
' dos comentarios tecnicos. O VBA NAO grava resultado nenhum: le.
'
' Antes (ate o ADR-056) a fonte era a aba DB_Resultados, gravada por formulario,
' pela aba Importar e pela exclusao logica. Tudo isso saiu: um resultado entra
' pelo interfaceamento ou pela tabela de resultados manuais, e sai da estatistica
' pela tabela de inativacao -- sempre por Power Query, nunca por VBA.
'
' CarregarDB devolve a mesma matriz POSICIONAL que o motor ja conhecia nas
' colunas 1..7 (para que o motor nao mude onde nao precisa) e acrescenta 8..13:
'   1 RUN        yymmddkk (data + corrida do dia), calculado no Power Query
'   2 DATA       DATA_HORA do resultado (o filtro de periodo compara por dia)
'   3 NIVEL      4 LOTE (nucleo, texto)   5 ANALITO (nome do cadastro)
'   6 RESULTADO
'   7 STATUS     "Ativo" quando PARTICIPA_ESTATISTICA = SIM, senao "Inativo".
'                A decisao e do Power Query; aqui so muda de nome.
'   8 TIPO_PLOTAGEM_LJ  NORMAL | X_VERMELHO | NAO_PLOTAR
'   9 EQUIPAMENTO       10 ID_REGISTRO       11 STATUS_ANALITICO
'  12 ORIGEM_RESULTADO  13 REGISTRAR_RESULTADO_NO_LJ
Public Const FONTE_CQ As String = "tblCQ_Final"
Public Const ABA_CQ As String = "Principal - Resultados"   ' tabela tblCQ_Final (DB_CQ_FINAL)
Public Const COL_RUN As Long = 1
Public Const COL_DATA As Long = 2
Public Const COL_NIVEL As Long = 3
Public Const COL_LOTE As Long = 4
Public Const COL_ANALITO As Long = 5
Public Const COL_RESULT As Long = 6
Public Const COL_STATUS As Long = 7
Public Const COL_PLOT As Long = 8
Public Const COL_EQUIP As Long = 9
Public Const COL_ID As Long = 10
Public Const COL_STATUS_AN As Long = 11
Public Const COL_ORIGEM As Long = 12
Public Const COL_REGLJ As Long = 13
Public Const DB_NCOL As Long = 13
Public Const ST_ATIVO As String = "Ativo"
Public Const ST_INATIVO As String = "Inativo"
Public Const PLOT_NORMAL As String = "NORMAL"
Public Const PLOT_X As String = "X_VERMELHO"

Public Function LoteAtivoCore() As String
    On Error Resume Next
    LoteAtivoCore = Trim$(CStr(ThisWorkbook.Names("loteAtivo").RefersToRange.Value))
End Function

' A tabela oficial; Nothing se a camada de dados ainda nao foi instalada.
Public Function TabelaCQ() As ListObject
    On Error Resume Next
    Set TabelaCQ = ThisWorkbook.Sheets(ABA_CQ).ListObjects(FONTE_CQ)
End Function

' Le a fonte oficial para memoria de uma vez: so as 13 colunas que o sistema usa,
' uma leitura em bloco por coluna (nunca celula a celula). Coluna achada pelo
' CABECALHO, nao pela posicao: o Power Query pode acrescentar campos sem
' quebrar o motor.
Public Function CarregarDB() As Variant
    Dim lo As ListObject, n As Long, i As Long, j As Long, out() As Variant
    Dim nomes As Variant, cols(1 To DB_NCOL) As Variant, v As Variant
    Set lo = TabelaCQ()
    If lo Is Nothing Then CarregarDB = Empty: Exit Function
    If lo.ListRows.Count = 0 Then CarregarDB = Empty: Exit Function
    nomes = Array("RUN", "DATA_HORA", "NIVEL", "LOTE", "ANALITO", "RESULTADO", "PARTICIPA_ESTATISTICA", _
                  "TIPO_PLOTAGEM_LJ", "EQUIPAMENTO", "ID_REGISTRO", "STATUS_ANALITICO", "ORIGEM_RESULTADO", _
                  "REGISTRAR_RESULTADO_NO_LJ")
    n = lo.ListRows.Count
    For j = 1 To DB_NCOL
        cols(j) = lo.ListColumns(nomes(j - 1)).DataBodyRange.Value     ' erro aqui = schema quebrado: falha alto
    Next j
    ReDim out(1 To n, 1 To DB_NCOL)
    For i = 1 To n
        For j = 1 To DB_NCOL
            If n = 1 Then v = cols(j) Else v = cols(j)(i, 1)
            out(i, j) = v
        Next j
        out(i, COL_LOTE) = Trim$(CStr(out(i, COL_LOTE)))
        If UCase$(Trim$(CStr(out(i, COL_STATUS)))) = "SIM" Then
            out(i, COL_STATUS) = ST_ATIVO
        Else
            out(i, COL_STATUS) = ST_INATIVO
        End If
    Next i
    CarregarDB = out
End Function

' Ultima data com resultado na fonte oficial (0 se vazia).
Public Function UltimaDataCQ() As Double
    Dim lo As ListObject
    On Error Resume Next
    Set lo = TabelaCQ()
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function
    UltimaDataCQ = Application.WorksheetFunction.Max(lo.ListColumns("DATA").DataBodyRange)
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
    ' extras em TODA atualizacao, para escrever os mesmos anos de sempre.
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
