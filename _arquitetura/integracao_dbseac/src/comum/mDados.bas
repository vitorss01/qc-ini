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

' Ultima data com resultado que PARTICIPA da estatistica (0 se nao houver). So o que participa:
' uma data futura digitada por engano num resultado manual vira MANUAL_INCOMPLETO no Power Query,
' mas continua na tabela -- o MAX da coluna inteira levava o ano padrao do Painel e a janela da
' incerteza para o futuro (auditoria 04/10/2026).
Public Function UltimaDataCQ() As Double
    Dim lo As ListObject
    On Error Resume Next
    Set lo = TabelaCQ()
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function
    UltimaDataCQ = Application.WorksheetFunction.MaxIfs(lo.ListColumns("DATA").DataBodyRange, _
                       lo.ListColumns("PARTICIPA_ESTATISTICA").DataBodyRange, "SIM")
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
    Dim anosCIQ As Object
    Set anosCIQ = CreateObject("Scripting.Dictionary")
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
    ' ADR-071: as listas do CEQ saem do mCEQ -- a MESMA regra de "rodada existente" do calculo
    ' (Uso_Analitico <> NAO, analito canonico, |bias| numerico): rodada de simulacao (Uso = NAO)
    ' e linha sem analito canonico nunca aparecem. Antes a lista de anos lia EQA_Base!B inteira.
    Dim lst(1 To 7) As Variant, cols As Variant, cabs As Variant, k As Long, mudou As Boolean, igual(1 To 7) As Boolean
    lst(1) = AnosOrdenados(anosCIQ)
    lst(2) = mCEQ.AnosExistentes("")
    If UBound(lst(2)) < 0 Then lst(2) = AnosOrdenados(CreateObject("Scripting.Dictionary"))  ' nenhum: um 0, como antes
    lst(3) = mCEQ.AnosExistentes("CAP")
    lst(4) = mCEQ.AnosExistentes("Controllab")
    lst(5) = ComFixos(mCEQ.RodadasExistentes("CAP"))
    lst(6) = ComFixos(mCEQ.RodadasExistentes("Controllab"))
    lst(7) = ComFixos(mCEQ.RodadasExistentes(""))
    '            Z    AA   AF   AG   AH   AI   AJ
    cols = Array(26, 27, 32, 33, 34, 35, 36)
    cabs = Array("lstAnosCIQ", "lstAnosCEQ", "lstAnosCAP", "lstAnosCTL", "lstRodadasCAP", "lstRodadasCTL", "lstRodadasEQA")
    ' ADR-050: so regrava a lista que MUDOU. Regravar igual invalidava
    ' lstAnosCEQ -> Estatistica!N4 (ano de EQA) -> ~480 funcoes de EQA: 10 s
    ' extras em TODA atualizacao, para escrever os mesmos anos de sempre.
    For k = 1 To 7
        igual(k) = ListaIgual(wsCfg.Range(wsCfg.Cells(2, cols(k - 1)), wsCfg.Cells(LinhasLista(k), cols(k - 1))).Value, lst(k))
        If Not igual(k) Then mudou = True
    Next k
    If Not mudou Then Exit Sub
    Dim protEstava As Boolean
    On Error GoTo restaura
    protEstava = LiberarEscrita(wsCfg)
    Dim j As Long
    For k = 1 To 7
        If Not igual(k) Then
            wsCfg.Range(wsCfg.Cells(1, cols(k - 1)), wsCfg.Cells(LinhasLista(k), cols(k - 1))).ClearContents
            wsCfg.Cells(1, cols(k - 1)).Value = cabs(k - 1) & " (auto - nao editar)"
            For j = 0 To UBound(lst(k))
                If 2 + j <= LinhasLista(k) Then wsCfg.Cells(2 + j, cols(k - 1)).Value = lst(k)(j)
            Next j
        End If
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

' Ultima linha de cada lista: Z/AA ate 50 (como antes); as do ADR-071 ate 101.
Private Function LinhasLista(ByVal k As Long) As Long
    LinhasLista = IIf(k <= 2, 50, 101)
End Function

' ADR-071: itens fixos da selecao de rodadas antes das rodadas existentes (ano|rodada).
Private Function ComFixos(ByVal rods As Variant) As Variant
    Dim out() As String, i As Long, n As Long
    n = UBound(rods) + 1
    ReDim out(0 To 3 + n)
    out(0) = "TODAS": out(1) = "ACUMULADAS": out(2) = "ULTIMAS 3": out(3) = "ULTIMAS 6"
    For i = 0 To n - 1
        out(4 + i) = rods(i)
    Next i
    ComFixos = out
End Function

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
