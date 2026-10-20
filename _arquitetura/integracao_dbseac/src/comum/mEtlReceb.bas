Attribute VB_Name = "mEtlReceb"
' ============================================================================
'  mEtlReceb -- recebimento do DB_SEAC em VBA (ADR-079)
'
'  Porte de pq/SEAC_ORIGEM.m + pq/DB_RECEBIMENTO.m. Le a tabela do setor no DB_SEAC pelo driver ACE (sem abrir o
'  arquivo no Excel: ~1 s contra 6-14 s do Power Query) e acrescenta a tblDB_Recebimento SO o que ainda nao foi
'  recebido -- diferenca de conjuntos pelo ID_REGISTRO, como o PQ (sem marca d'agua de data: a Hematologia tem
'  resultado que chega dias depois com DATA_HORA antiga).
'  Regras do PQ mantidas:
'    - ID_REGISTRO = PREFIXO_ID & "-" & ID_ORIGEM; sem ID_ORIGEM valido a linha nao entra (QA E06);
'    - ID repetido na mesma carga entra uma vez: o 1o em (ID_REGISTRO, ITEM_ID, DATA_HORA);
'    - linha recebida nunca e alterada; RECEBIDO_EM = um carimbo por execucao (com milissegundos);
'    - a tabela inteira fica ordenada por (DATA_HORA, ID_ORIGEM), nulos primeiro;
'    - MODO_FONTE = HISTORICO nao le o DB_SEAC; falha de leitura nao toca no recebimento.
'  Defesa contra a inferencia de tipo do ACE (ele olha as primeiras linhas): os tipos de cada coluna sao conferidos;
'  tipo inesperado = ERRO (o chamador NAO publica nada e o usuario ve a mensagem).
' ============================================================================
Option Explicit
Option Compare Binary

Private Const AD_TEXTO As String = ",8,129,130,200,201,202,203,"
Private Const AD_NUM As String = ",2,3,4,5,6,14,16,17,18,19,20,21,131,139,"
Private Const AD_DATA As String = ",7,133,134,135,"

' aba do DB_SEAC que guarda a tabela do setor (o ACE enxerga abas, nao tabelas)
Private Function AbaDaTabela(ByVal tabela As String) As String
    Select Case UCase$(Trim$(tabela))
        Case "TBBIOQUIMICA": AbaDaTabela = "DADOS_BIOQUIMICA"
        Case "TBHEMATOLOGIA": AbaDaTabela = "DADOS_HEMATOLOGIA"
        Case Else: AbaDaTabela = ""
    End Select
End Function

Private Function TipoEsperado(ByVal coluna As String) As String
    Select Case coluna
        Case "NIVEL", "ID_ORIGEM", "ITEM_ID", "VALOR": TipoEsperado = AD_NUM
        Case "DATA", "DATA_HORA": TipoEsperado = AD_DATA
        Case Else: TipoEsperado = AD_TEXTO
    End Select
End Function

' Le a tabela do setor no DB_SEAC. Devolve "" (ok) ou a mensagem de erro; nomes = colunas na ordem do DB_SEAC;
' dados = matriz (linha, coluna) 1-based; n = linhas.
Public Function LerOrigem(ByVal caminho As String, ByVal tabela As String, ByRef nomes As Variant, _
                          ByRef dados As Variant, ByRef n As Long) As String
    Dim cn As Object, rs As Object, aba As String, i As Long, j As Long, v As Variant, f As Object, tem As Boolean
    Dim nc As Long
    On Error GoTo falha
    n = 0
    aba = AbaDaTabela(tabela)
    If Len(aba) = 0 Then LerOrigem = "TABELA_ORIGEM desconhecida: " & tabela: Exit Function
    If Len(Dir$(caminho)) = 0 Then LerOrigem = "DB_SEAC nao encontrado: " & caminho: Exit Function
    ' le uma COPIA local: a tarefa agendada salva o DB_SEAC a cada ciclo (ler o arquivo no meio da gravacao perde a
    ' conexao) e abre o arquivo em modo exclusivo -- a nossa leitura nao pode travar a gravacao dela
    Dim tentativa As Long, copia As String, fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    copia = Environ$("TEMP") & "\qcini_db_seac_leitura.xlsm"
    For tentativa = 1 To 6
        Err.Clear
        On Error Resume Next
        fso.CopyFile caminho, copia, True          ' FileCopy recusa arquivo aberto por outro processo
        If Err.Number = 0 Then
            Set cn = CreateObject("ADODB.Connection")
            cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & copia & _
                    ";Extended Properties=""Excel 12.0 Macro;HDR=YES;IMEX=1"";Mode=Read;"
        End If
        If Err.Number = 0 Then Exit For
        If tentativa = 6 Then
            LerOrigem = "nao foi possivel ler o DB_SEAC em 6 tentativas (" & Err.Number & "): " & Err.Description
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
        Application.Wait Now + TimeSerial(0, 0, 5)
    Next tentativa
    On Error GoTo falha
    Set rs = cn.Execute("SELECT * FROM [" & aba & "$]")
    nc = rs.Fields.Count
    ReDim nomes(1 To nc)
    For j = 1 To nc
        Set f = rs.Fields(j - 1)
        nomes(j) = CStr(f.Name)
        If InStr(TipoEsperado(nomes(j)), "," & f.Type & ",") = 0 Then
            LerOrigem = "coluna " & nomes(j) & " do DB_SEAC lida com tipo inesperado (" & f.Type & "); recebimento NAO feito"
            GoTo fecha
        End If
    Next j
    If Not rs.EOF Then
        v = rs.GetRows                                   ' (coluna, linha), 0-based
        n = UBound(v, 2) + 1
        ReDim dados(1 To n, 1 To nc)
        For i = 1 To n
            For j = 1 To nc
                dados(i, j) = v(j - 1, i - 1)
                If IsNull(dados(i, j)) Then dados(i, j) = Empty
            Next j
        Next i
        ' a planilha pode ter linhas em branco no fim do intervalo usado: fora
        Do While n > 0
            tem = False
            For j = 1 To nc
                If Not IsEmpty(dados(n, j)) Then tem = True: Exit For
            Next j
            If tem Then Exit Do
            n = n - 1
        Loop
    End If
fecha:
    On Error Resume Next
    rs.Close
    cn.Close
    Kill copia
    Exit Function
falha:
    LerOrigem = "nao foi possivel ler o DB_SEAC (" & Err.Number & "): " & Err.Description
    On Error Resume Next
    If Not cn Is Nothing Then cn.Close
End Function

' Novos resultados para o recebimento: devolve a matriz (linha, coluna do RECEBIMENTO) so com os IDs ainda nao
' recebidos, ja carimbados. cabRec = cabecalho da tblDB_Recebimento (nome -> coluna); idsRec = IDs ja recebidos.
Public Function NovosParaReceber(ByRef nomes As Variant, ByRef dados As Variant, ByVal n As Long, _
                                 ByVal cabRec As Object, ByVal idsRec As Object, ByVal prefixo As String, _
                                 ByVal instante As Double, ByRef nNovos As Long, ByRef nSemId As Long) As Variant
    Dim cOrig As Object, j As Long, i As Long, ido As Variant, idr As String, cand() As Long, nCand As Long
    Dim chave() As String, ordem() As Long, k As Long, out() As Variant, vistos As Object, c As Variant, cDh As Long
    Dim cIt As Long, cIdo As Long
    nNovos = 0: nSemId = 0
    Set cOrig = CreateObject("Scripting.Dictionary")
    For j = 1 To UBound(nomes)
        cOrig(CStr(nomes(j))) = j
    Next j
    cIdo = cOrig("ID_ORIGEM"): cDh = cOrig("DATA_HORA"): cIt = cOrig("ITEM_ID")
    ReDim cand(1 To IIf(n > 0, n, 1)): ReDim chave(1 To IIf(n > 0, n, 1))
    For i = 1 To n
        ido = dados(i, cIdo)
        If IsEmpty(ido) Or Not IsNumeric(ido) Then
            nSemId = nSemId + 1
        Else
            idr = prefixo & "-" & Format$(CDbl(ido), "0")
            If Not idsRec.Exists(idr) Then
                nCand = nCand + 1: cand(nCand) = i
                chave(nCand) = idr & Chr$(0) & ChaveNum(dados(i, cIt)) & Chr$(0) & ChaveNum(dados(i, cDh))
            End If
        End If
    Next i
    If nCand = 0 Then Exit Function
    ordem = OrdenarTexto(chave, nCand)
    Set vistos = CreateObject("Scripting.Dictionary")
    ReDim out(1 To nCand, 1 To cabRec.Count)
    For k = 1 To nCand
        i = cand(ordem(k))
        idr = prefixo & "-" & Format$(CDbl(dados(i, cIdo)), "0")
        If Not vistos.Exists(idr) Then
            vistos.Add idr, 1
            nNovos = nNovos + 1
            For Each c In cabRec.Keys
                Select Case CStr(c)
                    Case "ID_REGISTRO": out(nNovos, cabRec(c)) = idr
                    Case "RECEBIDO_EM": out(nNovos, cabRec(c)) = instante
                    Case Else
                        If cOrig.Exists(CStr(c)) Then out(nNovos, cabRec(c)) = Valor2(dados(i, cOrig(CStr(c))))
                End Select
            Next c
        End If
    Next k
    NovosParaReceber = out
End Function

' data -> serial (Double), o resto como veio
Private Function Valor2(ByVal v As Variant) As Variant
    If VarType(v) = vbDate Then Valor2 = CDbl(v) Else Valor2 = v
End Function

Private Function ChaveNum(ByVal v As Variant) As String
    Dim s As String
    If IsEmpty(v) Then ChaveNum = Chr$(1): Exit Function
    s = CStr(Int(CDbl(v) * IIf(VarType(v) = vbDate, 86400000#, 1) + 0.5))
    ChaveNum = Chr$(2) & String$(15 - Len(s), "0") & s
End Function

' Junta recebido + novos e ordena por (DATA_HORA, ID_ORIGEM), nulos primeiro. Matrizes no layout do recebimento.
Public Function JuntarOrdenado(ByRef atual As Variant, ByVal nAtual As Long, ByRef novos As Variant, ByVal nNovos As Long, _
                               ByVal cabRec As Object) As Variant
    Dim n As Long, nc As Long, i As Long, j As Long, chave() As String, ordem() As Long, res() As Variant, cDh As Long
    Dim cIdo As Long, src As Long, v As Variant
    n = nAtual + nNovos: nc = cabRec.Count
    cDh = cabRec("DATA_HORA"): cIdo = cabRec("ID_ORIGEM")
    ReDim chave(1 To n)
    For i = 1 To n
        If i <= nAtual Then v = atual(i, cDh) Else v = novos(i - nAtual, cDh)
        If IsEmpty(v) Then chave(i) = Chr$(1) Else chave(i) = ChaveNum(CDbl(v) * 86400000#)
        If i <= nAtual Then v = atual(i, cIdo) Else v = novos(i - nAtual, cIdo)
        chave(i) = chave(i) & Chr$(0) & ChaveNum(v)
    Next i
    ordem = OrdenarTexto(chave, n)
    ReDim res(1 To n, 1 To nc)
    For i = 1 To n
        src = ordem(i)
        For j = 1 To nc
            If src <= nAtual Then res(i, j) = atual(src, j) Else res(i, j) = novos(src - nAtual, j)
        Next j
    Next i
    JuntarOrdenado = res
End Function

' indices ordenados por chave texto (binario), mesclagem
Private Function OrdenarTexto(ByRef chave() As String, ByVal n As Long) As Long()
    Dim a() As Long, b() As Long, w As Long, i As Long, lo As Long, md As Long, hi As Long, l As Long, r As Long, k As Long
    Dim t() As Long
    ReDim a(1 To IIf(n > 0, n, 1)): ReDim b(1 To IIf(n > 0, n, 1))
    For i = 1 To n: a(i) = i: Next i
    w = 1
    Do While w < n
        lo = 1
        Do While lo <= n
            md = lo + w - 1: If md > n Then md = n
            hi = lo + 2 * w - 1: If hi > n Then hi = n
            l = lo: r = md + 1: k = lo
            Do While l <= md And r <= hi
                If chave(a(r)) < chave(a(l)) Then b(k) = a(r): r = r + 1 Else b(k) = a(l): l = l + 1
                k = k + 1
            Loop
            Do While l <= md: b(k) = a(l): l = l + 1: k = k + 1: Loop
            Do While r <= hi: b(k) = a(r): r = r + 1: k = k + 1: Loop
            lo = lo + 2 * w
        Loop
        t = a: a = b: b = t
        w = w * 2
    Loop
    OrdenarTexto = a
End Function

' carimbo local com milissegundos (o Now do VBA so tem segundos)
Public Function Carimbo() As Double
    Dim s As Double
    s = Timer
    Carimbo = CDbl(Date) + Int(s * 1000# + 0.5) / 86400000#
End Function
