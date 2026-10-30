Attribute VB_Name = "mCEQForm"
Option Explicit
' ============================================================================
'  mCEQForm -- registro e consulta do Controle Externo da Qualidade (ADR-087)
'
'  A janela frmCEQ organiza o registro na sequencia natural de quem tem o
'  relatorio do provedor na mao: 1) a rodada (provedor, survey/rodada, ano);
'  2) o analito e a amostra; 3) o resultado do laboratorio e o do provedor
'  (alvo, SD, SDI, limites, n de laboratorios); 4) a avaliacao e a fonte.
'  Grava na MESMA tabela de digitacao do provedor de sempre (tblEQA_CAP_Dados /
'  tblEQA_Controllab_Dados, colunas A..N, Q, R -- as de formula nao sao tocadas)
'  e consolida a EQA_Base pelo mesmo mEQA.AtualizarEQABase. Nenhum campo mudou de
'  significado e nenhuma regra do CEQ mudou: a janela so organiza, valida e evita
'  duplicidade (a chave da EQA_Base: provedor|ano|rodada|analito|amostra).
'  Digitar direto na tabela continua possivel, como antes.
' ============================================================================
Private Const TB_CAP As String = "tblEQA_CAP_Dados"
Private Const TB_CTL As String = "tblEQA_Controllab_Dados"
' colunas digitadas da tabela do provedor (1-based)
Public Const CQ_PROV As Long = 1
Public Const CQ_RODADA As Long = 2
Public Const CQ_ANO As Long = 3
Public Const CQ_ANALITO As Long = 4
Public Const CQ_AMOSTRA As Long = 5
Public Const CQ_RES As Long = 6
Public Const CQ_ALVO As Long = 7
Public Const CQ_SD As Long = 8
Public Const CQ_SDI As Long = 9
Public Const CQ_LINF As Long = 10
Public Const CQ_LSUP As Long = 11
Public Const CQ_AVAL As Long = 12
Public Const CQ_NLABS As Long = 13
Public Const CQ_UNID As Long = 14
Public Const CQ_PAG As Long = 17
Public Const CQ_ARQ As Long = 18

Public Sub AbrirCEQ()
    On Error Resume Next
    Dim f As Object
    Set f = VBA.UserForms.Add("frmCEQ")
    If f Is Nothing Then MsgBox "A janela de registro do CEQ nao esta instalada.", vbExclamation: Exit Sub
    f.Show
    Unload f
End Sub

Public Function TabelaCEQ(ByVal provedor As String) As ListObject
    Dim nome As String, ws As Worksheet
    On Error Resume Next
    If UCase$(Trim$(provedor)) = "CAP" Then nome = TB_CAP Else nome = TB_CTL
    For Each ws In ThisWorkbook.Worksheets
        Set TabelaCEQ = ws.ListObjects(nome)
        If Not TabelaCEQ Is Nothing Then Exit Function
    Next ws
End Function

Private Function Txt(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    Txt = Trim$(CStr(v))
End Function

' Numero digitado (virgula ou ponto decimal). ok = False quando nao e numero; vazio = Empty e ok = True.
Public Function NumeroCEQ(ByVal s As String, ByRef ok As Boolean) As Variant
    Dim t As String
    ok = True
    t = Trim$(s)
    If Len(t) = 0 Then NumeroCEQ = Empty: Exit Function
    t = Replace(t, " ", "")
    If InStr(t, ",") > 0 And InStr(t, ".") > 0 Then t = Replace(t, ".", "")   ' 1.234,5
    t = Replace(t, ",", Application.International(xlDecimalSeparator))
    t = Replace(t, ".", Application.International(xlDecimalSeparator))
    If IsNumeric(t) Then NumeroCEQ = CDbl(t) Else ok = False
End Function

' Valores distintos (ordem alfabetica; rodadas e anos em ordem decrescente) de uma coluna da tabela do provedor.
' filtroCol/filtroVal: so as linhas com aquele valor (ex.: os analitos de uma rodada). Para as listas da janela.
Public Function DistintosCEQ(ByVal provedor As String, ByVal coluna As Long, Optional ByVal decrescente As Boolean = False, _
                             Optional ByVal filtroCol As Long = 0, Optional ByVal filtroVal As String = "") As Variant
    Dim lo As ListObject, v As Variant, d As Object, i As Long, s As String, a() As String, j As Long, t As String, n As Long
    Dim k As Variant
    DistintosCEQ = Array()
    On Error GoTo sai
    Set lo = TabelaCEQ(provedor)
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function
    v = lo.DataBodyRange.Value
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    For i = 1 To UBound(v, 1)
        s = Txt(v(i, coluna))
        If Len(s) > 0 And Len(Txt(v(i, CQ_ANALITO))) > 0 Then
            If filtroCol = 0 Or StrComp(Txt(v(i, IIf(filtroCol > 0, filtroCol, 1))), filtroVal, vbTextCompare) = 0 Then
                If Not d.Exists(s) Then d.Add s, True
            End If
        End If
    Next i
    n = d.Count
    If n = 0 Then Exit Function
    ReDim a(0 To n - 1)
    i = 0
    For Each k In d.Keys
        a(i) = CStr(k): i = i + 1
    Next k
    For i = 1 To n - 1
        t = a(i): j = i - 1
        Do While j >= 0
            If decrescente Then
                If StrComp(a(j), t, vbTextCompare) >= 0 Then Exit Do
            Else
                If StrComp(a(j), t, vbTextCompare) <= 0 Then Exit Do
            End If
            a(j + 1) = a(j): j = j - 1
        Loop
        a(j + 1) = t
    Next i
    DistintosCEQ = a
sai:
End Function

' Primeiro valor da coluna "coluna" na linha mais recente em que "chaveCol" = chaveVal (ex.: o ano de uma rodada que
' ja existe, a unidade de um analito). "" se nao houver. Reaproveita o que ja esta registrado, sem inventar.
Public Function ValorRegistrado(ByVal provedor As String, ByVal chaveCol As Long, ByVal chaveVal As String, _
                                ByVal coluna As Long) As String
    Dim lo As ListObject, v As Variant, i As Long
    On Error GoTo sai
    Set lo = TabelaCEQ(provedor)
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function
    v = lo.DataBodyRange.Value
    For i = UBound(v, 1) To 1 Step -1
        If StrComp(Txt(v(i, chaveCol)), chaveVal, vbTextCompare) = 0 And Len(Txt(v(i, coluna))) > 0 Then
            ValorRegistrado = Txt(v(i, coluna))
            Exit Function
        End If
    Next i
sai:
End Function

' Registros da rodada para a lista da janela: matriz (n x 4) Analito | Amostra | Resultado | Avaliacao; Empty = nenhum.
Public Function RegistrosDaRodada(ByVal provedor As String, ByVal rodada As String, Optional ByVal ano As String = "") As Variant
    Dim lo As ListObject, v As Variant, i As Long, n As Long, out() As Variant, idx() As Long
    On Error GoTo sai
    Set lo = TabelaCEQ(provedor)
    If lo Is Nothing Or Len(rodada) = 0 Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function
    v = lo.DataBodyRange.Value
    ReDim idx(1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        If StrComp(Txt(v(i, CQ_RODADA)), rodada, vbTextCompare) = 0 And Len(Txt(v(i, CQ_ANALITO))) > 0 Then
            If Len(ano) = 0 Or Txt(v(i, CQ_ANO)) = ano Then n = n + 1: idx(n) = i
        End If
    Next i
    If n = 0 Then Exit Function
    ReDim out(1 To n, 1 To 4)
    For i = 1 To n
        out(i, 1) = Txt(v(idx(i), CQ_ANALITO))
        out(i, 2) = Txt(v(idx(i), CQ_AMOSTRA))
        out(i, 3) = Txt(v(idx(i), CQ_RES))
        out(i, 4) = Txt(v(idx(i), CQ_AVAL))
    Next i
    RegistrosDaRodada = out
sai:
End Function

' Grava um resultado de CEQ (nucleo sem dialogo -- a janela e o QA chamam este). campos = Array de 16 textos na ordem
' Provedor, Rodada, Ano, Analito, Amostra, Resultado, Alvo, SD, SDI, LimInf, LimSup, Avaliacao, NLabs, Unidade,
' Pagina, Arquivo. Valida (obrigatorios, numeros, limites, duplicidade pela chave da EQA_Base) ANTES de gravar; grava
' na 1a linha vazia da tabela do provedor (ou acrescenta uma); registra CEQ_REGISTRADO no Audit_Log; consolida a
' EQA_Base. Devolve "OK|<linha da aba>|<aviso>" ou "ERRO|<campo>|<mensagem>" (nada gravado no ERRO).
Public Function SalvarCEQ(ByVal campos As Variant) As String
    Dim prov As String, rod As String, ano As String, an As String, am As String, i As Long
    Dim lo As ListObject, ws As Worksheet, v As Variant, r As Long, prot As Boolean, ev As Boolean, sE As String
    Dim nums(1 To 16) As Variant, ok As Boolean, aviso As String, livre As Long, nomes As Variant
    On Error GoTo falha
    nomes = Array("", "Provedor", "Rodada / Survey", "Ano", "Analito", "Amostra", "Seu resultado", "Valor-alvo", "SD", _
                  "SDI", "Limite inferior", "Limite superior", "Avaliação", "Nº de laboratórios", "Unidade", "Página", _
                  "Arquivo fonte")
    prov = Trim$(CStr(campos(0))): rod = Trim$(CStr(campos(1))): ano = Trim$(CStr(campos(2)))
    an = Trim$(CStr(campos(3))): am = Trim$(CStr(campos(4)))
    If UCase$(prov) <> "CAP" And UCase$(prov) <> "CONTROLLAB" Then SalvarCEQ = "ERRO|1|Escolha o provedor (CAP ou Controllab).": Exit Function
    If Len(rod) = 0 Then SalvarCEQ = "ERRO|2|Informe a rodada / survey.": Exit Function
    ' so os 4 algarismos (IsNumeric aceitava "2,02", "20.6", "-202" e "9e99", que estourava o CLng no meio da gravacao)
    If Not ano Like "####" Or Val(ano) < 1990 Or Val(ano) > 2100 Then SalvarCEQ = "ERRO|3|Ano com 4 dígitos (ex.: 2026).": Exit Function
    If Len(an) = 0 Then SalvarCEQ = "ERRO|4|Informe o analito (nome como no relatório do provedor).": Exit Function
    If Len(am) = 0 Then SalvarCEQ = "ERRO|5|Informe a amostra (ex.: CHM-01).": Exit Function
    Dim cn As Variant
    For Each cn In Array(6, 7, 8, 9, 10, 11, 13, 15)
        nums(cn) = NumeroCEQ(CStr(campos(cn - 1)), ok)
        If Not ok Then SalvarCEQ = "ERRO|" & cn & "|" & nomes(cn) & ": digite um número.": Exit Function
    Next cn
    If IsEmpty(nums(6)) Then SalvarCEQ = "ERRO|6|Informe o resultado do laboratório.": Exit Function
    If Len(Trim$(CStr(campos(11)))) = 0 Then SalvarCEQ = "ERRO|12|Informe a avaliação do provedor.": Exit Function
    If Not IsEmpty(nums(10)) And Not IsEmpty(nums(11)) Then
        If nums(10) > nums(11) Then SalvarCEQ = "ERRO|10|O limite inferior está maior que o superior.": Exit Function
    End If
    If Not IsEmpty(nums(8)) Then
        If nums(8) < 0 Then SalvarCEQ = "ERRO|8|O SD não pode ser negativo.": Exit Function
    End If
    Set lo = TabelaCEQ(prov)
    If lo Is Nothing Then SalvarCEQ = "ERRO|1|A tabela de digitação do " & prov & " não foi encontrada.": Exit Function
    Set ws = lo.Parent
    ' duplicidade: a chave da EQA_Base (provedor|ano|rodada|analito|amostra)
    If lo.ListRows.Count > 0 Then
        v = lo.DataBodyRange.Value
        For r = 1 To UBound(v, 1)
            If Len(Txt(v(r, CQ_ANALITO))) > 0 Then
                If StrComp(Txt(v(r, CQ_RODADA)), rod, vbTextCompare) = 0 And Txt(v(r, CQ_ANO)) = ano And _
                   StrComp(Txt(v(r, CQ_ANALITO)), an, vbTextCompare) = 0 And _
                   StrComp(Txt(v(r, CQ_AMOSTRA)), am, vbTextCompare) = 0 Then
                    SalvarCEQ = "ERRO|5|Já existe este registro (" & prov & " " & rod & " " & ano & " · " & an & " · " & am & _
                                ") na linha " & (lo.DataBodyRange.Row + r - 1) & " da aba " & ws.Name & _
                                ". Para corrigir, edite a linha existente."
                    Exit Function
                End If
            ElseIf livre = 0 Then
                If Len(Txt(v(r, CQ_RODADA)) & Txt(v(r, CQ_AMOSTRA)) & Txt(v(r, CQ_RES))) = 0 Then livre = r
            End If
        Next r
    End If
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = mSeguranca.LiberarEscrita(ws)
    If livre = 0 Then
        lo.ListRows.Add
        livre = lo.ListRows.Count
    End If
    r = lo.DataBodyRange.Row + livre - 1
    ws.Cells(r, lo.Range.Column + CQ_PROV - 1).Value = IIf(UCase$(prov) = "CAP", "CAP", "Controllab")
    ws.Cells(r, lo.Range.Column + CQ_RODADA - 1).Value = "'" & rod
    ws.Cells(r, lo.Range.Column + CQ_ANO - 1).Value = CLng(ano)
    ws.Cells(r, lo.Range.Column + CQ_ANALITO - 1).Value = "'" & an
    ws.Cells(r, lo.Range.Column + CQ_AMOSTRA - 1).Value = "'" & am
    For Each cn In Array(6, 7, 8, 9, 10, 11, 13)
        If IsEmpty(nums(cn)) Then
            ws.Cells(r, lo.Range.Column + cn - 1).ClearContents
        Else
            ws.Cells(r, lo.Range.Column + cn - 1).Value = nums(cn)
        End If
    Next cn
    ws.Cells(r, lo.Range.Column + CQ_AVAL - 1).Value = "'" & Trim$(CStr(campos(11)))
    ws.Cells(r, lo.Range.Column + CQ_UNID - 1).Value = "'" & Trim$(CStr(campos(13)))
    If IsEmpty(nums(15)) Then
        ws.Cells(r, lo.Range.Column + CQ_PAG - 1).ClearContents
    Else
        ws.Cells(r, lo.Range.Column + CQ_PAG - 1).Value = nums(15)
    End If
    ws.Cells(r, lo.Range.Column + CQ_ARQ - 1).Value = "'" & Trim$(CStr(campos(15)))
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    On Error Resume Next
    mAuditoria.RegistrarLog "CEQ_REGISTRADO", prov & " | " & rod & " | " & ano & " | " & an & " | " & am & " | resultado " & _
                            CStr(nums(6)) & " | avaliação " & Trim$(CStr(campos(11))) & " | aba " & ws.Name & " linha " & r
    Err.Clear
    mEQA.AtualizarEQABase
    If Err.Number <> 0 Then aviso = "Gravado, mas a consolidação da EQA_Base falhou (" & Err.Description & "). " & _
                                    "Clique em ATUALIZAR DADOS."
    On Error GoTo 0
    SalvarCEQ = "OK|" & r & "|" & aviso
    Exit Function
falha:
    sE = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then mSeguranca.RestaurarProtecao ws, prot
    If ev Then Application.EnableEvents = True
    SalvarCEQ = "ERRO|0|Não foi possível gravar (detalhe: " & sE & "). Nada foi alterado."
End Function
