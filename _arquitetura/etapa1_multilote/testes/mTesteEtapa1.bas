Attribute VB_Name = "mTesteEtapa1"
Option Explicit
' ============================================================================
'  SO EM COPIA DE TESTE (suite_etapa1.py). NUNCA instalar na producao.
'  Cada rotina aqui faz o que o USUARIO faria, pelo mesmo caminho de codigo:
'  cadastrar lote na Configuracao, escolher lote no Painel, digitar media/DP
'  na aba Analitos, gravar resultados pelo UpsertResultados.
' ============================================================================

' Zera os dados operacionais: banco, lotes, views e armazens.
Public Sub T_LimparBase()
    Dim ws As Worksheet, ult As Long
    Application.EnableEvents = False
    Set ws = ThisWorkbook.Sheets("DB_Resultados")
    ult = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    If ult >= 4 Then ws.Range(ws.Cells(4, 1), ws.Cells(ult, 56)).ClearContents
    mBanco.AtualizarFlagsBanco
    ThisWorkbook.Sheets("LotesStore").Range("A2:P5000").ClearContents
    ThisWorkbook.Sheets("Configuração").Range("C26:D125").ClearContents
    ThisWorkbook.Sheets("Configuração").Range("C20").Value = ""
    ThisWorkbook.Names("loteCarregado").RefersToRange.Value = ""
    ThisWorkbook.Names("loteParam").RefersToRange.Value = ""
    ThisWorkbook.Names("loteSel").RefersToRange.Value = ""
    ThisWorkbook.Sheets("LiberStore").Range("A2:D25000").ClearContents
    ThisWorkbook.Sheets("RegistrosStore").Range("A2:L25000").ClearContents
    ThisWorkbook.Names("regView").RefersToRange.ClearContents
    ThisWorkbook.Names("libView").RefersToRange.ClearContents
    ThisWorkbook.Sheets("Analitos").Range("E4:J43").ClearContents
    Set ws = ThisWorkbook.Sheets("Eventos_Westgard")
    ult = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    If ult >= 4 Then ws.Range(ws.Cells(4, 1), ws.Cells(ult, 14)).ClearContents
    mEstatistica.InvalidarCache
End Sub

' Cadastra o lote na primeira linha livre do registro (Configuracao!C26:C125).
Public Function T_Cadastrar(ByVal lote As String) As Long
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets("Configuração")
    For r = 26 To 125
        If Trim$(CStr(ws.Cells(r, 3).Value)) = lote Then T_Cadastrar = r: Exit Function
        If Trim$(CStr(ws.Cells(r, 3).Value)) = "" Then
            ws.Cells(r, 3).NumberFormat = "@"
            ws.Cells(r, 3).Value = lote
            T_Cadastrar = r
            Exit Function
        End If
    Next r
End Function

' Lote EM USO (Configuracao!C20) -- o caminho do Worksheet_Change da Configuracao.
Public Sub T_LoteEmUso(ByVal lote As String)
    ThisWorkbook.Sheets("Configuração").Range("C20").NumberFormat = "@"
    ThisWorkbook.Names("loteAtivo").RefersToRange.Value = lote
    mLotes.TrocarLote
End Sub

' Lote EM ANALISE (Painel!E3) -- o caminho do Worksheet_Change do Painel.
Public Sub T_Selecionar(ByVal lote As String)
    ' como o usuario: o valor entra do jeito que o Excel quiser (010 vira 10);
    ' quem tem de acertar o codigo e o TrocarLoteAnalise
    ThisWorkbook.Names("loteSel").RefersToRange.Value = lote
    mLotes.TrocarLoteAnalise
End Sub

Public Sub T_Analito(ByVal analito As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets("Analitos")
    For r = 4 To 43
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), analito, 1) = 0 Then
            ThisWorkbook.Sheets("Painel").Range("B3").Value = r - 3
            mEstatistica.PainelMudou
            Exit Sub
        End If
    Next r
    Err.Raise vbObjectError + 1, "T_Analito", "analito nao encontrado: " & analito
End Sub

' Digita media/DP do lote EM TELA (loteParam) na aba Analitos e dispara o mesmo
' tratamento do Worksheet_Change. Vazio = "" (apaga).
Public Sub T_Parametros(ByVal analito As String, ByVal m1 As Variant, ByVal s1 As Variant, _
                        ByVal m2 As Variant, ByVal s2 As Variant, _
                        Optional ByVal m3 As Variant, Optional ByVal s3 As Variant)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets("Analitos")
    For r = 4 To 43
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), analito, 1) = 0 Then
            Application.EnableEvents = False
            ws.Cells(r, 5).Value = m1
            ws.Cells(r, 6).Value = s1
            ws.Cells(r, 7).Value = m2
            ws.Cells(r, 8).Value = s2
            If Not IsMissing(m3) Then
                ws.Cells(r, 9).Value = m3
                ws.Cells(r, 10).Value = s3
                mLotes.ParametroEditado ws.Range(ws.Cells(r, 5), ws.Cells(r, 10))
                Exit Sub
            End If
            mLotes.ParametroEditado ws.Range(ws.Cells(r, 5), ws.Cells(r, 8))
            Exit Sub
        End If
    Next r
    Err.Raise vbObjectError + 2, "T_Parametros", "analito nao encontrado: " & analito
End Sub

' Grava resultados lidos da aba T_Entrada (A=Data B=Nivel C=LoteNucleo D=Analito E=Valor)
' pelo caminho real: RUN por (data, lote) e UpsertResultados.
Public Function T_Gravar() As String
    Dim wt As Worksheet, ult As Long, v As Variant, i As Long, n As Long
    Dim db As Variant, runDe As Object, mx As Long, k As String, regs() As Variant
    Set wt = ThisWorkbook.Sheets("T_Entrada")
    ult = wt.Cells(wt.rows.Count, 1).End(xlUp).Row
    If ult < 2 Then T_Gravar = "0|0": Exit Function
    v = wt.Range(wt.Cells(2, 1), wt.Cells(ult + 1, 5)).Value
    n = ult - 1
    Set runDe = CreateObject("Scripting.Dictionary")
    db = mDados.CarregarDB()
    If Not IsEmpty(db) Then
        For i = 1 To UBound(db, 1)
            If IsNumeric(db(i, 1)) And Len(Trim$(CStr(db(i, 1)))) > 0 Then
                If CLng(db(i, 1)) > mx Then mx = CLng(db(i, 1))
                If IsDate(db(i, 2)) Then
                    k = CStr(CLng(CDate(db(i, 2)))) & "|" & NucleoLote(CStr(db(i, 4)))
                    If Not runDe.Exists(k) Then runDe.Add k, CLng(db(i, 1))
                End If
            End If
        Next i
    End If
    ReDim regs(1 To n, 1 To 7)
    For i = 1 To n
        k = CStr(CLng(CDate(v(i, 1)))) & "|" & Trim$(CStr(v(i, 3)))
        If Not runDe.Exists(k) Then mx = mx + 1: runDe.Add k, mx
        regs(i, 1) = runDe(k)
        regs(i, 2) = CDate(v(i, 1))
        regs(i, 3) = CLng(v(i, 2))
        regs(i, 4) = mEntrada.CodigoLote(Trim$(CStr(v(i, 3))), CLng(v(i, 2)))
        regs(i, 5) = CStr(v(i, 4))
        regs(i, 6) = CDbl(v(i, 5))
        regs(i, 7) = "Ativo"
    Next i
    T_Gravar = mDados.UpsertResultados(regs)
    wt.Range(wt.Cells(2, 1), wt.Cells(ult, 5)).ClearContents
    mEstatistica.InvalidarCache
End Function

' Operacao diaria SEM a aba Importar (Hematologia: frmCorrida/frmMassa fazem
' exatamente isto -- UpsertResultados e, em seguida, AtualizarOperacao).
Public Function T_GravarOperacao() As String
    T_GravarOperacao = T_Gravar()
    mOperacao.AtualizarOperacao
End Function

' Recalcula o que o usuario ve: motor do analito em tela + formulas.
Public Sub T_Refazer()
    mEstatistica.InvalidarCache
    mEstatistica.RecalcularAnalitoAtual
    Application.Calculate
End Sub

' Resumo dos eventos de Westgard em tela: "lote|n|RUNs separados por ,"
Public Function T_Eventos() As String
    Dim ws As Worksheet, ult As Long, i As Long, s As String, n As Long
    Set ws = ThisWorkbook.Sheets("Eventos_Westgard")
    ult = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    For i = 4 To ult
        If Len(Trim$(CStr(ws.Cells(i, 2).Value))) > 0 Then
            n = n + 1
            s = s & "," & ws.Cells(i, 2).Value & ":" & ws.Cells(i, 3).Value & ":" & ws.Cells(i, 5).Value & ":" & ws.Cells(i, 4).Value
        End If
    Next i
    T_Eventos = ws.Range("L2").Value & "|" & n & "|" & Mid$(s, 2)
End Function
