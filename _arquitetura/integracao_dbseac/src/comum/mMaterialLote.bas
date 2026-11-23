Attribute VB_Name = "mMaterialLote"
Option Explicit
' CAMADA: Lotes -- ADR-099: coluna MATERIAL no cadastro de lotes (Configuracao!F26:F125)
'
' Pedido do usuario (22/11/2026): informar, ao lado de cada lote e da sua validade, o material do controle (soro,
' urina, ...), para o sistema direcionar cada resultado. A FONTE dos motores continua UNICA: a regra MATRIZ da
' tblLotesMaterial (ADR-091), lida igual pelo VBA (mEtlCQ) e pelo PQ (DB_CQ_FINAL). Esta coluna e a entrada amigavel:
'   - digitar o material grava/troca a 1a regra MATRIZ do lote (e audita);
'   - apagar remove as regras MATRIZ do lote (SO_ANALITOS e FORA_DO_CQ nunca sao tocadas);
'   - a coluna mostra sempre a regra que vale (espelho refeito no ATUALIZAR, ao abrir e ao cadastrar lote).
' Lista permitida = as matrizes do de/para (+ MATRIZ_PADRAO), em maiusculas como o de/para: um material sem de/para
' tiraria o lote inteiro do catalogo. Vazio = o padrao de hoje (a matriz que vem com o resultado / MATRIZ_PADRAO).

Private Const ABA_CAD As String = "Configuração"
Private Const R0 As Long = 26
Private Const R1 As Long = 125
Private Const C_LOTE As Long = 3            ' C codigo (regLoteCol)
Private Const C_MAT As Long = 6             ' F material (ADR-099)
Private Const TB As String = "tblLotesMaterial"
Private Const TB_DP As String = "tblDeParaAnalitos"

Private Function Chave(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    Chave = UCase$(Trim$(CStr(v)))
End Function

Private Sub Avisar(ByVal msg As String, ByVal estilo As VbMsgBoxStyle)
    If Application.Interactive Then MsgBox msg, estilo, "Material do lote"
End Sub

' "SORO;URINA" (Bio) / "SANGUE TOTAL" (Hema): matrizes nao vazias do de/para + MATRIZ_PADRAO, sem repetir.
Public Function MatrizesPermitidas() As String
    Dim lo As ListObject, v As Variant, i As Long, s As String, k As String, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set lo = mIntegracao.AcharTabela(TB_DP)
    If Not lo Is Nothing Then
        If lo.ListRows.Count > 0 Then
            v = lo.ListColumns("MATRIZ").DataBodyRange.Value2
            If Not IsArray(v) Then
                k = Chave(v): If Len(k) > 0 Then d(k) = 1
            Else
                For i = 1 To UBound(v, 1)
                    k = Chave(v(i, 1)): If Len(k) > 0 Then d(k) = 1
                Next i
            End If
        End If
    End If
    k = Chave(mIntegracao.Cfg("MATRIZ_PADRAO")): If Len(k) > 0 Then d(k) = 1
    On Error GoTo 0
    For Each v In d.Keys
        s = s & IIf(Len(s) > 0, ";", "") & v
    Next v
    MatrizesPermitidas = s
End Function

Private Function Permitida(ByVal m As String) As Boolean
    Permitida = InStr(1, ";" & MatrizesPermitidas() & ";", ";" & m & ";", vbBinaryCompare) > 0
End Function

' A 1a regra MATRIZ do lote (a que vale nos motores), ou "".
Public Function MatrizRegraDoLote(ByVal lote As String) As String
    Dim lo As ListObject, i As Long, v As Variant, cL As Long, cR As Long, cM As Long, k As String
    Set lo = mIntegracao.AcharTabela(TB)
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function
    k = Chave(lote)
    If Len(k) = 0 Then Exit Function
    cL = lo.ListColumns("LOTE").Index: cR = lo.ListColumns("REGRA").Index: cM = lo.ListColumns("MATRIZ").Index
    v = lo.DataBodyRange.Value2
    For i = 1 To UBound(v, 1)
        If Chave(v(i, cL)) = k And Chave(v(i, cR)) = "MATRIZ" And Len(Chave(v(i, cM))) > 0 Then
            MatrizRegraDoLote = Chave(v(i, cM)): Exit Function
        End If
    Next i
End Function

' "<linhas do lote>|<mudam de MATRIZ>|<mudam e hoje participam>" na tblCQ_Final (vazio = so as que tem outra matriz
' que nao a padrao nao entram na conta: o resultado volta a matriz que a origem mandou, que a tabela ja nao guarda).
Public Function PreviaMaterial(ByVal lote As String, ByVal matriz As String) As String
    Dim lo As ListObject, vL As Variant, vM As Variant, vP As Variant, i As Long, n As Long, mu As Long, mp As Long, k As String
    On Error GoTo fim
    Set lo = mIntegracao.AcharTabela("tblCQ_Final")
    k = Chave(lote): matriz = Chave(matriz)
    If Not lo Is Nothing Then
        If lo.ListRows.Count > 0 Then
            vL = lo.ListColumns("LOTE").DataBodyRange.Value2
            vM = lo.ListColumns("MATRIZ").DataBodyRange.Value2
            vP = lo.ListColumns("PARTICIPA_ESTATISTICA").DataBodyRange.Value2
            For i = 1 To UBound(vL, 1)
                If Chave(vL(i, 1)) = k Then
                    n = n + 1
                    If Len(matriz) > 0 And Chave(vM(i, 1)) <> matriz Then
                        mu = mu + 1
                        If Chave(vP(i, 1)) = "SIM" Then mp = mp + 1
                    End If
                End If
            Next i
        End If
    End If
fim:
    PreviaMaterial = n & "|" & mu & "|" & mp
End Function

' Grava a regra: m = "" remove as regras MATRIZ do lote; senao troca a 1a (ou cria uma). Devolve o valor anterior.
Private Function GravarRegra(ByVal lote As String, ByVal m As String) As String
    Dim lo As ListObject, ws As Worksheet, prot As Boolean, i As Long, cL As Long, cR As Long, cM As Long, cN As Long
    Dim k As String, achou As Boolean, lr As ListRow
    GravarRegra = MatrizRegraDoLote(lote)
    Set lo = mIntegracao.AcharTabela(TB)
    If lo Is Nothing Then Err.Raise vbObjectError + 991, , "tblLotesMaterial nao encontrada (ADR-091)"
    Set ws = lo.Parent
    prot = mSeguranca.LiberarEscrita(ws)
    On Error GoTo restaura
    cL = lo.ListColumns("LOTE").Index: cR = lo.ListColumns("REGRA").Index
    cM = lo.ListColumns("MATRIZ").Index: cN = lo.ListColumns("NOTA").Index
    k = Chave(lote)
    If Len(m) = 0 Then
        For i = lo.ListRows.Count To 1 Step -1
            If Chave(lo.ListRows(i).Range.Cells(1, cL).Value) = k And _
               Chave(lo.ListRows(i).Range.Cells(1, cR).Value) = "MATRIZ" Then lo.ListRows(i).Delete
        Next i
    Else
        ' a 1a regra MATRIZ e a que vale: so ela troca (uma repetida continua A14, como antes)
        For i = 1 To lo.ListRows.Count
            With lo.ListRows(i).Range
                If Chave(.Cells(1, cL).Value) = k And Chave(.Cells(1, cR).Value) = "MATRIZ" Then
                    .Cells(1, cM).Value = m
                    achou = True
                    Exit For
                End If
            End With
        Next i
        If Not achou Then
            Set lr = lo.ListRows.Add
            lr.Range.NumberFormat = "@"
            lr.Range.Cells(1, cL).Value = Trim$(lote)
            lr.Range.Cells(1, cR).Value = "MATRIZ"
            lr.Range.Cells(1, cM).Value = m
            lr.Range.Cells(1, cN).Value = "ADR-099: informado na Configuracao por " & mAuditoria.UsuarioSistema() & _
                                         " em " & Format$(Date, "dd/mm/yyyy")
        End If
    End If
restaura:
    Dim nE As Long, sE As String
    nE = Err.Number: sE = Err.Description
    mSeguranca.RestaurarProtecao ws, prot
    If nE <> 0 Then Err.Raise nE, , sE
End Function

' Worksheet_Change da Configuracao, F26:F125.
Public Sub MaterialEditado(ByVal Target As Range)
    Dim ws As Worksheet, alvo As Range, c As Range, lote As String, novo As String, antes As String, prev As Variant
    Dim prot As Boolean, ev As Boolean, msg As String, muda As Boolean, bruto As String
    Set ws = ThisWorkbook.Sheets(ABA_CAD)
    Set alvo = Intersect(Target, ws.Range(ws.Cells(R0, C_MAT), ws.Cells(R1, C_MAT)))
    If alvo Is Nothing Then Exit Sub
    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo fim
    prot = mSeguranca.LiberarEscrita(ws)
    For Each c In alvo.Cells
        lote = Trim$(CStr(ws.Cells(c.Row, C_LOTE).Value))
        bruto = CStr(c.Value)
        novo = Chave(c.Value)
        antes = MatrizRegraDoLote(lote)
        If novo = antes Then
            c.Value = IIf(Len(antes) > 0, antes, Empty)            ' normaliza a grafia (ex.: "urina" -> "URINA")
        ElseIf Len(lote) = 0 Then
            c.Value = Empty
            Avisar "Informe o material só numa linha com o código do lote.", vbExclamation
        ElseIf Len(novo) > 0 And Not Permitida(novo) Then
            c.Value = IIf(Len(antes) > 0, antes, Empty)
            Avisar "Material '" & bruto & "' não reconhecido para o lote " & lote & "." & vbCrLf & _
                   "Use: " & Replace(MatrizesPermitidas(), ";", ", ") & " (ou deixe vazio = padrão do equipamento)." & _
                   vbCrLf & "Outro material precisa de linhas no de/para (aba Cfg_Integracao).", vbExclamation
        Else
            prev = Split(PreviaMaterial(lote, novo), "|")
            muda = True
            If Application.Interactive Then
                msg = "Lote " & lote & ": material " & IIf(Len(antes) > 0, antes, "(padrão)") & " -> " & _
                      IIf(Len(novo) > 0, novo, "(padrão)") & "." & vbCrLf & vbCrLf & _
                      "O material do lote vale para TODOS os resultados do lote (do interfaceamento e manuais), " & _
                      "inclusive os já recebidos: tudo é reprocessado no próximo ATUALIZAR." & vbCrLf & _
                      "Resultados do lote: " & prev(0) & IIf(Len(novo) > 0, " · mudam de material: " & prev(1) & _
                      " (" & prev(2) & " hoje na estatística)", "") & vbCrLf & vbCrLf & "Confirmar?"
                muda = (MsgBox(msg, vbYesNo + vbQuestion, "Material do lote") = vbYes)
            End If
            If Not muda Then
                c.Value = IIf(Len(antes) > 0, antes, Empty)
            Else
                GravarRegra lote, novo
                c.Value = IIf(Len(novo) > 0, novo, Empty)
                mAuditoria.Auditar mAuditoria.CAT_CONFIG, "LOTE_MATERIAL_ALTERADO", "mMaterialLote", 0, Empty, "", lote, 0, "", _
                    IIf(Len(antes) > 0, antes, "(padrao)"), IIf(Len(novo) > 0, novo, "(padrao)"), "", "", "", _
                    "Material do lote " & lote & " informado na Configuracao (regra MATRIZ da tblLotesMaterial); " & _
                    prev(0) & " resultado(s) do lote, " & prev(1) & " mudam de material no proximo ATUALIZAR"
                Avisar "Material do lote " & lote & " gravado. Ele vale a partir do próximo ATUALIZAR DADOS.", vbInformation
            End If
        End If
    Next c
fim:
    If Err.Number <> 0 Then Avisar "Não foi possível gravar o material: " & Err.Description, vbExclamation
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
End Sub

' Espelho: F de cada lote do cadastro = a 1a regra MATRIZ da tabela (ou vazio). Audita cada diferenca.
' "OK|<n alterados>|<lotes alterados>"
Public Function SincronizarMaterialCadastro() As String
    Dim ws As Worksheet, r As Long, lote As String, m As String, atual As String, n As Long, lst As String
    Dim prot As Boolean, ev As Boolean
    On Error GoTo fim
    Set ws = ThisWorkbook.Sheets(ABA_CAD)
    If CStr(ws.Cells(R0 - 1, C_MAT).Value) <> "Material" Then SincronizarMaterialCadastro = "OK|0|": Exit Function
    ev = Application.EnableEvents
    Application.EnableEvents = False
    prot = mSeguranca.LiberarEscrita(ws)
    For r = R0 To R1
        lote = Trim$(CStr(ws.Cells(r, C_LOTE).Value))
        If Len(lote) > 0 Then m = MatrizRegraDoLote(lote) Else m = ""
        atual = Chave(ws.Cells(r, C_MAT).Value)
        If atual <> m Then
            ws.Cells(r, C_MAT).Value = IIf(Len(m) > 0, m, Empty)
            n = n + 1
            lst = lst & IIf(Len(lst) > 0, ",", "") & IIf(Len(lote) > 0, lote, "linha " & r)
            If Len(lote) > 0 Then
                mAuditoria.Auditar mAuditoria.CAT_CONFIG, "LOTE_MATERIAL_SINCRONIZADO", "mMaterialLote", 0, Empty, "", _
                    lote, 0, "", IIf(Len(atual) > 0, atual, "(vazio)"), IIf(Len(m) > 0, m, "(padrao)"), "", "", "", _
                    "Coluna Material do cadastro refeita pela regra MATRIZ da tblLotesMaterial"
            End If
        End If
    Next r
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
    SincronizarMaterialCadastro = "OK|" & n & "|" & lst
    Exit Function
fim:
    SincronizarMaterialCadastro = "ERRO|" & Err.Description
    On Error Resume Next
    mSeguranca.RestaurarProtecao ws, prot
    Application.EnableEvents = ev
End Function

' Instalacao (ADR-099): so espelha. Nenhuma regra e criada. "OK|<n espelhados>|<lotes do cadastro sem regra MATRIZ>"
Public Function CargaInicial() As String
    Dim s As String, ws As Worksheet, r As Long, lote As String, sem As String
    s = SincronizarMaterialCadastro()
    If Left$(s, 3) <> "OK|" Then CargaInicial = s: Exit Function
    Set ws = ThisWorkbook.Sheets(ABA_CAD)
    For r = R0 To R1
        lote = Trim$(CStr(ws.Cells(r, C_LOTE).Value))
        If Len(lote) > 0 Then
            If Len(MatrizRegraDoLote(lote)) = 0 Then sem = sem & IIf(Len(sem) > 0, ",", "") & lote
        End If
    Next r
    CargaInicial = "OK|" & Split(s, "|")(1) & "|" & sem
End Function
