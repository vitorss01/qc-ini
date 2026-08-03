Attribute VB_Name = "mSeguranca"
Option Explicit
' CAMADA: Seguranca / Auditoria de acesso — hash SHA-256, login, papeis,
' assinatura eletronica, protecao de planilhas e Modo Desenvolvedor.
Public Const SHEET_LOGIN As String = "Login"
Public Const SHEET_START As String = "Início"
Private Const DEV_HASH As String = "edbed0fcb2dac08503c515de40f4c771a137c4905cbe01ff60da11546f6e9c4f"
Private m_K(0 To 63) As Long
Private m_ready As Boolean


Private Sub InitK()
    Dim v As Variant, i As Long
    v = Array(&H428A2F98, &H71374491, &HB5C0FBCF, &HE9B5DBA5, &H3956C25B, &H59F111F1, &H923F82A4, &HAB1C5ED5, _
      &HD807AA98, &H12835B01, &H243185BE, &H550C7DC3, &H72BE5D74, &H80DEB1FE, &H9BDC06A7, &HC19BF174, _
      &HE49B69C1, &HEFBE4786, &HFC19DC6, &H240CA1CC, &H2DE92C6F, &H4A7484AA, &H5CB0A9DC, &H76F988DA, _
      &H983E5152, &HA831C66D, &HB00327C8, &HBF597FC7, &HC6E00BF3, &HD5A79147, &H6CA6351, &H14292967, _
      &H27B70A85, &H2E1B2138, &H4D2C6DFC, &H53380D13, &H650A7354, &H766A0ABB, &H81C2C92E, &H92722C85, _
      &HA2BFE8A1, &HA81A664B, &HC24B8B70, &HC76C51A3, &HD192E819, &HD6990624, &HF40E3585, &H106AA070, _
      &H19A4C116, &H1E376C08, &H2748774C, &H34B0BCB5, &H391C0CB3, &H4ED8AA4A, &H5B9CCA4F, &H682E6FF3, _
      &H748F82EE, &H78A5636F, &H84C87814, &H8CC70208, &H90BEFFFA, &HA4506CEB, &HBEF9A3F7, &HC67178F2)
    For i = 0 To 63: m_K(i) = v(i): Next
    m_ready = True
End Sub

Private Function ToU(ByVal L As Long) As Double
    If L < 0 Then ToU = L + 4294967296# Else ToU = L
End Function

Private Function ToS(ByVal d As Double) As Long
    d = d - Int(d / 4294967296#) * 4294967296#
    If d >= 2147483648# Then ToS = CLng(d - 4294967296#) Else ToS = CLng(d)
End Function

Private Function AddM(ParamArray vals() As Variant) As Long
    Dim s As Double, i As Long
    For i = LBound(vals) To UBound(vals): s = s + ToU(CLng(vals(i))): Next
    AddM = ToS(s)
End Function

Private Function ShR(ByVal X As Long, ByVal n As Integer) As Long
    ShR = ToS(Int(ToU(X) / (2 ^ n)))
End Function

Private Function ShL(ByVal X As Long, ByVal n As Integer) As Long
    Dim u As Double: u = ToU(X)
    u = u - Int(u / (2 ^ (32 - n))) * (2 ^ (32 - n))
    ShL = ToS(u * (2 ^ n))
End Function

Private Function RotR(ByVal X As Long, ByVal n As Integer) As Long
    RotR = ShR(X, n) Or ShL(X, 32 - n)
End Function

Public Function SHA256Hex(ByVal message As String) As String
    Dim h(0 To 7) As Long, i As Long, j As Long
    If Not m_ready Then InitK
    h(0) = &H6A09E667: h(1) = &HBB67AE85: h(2) = &H3C6EF372: h(3) = &HA54FF53A
    h(4) = &H510E527F: h(5) = &H9B05688C: h(6) = &H1F83D9AB: h(7) = &H5BE0CD19
    Dim mLen As Long: mLen = Len(message)
    Dim total As Long: total = mLen + 1
    Do While (total Mod 64) <> 56: total = total + 1: Loop
    total = total + 8
    Dim by() As Byte: ReDim by(0 To total - 1)
    For i = 1 To mLen: by(i - 1) = Asc(Mid$(message, i, 1)) And 255: Next
    by(mLen) = &H80
    Dim bl As Double: bl = mLen * 8#
    For i = 0 To 7
        by(total - 1 - i) = CByte(bl - Int(bl / 256#) * 256#): bl = Int(bl / 256#)
    Next
    Dim wArr(0 To 63) As Long, a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long, hh As Long
    Dim s0 As Long, s1 As Long, ch As Long, mj As Long, t1 As Long, t2 As Long, blk As Long
    For blk = 0 To (total \ 64) - 1
        For i = 0 To 15
            j = blk * 64 + i * 4
            wArr(i) = ToS(by(j) * 16777216# + by(j + 1) * 65536# + by(j + 2) * 256# + by(j + 3))
        Next
        For i = 16 To 63
            s0 = RotR(wArr(i - 15), 7) Xor RotR(wArr(i - 15), 18) Xor ShR(wArr(i - 15), 3)
            s1 = RotR(wArr(i - 2), 17) Xor RotR(wArr(i - 2), 19) Xor ShR(wArr(i - 2), 10)
            wArr(i) = AddM(wArr(i - 16), s0, wArr(i - 7), s1)
        Next
        a = h(0): b = h(1): c = h(2): d = h(3): e = h(4): f = h(5): g = h(6): hh = h(7)
        For i = 0 To 63
            s1 = RotR(e, 6) Xor RotR(e, 11) Xor RotR(e, 25)
            ch = (e And f) Xor ((Not e) And g)
            t1 = AddM(hh, s1, ch, m_K(i), wArr(i))
            s0 = RotR(a, 2) Xor RotR(a, 13) Xor RotR(a, 22)
            mj = (a And b) Xor (a And c) Xor (b And c)
            t2 = AddM(s0, mj)
            hh = g: g = f: f = e: e = AddM(d, t1): d = c: c = b: b = a: a = AddM(t1, t2)
        Next
        h(0) = AddM(h(0), a): h(1) = AddM(h(1), b): h(2) = AddM(h(2), c): h(3) = AddM(h(3), d)
        h(4) = AddM(h(4), e): h(5) = AddM(h(5), f): h(6) = AddM(h(6), g): h(7) = AddM(h(7), hh)
    Next
    Dim res As String
    For i = 0 To 7: res = res & Right$("00000000" & LCase$(Hex$(h(i))), 8): Next
    SHA256Hex = res
End Function

Private Function FindUserRow(ByVal login As String) As Long
    Dim ws As Worksheet, i As Long
    Set ws = ThisWorkbook.Sheets("Usuarios")
    FindUserRow = 0
    For i = 4 To 53
        If Trim$(CStr(ws.Cells(i, 1).Value)) <> "" Then
            If UCase$(Trim$(CStr(ws.Cells(i, 1).Value))) = UCase$(Trim$(login)) Then
                FindUserRow = i: Exit Function
            End If
        End If
    Next i
End Function

Public Sub DoLogin()
    Dim u As String, p As String, r As Long
    On Error Resume Next
    u = Trim$(CStr(ThisWorkbook.Names("loginUser").RefersToRange.Value))
    p = CStr(ThisWorkbook.Names("loginPass").RefersToRange.Value)
    ThisWorkbook.Names("loginPass").RefersToRange.Value = ""
    r = FindUserRow(u)
    If r > 0 And p <> "" Then
        If LCase$(CStr(ThisWorkbook.Sheets("Usuarios").Cells(r, 4).Value)) = SHA256Hex(p) Then
            ThisWorkbook.Names("currentUser").RefersToRange.Value = ThisWorkbook.Sheets("Usuarios").Cells(r, 1).Value
            ThisWorkbook.Names("currentPapel").RefersToRange.Value = ThisWorkbook.Sheets("Usuarios").Cells(r, 3).Value
            ThisWorkbook.Names("loginMsg").RefersToRange.Value = ""
            UnlockApp
            Exit Sub
        End If
    End If
    ThisWorkbook.Names("loginMsg").RefersToRange.Value = "Usuario ou senha invalidos."
End Sub

Public Sub CadastrarUsuario()
    Dim ws As Worksheet, lg$, nm$, sn$, pp$, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Usuarios")
    Dim ppAtual$
    ppAtual = UCase$(Trim$(CStr(ThisWorkbook.Names("currentPapel").RefersToRange.Value)))
    If ppAtual <> "ANALISTA" And ppAtual <> "ADM" Then
        ThisWorkbook.Names("cadMsg").RefersToRange.Value = "Apenas ANALISTA/ADM pode cadastrar usuarios."
        Exit Sub
    End If
    lg = Trim$(CStr(ThisWorkbook.Names("cadLogin").RefersToRange.Value))
    nm = Trim$(CStr(ThisWorkbook.Names("cadNome").RefersToRange.Value))
    sn = CStr(ThisWorkbook.Names("cadSenha").RefersToRange.Value)
    pp = Trim$(CStr(ThisWorkbook.Names("cadPapel").RefersToRange.Value))
    If lg = "" Or sn = "" Then ThisWorkbook.Names("cadMsg").RefersToRange.Value = "Preencha login e senha.": Exit Sub
    If pp = "" Then pp = "TÉCNICO"
    r = FindUserRow(lg)
    If r = 0 Then
        For r = 4 To 53
            If Trim$(CStr(ws.Cells(r, 1).Value)) = "" Then Exit For
        Next r
    End If
    ws.Cells(r, 1).Value = lg: ws.Cells(r, 2).Value = nm: ws.Cells(r, 3).Value = pp
    ws.Cells(r, 4).Value = SHA256Hex(sn)
    ThisWorkbook.Names("cadSenha").RefersToRange.Value = ""
    ThisWorkbook.Names("cadMsg").RefersToRange.Value = "Usuario '" & lg & "' salvo (" & pp & ")."
End Sub

Public Function AssinarCom(ByVal tcell As Range, ByVal login As String, ByVal senha As String, ByVal tipo As String) As String
    Dim r As Long, papel$, ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Usuarios")
    r = FindUserRow(login)
    If r = 0 Then AssinarCom = "Usuario nao encontrado.": Exit Function
    If LCase$(CStr(ws.Cells(r, 4).Value)) <> SHA256Hex(senha) Or senha = "" Then AssinarCom = "Senha invalida.": Exit Function
    papel = UCase$(Trim$(CStr(ws.Cells(r, 3).Value)))
    If UCase$(tipo) = "FINALIZADO" And papel <> "ANALISTA" And papel <> "ADM" Then
        AssinarCom = "Apenas ANALISTA/ADM pode finalizar/liberar o equipamento."
        Exit Function
    End If
    CarimbarRubrica ws, r, tcell
    tcell.Offset(0, 1).Value = ws.Cells(r, 1).Value & " · " & Format(Now, "dd/mm/yyyy hh:mm")
    On Error Resume Next
    tcell.Offset(0, 1).EntireColumn.AutoFit
    On Error GoTo 0
    AssinarCom = "OK"
End Function

Private Sub CarimbarRubrica(ByVal ws As Worksheet, ByVal userRow As Long, ByVal tcell As Range)
    ' rubrica = conteudo da coluna E do usuario (PROCV login->E): texto OU Imagem na Celula (365)
    On Error Resume Next
    Application.CutCopyMode = False
    tcell.Parent.Activate
    ws.Cells(userRow, 5).Copy
    tcell.PasteSpecial -4104   ' xlPasteAll
    Application.CutCopyMode = False
    tcell.Locked = True
End Sub

Public Sub Assinar(ByVal tcell As Range, ByVal tipo As String)
    Dim u$, p$, res$
    On Error Resume Next
    frmAssinar.lblTipo.Caption = "Assinatura eletronica - " & tipo
    frmAssinar.txtLogin.Value = ""
    frmAssinar.txtSenha.Value = ""
    frmAssinar.Show
    If Not frmAssinar.Confirmado Then Unload frmAssinar: Exit Sub
    u = Trim$(frmAssinar.txtLogin.Value): p = frmAssinar.txtSenha.Value
    Unload frmAssinar
    res = AssinarCom(tcell, u, p, tipo)
    If res <> "OK" Then MsgBox res, vbExclamation, "Assinatura"
End Sub

Public Sub LockApp()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    On Error Resume Next
    Sheets(SHEET_LOGIN).Visible = xlSheetVisible
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> SHEET_LOGIN Then ws.Visible = xlSheetVeryHidden
    Next ws
    Sheets(SHEET_LOGIN).Activate
    ReprotectAll
    SystemLook True
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

Public Sub UnlockApp()
    Dim ws As Worksheet, isAdm As Boolean
    Application.ScreenUpdating = False
    On Error Resume Next
    isAdm = (UCase$(Trim$(CStr(ThisWorkbook.Names("currentPapel").RefersToRange.Value))) = "ADM")
    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Name
            Case SHEET_LOGIN
                ws.Visible = xlSheetVeryHidden
            Case "Calc", "LotesStore", "LiberStore", "RegistrosStore"
                ws.Visible = IIf(isAdm, xlSheetVisible, xlSheetVeryHidden)
            Case Else
                ws.Visible = xlSheetVisible
        End Select
    Next ws
    If isAdm Then
        UnprotectAll
    Else
        ReprotectAll
    End If
    HookCharts
    Sheets(SHEET_START).Activate
    SystemLook True
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

Public Sub UnprotectAll()
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:="qcini2025"
    Next ws
End Sub

Public Sub ReprotectAll()
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect Password:="qcini2025", UserInterfaceOnly:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True
    Next ws
End Sub

Public Sub Logout()
    LockApp
End Sub


' ===================== MODO DESENVOLVEDOR (Etapa 6) =====================
' Acionado por um Shape 100% transparente na aba Painel (sem atalho de teclado).
Public Sub ModoDesenvolvedor()
    Dim s As String, ws As Worksheet
    frmDev.Confirmado = False
    frmDev.txtSenha.Value = ""
    frmDev.Show
    If Not frmDev.Confirmado Then
        Unload frmDev
        Exit Sub
    End If
    s = frmDev.txtSenha.Value
    Unload frmDev
    If LCase$(SHA256Hex(s)) <> LCase$(DEV_HASH) Then
        MsgBox "Senha incorreta.", vbExclamation, "Modo Desenvolvedor"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:="qcini2025"
        ws.Visible = xlSheetVisible
    Next ws
    On Error GoTo 0
    SystemLook False
    Application.ScreenUpdating = True
    MsgBox "Modo Desenvolvedor ATIVO." & vbCrLf & vbCrLf & _
           "Todas as abas visiveis e desprotegidas." & vbCrLf & _
           "Use Sair (logout) para voltar ao modo normal.", vbInformation, "Modo Desenvolvedor"
End Sub


