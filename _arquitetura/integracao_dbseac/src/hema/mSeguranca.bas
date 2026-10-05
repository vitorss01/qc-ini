Attribute VB_Name = "mSeguranca"
Option Explicit
' CAMADA: Seguranca / Auditoria de acesso — hash SHA-256, login, papeis,
' assinatura eletronica, protecao de planilhas e Modo Desenvolvedor.
Public Const SHEET_LOGIN As String = "Login"
Public Const SHEET_START As String = "Início"
Private Const DEV_HASH As String = "edbed0fcb2dac08503c515de40f4c771a137c4905cbe01ff60da11546f6e9c4f"
Private m_K(0 To 63) As Long
Private m_ready As Boolean
' Papeis da lista fechada (ADR-059). Declaracao de modulo so pode ficar aqui,
' antes do primeiro procedimento -- no meio do modulo e erro de COMPILACAO.
Private Const PAPEL_ADM As String = "ADM"
Private Const PAPEL_ANALISTA As String = "ANALISTA"
Private Const PAPEL_TECNICO As String = "TÉCNICO"
' Ancora da sessao EM MEMORIA (03/10/2026): gravada so pelo DoLogin depois da
' senha conferida. As celulas currentUser/currentPapel podem ser forcadas por
' quem desproteger a aba (a senha esta no VBA); esta variavel, nao. Se o VBA for
' reiniciado ela some -- e o cadastro pede login de novo, que e o certo.
Private mSessaoLogin As String
' Estado gravado no disco (auditoria 04/10/2026): o BeforeSave grava TRANCADO e o AfterSave devolve
' a vista de quem estava trabalhando -- sessao comum, ADM ou Modo Desenvolvedor.
Private mModoDev As Boolean
Private mTrancadoParaGravar As Boolean
Private mAbaAntesGravar As String
Private mUsuarioAntesGravar As String
Private mPapelAntesGravar As String


' ============== ESCRITA EM ABA PROTEGIDA (ADR-046) ==============
'
' ReprotectAll aplica UserInterfaceOnly:=True, que e a configuracao certa, mas
' o Excel NAO PERSISTE essa flag ao salvar. Reaberto o arquivo, a aba volta
' protegida tambem para o VBA ate Workbook_Open rodar LockApp -- e Workbook_Open
' nao roda em automacao, porque todo script abre com EnableEvents = False.
'
' Por isso a garantia nao pode ser "alguem rodou LockApp antes": cada rotina que
' escreve trata a propria janela.
'
' SENHA EM LITERAL, e nao a constante SENHA_PROT do trecho de referencia: essa
' constante nao existe neste modulo, e uma referencia a nome inexistente e erro
' de COMPILACAO -- que derruba o projeto inteiro, nao so a rotina.
Public Function LiberarEscrita(ByVal ws As Worksheet) As Boolean
    LiberarEscrita = ws.ProtectContents
    If LiberarEscrita Then ws.Unprotect Password:="qcini2025"
End Function


' ADR-062: no caminho do CLIQUE (troca de analito), Protect/Unprotect com senha custa ~70-80 ms
' CADA (hash da senha com 100 mil voltas; medido em 03/10/2026) e o motor fazia 2 ciclos por clique.
' Com UserInterfaceOnly em vigor NESTA sessao (ProtectionMode, aplicado no login pelo ReprotectAll),
' o VBA ja escreve sem desproteger: devolve False e o RestaurarProtecao do chamador nao faz nada.
' Sem UserInterfaceOnly (automacao, arquivo recem-aberto) cai no LiberarEscrita normal (ADR-046).
Public Function LiberarEscritaRapida(ByVal ws As Worksheet) As Boolean
    If ws.ProtectContents And ws.ProtectionMode Then Exit Function
    LiberarEscritaRapida = LiberarEscrita(ws)
End Function


' Nao propaga erro: e chamada no caminho de limpeza, onde mascarar a excecao
' original seria pior do que falhar em reproteger.
Public Sub RestaurarProtecao(ByVal ws As Worksheet, ByVal estava As Boolean)
    If Not estava Then Exit Sub
    On Error Resume Next
    ws.Protect Password:="qcini2025", UserInterfaceOnly:=True, _
               DrawingObjects:=False, Contents:=True, Scenarios:=True
End Sub



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

Private Function ShR(ByVal x As Long, ByVal n As Integer) As Long
    ShR = ToS(Int(ToU(x) / (2 ^ n)))
End Function

Private Function ShL(ByVal x As Long, ByVal n As Integer) As Long
    Dim u As Double: u = ToU(x)
    u = u - Int(u / (2 ^ (32 - n))) * (2 ^ (32 - n))
    ShL = ToS(u * (2 ^ n))
End Function

Private Function RotR(ByVal x As Long, ByVal n As Integer) As Long
    RotR = ShR(x, n) Or ShL(x, 32 - n)
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

' ===================== PAPEIS E SESSAO (ADR-059) =====================
'
' FASE3A, achado CRITICO "escalacao de privilegio e sequestro de identidade":
' o papel da sessao era lido da celula currentPapel, e CadastrarUsuario deixava
' um ANALISTA criar um ADM ou regravar o hash de QUALQUER usuario existente --
' inclusive o do ADM, e com ele assinar em nome de outro. Nada era registrado.
'
' Regras agora:
'   1. o papel de quem age vem da TABELA de usuarios (pelo login da sessao),
'      nunca da celula -- a celula so mostra;
'   2. ninguem concede papel acima do proprio;
'   3. SOMENTE O ADM CADASTRA USUARIOS E DEFINE OU REDEFINE SENHAS (decisao do
'      usuario, 03/10/2026), inclusive a propria. Criar usuario tambem e definir
'      senha: quem define a senha inicial a conhece e entraria como o novo usuario
'      -- o mesmo vetor do achado critico. TECNICO e ANALISTA nao gravam hash;
'   4. o ultimo ADM nao pode ser rebaixado (o sistema ficaria sem administrador);
'   5. a sessao e LIMPA no logout e na abertura: quem entra nunca herda a
'      identidade de quem usou o arquivo antes;
'   6. login, falha de login, logout, cadastro, recusa, assinatura e Modo
'      Desenvolvedor vao para o Audit_Log (categoria SEGURANCA). Senha e hash
'      NUNCA entram no log.

' Papel da lista fechada, ou "" se o texto nao for um papel. Aceita a grafia sem
' acento: UCase$("Técnico") = "TÉCNICO" (mesma armadilha do item 7.4 do gate).
Private Function NormalizarPapel(ByVal p As String) As String
    Select Case UCase$(Trim$(p))
        Case "ADM": NormalizarPapel = PAPEL_ADM
        Case "ANALISTA": NormalizarPapel = PAPEL_ANALISTA
        Case "TÉCNICO", "TECNICO": NormalizarPapel = PAPEL_TECNICO
    End Select
End Function

Private Function NivelPapel(ByVal p As String) As Long
    Select Case p
        Case PAPEL_TECNICO: NivelPapel = 1
        Case PAPEL_ANALISTA: NivelPapel = 2
        Case PAPEL_ADM: NivelPapel = 3
    End Select
End Function

' Papel do usuario segundo a tabela (aba Usuarios, coluna C).
Public Function PapelDoUsuario(ByVal login As String) As String
    Dim r As Long
    r = FindUserRow(login)
    If r > 0 Then PapelDoUsuario = NormalizarPapel(CStr(ThisWorkbook.Sheets("Usuarios").Cells(r, 3).Value))
End Function

Private Function ContarAdm() As Long
    Dim ws As Worksheet, i As Long
    Set ws = ThisWorkbook.Sheets("Usuarios")
    For i = 4 To 53
        If Trim$(CStr(ws.Cells(i, 1).Value)) <> "" Then
            If NormalizarPapel(CStr(ws.Cells(i, 3).Value)) = PAPEL_ADM Then ContarAdm = ContarAdm + 1
        End If
    Next i
End Function

Private Function PrimeiraLinhaLivre() As Long
    Dim ws As Worksheet, i As Long
    Set ws = ThisWorkbook.Sheets("Usuarios")
    For i = 4 To 53
        If Trim$(CStr(ws.Cells(i, 1).Value)) = "" Then PrimeiraLinhaLivre = i: Exit Function
    Next i
End Function

' Escreve num nome definido mesmo com a aba protegida (ADR-046): as celulas da
' sessao e das mensagens ficam TRAVADAS, e UserInterfaceOnly pode nao estar em
' vigor (nao persiste ao salvar). Quem precisa de garantia confere lendo de volta.
Private Sub EscreverNome(ByVal nome As String, ByVal valor As Variant)
    Dim rg As Range, prot As Boolean
    On Error Resume Next
    Set rg = ThisWorkbook.Names(nome).RefersToRange
    If rg Is Nothing Then Exit Sub
    prot = LiberarEscrita(rg.Worksheet)
    rg.Value = valor
    RestaurarProtecao rg.Worksheet, prot
End Sub

Private Sub EncerrarSessao()
    mSessaoLogin = ""
    EscreverNome "currentUser", ""
    EscreverNome "currentPapel", ""
End Sub

' Evento de seguranca no Audit_Log. Mesmo criterio do mIntegracao: a trilha
' nunca impede o login nem o cadastro -- falhar em registrar nao bloqueia o
' laboratorio. stAnt/stNovo carregam o papel antes/depois.
Private Sub AuditarSeg(ByVal acao As String, ByVal papelAnt As String, ByVal papelNovo As String, ByVal detalhe As String)
    On Error Resume Next
    mAuditoria.Auditar mAuditoria.CAT_SEG, acao, "mSeguranca", 0, Empty, "", "", 0, "", Empty, Empty, _
                       papelAnt, papelNovo, "", detalhe
End Sub

Public Sub DoLogin()
    Dim u As String, p As String, r As Long, ok As Boolean, papel As String, login As String
    On Error Resume Next
    u = Trim$(CStr(ThisWorkbook.Names("loginUser").RefersToRange.Value))
    p = CStr(ThisWorkbook.Names("loginPass").RefersToRange.Value)
    On Error GoTo 0
    EscreverNome "loginPass", ""
    EncerrarSessao                       ' nada da sessao anterior sobrevive a uma tentativa
    r = FindUserRow(u)
    If r > 0 And p <> "" Then ok = (LCase$(CStr(ThisWorkbook.Sheets("Usuarios").Cells(r, 4).Value)) = SHA256Hex(p))
    p = ""
    If Not ok Then
        EscreverNome "loginMsg", "Usuario ou senha invalidos."
        AuditarSeg "LOGIN_FALHOU", "", "", "login informado: " & Left$(u, 60)
        Exit Sub
    End If
    login = CStr(ThisWorkbook.Sheets("Usuarios").Cells(r, 1).Value)
    papel = NormalizarPapel(CStr(ThisWorkbook.Sheets("Usuarios").Cells(r, 3).Value))
    If Len(papel) = 0 Then
        EscreverNome "loginMsg", "Usuario sem funcao valida (ADM, ANALISTA ou TÉCNICO). Procure o administrador."
        AuditarSeg "LOGIN_FALHOU", "", "", "login " & login & " com funcao invalida na tabela"
        Exit Sub
    End If
    EscreverNome "currentUser", login
    EscreverNome "currentPapel", papel
    ' a sessao tem de ser A DA TABELA: se a escrita falhou em silencio, o login
    ' nao segue com o que estava na celula
    If UsuarioSistema() <> login Or PapelSistema() <> papel Then
        EncerrarSessao
        EscreverNome "loginMsg", "Nao foi possivel abrir a sessao. Feche e abra o arquivo novamente."
        AuditarSeg "LOGIN_FALHOU", "", papel, "login " & login & ": sessao nao gravada"
        Exit Sub
    End If
    mSessaoLogin = login
    EscreverNome "loginMsg", ""
    AuditarSeg "LOGIN", "", papel, ""
    UnlockApp
    ' ADR-060: toda abertura passa pelo login -- lote sem validade avisa aqui
    ' (so leitura; no Workbook_Open nao ha usuario nem a aba visivel para levar)
    On Error Resume Next
    mLotes.AvisarLotesSemValidade
    ' ADR-062: o 1o clique no Painel nao paga mais a leitura do banco e o indice
    Application.StatusBar = "Preparando o Painel..."
    mEstatistica.AquecerMotor
    Application.StatusBar = False
End Sub

Public Sub CadastrarUsuario()
    Dim ws As Worksheet, lg As String, nm As String, sN As String, pedido As String
    Dim ator As String, papelAtor As String, papelAnt As String, papelNovo As String
    Dim r As Long, novo As Boolean, recusa As String, acao As String
    Dim prot As Boolean, nE As Long, sE As String
    Set ws = ThisWorkbook.Sheets("Usuarios")
    On Error Resume Next
    lg = Trim$(CStr(ThisWorkbook.Names("cadLogin").RefersToRange.Value))
    nm = Trim$(CStr(ThisWorkbook.Names("cadNome").RefersToRange.Value))
    sN = CStr(ThisWorkbook.Names("cadSenha").RefersToRange.Value)
    pedido = Trim$(CStr(ThisWorkbook.Names("cadPapel").RefersToRange.Value))
    On Error GoTo 0
    EscreverNome "cadSenha", ""          ' a senha digitada nunca fica no arquivo, nem na recusa

    ator = UsuarioSistema()
    papelAtor = PapelDoUsuario(ator)     ' da TABELA, nunca da celula da sessao
    r = FindUserRow(lg)
    novo = (r = 0)
    If Not novo Then papelAnt = NormalizarPapel(CStr(ws.Cells(r, 3).Value))
    If Len(pedido) = 0 Then
        papelNovo = IIf(novo, PAPEL_TECNICO, papelAnt)
    Else
        papelNovo = NormalizarPapel(pedido)
    End If

    If papelAtor <> PAPEL_ADM Then
        recusa = "Somente o ADM cadastra usuarios e define ou redefine senhas (inclusive a propria). Procure o administrador."
    ElseIf Len(mSessaoLogin) = 0 Or UCase$(mSessaoLogin) <> UCase$(ator) Then
        recusa = "Sessao invalida: faca login novamente."
    ElseIf lg = "" Or sN = "" Then
        recusa = "Preencha login e senha."
    ElseIf Len(papelNovo) = 0 Then
        recusa = "Funcao invalida: use ADM, ANALISTA ou TÉCNICO."
    ElseIf NivelPapel(papelNovo) > NivelPapel(papelAtor) Then
        recusa = "Um " & papelAtor & " nao pode conceder a funcao " & papelNovo & "."
    ElseIf Not novo And papelAnt = PAPEL_ADM And papelNovo <> PAPEL_ADM And ContarAdm() <= 1 Then
        recusa = "O usuario '" & lg & "' e o unico ADM. Cadastre outro ADM antes de mudar esta funcao."
    ElseIf novo Then
        r = PrimeiraLinhaLivre()
        If r = 0 Then recusa = "A tabela de usuarios esta cheia (50 linhas)."
    End If
    If Len(recusa) > 0 Then
        sN = ""
        EscreverNome "cadMsg", recusa
        AuditarSeg "CADASTRO_RECUSADO", papelAnt, papelNovo, "alvo: " & Left$(lg, 60) & " | " & recusa
        Exit Sub
    End If

    On Error GoTo restaura
    prot = LiberarEscrita(ws)
    ws.Cells(r, 1).Value = lg
    If novo Or Len(nm) > 0 Then ws.Cells(r, 2).Value = nm
    ws.Cells(r, 3).Value = papelNovo
    ws.Cells(r, 4).Value = SHA256Hex(sN)
restaura:
    nE = Err.Number: sE = Err.Description
    RestaurarProtecao ws, prot
    On Error GoTo 0
    sN = ""
    If nE <> 0 Then
        EscreverNome "cadMsg", "Erro ao salvar o usuario: " & sE
        AuditarSeg "CADASTRO_RECUSADO", papelAnt, papelNovo, "alvo: " & Left$(lg, 60) & " | erro " & nE & ": " & sE
        Exit Sub
    End If
    If novo Then
        acao = "USUARIO_CRIADO"
    ElseIf UCase$(lg) = UCase$(ator) Then
        acao = "SENHA_PROPRIA_ALTERADA"
    Else
        acao = "USUARIO_ALTERADO"
    End If
    EscreverNome "cadMsg", "Usuario '" & lg & "' salvo (" & papelNovo & ")."
    AuditarSeg acao, papelAnt, papelNovo, "alvo: " & lg
End Sub

Public Function AssinarCom(ByVal tcell As Range, ByVal login As String, ByVal senha As String, ByVal tipo As String) As String
    Dim r As Long, papel$, ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Usuarios")
    r = FindUserRow(login)
    If r = 0 Then
        AssinarCom = "Usuario nao encontrado."
        AuditarSeg "ASSINATURA_RECUSADA", "", "", tipo & " | login informado: " & Left$(login, 60) & " | " & AssinarCom
        Exit Function
    End If
    If LCase$(CStr(ws.Cells(r, 4).Value)) <> SHA256Hex(senha) Or senha = "" Then
        AssinarCom = "Senha invalida."
        AuditarSeg "ASSINATURA_RECUSADA", "", "", tipo & " | login " & ws.Cells(r, 1).Value & " | " & AssinarCom
        Exit Function
    End If
    papel = NormalizarPapel(CStr(ws.Cells(r, 3).Value))
    If UCase$(tipo) = "FINALIZADO" And papel <> PAPEL_ANALISTA And papel <> PAPEL_ADM Then
        AssinarCom = "Apenas ANALISTA/ADM pode finalizar/liberar o equipamento."
        AuditarSeg "ASSINATURA_RECUSADA", "", papel, tipo & " | login " & ws.Cells(r, 1).Value & " | " & AssinarCom
        Exit Function
    End If
    CarimbarRubrica ws, r, tcell
    tcell.Offset(0, 1).Value = ws.Cells(r, 1).Value & " · " & Format(Now, "dd/mm/yyyy hh:mm")
    On Error Resume Next
    tcell.Offset(0, 1).EntireColumn.AutoFit
    On Error GoTo 0
    ' quem assinou pode nao ser quem esta logado (assinatura pede login e senha
    ' proprios): o signatario vai no detalhe, a sessao fica no usuario do evento
    AuditarSeg "ASSINATURA_" & UCase$(tipo), "", papel, "signatario: " & ws.Cells(r, 1).Value & _
               " | " & tcell.Worksheet.Name & "!" & tcell.Address(False, False)
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

' ESTRUTURA DA PASTA (ADR-052). A Hematologia chegou com a estrutura protegida,
' e com ela LockApp/UnlockApp NAO conseguiam mostrar nem ocultar abas -- o
' On Error Resume Next escondia a falha e o login deixava tudo a mostra. A
' estrutura agora e liberada so durante a troca de visibilidade e volta ao que
' o nome "protegerEstrutura" manda (VERDADEIRO na Hematologia).
Private Sub LiberarEstrutura()
    On Error Resume Next
    ThisWorkbook.Unprotect Password:="qcini2025"
End Sub

Private Sub RestaurarEstrutura()
    Dim v As Variant
    On Error Resume Next
    v = Application.Evaluate("protegerEstrutura")
    If VarType(v) = vbBoolean Then
        If v Then ThisWorkbook.Protect Password:="qcini2025", Structure:=True
    End If
End Sub

Public Sub LockApp()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    On Error Resume Next
    EncerrarSessao                  ' ADR-059: a tela de login nunca guarda a identidade de ninguem
    mModoDev = False
    LiberarEstrutura
    Sheets(SHEET_LOGIN).Visible = xlSheetVisible
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> SHEET_LOGIN Then ws.Visible = xlSheetVeryHidden
    Next ws
    Sheets(SHEET_LOGIN).Activate
    ReprotectAll
    RestaurarEstrutura
    SystemLook True
    mApp.Reset                      ' ADR-050: calculo/eventos/tela sempre limpos no login
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

Public Sub UnlockApp()
    Dim ws As Worksheet, isAdm As Boolean
    ' Rodar pela lista de macros (Alt+F8) sem login revelava as abas: sem sessao
    ' valida na tabela, volta para a tela de login.
    If Len(PapelDoUsuario(UsuarioSistema())) = 0 Then
        LockApp
        Exit Sub
    End If
    Application.ScreenUpdating = False
    On Error Resume Next
    isAdm = (PapelDoUsuario(UsuarioSistema()) = PAPEL_ADM)      ' ADR-059: da tabela, nao da celula
    LiberarEstrutura
    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Name
            Case SHEET_LOGIN
                ws.Visible = xlSheetVeryHidden
            Case "Calc", "LotesStore", "LiberStore", "RegistrosStore", "Cfg_Integracao"
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
    RestaurarEstrutura
    HookCharts
    Sheets(SHEET_START).Activate
    SystemLook True
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

Public Sub UnprotectAll()
    ' Pela lista de macros (Alt+F8) qualquer usuario desprotegia o sistema inteiro:
    ' so uma sessao ADM, conferida pela tabela e pela ancora em memoria.
    If PapelDoUsuario(UsuarioSistema()) <> PAPEL_ADM Then Exit Sub
    If UCase$(mSessaoLogin) <> UCase$(UsuarioSistema()) Then Exit Sub
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' a trilha (Audit_*) continua protegida tambem para o ADM: desprotegida, a sessao ADM
        ' editava o Audit_Log sem deixar rastro (auditoria 04/10/2026). Quem grava nela e o mAuditoria.
        If Not ws.Name Like "Audit_*" Then ws.Unprotect Password:="qcini2025"
    Next ws
End Sub

' Sessao aberta pelo DoLogin NESTA execucao do VBA: a ancora em memoria confere com a celula e o
' papel vem da tabela. Macro que altera dados e roda pela lista (Alt+F8) na tela de login confere
' isto antes de fazer qualquer coisa (auditoria 04/10/2026).
Public Function SessaoAtiva() As Boolean
    On Error Resume Next
    If Len(mSessaoLogin) = 0 Then Exit Function
    If UCase$(mSessaoLogin) <> UCase$(UsuarioSistema()) Then Exit Function
    SessaoAtiva = (Len(PapelDoUsuario(UsuarioSistema())) > 0)
End Function

' ===================== ESTADO GRAVADO NO DISCO (auditoria 04/10/2026) =====================
' Com macros desabilitadas (arquivo baixado, recebido por e-mail) o Workbook_Open nao roda: o que
' protege o arquivo e o estado em que ele foi GRAVADO. A entrega sai blindada (blindar_entrega.py),
' mas o primeiro Ctrl+S de uma sessao gravava as abas visiveis -- e, na sessao ADM ou no Modo
' Desenvolvedor, desprotegidas (FASE3A #2, recomendacao P1). Agora o BeforeSave grava TRANCADO
' (so o Login visivel, toda aba protegida, estrutura travada) sem encerrar a sessao, e o AfterSave
' devolve a vista. Se o "Salvar como" for cancelado o AfterSave nao vem: fica a tela de login.
Public Sub TrancarParaGravar()
    Dim ws As Worksheet
    On Error Resume Next
    mAbaAntesGravar = ""
    If ActiveWorkbook Is ThisWorkbook Then mAbaAntesGravar = ActiveSheet.Name
    mTrancadoParaGravar = True
    ' a identidade da sessao tambem nao vai para o disco (FASE3A: N1:N2 gravados com o ultimo login)
    mUsuarioAntesGravar = Trim$(CStr(ThisWorkbook.Names("currentUser").RefersToRange.Value))   ' a celula, sem o
    mPapelAntesGravar = Trim$(CStr(ThisWorkbook.Names("currentPapel").RefersToRange.Value))   ' "(sem login)"
    EscreverNome "currentUser", ""
    EscreverNome "currentPapel", ""
    Application.ScreenUpdating = False
    LiberarEstrutura
    ThisWorkbook.Sheets(SHEET_LOGIN).Visible = xlSheetVisible
    ThisWorkbook.Sheets(SHEET_LOGIN).Activate
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> SHEET_LOGIN Then ws.Visible = xlSheetVeryHidden
    Next ws
    ReprotectAll
    ThisWorkbook.Protect Password:="qcini2025", Structure:=True     ' no disco, sempre travada
End Sub

Public Sub RestaurarAposGravar()
    Dim ws As Worksheet
    If Not mTrancadoParaGravar Then Exit Sub
    mTrancadoParaGravar = False
    On Error Resume Next
    If Len(mUsuarioAntesGravar) > 0 Then
        EscreverNome "currentUser", mUsuarioAntesGravar
        EscreverNome "currentPapel", mPapelAntesGravar
    End If
    mUsuarioAntesGravar = ""
    mPapelAntesGravar = ""
    If mModoDev Then
        LiberarEstrutura
        For Each ws In ThisWorkbook.Worksheets
            ws.Unprotect Password:="qcini2025"
            ws.Visible = xlSheetVisible
        Next ws
    ElseIf SessaoAtiva() Then
        UnlockApp                    ' a mesma vista do login (o ADM ve as abas tecnicas)
    Else
        RestaurarEstrutura           ' sem sessao: fica a tela de login, como no repouso
        Exit Sub
    End If
    Application.ScreenUpdating = False
    If Len(mAbaAntesGravar) > 0 Then ThisWorkbook.Sheets(mAbaAntesGravar).Activate
    ThisWorkbook.Saved = True        ' o disco ja tem tudo; so a vista mudou
    Application.ScreenUpdating = True
End Sub

Public Sub ReprotectAll()
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Audit_*" Then
            ' a trilha precisa de filtro e ordenacao para o auditor (item 3.1); a
            ' protecao generica tirava as duas a cada login, ate o proximo evento
            mAuditoria.ProtegerAudit ws
        Else
            ws.Protect Password:="qcini2025", UserInterfaceOnly:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True
        End If
    Next ws
End Sub

Public Sub Logout()
    AuditarSeg "LOGOUT", PapelSistema(), "", ""
    LockApp
End Sub


' ===================== MODO DESENVOLVEDOR (Etapa 6) =====================
' Acionado por um Shape 100% transparente na aba Painel (sem atalho de teclado).
Public Sub ModoDesenvolvedor()
    Dim s As String
    frmDev.Confirmado = False
    frmDev.txtSenha.Value = ""
    frmDev.Show
    If Not frmDev.Confirmado Then
        Unload frmDev
        Exit Sub
    End If
    s = frmDev.txtSenha.Value
    Unload frmDev
    If Not AtivarModoDesenvolvedor(s) Then
        s = ""
        MsgBox "Senha incorreta.", vbExclamation, "Modo Desenvolvedor"
        Exit Sub
    End If
    s = ""
    MsgBox "Modo Desenvolvedor ATIVO." & vbCrLf & vbCrLf & _
           "Todas as abas visiveis e desprotegidas." & vbCrLf & _
           "Use Sair (logout) para voltar ao modo normal.", vbInformation, "Modo Desenvolvedor"
End Sub

' Nucleo do Modo Desenvolvedor, sem formulario (o QA chama este). Achado FASE3A:
' "desprotege todo o sistema e nao deixa nenhuma marca" -- agora deixa, tanto a
' ativacao quanto a tentativa com senha errada.
Public Function AtivarModoDesenvolvedor(ByVal senha As String) As Boolean
    Dim ws As Worksheet
    If LCase$(SHA256Hex(senha)) <> LCase$(DEV_HASH) Then
        AuditarSeg "MODO_DESENVOLVEDOR_RECUSADO", PapelSistema(), "", "senha incorreta"
        Exit Function
    End If
    AuditarSeg "MODO_DESENVOLVEDOR_ATIVADO", PapelSistema(), "", "todas as abas visiveis e desprotegidas ate o logout"
    Application.ScreenUpdating = False
    On Error Resume Next
    LiberarEstrutura
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:="qcini2025"
        ws.Visible = xlSheetVisible
    Next ws
    On Error GoTo 0
    SystemLook False
    Application.ScreenUpdating = True
    mModoDev = True                 ' o AfterSave refaz esta vista depois de gravar trancado
    AtivarModoDesenvolvedor = True
End Function


